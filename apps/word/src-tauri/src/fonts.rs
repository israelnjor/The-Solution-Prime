//! Finds installed font files by family and style, for embedding fonts in exported PDFs.
//!
//! Only each file's table directory and `name` table are read (a few KB), not whole fonts, so
//! indexing the system's fonts stays quick. Font collections (.ttc) are indexed per face.

use std::collections::HashMap;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::PathBuf;
use std::sync::Mutex;

use tauri::ipc::Response;
use tauri::State;

#[derive(Clone)]
struct Face {
    path: PathBuf,
    index: u32,
    bold: bool,
    italic: bool,
}

#[derive(Default)]
pub struct FontIndex(Mutex<Option<HashMap<String, Vec<Face>>>>);

/// The folders fonts are installed in: Windows's (and Office's cloud fonts), macOS's and Linux's.
/// Folders that do not exist on this system are skipped when indexing.
fn font_dirs() -> Vec<PathBuf> {
    let mut dirs = Vec::new();
    if let Ok(windir) = std::env::var("WINDIR") {
        dirs.push(PathBuf::from(windir).join("Fonts"));
    }
    if let Ok(local) = std::env::var("LOCALAPPDATA") {
        dirs.push(PathBuf::from(&local).join("Microsoft").join("Windows").join("Fonts"));
        // Office cloud fonts (Aptos etc.), one folder per family.
        let cloud = PathBuf::from(&local).join("Microsoft").join("FontCache").join("4").join("CloudFonts");
        if let Ok(entries) = std::fs::read_dir(&cloud) {
            dirs.extend(entries.flatten().map(|e| e.path()).filter(|p| p.is_dir()));
        }
    }
    // macOS: the system's, all users', the user's, and Office's (Calibri, Cambria… when Office is installed).
    dirs.push(PathBuf::from("/System/Library/Fonts"));
    dirs.push(PathBuf::from("/Library/Fonts"));
    dirs.push(PathBuf::from("/Applications/Microsoft Word.app/Contents/Resources/DFonts"));
    // Linux: the system's and the user's (nested in folders per package).
    dirs.push(PathBuf::from("/usr/share/fonts"));
    dirs.push(PathBuf::from("/usr/local/share/fonts"));
    if let Ok(home) = std::env::var("HOME") {
        let home = PathBuf::from(home);
        dirs.push(home.join("Library").join("Fonts"));
        dirs.push(home.join(".local").join("share").join("fonts"));
        dirs.push(home.join(".fonts"));
    }
    dirs
}

/// Font files in a folder and its subfolders (Linux keeps fonts nested a few levels deep).
fn font_files(dir: &PathBuf, depth: u32, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else { return };
    for path in entries.flatten().map(|e| e.path()) {
        if path.is_dir() {
            if depth < 4 {
                font_files(&path, depth + 1, out);
            }
        } else {
            let ext = path.extension().map(|e| e.to_string_lossy().to_lowercase()).unwrap_or_default();
            if ["ttf", "otf", "ttc"].contains(&ext.as_str()) {
                out.push(path);
            }
        }
    }
}

fn u16_at(b: &[u8], at: usize) -> Option<u16> {
    b.get(at..at + 2).map(|s| u16::from_be_bytes([s[0], s[1]]))
}

fn u32_at(b: &[u8], at: usize) -> Option<u32> {
    b.get(at..at + 4).map(|s| u32::from_be_bytes([s[0], s[1], s[2], s[3]]))
}

fn read_at(f: &mut File, offset: u64, len: usize) -> Option<Vec<u8>> {
    f.seek(SeekFrom::Start(offset)).ok()?;
    let mut buf = vec![0; len];
    f.read_exact(&mut buf).ok()?;
    Some(buf)
}

/// Family and subfamily names of the face whose table directory starts at `offset`.
fn names(f: &mut File, offset: u64) -> Option<(String, String)> {
    let head = read_at(f, offset, 12)?;
    let count = u16_at(&head, 4)? as usize;
    let dir = read_at(f, offset + 12, 16 * count)?;
    let (mut name_at, mut name_len) = (0u32, 0u32);
    for i in 0..count {
        if &dir[16 * i..16 * i + 4] == b"name" {
            name_at = u32_at(&dir, 16 * i + 8)?;
            name_len = u32_at(&dir, 16 * i + 12)?;
        }
    }
    if name_len == 0 || name_len > 1 << 20 {
        return None;
    }
    let table = read_at(f, name_at as u64, name_len as usize)?;
    let records = u16_at(&table, 2)? as usize;
    let storage = u16_at(&table, 4)? as usize;
    let (mut family, mut sub) = (None, None);
    for r in 0..records {
        let at = 6 + 12 * r;
        let (platform, encoding, language, id) = (u16_at(&table, at)?, u16_at(&table, at + 2)?, u16_at(&table, at + 4)?, u16_at(&table, at + 6)?);
        // Windows, Unicode BMP, US English: the names Windows and Word use.
        if platform != 3 || encoding != 1 || language != 0x409 || (id != 1 && id != 2) {
            continue;
        }
        let (len, off) = (u16_at(&table, at + 8)? as usize, u16_at(&table, at + 10)? as usize);
        let raw = table.get(storage + off..storage + off + len)?;
        let utf16: Vec<u16> = raw.chunks(2).filter(|c| c.len() == 2).map(|c| u16::from_be_bytes([c[0], c[1]])).collect();
        let text = String::from_utf16_lossy(&utf16);
        if id == 1 { family = Some(text) } else { sub = Some(text) }
    }
    Some((family?, sub.unwrap_or_default()))
}

fn build() -> HashMap<String, Vec<Face>> {
    let mut index: HashMap<String, Vec<Face>> = HashMap::new();
    let mut files = Vec::new();
    for dir in font_dirs() {
        font_files(&dir, 0, &mut files);
    }
    for path in files {
        let Ok(mut f) = File::open(&path) else { continue };
        let Some(head) = read_at(&mut f, 0, 12) else { continue };
        // A collection lists the offsets of its faces' table directories.
        let offsets: Vec<u64> = if &head[0..4] == b"ttcf" {
            let n = u32_at(&head, 8).unwrap_or(0).min(64) as usize;
            read_at(&mut f, 12, 4 * n).map_or(vec![], |b| (0..n).filter_map(|i| u32_at(&b, 4 * i).map(u64::from)).collect())
        } else {
            vec![0]
        };
        for (i, offset) in offsets.into_iter().enumerate() {
            if let Some((family, sub)) = names(&mut f, offset) {
                let s = sub.to_lowercase();
                index.entry(family.to_lowercase()).or_default().push(Face {
                    path: path.clone(),
                    index: i as u32,
                    bold: s.contains("bold") || s.contains("black") || s.contains("heavy"),
                    italic: s.contains("italic") || s.contains("oblique"),
                });
            }
        }
    }
    index
}

/// The best installed face for a family and style: its file path and index within a collection.
#[tauri::command]
pub fn font_lookup(state: State<'_, FontIndex>, family: String, bold: bool, italic: bool) -> Option<(String, u32)> {
    let mut guard = state.0.lock().ok()?;
    let index = guard.get_or_insert_with(build);
    let faces = index.get(&family.to_lowercase())?;
    let score = |f: &Face| (f.bold == bold) as u8 * 2 + (f.italic == italic) as u8;
    let best = faces.iter().max_by_key(|f| score(f))?;
    Some((best.path.to_string_lossy().into_owned(), best.index))
}

/// The bytes of a font file previously returned by `font_lookup` (only indexed font files).
#[tauri::command]
pub fn font_bytes(state: State<'_, FontIndex>, path: String) -> Result<Response, String> {
    let known = state
        .0
        .lock()
        .ok()
        .and_then(|g| g.as_ref().map(|idx| idx.values().flatten().any(|f| f.path.to_string_lossy() == path)))
        .unwrap_or(false);
    if !known {
        return Err("not an indexed font".into());
    }
    std::fs::read(&path).map(Response::new).map_err(|e| e.to_string())
}
