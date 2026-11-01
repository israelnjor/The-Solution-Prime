//! Prime Word desktop shell (ADR-003): a native window around the web editor.
//! Native work goes through official Tauri plugins and the few commands below, limited by
//! capabilities/default.json: documents the user picks, the app's own recovery folder, and fonts.

mod fonts;

use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{AppHandle, Manager};
use tauri_plugin_fs::FsExt;

/// The .docx passed on the command line ("Open with", double-click), if any: its path and bytes.
/// That one file is added to the app's file scope so Save can write back to it.
#[tauri::command]
fn startup_document(app: AppHandle) -> Option<(String, Vec<u8>)> {
    let path = std::env::args().skip(1).find(|a| a.to_lowercase().ends_with(".docx"))?;
    let bytes = fs::read(&path).ok()?;
    let _ = app.fs_scope().allow_file(&path);
    remember(&app, &path);
    Some((path, bytes))
}

const RECENT_MAX: usize = 12;

/// Recently opened or saved documents, most recent first, in the app's data folder (paths only).
fn recent_file(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().app_local_data_dir().map_err(|e| e.to_string())?;
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.join("recent.json"))
}

fn recent_list(app: &AppHandle) -> Vec<String> {
    recent_file(app)
        .ok()
        .and_then(|p| fs::read(p).ok())
        .and_then(|b| serde_json::from_slice(&b).ok())
        .unwrap_or_default()
}

fn store_recent(app: &AppHandle, list: &[String]) {
    if let (Ok(file), Ok(json)) = (recent_file(app), serde_json::to_vec(list)) {
        let _ = atomic_write(&file, &json);
    }
}

/// Windows paths are case-insensitive.
fn same_path(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b)
}

fn remember(app: &AppHandle, path: &str) {
    if !path.to_lowercase().ends_with(".docx") {
        return; // e.g. a PDF export written through save_document
    }
    let mut list = recent_list(app);
    list.retain(|p| !same_path(p, path));
    list.insert(0, path.to_string());
    list.truncate(RECENT_MAX);
    store_recent(app, &list);
}

fn is_recent(app: &AppHandle, path: &str) -> bool {
    recent_list(app).iter().any(|p| same_path(p, path))
}

/// A file's last-modified time in milliseconds since 1970, if it exists.
fn modified_ms(path: &Path) -> Option<u64> {
    let t = fs::metadata(path).ok()?.modified().ok()?;
    Some(t.duration_since(std::time::UNIX_EPOCH).ok()?.as_millis() as u64)
}

/// Files the user chose (dialog, "Open with") or documents opened before (recent list) — nothing else.
fn may_access(app: &AppHandle, path: &str) -> bool {
    app.fs_scope().is_allowed(path) || (path.to_lowercase().ends_with(".docx") && is_recent(app, path))
}

#[derive(serde::Serialize)]
struct RecentDocument {
    path: String,
    name: String,
    exists: bool,
}

#[tauri::command]
fn recent_documents(app: AppHandle) -> Vec<RecentDocument> {
    recent_list(&app)
        .into_iter()
        .map(|path| {
            let p = Path::new(&path);
            RecentDocument {
                name: p.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default(),
                exists: p.is_file(),
                path,
            }
        })
        .collect()
}

#[tauri::command]
fn forget_recent(app: AppHandle, path: String) {
    let mut list = recent_list(&app);
    list.retain(|p| !same_path(p, &path));
    store_recent(&app, &list);
}

/// Reads a document the user chose or opened before, and records it as recent.
#[tauri::command]
fn read_document(app: AppHandle, path: String) -> Result<Response, String> {
    if !may_access(&app, &path) {
        return Err("Prime Word may only open files you chose".into());
    }
    let bytes = fs::read(&path).map_err(|e| e.to_string())?;
    let _ = app.fs_scope().allow_file(&path);
    remember(&app, &path);
    Ok(Response::new(bytes))
}

/// When a document was last changed on disk (to notice changes by other programs before saving).
#[tauri::command]
fn document_modified(app: AppHandle, path: String) -> Option<u64> {
    if !may_access(&app, &path) {
        return None;
    }
    modified_ms(Path::new(&path))
}

/// Automation switches from the command line (`--pw-export-pdf=<path>`, `--pw-type=<text>`), used for
/// end-to-end tests and batch conversion. An export target is added to the file scope.
#[tauri::command]
fn automation_args(app: AppHandle) -> Vec<String> {
    let args: Vec<String> = std::env::args().filter(|a| a.starts_with("--pw-")).collect();
    for a in &args {
        if let Some(path) = a.strip_prefix("--pw-export-pdf=") {
            let _ = app.fs_scope().allow_file(path);
        }
    }
    args
}

/// Writes atomically: a temporary file in the same folder, flushed to disk, then renamed over the
/// target. A crash or power cut leaves either the old file or the new one — never half of one.
fn atomic_write(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let dir = path.parent().unwrap_or_else(|| Path::new("."));
    let name = path.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_default();
    let tmp = dir.join(format!("~${name}.prime-tmp"));
    {
        let mut f = fs::File::create(&tmp)?;
        f.write_all(bytes)?;
        f.sync_all()?;
    }
    fs::rename(&tmp, path).inspect_err(|_| {
        let _ = fs::remove_file(&tmp);
    })
}

/// Percent-decoding for header values (paths travel percent-encoded: headers are ASCII-only).
fn decode(s: &str) -> String {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 2 < bytes.len() {
            if let Ok(b) = u8::from_str_radix(&s[i + 1..i + 3], 16) {
                out.push(b);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

fn header(request: &Request, name: &str) -> Result<String, String> {
    request
        .headers()
        .get(name)
        .and_then(|v| v.to_str().ok())
        .map(decode)
        .ok_or_else(|| format!("missing {name}"))
}

fn raw_body<'a>(request: &'a Request) -> Result<&'a [u8], String> {
    match request.body() {
        InvokeBody::Raw(bytes) => Ok(bytes),
        _ => Err("expected raw bytes".into()),
    }
}

/// Saves a document the user chose (via a dialog or by opening it). Paths outside that scope are refused.
/// With `x-known-modified`, refuses ("CHANGED_ON_DISK") when the file changed after that time — another
/// program saved it — so a newer version is never overwritten silently. Returns the new modified time.
#[tauri::command]
fn save_document(app: AppHandle, request: Request) -> Result<u64, String> {
    let name = header(&request, "x-path")?;
    if !may_access(&app, &name) {
        return Err("Prime Word may only save to files you chose".into());
    }
    let path = PathBuf::from(&name);
    if let Some(known) = header(&request, "x-known-modified").ok().and_then(|v| v.parse::<u64>().ok()) {
        if modified_ms(&path).is_some_and(|now| now > known) {
            return Err("CHANGED_ON_DISK".into());
        }
    }
    atomic_write(&path, raw_body(&request)?).map_err(|e| e.to_string())?;
    remember(&app, &name);
    Ok(modified_ms(&path).unwrap_or(0))
}

fn recovery_dir(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().app_local_data_dir().map_err(|e| e.to_string())?.join("recovery");
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir)
}

fn recovery_file(app: &AppHandle, id: &str) -> Result<PathBuf, String> {
    if id.is_empty() || !id.chars().all(|c| c.is_ascii_alphanumeric() || c == '-') {
        return Err("bad recovery id".into());
    }
    Ok(recovery_dir(app)?.join(format!("{id}.json")))
}

/// Stores a recovery snapshot (the document model as JSON) for this editing session.
#[tauri::command]
fn recovery_save(app: AppHandle, request: Request) -> Result<(), String> {
    let path = recovery_file(&app, &header(&request, "x-id")?)?;
    atomic_write(&path, raw_body(&request)?).map_err(|e| e.to_string())
}

/// Recovery snapshots on disk: (id, last written, seconds since the Unix epoch).
#[tauri::command]
fn recovery_list(app: AppHandle) -> Result<Vec<(String, u64)>, String> {
    let mut out = Vec::new();
    for entry in fs::read_dir(recovery_dir(&app)?).map_err(|e| e.to_string())?.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        if let Some(id) = name.strip_suffix(".json") {
            let modified = entry
                .metadata()
                .and_then(|m| m.modified())
                .ok()
                .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
                .map_or(0, |d| d.as_secs());
            out.push((id.to_string(), modified));
        }
    }
    Ok(out)
}

#[tauri::command]
fn recovery_load(app: AppHandle, id: String) -> Result<Response, String> {
    fs::read(recovery_file(&app, &id)?).map(Response::new).map_err(|e| e.to_string())
}

#[tauri::command]
fn recovery_delete(app: AppHandle, id: String) -> Result<(), String> {
    let path = recovery_file(&app, &id)?;
    if path.exists() {
        fs::remove_file(path).map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// Writes a PDF to Prime Word's temporary folder and opens it in the default PDF viewer (to print).
#[tauri::command]
fn open_pdf(request: Request) -> Result<(), String> {
    let name: String = header(&request, "x-name")?
        .chars()
        .map(|c| if c.is_alphanumeric() || " -_.".contains(c) { c } else { '_' })
        .collect();
    let dir = std::env::temp_dir().join("Prime Word");
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    let path = dir.join(format!("{}.pdf", name.trim_end_matches(".docx")));
    atomic_write(&path, raw_body(&request)?).map_err(|e| e.to_string())?;
    // Explorer opens a file with its default application; the path is our own temp file.
    std::process::Command::new("explorer").arg(&path).spawn().map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .manage(fonts::FontIndex::default())
        .invoke_handler(tauri::generate_handler![
            startup_document,
            recent_documents,
            forget_recent,
            read_document,
            document_modified,
            automation_args,
            save_document,
            recovery_save,
            recovery_list,
            recovery_load,
            recovery_delete,
            open_pdf,
            fonts::font_lookup,
            fonts::font_bytes,
        ])
        .run(tauri::generate_context!())
        .expect("error while running Prime Word");
}
