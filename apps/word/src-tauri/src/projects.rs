//! Projects: ordinary folders on the person's disk that Prime Word keeps together (a thesis, a client's
//! job). The files stay where the person can see, back up and sync them; Prime Word only keeps a list of
//! the project folders (in its data folder) and, in each one, a small hidden `.primeword/project.json`
//! with the project's name, colour and the like — so the project travels with its folder.
//!
//! Files inside a project folder may be opened and saved by the app (the person chose the folder).

use std::fs;
use std::path::{Path, PathBuf};

use serde::Serialize;
use tauri::{AppHandle, Manager};

/// The project's own folder, inside the project folder (hidden).
const META_DIR: &str = ".primeword";
const META_FILE: &str = "project.json";
/// How deep inside a project folder files are listed, and how many at most.
const MAX_DEPTH: usize = 4;
const MAX_FILES: usize = 3000;

fn registry_file(app: &AppHandle) -> Result<PathBuf, String> {
    let dir = app.path().app_local_data_dir().map_err(|e| e.to_string())?;
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    Ok(dir.join("projects.json"))
}

/// The project folders, most recently used first.
pub fn registered(app: &AppHandle) -> Vec<String> {
    registry_file(app)
        .ok()
        .and_then(|p| fs::read(p).ok())
        .and_then(|b| serde_json::from_slice(&b).ok())
        .unwrap_or_default()
}

fn store(app: &AppHandle, list: &[String]) -> Result<(), String> {
    let json = serde_json::to_vec(list).map_err(|e| e.to_string())?;
    crate::atomic_write(&registry_file(app)?, &json).map_err(|e| e.to_string())
}

fn same(a: &str, b: &str) -> bool {
    crate::same_path(a.trim_end_matches(['\\', '/']), b.trim_end_matches(['\\', '/']))
}

fn is_registered(app: &AppHandle, folder: &str) -> bool {
    registered(app).iter().any(|p| same(p, folder))
}

/// A path with `..` and links resolved; for a file not yet written, its folder's resolved path joined
/// with its name.
fn resolved(path: &Path) -> Option<PathBuf> {
    if let Ok(p) = fs::canonicalize(path) {
        return Some(p);
    }
    let parent = fs::canonicalize(path.parent()?).ok()?;
    let name = path.file_name()?;
    if name == ".." || name == "." {
        return None;
    }
    Some(parent.join(name))
}

/// Whether `path` lies inside `folder` (after resolving both: no way out with `..` or links).
pub fn contains(folder: &Path, path: &Path) -> bool {
    // Resolved paths carry the names as they are on disk, so comparing them part by part is exact.
    match (fs::canonicalize(folder), resolved(path)) {
        (Ok(f), Some(p)) => p != f && p.starts_with(&f),
        _ => false,
    }
}

/// Whether `path` is inside one of the person's project folders.
pub fn in_a_project(app: &AppHandle, path: &str) -> bool {
    registered(app).iter().any(|f| contains(Path::new(f), Path::new(path)))
}

fn meta_path(folder: &Path) -> PathBuf {
    folder.join(META_DIR).join(META_FILE)
}

fn read_meta(folder: &Path) -> serde_json::Value {
    fs::read(meta_path(folder))
        .ok()
        .and_then(|b| serde_json::from_slice(&b).ok())
        .unwrap_or_else(|| serde_json::json!({}))
}

fn write_meta(folder: &Path, meta: &serde_json::Value) -> Result<(), String> {
    let dir = folder.join(META_DIR);
    let new = !dir.exists();
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    if new {
        hide(&dir);
    }
    let json = serde_json::to_vec_pretty(meta).map_err(|e| e.to_string())?;
    crate::atomic_write(&meta_path(folder), &json).map_err(|e| e.to_string())
}

/// Hides the project's own folder in Explorer (folders starting with a dot are hidden elsewhere).
#[cfg(windows)]
fn hide(dir: &Path) {
    use std::os::windows::process::CommandExt;
    const CREATE_NO_WINDOW: u32 = 0x0800_0000;
    let _ = std::process::Command::new("attrib").arg("+h").arg(dir).creation_flags(CREATE_NO_WINDOW).status();
}

#[cfg(not(windows))]
fn hide(_dir: &Path) {}

fn folder_name(folder: &Path) -> String {
    folder.file_name().map(|n| n.to_string_lossy().into_owned()).unwrap_or_else(|| folder.to_string_lossy().into_owned())
}

/// A file in a project.
#[derive(Serialize)]
pub struct ProjectFile {
    pub path: String,
    /// Its folder within the project ("" at the top).
    pub folder: String,
    pub name: String,
    pub size: u64,
    pub modified: u64,
}

fn ms(t: std::io::Result<std::time::SystemTime>) -> u64 {
    t.ok().and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok()).map(|d| d.as_millis() as u64).unwrap_or(0)
}

/// The files of a project folder (not hidden ones, Office's "~$" lock files or the project's own folder),
/// at most four folders deep.
pub fn files(folder: &Path) -> Vec<ProjectFile> {
    let mut out = Vec::new();
    let mut stack = vec![(folder.to_path_buf(), 0usize)];
    while let Some((dir, depth)) = stack.pop() {
        let Ok(entries) = fs::read_dir(&dir) else { continue };
        for entry in entries.flatten() {
            let name = entry.file_name().to_string_lossy().into_owned();
            if name.starts_with('.') || name.starts_with("~$") || name.eq_ignore_ascii_case("desktop.ini") || name.eq_ignore_ascii_case("thumbs.db") {
                continue;
            }
            let Ok(meta) = entry.metadata() else { continue };
            let path = entry.path();
            if meta.is_dir() {
                if depth + 1 < MAX_DEPTH {
                    stack.push((path, depth + 1));
                }
            } else if meta.is_file() {
                let rel = path.parent().and_then(|p| p.strip_prefix(folder).ok()).map(|p| p.to_string_lossy().into_owned()).unwrap_or_default();
                out.push(ProjectFile { path: path.to_string_lossy().into_owned(), folder: rel, name, size: meta.len(), modified: ms(meta.modified()) });
                if out.len() >= MAX_FILES {
                    return out;
                }
            }
        }
    }
    out
}

/// A project as listed on the home screen.
#[derive(Serialize)]
pub struct ProjectSummary {
    pub path: String,
    pub exists: bool,
    /// The project's own settings (name, colour…), as the app wrote them.
    pub meta: serde_json::Value,
    pub files: usize,
    /// When a file in it last changed (ms since 1970).
    pub modified: u64,
}

fn summary(folder: &str) -> ProjectSummary {
    let p = Path::new(folder);
    let exists = p.is_dir();
    let mut meta = if exists { read_meta(p) } else { serde_json::json!({}) };
    if meta.get("name").and_then(|n| n.as_str()).unwrap_or("").is_empty() {
        meta["name"] = serde_json::Value::String(folder_name(p));
    }
    let list = if exists { files(p) } else { Vec::new() };
    ProjectSummary {
        path: folder.to_string(),
        exists,
        meta,
        files: list.len(),
        modified: list.iter().map(|f| f.modified).max().unwrap_or(0),
    }
}

#[tauri::command]
pub fn project_list(app: AppHandle) -> Vec<ProjectSummary> {
    registered(&app).iter().map(|f| summary(f)).collect()
}

/// Makes a folder a project (creating it when `create`), with its settings; it goes to the top of the
/// list. A folder that already is a project (moved here, or from another computer) keeps its settings.
#[tauri::command]
pub fn project_add(app: AppHandle, folder: String, meta: serde_json::Value, create: bool) -> Result<ProjectSummary, String> {
    let p = PathBuf::from(&folder);
    if !p.is_absolute() {
        return Err("Choose a folder on your computer for the project.".into());
    }
    if create {
        if p.exists() {
            return Err(format!("A folder named “{}” is already there.", folder_name(&p)));
        }
        fs::create_dir_all(&p).map_err(|e| e.to_string())?;
    } else if !p.is_dir() {
        return Err("That folder is no longer there.".into());
    }
    let existing = read_meta(&p);
    if existing.get("name").is_none() {
        write_meta(&p, &meta)?;
    }
    touch(&app, &folder)?;
    Ok(summary(&folder))
}

fn touch(app: &AppHandle, folder: &str) -> Result<(), String> {
    let mut list = registered(app);
    list.retain(|p| !same(p, folder));
    list.insert(0, folder.to_string());
    store(app, &list)
}

/// A project's files; it moves to the top of the list (most recently used).
#[tauri::command]
pub fn project_files(app: AppHandle, folder: String) -> Result<Vec<ProjectFile>, String> {
    if !is_registered(&app, &folder) {
        return Err("Not a project".into());
    }
    touch(&app, &folder)?;
    Ok(files(Path::new(&folder)))
}

/// Changes a project's settings (name, colour…).
#[tauri::command]
pub fn project_save(app: AppHandle, folder: String, meta: serde_json::Value) -> Result<(), String> {
    if !is_registered(&app, &folder) {
        return Err("Not a project".into());
    }
    write_meta(Path::new(&folder), &meta)
}

/// Takes a project off Prime Word's list. Its folder and files stay as they are.
#[tauri::command]
pub fn project_forget(app: AppHandle, folder: String) -> Result<(), String> {
    let mut list = registered(&app);
    list.retain(|p| !same(p, &folder));
    store(&app, &list)
}

/// A free name in `dir` for `name`: "Report.docx", else "Report (2).docx", "Report (3).docx"…
pub fn free_name(dir: &Path, name: &str) -> PathBuf {
    let first = dir.join(name);
    if !first.exists() {
        return first;
    }
    let (stem, ext) = match name.rfind('.') {
        Some(i) if i > 0 => (&name[..i], &name[i..]),
        _ => (name, ""),
    };
    (2..).map(|n| dir.join(format!("{stem} ({n}){ext}"))).find(|p| !p.exists()).unwrap()
}

/// A safe file name: no folder separators or characters Windows refuses.
fn clean_name(name: &str) -> String {
    let n: String = name.chars().map(|c| if "\\/:*?\"<>|".contains(c) || c.is_control() { '_' } else { c }).collect();
    let n = n.trim().trim_end_matches('.').to_string();
    if n.is_empty() || n == "." || n == ".." { "Untitled".into() } else { n }
}

/// A path in the project folder for a new file named `name` (not taken: "(2)" added if it is).
#[tauri::command]
pub fn project_new_path(app: AppHandle, folder: String, name: String) -> Result<String, String> {
    if !is_registered(&app, &folder) {
        return Err("Not a project".into());
    }
    Ok(free_name(Path::new(&folder), &clean_name(&name)).to_string_lossy().into_owned())
}

/// Copies files the person chose into the project folder (keeping their names; "(2)" when taken).
#[tauri::command]
pub fn project_import(app: AppHandle, folder: String, paths: Vec<String>) -> Result<Vec<String>, String> {
    use tauri_plugin_fs::FsExt;
    if !is_registered(&app, &folder) {
        return Err("Not a project".into());
    }
    let mut out = Vec::new();
    for src in paths {
        if !app.fs_scope().is_allowed(&src) {
            return Err("Prime Word may only copy files you chose".into());
        }
        let from = Path::new(&src);
        let name = from.file_name().map(|n| n.to_string_lossy().into_owned()).ok_or("not a file")?;
        let to = free_name(Path::new(&folder), &clean_name(&name));
        fs::copy(from, &to).map_err(|e| format!("{name}: {e}"))?;
        out.push(to.to_string_lossy().into_owned());
    }
    Ok(out)
}

/// Opens a file of a project in its own program (a spreadsheet, a picture), or shows a project folder.
#[tauri::command]
pub fn project_open_external(app: AppHandle, path: String) -> Result<(), String> {
    let p = Path::new(&path);
    if !(in_a_project(&app, &path) || (p.is_dir() && is_registered(&app, &path))) {
        return Err("Not in a project".into());
    }
    let opener = if cfg!(target_os = "windows") { "explorer" } else if cfg!(target_os = "macos") { "open" } else { "xdg-open" };
    std::process::Command::new(opener).arg(p).spawn().map_err(|e| e.to_string())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn temp(name: &str) -> PathBuf {
        let d = std::env::temp_dir().join(format!("prime-projects-test-{name}-{}", std::process::id()));
        let _ = fs::remove_dir_all(&d);
        fs::create_dir_all(&d).unwrap();
        d
    }

    #[test]
    fn containment() {
        let root = temp("contain");
        let project = root.join("Thesis");
        fs::create_dir_all(project.join("Chapters")).unwrap();
        fs::write(project.join("Chapters").join("One.docx"), b"x").unwrap();
        fs::write(root.join("Outside.docx"), b"x").unwrap();
        assert!(contains(&project, &project.join("Chapters").join("One.docx")));
        assert!(contains(&project, &project.join("New file.docx"))); // not written yet
        assert!(!contains(&project, &root.join("Outside.docx")));
        assert!(!contains(&project, &project.join("..").join("Outside.docx")));
        assert!(!contains(&project, &project));
        // A sibling whose name starts the same is not inside.
        fs::create_dir_all(root.join("Thesis old")).unwrap();
        assert!(!contains(&project, &root.join("Thesis old").join("a.docx")));
        let _ = fs::remove_dir_all(&root);
    }

    #[test]
    fn listing_and_names() {
        let root = temp("files");
        fs::create_dir_all(root.join("Sub")).unwrap();
        fs::create_dir_all(root.join(META_DIR)).unwrap();
        fs::write(root.join("Report.docx"), b"x").unwrap();
        fs::write(root.join("~$Report.docx"), b"x").unwrap();
        fs::write(root.join("Sub").join("Scan.pdf"), b"x").unwrap();
        fs::write(root.join(META_DIR).join(META_FILE), b"{}").unwrap();
        let mut names: Vec<_> = files(&root).into_iter().map(|f| (f.folder, f.name)).collect();
        names.sort();
        assert_eq!(names, vec![("".into(), "Report.docx".into()), ("Sub".into(), "Scan.pdf".into())]);
        assert_eq!(free_name(&root, "Report.docx").file_name().unwrap(), "Report (2).docx");
        assert_eq!(free_name(&root, "New.docx").file_name().unwrap(), "New.docx");
        assert_eq!(clean_name("a/b:c?.docx"), "a_b_c_.docx");
        assert_eq!(clean_name(".."), "Untitled");
        let _ = fs::remove_dir_all(&root);
    }
}
