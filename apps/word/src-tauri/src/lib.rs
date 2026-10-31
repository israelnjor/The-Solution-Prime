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
    Some((path, bytes))
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
#[tauri::command]
fn save_document(app: AppHandle, request: Request) -> Result<(), String> {
    let path = PathBuf::from(header(&request, "x-path")?);
    if !app.fs_scope().is_allowed(&path) {
        return Err("Prime Word may only save to files you chose".into());
    }
    atomic_write(&path, raw_body(&request)?).map_err(|e| e.to_string())
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
