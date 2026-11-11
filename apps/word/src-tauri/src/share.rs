//! Sharing a document through the system's share panel (Windows: the panel with Mail, Outlook, WhatsApp,
//! Nearby Share, Phone Link…). What is shared is a file — a copy of the document as .docx or .pdf, written
//! by lib.rs to Prime Word's cache folder — never a link, so it works offline and the person receiving it
//! needs nothing but the file.

use std::path::PathBuf;
use std::sync::Mutex;

/// What the share panel asks for when it opens (it asks after `show`, on the window's thread).
struct Pending {
    title: String,
    path: PathBuf,
}

static PENDING: Mutex<Option<Pending>> = Mutex::new(None);

/// Opens the share panel for `path`. Must run on the window's (main) thread.
#[cfg(windows)]
pub fn show(hwnd: isize, title: String, path: PathBuf) -> Result<(), String> {
    use std::sync::atomic::{AtomicBool, Ordering};
    use windows::core::{factory, Interface, Ref, HSTRING};
    use windows::ApplicationModel::DataTransfer::{DataRequestedEventArgs, DataTransferManager};
    use windows::Foundation::TypedEventHandler;
    use windows::Storage::{IStorageItem, StorageFile};
    use windows::Win32::Foundation::HWND;
    use windows::Win32::UI::Shell::IDataTransferManagerInterop;
    use windows_collections::IIterable;

    static REGISTERED: AtomicBool = AtomicBool::new(false);
    let err = |e: windows::core::Error| e.message().to_string();

    *PENDING.lock().map_err(|e| e.to_string())? = Some(Pending { title, path });
    let hwnd = HWND(hwnd as *mut core::ffi::c_void);
    let interop = factory::<DataTransferManager, IDataTransferManagerInterop>().map_err(err)?;
    if !REGISTERED.load(Ordering::SeqCst) {
        let manager: DataTransferManager = unsafe { interop.GetForWindow(hwnd) }.map_err(err)?;
        manager
            .DataRequested(&TypedEventHandler::new(|_: Ref<DataTransferManager>, args: Ref<DataRequestedEventArgs>| {
                let request = args.ok()?.Request()?;
                let pending = PENDING.lock().ok().and_then(|p| p.as_ref().map(|p| (p.title.clone(), p.path.clone())));
                let Some((title, path)) = pending else {
                    return request.FailWithDisplayText(&HSTRING::from("There is nothing to share."));
                };
                let data = request.Data()?;
                let props = data.Properties()?;
                props.SetTitle(&HSTRING::from(title))?;
                props.SetDescription(&HSTRING::from("Shared from Prime Word"))?;
                match StorageFile::GetFileFromPathAsync(&HSTRING::from(path.as_os_str()))?.join() {
                    Ok(file) => {
                        let items: Vec<Option<IStorageItem>> = vec![Some(file.cast()?)];
                        data.SetStorageItems(&IIterable::<IStorageItem>::from(items), true)
                    }
                    Err(e) => request.FailWithDisplayText(&HSTRING::from(format!("Prime Word could not share the file: {}", e.message()))),
                }
            }))
            .map_err(err)?;
        // The panel keeps asking this window's manager; it stays registered for the app's life.
        std::mem::forget(manager);
        REGISTERED.store(true, Ordering::SeqCst);
    }
    unsafe { interop.ShowShareUIForWindow(hwnd) }.map_err(err)
}

#[cfg(not(windows))]
pub fn show(_hwnd: isize, _title: String, _path: PathBuf) -> Result<(), String> {
    Err("Sharing is available in Prime Word for Windows for now.".into())
}
