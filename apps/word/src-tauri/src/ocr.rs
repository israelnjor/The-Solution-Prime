//! Text recognition for PDF pages that carry no text (printed as outlines, or scanned): the operating
//! system's own recogniser, offline — Windows.Media.Ocr on Windows. Elsewhere it is not available yet.

use serde::Serialize;

/// A recognised word: its text and box in the image's pixels.
#[derive(Serialize)]
pub struct Word {
    pub text: String,
    pub x: f32,
    pub y: f32,
    pub w: f32,
    pub h: f32,
}

/// A recognised line of words (as the recogniser grouped them).
#[derive(Serialize)]
pub struct Line {
    pub words: Vec<Word>,
}

/// What the recogniser found: the lines, and how far the text is turned (degrees), if it is.
#[derive(Serialize)]
pub struct Recognized {
    pub lines: Vec<Line>,
    pub angle: f64,
    pub language: String,
}

/// The largest side of an image the recogniser takes (pixels).
#[cfg(windows)]
pub fn max_side() -> u32 {
    windows::Media::Ocr::OcrEngine::MaxImageDimension().unwrap_or(10_000)
}

#[cfg(not(windows))]
pub fn max_side() -> u32 {
    0
}

/// Recognises the text in an 8-bit grey image (`width` × `height`, one byte a pixel, rows top down).
#[cfg(windows)]
pub fn recognize(gray: &[u8], width: u32, height: u32) -> Result<Recognized, String> {
    use windows::Graphics::Imaging::{BitmapPixelFormat, SoftwareBitmap};
    use windows::Media::Ocr::OcrEngine;
    use windows::Storage::Streams::DataWriter;
    use windows::Win32::System::WinRT::{RoInitialize, RO_INIT_MULTITHREADED};

    if gray.len() != (width as usize) * (height as usize) {
        return Err("the image's size does not match its pixels".into());
    }
    // This worker thread may not have joined a COM apartment yet ("already initialised" is fine).
    unsafe {
        let _ = RoInitialize(RO_INIT_MULTITHREADED);
    }
    let err = |e: windows::core::Error| e.message().to_string();
    let engine = OcrEngine::TryCreateFromUserProfileLanguages()
        .map_err(|_| "Text recognition is not installed for your Windows display language".to_string())?;
    let writer = DataWriter::new().map_err(err)?;
    writer.WriteBytes(gray).map_err(err)?;
    let buffer = writer.DetachBuffer().map_err(err)?;
    let bitmap = SoftwareBitmap::CreateCopyFromBuffer(&buffer, BitmapPixelFormat::Gray8, width as i32, height as i32).map_err(err)?;
    let result = engine.RecognizeAsync(&bitmap).map_err(err)?.join().map_err(err)?;
    let angle = result.TextAngle().ok().and_then(|a| a.Value().ok()).unwrap_or(0.0);
    let language = engine.RecognizerLanguage().and_then(|l| l.LanguageTag()).map(|t| t.to_string()).unwrap_or_default();
    let mut lines = Vec::new();
    for line in result.Lines().map_err(err)? {
        let mut words = Vec::new();
        for word in line.Words().map_err(err)? {
            let r = word.BoundingRect().map_err(err)?;
            words.push(Word { text: word.Text().map_err(err)?.to_string(), x: r.X, y: r.Y, w: r.Width, h: r.Height });
        }
        lines.push(Line { words });
    }
    Ok(Recognized { lines, angle, language })
}

#[cfg(not(windows))]
pub fn recognize(_gray: &[u8], _width: u32, _height: u32) -> Result<Recognized, String> {
    Err("Text recognition is not available on this system yet".into())
}
