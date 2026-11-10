//! Developer check of text recognition: `cargo run --example ocr -- image.gray WIDTH HEIGHT`
//! (an 8-bit grey image, one byte a pixel). Prints what is recognised in it, and in it inverted, as JSON.
fn main() {
    let args: Vec<String> = std::env::args().collect();
    let gray = std::fs::read(&args[1]).expect("image");
    let (w, h): (u32, u32) = (args[2].parse().unwrap(), args[3].parse().unwrap());
    let first = prime_word_lib::ocr::recognize(&gray, w, h).expect("recognize");
    let inverted: Vec<u8> = gray.iter().map(|v| 255 - v).collect();
    let second = prime_word_lib::ocr::recognize(&inverted, w, h).expect("recognize inverted");
    println!("{}", serde_json::json!({ "first": first, "second": second }));
}
