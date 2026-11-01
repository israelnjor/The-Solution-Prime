export * from "./extract.js";
export { type ConversionResult, reconstruct } from "./reconstruct.js";
import { extractPages, type PdfJs } from "./extract.js";
import { type ConversionResult, reconstruct } from "./reconstruct.js";

/** PDF bytes → an editable document (and what may need checking). */
export async function convertPdf(data: Uint8Array, pdfjs: PdfJs): Promise<ConversionResult> {
  return reconstruct(await extractPages(data, pdfjs));
}
