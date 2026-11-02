export * from "./extract.js";
export { type ConversionResult, reconstruct } from "./reconstruct.js";
import { extractPages, type PdfJs, type PdfJsFiles } from "./extract.js";
import { type ConversionResult, reconstruct } from "./reconstruct.js";

/** PDF bytes → an editable document (and what may need checking). */
export async function convertPdf(data: Uint8Array, pdfjs: PdfJs, files?: PdfJsFiles,
  onPage?: (done: number, total: number) => void): Promise<ConversionResult> {
  return reconstruct(await extractPages(data, pdfjs, files, onPage));
}
