/**
 * File formats besides .docx: which one a file is (by its contents first, then its name), reading any
 * of them into a document, and writing a document in one of those that can be written.
 */
import { isCfb, readDocx } from "@prime/docx";
import type { PrimeDocument } from "@prime/pdm";
import { readDoc } from "./doc.js";
import { readOdt, writeOdt } from "./odt.js";
import { readRtf, writeRtf } from "./rtf.js";
import { readTxt, writeTxt } from "./txt.js";

export type FileFormat = "docx" | "doc" | "rtf" | "odt" | "txt";
/** Formats a document can be saved in. */
export const WRITABLE: FileFormat[] = ["docx", "rtf", "odt", "txt"];
export const FORMAT_NAMES: Record<FileFormat, string> = {
  docx: "Word document", doc: "Word 97–2003 document", rtf: "Rich Text Format", odt: "OpenDocument text", txt: "Plain text",
};

export const extensionOf = (name: string) => (/\.([a-z0-9]+)$/i.exec(name)?.[1] ?? "").toLowerCase();

/** The format of a file: from its first bytes (zip, compound file, RTF), else from its name. */
export function formatOf(bytes: Uint8Array, name: string): FileFormat {
  const ext = extensionOf(name);
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    // A zip: OpenDocument says so in its first entry ("mimetype").
    const head = new TextDecoder("latin1").decode(bytes.subarray(0, 100));
    return head.includes("application/vnd.oasis.opendocument.text") || ext === "odt" ? "odt" : "docx";
  }
  if (isCfb(bytes)) return "doc";
  if (new TextDecoder("latin1").decode(bytes.subarray(0, 6)) === "{\\rtf1" || new TextDecoder("latin1").decode(bytes.subarray(0, 5)) === "{\\rtf") return "rtf";
  if (ext === "docx" || ext === "doc" || ext === "rtf" || ext === "odt") return ext;
  return "txt";
}

/** Reads a document in any of the formats. */
export function readAnyFormat(bytes: Uint8Array, name: string): { doc: PrimeDocument; warnings: string[]; format: FileFormat } {
  const format = formatOf(bytes, name);
  switch (format) {
    case "docx": return { ...readDocx(bytes), format };
    case "doc": return { ...readDoc(bytes), format };
    case "rtf": return { ...readRtf(bytes), format };
    case "odt": return { ...readOdt(bytes), format };
    default: return { doc: readTxt(bytes), warnings: [], format };
  }
}

/** Writes a document as .rtf, .odt or .txt (for .docx, writeDocx). */
export function writeFormat(doc: PrimeDocument, format: Exclude<FileFormat, "docx" | "doc">, props: { title?: string; author?: string } = {}): Uint8Array {
  if (format === "rtf") return writeRtf(doc, props);
  if (format === "odt") return writeOdt(doc, props);
  return writeTxt(doc);
}
