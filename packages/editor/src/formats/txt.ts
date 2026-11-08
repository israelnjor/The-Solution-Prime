/** Plain text (.txt): each line a paragraph; written as UTF-8 (with Windows line ends), tables as tab-separated rows. */
import type { PrimeDocument } from "@prime/pdm";
import { newDocument, paragraph, plainText, type Shown, shownBody, textInlines } from "./common.js";

/** The text of a file: UTF-8 or UTF-16 (by their byte-order marks), else UTF-8, else Windows-1252. */
export function decodeText(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return new TextDecoder("utf-8").decode(bytes.subarray(3));
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

export function readTxt(bytes: Uint8Array): PrimeDocument {
  const doc = newDocument();
  const text = decodeText(bytes).replace(/\r\n?/g, "\n");
  const lines = text.split("\n");
  if (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
  // A form feed (page break in some text files) starts a new page.
  doc.sections[0].blocks = lines.map((line) => {
    const parts = line.split("\f");
    const content = parts.flatMap((part, i) => [...(i ? [{ type: "break" as const, kind: "page" as const }] : []), ...textInlines(part)]);
    return paragraph(content);
  });
  if (!doc.sections[0].blocks.length) doc.sections[0].blocks = [paragraph([])];
  return doc;
}

export function writeTxt(doc: PrimeDocument): Uint8Array {
  const lines: string[] = [];
  const add = (blocks: Shown[]) => {
    for (const b of blocks) {
      if (b.kind === "paragraph") lines.push(...plainText(b).split("\n"));
      else for (const row of b.rows) lines.push(row.cells.map((c) => c.blocks.map((x) => (x.kind === "paragraph" ? plainText(x).replace(/\n/g, " ") : "")).join(" ")).join("\t"));
    }
  };
  add(shownBody(doc));
  const body = new TextEncoder().encode(lines.join("\r\n") + "\r\n");
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf]);
  out.set(body, 3);
  return out;
}
