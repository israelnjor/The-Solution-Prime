/** Node measurer: reads installed Windows fonts with fontkit. Not used in the browser. */
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import * as fontkit from "fontkit";
import { extractFont } from "../fontfile.js";
import type { FontKey, FontMetrics, Measurer } from "../measure.js";

const LOCAL = process.env.LOCALAPPDATA ?? "";
const FONT_DIRS = [
  join(process.env.WINDIR ?? "C:\Windows", "Fonts"),
  join(LOCAL, "Microsoft", "Windows", "Fonts"),
  // Office downloads "cloud fonts" (e.g. Aptos, the default since 2023) here, one folder per family.
  join(LOCAL, "Microsoft", "FontCache", "4", "CloudFonts"),
];
/** Word's own substitute when a requested font is not installed. */
const FALLBACK = "times new roman";
/**
 * Fonts tried, in order, for characters the requested font lacks. Chosen by matching Word's line
 * heights in Experiment 2 (symbols -> Segoe UI Symbol, CJK -> Yu Gothic on this system).
 */
const GLYPH_FALLBACKS = ["Segoe UI Symbol", "Yu Gothic", "Microsoft YaHei", "Malgun Gothic", "Segoe UI Emoji", "Segoe UI"];

type Face = { font: fontkit.Font; bold: boolean; italic: boolean; file: string; index: number };

export class SystemFontMeasurer implements Measurer {
  private families = new Map<string, Face[]>();
  private resolved = new Map<string, Face>();
  private files = new Map<string, Uint8Array>();
  private widths = new Map<string, number>();

  constructor(dirs = FONT_DIRS) {
    for (const dir of dirs) {
      let names: string[] = [];
      try { names = readdirSync(dir); } catch { continue; }
      const files = names.flatMap((name) => {
        if (/\.(ttf|otf|ttc)$/i.test(name)) return [join(dir, name)];
        try { return readdirSync(join(dir, name)).filter((f) => /\.(ttf|otf|ttc)$/i.test(f)).map((f) => join(dir, name, f)); }
        catch { return []; }
      });
      for (const file of files) {
        let opened: fontkit.Font | fontkit.FontCollection;
        try { opened = fontkit.openSync(file); } catch { continue; }
        ("fonts" in opened ? opened.fonts : [opened]).forEach((font, index) => {
          const sub = (font.subfamilyName ?? "").toLowerCase();
          const face = { font, bold: /bold|black|heavy/.test(sub), italic: /italic|oblique/.test(sub), file, index };
          const key = (font.familyName ?? "").toLowerCase();
          this.families.set(key, [...(this.families.get(key) ?? []), face]);
        });
      }
    }
  }

  private record(f: FontKey): Face {
    const key = `${f.family}|${f.bold}|${f.italic}`;
    let face = this.resolved.get(key);
    if (!face) {
      const faces = this.families.get(f.family.toLowerCase()) ?? this.families.get(FALLBACK) ?? [...this.families.values()][0];
      const score = (x: Face) => (x.bold === f.bold ? 2 : 0) + (x.italic === f.italic ? 1 : 0);
      face = [...faces].sort((a, b) => score(b) - score(a))[0];
      this.resolved.set(key, face);
    }
    return face;
  }

  private face(f: FontKey): fontkit.Font {
    return this.record(f).font;
  }

  /** The font file for embedding (a single font, extracted when it lives in a .ttc collection). */
  fontData(f: FontKey): Uint8Array {
    const face = this.record(f);
    const id = `${face.file}#${face.index}`;
    let data = this.files.get(id);
    if (!data) {
      data = extractFont(readFileSync(face.file), face.index);
      this.files.set(id, data);
    }
    return data;
  }

  /** Splits text into runs drawn by one font each (the requested font, or glyph fallbacks). */
  split(text: string, f: FontKey): { text: string; font: FontKey }[] {
    const out: { text: string; font: FontKey }[] = [];
    for (const ch of text) {
      const font = this.drawer(ch.codePointAt(0)!, f);
      const last = out[out.length - 1];
      if (last && last.font.family === font.family) last.text += ch;
      else out.push({ text: ch, font });
    }
    return out;
  }

  /** The font that draws `cp`: the requested one, else the first fallback that has the glyph. */
  private drawer(cp: number, f: FontKey): FontKey {
    if (this.face(f).hasGlyphForCodePoint(cp)) return f;
    for (const family of GLYPH_FALLBACKS) {
      const fb = { family, bold: f.bold, italic: f.italic };
      if (this.families.has(family.toLowerCase()) && this.face(fb).hasGlyphForCodePoint(cp)) return fb;
    }
    return f;
  }

  width(text: string, f: FontKey): number {
    let total = 0;
    for (const ch of text) {
      const k = `${f.family}|${f.bold}|${f.italic}|${ch}`;
      let w = this.widths.get(k);
      if (w === undefined) {
        const cp = ch.codePointAt(0)!;
        const font = this.face(this.drawer(cp, f));
        w = font.hasGlyphForCodePoint(cp) ? font.glyphForCodePoint(cp).advanceWidth / font.unitsPerEm : 0.5;
        this.widths.set(k, w);
      }
      total += w;
    }
    return total;
  }

  faces(text: string, f: FontKey): FontKey[] {
    const out = new Map<string, FontKey>([[f.family, f]]);
    for (const ch of text) {
      const d = this.drawer(ch.codePointAt(0)!, f);
      out.set(d.family, d);
    }
    return [...out.values()];
  }

  metrics(f: FontKey): FontMetrics {
    const font = this.face(f);
    const os2 = font["OS/2"];
    // Fonts flagged USE_TYPO_METRICS (e.g. Aptos) are spaced by their typographic metrics; others by
    // the Windows ascent/descent — matching Word's line heights in Experiment 2.
    if (os2?.fsSelection?.useTypoMetrics) {
      const ascent = os2.typoAscender + os2.typoLineGap / 2;
      return { lineHeight: (os2.typoAscender - os2.typoDescender + os2.typoLineGap) / font.unitsPerEm, ascent: ascent / font.unitsPerEm };
    }
    const ascent = os2?.winAscent ?? font.ascent;
    const descent = os2?.winDescent ?? Math.abs(font.descent);
    return { lineHeight: (ascent + descent) / font.unitsPerEm, ascent: ascent / font.unitsPerEm };
  }
}
