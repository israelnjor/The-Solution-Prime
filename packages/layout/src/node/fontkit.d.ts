declare module "fontkit" {
  interface Glyph { advanceWidth: number }
  interface Font {
    familyName: string;
    subfamilyName: string;
    unitsPerEm: number;
    ascent: number;
    descent: number;
    "OS/2"?: {
      winAscent: number;
      winDescent: number;
      typoAscender: number;
      typoDescender: number;
      typoLineGap: number;
      fsSelection?: { useTypoMetrics?: boolean };
    };
    hasGlyphForCodePoint(cp: number): boolean;
    glyphForCodePoint(cp: number): Glyph;
  }
  interface FontCollection { fonts: Font[] }
  export function openSync(path: string): Font | FontCollection;
}
