/**
 * Text measurement interface. The layout engine never touches fonts directly, so it runs the same
 * in Node (font files via fontkit) and later in the editor (browser canvas or HarfBuzz/WASM).
 */
export interface FontKey {
  family: string;
  bold: boolean;
  italic: boolean;
}

export interface FontMetrics {
  /** Height of a single-spaced line at 1pt, as Word computes it (Windows ascent + descent). */
  lineHeight: number;
  /** Distance from line top to baseline at 1pt. */
  ascent: number;
}

export interface Measurer {
  /** Advance width of the text at 1pt (multiply by the font size). */
  width(text: string, font: FontKey): number;
  metrics(font: FontKey): FontMetrics;
  /** Fonts used to draw `text`: the requested font plus fallbacks for characters it lacks. */
  faces?(text: string, font: FontKey): FontKey[];
}
