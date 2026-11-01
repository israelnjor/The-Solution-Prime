/**
 * Ink (InkML, W3C) as Word stores it in word/ink/*.xml: pen strokes ("traces") of X, Y and pressure
 * samples. Values are written compactly: after a value, "'" means "change from the last value" and
 * '"' "change in that change"; the kind sticks to its channel until another is given, and numbers
 * may run together where a sign or one of these marks separates them ("-7'0'0").
 */

export interface Ink {
  /** Strokes as x, y pairs, relative to the ink's bounding box (0–1 each way). */
  strokes: number[][];
  /** Pen width in pt; colour as hex (RRGGBB). */
  width: number;
  color: string;
}

const PT_PER_CM = 72 / 2.54;

function attr(tag: string, name: string): string | undefined {
  return new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
}

/** One trace's points: [x, y] per sample, in the ink's own units. */
export function decodeTrace(text: string, channels: number): number[][] {
  const kind: string[] = Array(channels).fill("!");
  const last: number[] = Array(channels).fill(0);
  const velocity: number[] = Array(channels).fill(0);
  const points: number[][] = [];
  for (const sample of text.split(",")) {
    const values = [...sample.matchAll(/([!'"]?)\s*(-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)/g)];
    if (!values.length) continue;
    values.slice(0, channels).forEach(([, mark, num], c) => {
      if (mark) kind[c] = mark;
      const v = Number(num);
      if (kind[c] === "'") velocity[c] = v;
      else if (kind[c] === '"') velocity[c] += v;
      else velocity[c] = v - last[c];
      last[c] = kind[c] === "!" ? v : last[c] + velocity[c];
    });
    points.push([last[0], last[1]]);
  }
  return points;
}

/** The strokes of an InkML document, scaled into its bounding box; undefined if it has none. */
export function readInk(xml: string): Ink | undefined {
  const channels = Math.max(2, [...xml.matchAll(/<inkml:channel\b/g)].length);
  const brush = /<inkml:brush\b[\s\S]*?<\/inkml:brush>/.exec(xml)?.[0] ?? "";
  const prop = (name: string) => {
    const tag = new RegExp(`<inkml:brushProperty\\b[^>]*name="${name}"[^>]*>`).exec(brush)?.[0];
    return tag ? { value: attr(tag, "value"), units: attr(tag, "units") } : undefined;
  };
  const w = prop("width");
  const width = w?.value ? Number(w.value) * (w.units === "mm" ? PT_PER_CM / 10 : w.units === "in" ? 72 : PT_PER_CM) : 1;
  const color = (prop("color")?.value ?? "#000000").replace("#", "").toUpperCase();
  const traces = [...xml.matchAll(/<inkml:trace\b[^>]*>([\s\S]*?)<\/inkml:trace>/g)].map((m) => decodeTrace(m[1], channels));
  const all = traces.flat();
  if (!all.length) return undefined;
  const xs = all.map((p) => p[0]);
  const ys = all.map((p) => p[1]);
  const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
  // A single dot (zero-size ink) still gets a box.
  const sx = x1 - x0 || 1;
  const sy = y1 - y0 || 1;
  // Coordinates are scaled into the drawing's box (Word sizes it to the ink), so their unit does not matter.
  return {
    strokes: traces.filter((t) => t.length).map((t) => t.flatMap(([x, y]) => [(x - x0) / sx, (y - y0) / sy])),
    width: Math.max(0.25, width),
    color: /^[0-9A-F]{6}$/.test(color) ? color : "000000",
  };
}
