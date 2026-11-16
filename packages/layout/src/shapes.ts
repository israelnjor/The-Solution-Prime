/**
 * Shapes as drawings in a w×h box (origin top-left, y down), for the screen (SVG) and the PDF (pdf-lib's
 * drawSvgPath): Word's preset geometries (a:prstGeom prst) with their default adjustments, their
 * details (a cylinder's top, a cube's sides, a flowchart symbol's inner lines), dashed lines and arrow
 * heads, and where text sits inside each. Unknown geometries are drawn as rectangles.
 */
import type { LineEnd, ShapeLine, ShapeStyle } from "@prime/pdm";

const r = (n: number) => Math.round(n * 100) / 100;
type Pt = [number, number];
/** A closed polygon. */
const P = (pts: Pt[]) => `M${pts.map(([x, y]) => `${r(x)},${r(y)}`).join(" L")} Z`;
/** A full ellipse (clockwise; reversed: anticlockwise, to cut a hole in a shape drawn clockwise). */
function E(cx: number, cy: number, rx: number, ry: number, reversed = false): string {
  const s = reversed ? 0 : 1;
  return `M${r(cx - rx)},${r(cy)} A${r(rx)},${r(ry)} 0 1,${s} ${r(cx + rx)},${r(cy)} A${r(rx)},${r(ry)} 0 1,${s} ${r(cx - rx)},${r(cy)} Z`;
}
/**
 * Points of a star with n points (the first at the top); `inner` is the inner radius as a share of the
 * outer. Word stretches some stars to fill their box (its hf / vf factors, the centre moved down).
 */
function star(n: number, inner: number, w: number, h: number, hf = 1, vf = 1): string {
  const pts: Pt[] = [];
  const [cx, cy, rx, ry] = [w / 2, (h / 2) * vf, (w / 2) * hf, (h / 2) * vf];
  for (let i = 0; i < n * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / n;
    const k = i % 2 ? inner : 1;
    pts.push([cx + Math.cos(a) * rx * k, cy + Math.sin(a) * ry * k]);
  }
  return P(pts);
}
/** A regular polygon with n corners (the first at the top). */
function polygon(n: number, w: number, h: number): string {
  const pts: Pt[] = [];
  for (let i = 0; i < n; i++) {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
    pts.push([w / 2 + (Math.cos(a) * w) / 2, h / 2 + (Math.sin(a) * h) / 2]);
  }
  return P(pts);
}
/** Points given in Word's 21600-unit box, scaled to w×h. */
const U = (w: number, h: number, pts: number[]) => P(pts.reduce<Pt[]>((out, v, i) => (i % 2 ? out : [...out, [(v * w) / 21600, (pts[i + 1] * h) / 21600]]), []));
/** A document's wavy bottom (flowchart document), in a box at x, y. */
const documentPath = (x: number, y: number, w: number, h: number) =>
  `M${r(x)},${r(y)} L${r(x + w)},${r(y)} L${r(x + w)},${r(y + h * 0.802)} C${r(x + w / 2)},${r(y + h * 0.802)} ${r(x + w / 2)},${r(y + h * 1.107)} ${r(x)},${r(y + h * 0.934)} Z`;
/** A cloud: round puffs of uneven sizes round an ellipse (each a cubic bulge between two dips). */
function cloud(cx: number, cy: number, rx: number, ry: number): string {
  const sizes = [1.1, 0.8, 1.2, 0.9, 1.15, 0.85, 1.05, 0.95];
  const total = sizes.reduce((a, b) => a + b, 0);
  const at = (a: number, k: number): Pt => [cx + Math.cos(a) * rx * k, cy + Math.sin(a) * ry * k];
  let a0 = -Math.PI * 0.6;
  let d = "";
  for (let i = 0; i < sizes.length; i++) {
    const a1 = a0 + (sizes[i] / total) * Math.PI * 2;
    const [x0, y0] = at(a0, 0.88);
    const [x1, y1] = at(a1, 0.88);
    const bulge = 1.12 + sizes[i] * 0.22;
    const [c1x, c1y] = at(a0 + (a1 - a0) * 0.05, bulge);
    const [c2x, c2y] = at(a1 - (a1 - a0) * 0.05, bulge);
    d += `${i ? " " : `M${r(x0)},${r(y0)}`} C${r(c1x)},${r(c1y)} ${r(c2x)},${r(c2y)} ${r(x1)},${r(y1)}`;
    a0 = a1;
  }
  return `${d} Z`;
}

/** A part of a shape's drawing: its outline and how it is filled and stroked. */
export interface ShapePart {
  d: string;
  /** "shape": the shape's fill; "darken"/"lighten": shaded from it (a cube's side); "line": the line's colour (an arrow head); "none". */
  fill: "shape" | "darken" | "lighten" | "line" | "none";
  /** Drawn with the shape's line (its dash too, except arrow heads). */
  stroke: boolean;
  /** An arrow head: solid line, never dashed. */
  head?: boolean;
}

/** Shapes that are lines (nothing to fill): straight lines, arcs, brackets and braces. */
const LINES = new Set(["line", "straightConnector1", "arc", "leftBracket", "rightBracket", "leftBrace", "rightBrace", "bracketPair", "bracePair"]);
export const isLineShape = (prst: string) => LINES.has(prst);

/** A shape's adjustment (Word's a:avLst gd "adj", "adj1"…: in 100000ths), its default, kept within its range. */
type Adjust = Record<string, number>;
const adjust = (adj: Adjust | undefined) => (name: string, fallback: number, lo: number, hi: number) =>
  Math.min(Math.max(adj?.[name] ?? fallback, Math.min(lo, hi)), Math.max(lo, hi)) / 100000;

/** The drawing of a geometry in a w×h box: its outline and details (without line ends). */
function parts(prst: string, w: number, h: number, adj?: Adjust): ShapePart[] {
  const m = Math.min(w, h);
  const A = adjust(adj);
  /** A block arrow's shaft (half its thickness, from the edge) and head length; `along` is its length, `heads` how many. */
  const arrowSizes = (across: number, along: number, heads = 1) => {
    const shaft = across * A("adj1", 50000, 0, 100000);
    const hd = Math.min(along / heads, m * A("adj2", 50000, 0, (100000 * along) / (heads * m)));
    return { t: (across - shaft) / 2, hd };
  };
  /** A callout's pointer tip: from the centre by adj1 (across) and adj2 (down), shares of the box. */
  const calloutTip = (bw: number, bh: number): Pt => [bw / 2 + bw * A("adj1", -20833, -10000000, 10000000), bh / 2 + bh * A("adj2", 62500, -10000000, 10000000)];
  /** A rectangle (corners rounded by `k`) with a pointer to `tip` from the side the tip is beyond. */
  const rectCallout = (bw: number, bh: number, k: number, tip: Pt): string => {
    const [dx, dy] = [(tip[0] - bw / 2) / bw, (tip[1] - bh / 2) / bh];
    const inside = Math.abs(dx) <= 0.5 && Math.abs(dy) <= 0.5;
    const pts: string[] = [];
    const L = (x: number, y: number) => pts.push(`L${r(x)},${r(y)}`);
    const corner = (cx: number, cy: number, x: number, y: number) => pts.push(k ? `Q${r(cx)},${r(cy)} ${r(x)},${r(y)}` : `L${r(cx)},${r(cy)} L${r(x)},${r(y)}`);
    // Which side the pointer leaves from, and where along it (toward the tip's side).
    const side = inside ? "" : Math.abs(dy) >= Math.abs(dx) ? (dy > 0 ? "b" : "t") : dx > 0 ? "r" : "l";
    const span = (len: number, toward: number) => (toward < 0 ? [len * 0.167, len * 0.417] : [len * 0.583, len * 0.833]);
    const [hx1, hx2] = span(bw, dx);
    const [vy1, vy2] = span(bh, dy);
    pts.push(`M${r(k)},0`);
    if (side === "t") { L(hx1, 0); L(tip[0], tip[1]); L(hx2, 0); }
    L(bw - k, 0); corner(bw, 0, bw, k);
    if (side === "r") { L(bw, vy1); L(tip[0], tip[1]); L(bw, vy2); }
    L(bw, bh - k); corner(bw, bh, bw - k, bh);
    if (side === "b") { L(hx2, bh); L(tip[0], tip[1]); L(hx1, bh); }
    L(k, bh); corner(0, bh, 0, bh - k);
    if (side === "l") { L(0, vy2); L(tip[0], tip[1]); L(0, vy1); }
    L(0, k); corner(0, 0, k, 0);
    return `${pts.join(" ")} Z`;
  };
  const one = (d: string): ShapePart[] => [{ d, fill: "shape", stroke: true }];
  const detail = (d: string): ShapePart => ({ d, fill: "none", stroke: true });
  const shaded = (d: string, fill: "darken" | "lighten"): ShapePart => ({ d, fill, stroke: true });
  switch (prst) {
    // Lines.
    case "line": case "straightConnector1": return [{ d: `M0,0 L${r(w)},${r(h)}`, fill: "none", stroke: true }];
    case "arc": return [{ d: `M${r(w / 2)},0 A${r(w / 2)},${r(h / 2)} 0 0,1 ${r(w)},${r(h / 2)}`, fill: "none", stroke: true }];
    case "leftBracket": { const k = Math.min(h / 2, m * A("adj", 8333, 0, 50000 * h / m)); return [detail(`M${r(w)},0 Q0,0 0,${r(k)} L0,${r(h - k)} Q0,${r(h)} ${r(w)},${r(h)}`)]; }
    case "rightBracket": { const k = Math.min(h / 2, m * A("adj", 8333, 0, 50000 * h / m)); return [detail(`M0,0 Q${r(w)},0 ${r(w)},${r(k)} L${r(w)},${r(h - k)} Q${r(w)},${r(h)} 0,${r(h)}`)]; }
    case "leftBrace": {
      const k = Math.min(h / 4, m * A("adj1", 8333, 0, 25000 * h / m)), c = h * A("adj2", 50000, 0, 100000);
      return [detail(`M${r(w)},0 Q${r(w / 2)},0 ${r(w / 2)},${r(k)} L${r(w / 2)},${r(c - k)} Q${r(w / 2)},${r(c)} 0,${r(c)} Q${r(w / 2)},${r(c)} ${r(w / 2)},${r(c + k)} L${r(w / 2)},${r(h - k)} Q${r(w / 2)},${r(h)} ${r(w)},${r(h)}`)];
    }
    case "rightBrace": {
      const k = Math.min(h / 4, m * A("adj1", 8333, 0, 25000 * h / m)), c = h * A("adj2", 50000, 0, 100000);
      return [detail(`M0,0 Q${r(w / 2)},0 ${r(w / 2)},${r(k)} L${r(w / 2)},${r(c - k)} Q${r(w / 2)},${r(c)} ${r(w)},${r(c)} Q${r(w / 2)},${r(c)} ${r(w / 2)},${r(c + k)} L${r(w / 2)},${r(h - k)} Q${r(w / 2)},${r(h)} 0,${r(h)}`)];
    }
    case "bracketPair": { const k = m * A("adj", 16667, 0, 50000); return [detail(`M${r(k)},${r(h)} Q0,${r(h)} 0,${r(h - k)} L0,${r(k)} Q0,0 ${r(k)},0 M${r(w - k)},0 Q${r(w)},0 ${r(w)},${r(k)} L${r(w)},${r(h - k)} Q${r(w)},${r(h)} ${r(w - k)},${r(h)}`)]; }
    case "bracePair": { const k = m * 0.083; return [detail(`M${r(2 * k)},${r(h)} Q${r(k)},${r(h)} ${r(k)},${r(h - k)} L${r(k)},${r(h / 2 + k)} Q${r(k)},${r(h / 2)} 0,${r(h / 2)} Q${r(k)},${r(h / 2)} ${r(k)},${r(h / 2 - k)} L${r(k)},${r(k)} Q${r(k)},0 ${r(2 * k)},0 M${r(w - 2 * k)},0 Q${r(w - k)},0 ${r(w - k)},${r(k)} L${r(w - k)},${r(h / 2 - k)} Q${r(w - k)},${r(h / 2)} ${r(w)},${r(h / 2)} Q${r(w - k)},${r(h / 2)} ${r(w - k)},${r(h / 2 + k)} L${r(w - k)},${r(h - k)} Q${r(w - k)},${r(h)} ${r(w - 2 * k)},${r(h)}`)]; }

    // Rectangles.
    case "roundRect": case "flowChartAlternateProcess": {
      const k = m * (prst === "roundRect" ? A("adj", 16667, 0, 50000) : 0.1667);
      return one(`M${r(k)},0 L${r(w - k)},0 Q${r(w)},0 ${r(w)},${r(k)} L${r(w)},${r(h - k)} Q${r(w)},${r(h)} ${r(w - k)},${r(h)} L${r(k)},${r(h)} Q0,${r(h)} 0,${r(h - k)} L0,${r(k)} Q0,0 ${r(k)},0 Z`);
    }
    case "snip1Rect": { const k = m * A("adj", 16667, 0, 50000); return one(P([[0, 0], [w - k, 0], [w, k], [w, h], [0, h]])); }
    case "snip2SameRect": { const k = m * A("adj1", 16667, 0, 50000), b = m * A("adj2", 0, 0, 50000); return one(P([[k, 0], [w - k, 0], [w, k], [w, h - b], [w - b, h], [b, h], [0, h - b], [0, k]])); }
    case "round1Rect": { const k = m * A("adj", 16667, 0, 50000); return one(`M0,0 L${r(w - k)},0 Q${r(w)},0 ${r(w)},${r(k)} L${r(w)},${r(h)} L0,${r(h)} Z`); }
    case "round2SameRect": {
      const k = m * A("adj1", 16667, 0, 50000), b = m * A("adj2", 0, 0, 50000);
      return one(`M${r(k)},0 L${r(w - k)},0 Q${r(w)},0 ${r(w)},${r(k)} L${r(w)},${r(h - b)} Q${r(w)},${r(h)} ${r(w - b)},${r(h)} L${r(b)},${r(h)} Q0,${r(h)} 0,${r(h - b)} L0,${r(k)} Q0,0 ${r(k)},0 Z`);
    }
    case "snipRoundRect": { const k = m * A("adj1", 16667, 0, 50000), s = m * A("adj2", 16667, 0, 50000); return one(`M${r(k)},0 L${r(w - s)},0 L${r(w)},${r(s)} L${r(w)},${r(h)} L0,${r(h)} L0,${r(k)} Q0,0 ${r(k)},0 Z`); }

    // Basic shapes.
    case "ellipse": case "flowChartConnector": return one(E(w / 2, h / 2, w / 2, h / 2));
    case "triangle": return one(P([[w * A("adj", 50000, 0, 100000), 0], [w, h], [0, h]]));
    case "flowChartExtract": return one(P([[w / 2, 0], [w, h], [0, h]]));
    case "flowChartMerge": return one(P([[0, 0], [w, 0], [w / 2, h]]));
    case "rtTriangle": return one(P([[0, 0], [w, h], [0, h]]));
    case "diamond": case "flowChartDecision": return one(P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]));
    case "parallelogram": { const k = m * A("adj", 25000, 0, 100000 * w / m); return one(P([[k, 0], [w, 0], [w - k, h], [0, h]])); }
    case "flowChartInputOutput": { const k = w * 0.2; return one(P([[k, 0], [w, 0], [w - k, h], [0, h]])); }
    case "trapezoid": { const k = m * A("adj", 25000, 0, 50000 * w / m); return one(P([[k, 0], [w - k, 0], [w, h], [0, h]])); }
    case "pentagon": return one(P([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]));
    case "hexagon": { const k = m * A("adj", 25000, 0, 50000 * w / m); return one(P([[k, 0], [w - k, 0], [w, h / 2], [w - k, h], [k, h], [0, h / 2]])); }
    case "heptagon": return one(polygon(7, w, h));
    case "octagon": { const k = m * A("adj", 29289, 0, 50000); return one(P([[k, 0], [w - k, 0], [w, k], [w, h - k], [w - k, h], [k, h], [0, h - k], [0, k]])); }
    case "decagon": return one(polygon(10, w, h));
    case "dodecagon": return one(polygon(12, w, h));
    case "plus": { const k = m * A("adj", 25000, 0, 50000); return one(P([[k, 0], [w - k, 0], [w - k, k], [w, k], [w, h - k], [w - k, h - k], [w - k, h], [k, h], [k, h - k], [0, h - k], [0, k], [k, k]])); }
    case "frame": { const k = m * A("adj1", 12500, 0, 50000); return one(`${P([[0, 0], [w, 0], [w, h], [0, h]])} ${P([[k, k], [k, h - k], [w - k, h - k], [w - k, k]])}`); }
    case "donut": { const k = m * A("adj", 25000, 0, 50000); return one(`${E(w / 2, h / 2, w / 2, h / 2)} ${E(w / 2, h / 2, Math.max(0.01, w / 2 - k), Math.max(0.01, h / 2 - k), true)}`); }
    case "noSmoking": {
      const k = m * A("adj", 18750, 0, 50000);
      const [rx, ry] = [Math.max(0.01, w / 2 - k), Math.max(0.01, h / 2 - k)];
      // The ring, and the bar across it (top left to bottom right, inside the ring).
      const a = Math.PI / 4;
      const dx = Math.cos(a) * rx, dy = Math.sin(a) * ry, t = k * 0.5;
      return one(`${E(w / 2, h / 2, w / 2, h / 2)} ${E(w / 2, h / 2, rx, ry, true)} ${P([[w / 2 - dx - t, h / 2 - dy + t], [w / 2 - dx + t, h / 2 - dy - t], [w / 2 + dx + t, h / 2 + dy - t], [w / 2 + dx - t, h / 2 + dy + t]])}`);
    }
    case "teardrop": return one(`M0,${r(h / 2)} A${r(w / 2)},${r(h / 2)} 0 0,1 ${r(w / 2)},0 L${r(w)},0 L${r(w)},${r(h / 2)} A${r(w / 2)},${r(h / 2)} 0 1,1 0,${r(h / 2)} Z`);
    case "heart": return one(`M${r(w / 2)},${r(h * 0.25)} C${r(w * 0.5)},${r(-h * 0.08)} ${r(-w * 0.08)},${r(-h * 0.02)} ${r(w * 0.02)},${r(h * 0.33)} C${r(w * 0.08)},${r(h * 0.62)} ${r(w * 0.42)},${r(h * 0.82)} ${r(w / 2)},${r(h)} C${r(w * 0.58)},${r(h * 0.82)} ${r(w * 0.92)},${r(h * 0.62)} ${r(w * 0.98)},${r(h * 0.33)} C${r(w * 1.08)},${r(-h * 0.02)} ${r(w * 0.5)},${r(-h * 0.08)} ${r(w / 2)},${r(h * 0.25)} Z`);
    case "lightningBolt": return one(U(w, h, [8472, 0, 12860, 6080, 11050, 6797, 16577, 12007, 14767, 12877, 21600, 21600, 10012, 14915, 12222, 13987, 5022, 9705, 7602, 8382, 0, 3890]));
    case "sun": {
      const rays: string[] = [];
      for (let i = 0; i < 8; i++) {
        const a = (i * Math.PI) / 4;
        const tip: Pt = [w / 2 + Math.cos(a) * w / 2, h / 2 + Math.sin(a) * h / 2];
        const s = 0.16;
        const b1: Pt = [w / 2 + Math.cos(a - s) * w * 0.36, h / 2 + Math.sin(a - s) * h * 0.36];
        const b2: Pt = [w / 2 + Math.cos(a + s) * w * 0.36, h / 2 + Math.sin(a + s) * h * 0.36];
        rays.push(P([b1, tip, b2]));
      }
      return one(`${E(w / 2, h / 2, w * 0.25, h * 0.25)} ${rays.join(" ")}`);
    }
    case "moon": { const k = w * (1 - A("adj", 50000, 0, 87500)); return one(`M${r(w)},0 A${r(w)},${r(h / 2)} 0 0,0 ${r(w)},${r(h)} A${r(Math.max(0.01, k))},${r(h / 2)} 0 0,1 ${r(w)},0 Z`); }
    case "smileyFace": {
      const c = h * 0.7 + h * A("adj", 4653, -4653, 4653) * 3.44;
      return [
        { d: E(w / 2, h / 2, w / 2, h / 2), fill: "shape", stroke: true },
        shaded(`${E(w * 0.33, h * 0.37, w * 0.06, h * 0.06)} ${E(w * 0.67, h * 0.37, w * 0.06, h * 0.06)}`, "darken"),
        detail(`M${r(w * 0.23)},${r(h * 0.7)} Q${r(w / 2)},${r(c)} ${r(w * 0.77)},${r(h * 0.7)}`),
      ];
    }
    case "cloud": return one(cloud(w / 2, h / 2, w / 2 * 0.86, h / 2 * 0.86));
    case "can": case "flowChartMagneticDisk": {
      const k = (prst === "can" ? m * A("adj", 25000, 0, 50000 * h / m) : Math.min(h * 0.25, m * 0.25)) / 2;
      return [
        { d: `M0,${r(k)} A${r(w / 2)},${r(k)} 0 0,0 ${r(w)},${r(k)} L${r(w)},${r(h - k)} A${r(w / 2)},${r(k)} 0 0,1 0,${r(h - k)} Z`, fill: "shape", stroke: true },
        shaded(E(w / 2, k, w / 2, Math.max(0.01, k)), "lighten"),
      ];
    }
    case "flowChartMagneticDrum": {
      const k = Math.min(w * 0.25, m * 0.25) / 2;
      return [
        { d: `M${r(k)},0 L${r(w - k)},0 A${r(k)},${r(h / 2)} 0 0,1 ${r(w - k)},${r(h)} L${r(k)},${r(h)} A${r(k)},${r(h / 2)} 0 0,1 ${r(k)},0 Z`, fill: "shape", stroke: true },
        shaded(E(w - k, h / 2, k, h / 2), "lighten"),
      ];
    }
    case "cube": {
      const k = m * A("adj", 25000, 0, 100000);
      return [
        { d: P([[0, k], [w - k, k], [w - k, h], [0, h]]), fill: "shape", stroke: true },
        shaded(P([[0, k], [k, 0], [w, 0], [w - k, k]]), "lighten"),
        shaded(P([[w - k, k], [w, 0], [w, h - k], [w - k, h]]), "darken"),
      ];
    }
    case "bevel": {
      const k = m * A("adj", 12500, 0, 50000);
      return [
        { d: P([[0, 0], [w, 0], [w, h], [0, h]]), fill: "shape", stroke: true },
        shaded(P([[0, 0], [w, 0], [w - k, k], [k, k]]), "lighten"),
        shaded(P([[0, 0], [k, k], [k, h - k], [0, h]]), "lighten"),
        shaded(P([[w, 0], [w, h], [w - k, h - k], [w - k, k]]), "darken"),
        shaded(P([[0, h], [k, h - k], [w - k, h - k], [w, h]]), "darken"),
        detail(P([[k, k], [w - k, k], [w - k, h - k], [k, h - k]])),
      ];
    }
    case "foldedCorner": {
      const k = m * A("adj", 16667, 0, 50000);
      return [
        { d: P([[0, 0], [w, 0], [w, h - k], [w - k, h], [0, h]]), fill: "shape", stroke: true },
        shaded(`M${r(w - k)},${r(h)} L${r(w - k * 0.8)},${r(h - k * 0.8)} L${r(w)},${r(h - k)} Z`, "darken"),
      ];
    }
    case "chord": return one(`M${r(w * 0.854)},${r(h * 0.854)} A${r(w / 2)},${r(h / 2)} 0 1,1 ${r(w / 2)},0 Z`);
    case "pie": return one(`M${r(w / 2)},${r(h / 2)} L${r(w)},${r(h / 2)} A${r(w / 2)},${r(h / 2)} 0 1,1 ${r(w / 2)},0 Z`);

    // Block arrows (shaft half the height, head as long as half the shorter side).
    case "rightArrow": { const { t, hd } = arrowSizes(h, w); return one(P([[0, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [0, h - t]])); }
    case "leftArrow": { const { t, hd } = arrowSizes(h, w); return one(P([[w, t], [hd, t], [hd, 0], [0, h / 2], [hd, h], [hd, h - t], [w, h - t]])); }
    case "upArrow": { const { t, hd } = arrowSizes(w, h); return one(P([[t, h], [t, hd], [0, hd], [w / 2, 0], [w, hd], [w - t, hd], [w - t, h]])); }
    case "downArrow": { const { t, hd } = arrowSizes(w, h); return one(P([[t, 0], [w - t, 0], [w - t, h - hd], [w, h - hd], [w / 2, h], [0, h - hd], [t, h - hd]])); }
    case "leftRightArrow": { const { t, hd } = arrowSizes(h, w, 2); return one(P([[0, h / 2], [hd, 0], [hd, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [hd, h - t], [hd, h]])); }
    case "upDownArrow": { const { t, hd } = arrowSizes(w, h, 2); return one(P([[w / 2, 0], [w, hd], [w - t, hd], [w - t, h - hd], [w, h - hd], [w / 2, h], [0, h - hd], [t, h - hd], [t, hd], [0, hd]])); }
    case "quadArrow": {
      const a = m * 0.225, s = m * 0.0625, hd = m * 0.225;
      const cx = w / 2, cy = h / 2;
      return one(P([[cx, 0], [cx + a, hd], [cx + s, hd], [cx + s, cy - s], [w - hd, cy - s], [w - hd, cy - a], [w, cy], [w - hd, cy + a], [w - hd, cy + s], [cx + s, cy + s], [cx + s, h - hd], [cx + a, h - hd], [cx, h], [cx - a, h - hd], [cx - s, h - hd], [cx - s, cy + s], [hd, cy + s], [hd, cy + a], [0, cy], [hd, cy - a], [hd, cy - s], [cx - s, cy - s], [cx - s, hd], [cx - a, hd]]));
    }
    case "notchedRightArrow": { const { t, hd } = arrowSizes(h, w); const n = (hd * (h / 2 - t)) / (h / 2 || 1); return one(P([[0, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [0, h - t], [n, h / 2]])); }
    case "stripedRightArrow": {
      const { t, hd } = arrowSizes(h, w); const u = m * 0.03125;
      return one(`${P([[0, t], [u, t], [u, h - t], [0, h - t]])} ${P([[2 * u, t], [4 * u, t], [4 * u, h - t], [2 * u, h - t]])} ${P([[5 * u, t], [w - hd, t], [w - hd, 0], [w, h / 2], [w - hd, h], [w - hd, h - t], [5 * u, h - t]])}`);
    }
    case "homePlate": { const k = m * A("adj", 50000, 0, 100000 * w / m); return one(P([[0, 0], [w - k, 0], [w, h / 2], [w - k, h], [0, h]])); }
    case "chevron": { const k = m * A("adj", 50000, 0, 100000 * w / m); return one(P([[0, 0], [w - k, 0], [w, h / 2], [w - k, h], [0, h], [k, h / 2]])); }

    // Flowchart.
    case "flowChartProcess": return one(P([[0, 0], [w, 0], [w, h], [0, h]]));
    case "flowChartPredefinedProcess": return [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]), fill: "shape", stroke: true }, detail(`M${r(w / 8)},0 L${r(w / 8)},${r(h)} M${r(w * 7 / 8)},0 L${r(w * 7 / 8)},${r(h)}`)];
    case "flowChartInternalStorage": return [{ d: P([[0, 0], [w, 0], [w, h], [0, h]]), fill: "shape", stroke: true }, detail(`M${r(w / 8)},0 L${r(w / 8)},${r(h)} M0,${r(h / 8)} L${r(w)},${r(h / 8)}`)];
    case "flowChartDocument": return one(documentPath(0, 0, w, h));
    case "flowChartMultidocument": {
      const dw = w * 0.82, dh = h * 0.85, o = w * 0.09, oy = h * 0.075;
      return [
        { d: documentPath(2 * o, 0, dw, dh), fill: "shape", stroke: true },
        { d: documentPath(o, oy, dw, dh), fill: "shape", stroke: true },
        { d: documentPath(0, 2 * oy, dw, dh), fill: "shape", stroke: true },
      ];
    }
    case "flowChartTerminator": { const k = Math.min(h / 2, w / 2); return one(`M${r(k)},0 L${r(w - k)},0 A${r(k)},${r(h / 2)} 0 0,1 ${r(w - k)},${r(h)} L${r(k)},${r(h)} A${r(k)},${r(h / 2)} 0 0,1 ${r(k)},0 Z`); }
    case "flowChartPreparation": { const k = w * 0.2; return one(P([[k, 0], [w - k, 0], [w, h / 2], [w - k, h], [k, h], [0, h / 2]])); }
    case "flowChartManualInput": return one(P([[0, h * 0.2], [w, 0], [w, h], [0, h]]));
    case "flowChartManualOperation": return one(P([[0, 0], [w, 0], [w * 0.8, h], [w * 0.2, h]]));
    case "flowChartOffpageConnector": return one(P([[0, 0], [w, 0], [w, h * 0.8], [w / 2, h], [0, h * 0.8]]));
    case "flowChartPunchedCard": return one(P([[w * 0.2, 0], [w, 0], [w, h], [0, h], [0, h * 0.2]]));
    case "flowChartPunchedTape": return one(`M0,${r(h * 0.1)} Q${r(w * 0.25)},${r(h * 0.3)} ${r(w / 2)},${r(h * 0.1)} Q${r(w * 0.75)},${r(-h * 0.1)} ${r(w)},${r(h * 0.1)} L${r(w)},${r(h * 0.9)} Q${r(w * 0.75)},${r(h * 0.7)} ${r(w / 2)},${r(h * 0.9)} Q${r(w * 0.25)},${r(h * 1.1)} 0,${r(h * 0.9)} Z`);
    case "flowChartSummingJunction": {
      const dx = (w / 2) * Math.SQRT1_2, dy = (h / 2) * Math.SQRT1_2;
      return [{ d: E(w / 2, h / 2, w / 2, h / 2), fill: "shape", stroke: true }, detail(`M${r(w / 2 - dx)},${r(h / 2 - dy)} L${r(w / 2 + dx)},${r(h / 2 + dy)} M${r(w / 2 + dx)},${r(h / 2 - dy)} L${r(w / 2 - dx)},${r(h / 2 + dy)}`)];
    }
    case "flowChartOr": return [{ d: E(w / 2, h / 2, w / 2, h / 2), fill: "shape", stroke: true }, detail(`M${r(w / 2)},0 L${r(w / 2)},${r(h)} M0,${r(h / 2)} L${r(w)},${r(h / 2)}`)];
    case "flowChartCollate": return one(P([[0, 0], [w, 0], [0, h], [w, h]]));
    case "flowChartSort": return [{ d: P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]), fill: "shape", stroke: true }, detail(`M0,${r(h / 2)} L${r(w)},${r(h / 2)}`)];
    case "flowChartOnlineStorage": return one(`M${r(w / 6)},0 L${r(w)},0 A${r(w / 6)},${r(h / 2)} 0 0,0 ${r(w)},${r(h)} L${r(w / 6)},${r(h)} A${r(w / 6)},${r(h / 2)} 0 0,1 ${r(w / 6)},0 Z`);
    case "flowChartDelay": return one(`M0,0 L${r(w / 2)},0 A${r(w / 2)},${r(h / 2)} 0 0,1 ${r(w / 2)},${r(h)} L0,${r(h)} Z`);
    case "flowChartMagneticTape": return one(`M${r(w / 2)},${r(h)} A${r(w / 2)},${r(h / 2)} 0 1,1 ${r(w * 0.854)},${r(h * 0.854)} L${r(w)},${r(h * 0.854)} L${r(w)},${r(h)} Z`);
    case "flowChartDisplay": return one(`M0,${r(h / 2)} L${r(w / 6)},0 L${r(w * 5 / 6)},0 A${r(w / 6)},${r(h / 2)} 0 0,1 ${r(w * 5 / 6)},${r(h)} L${r(w / 6)},${r(h)} Z`);

    // Stars and banners.
    case "star4": return one(star(4, A("adj", 12500, 0, 50000) * 2, w, h, 1, 1));
    case "star5": return one(star(5, A("adj", 19098, 0, 50000) * 2, w, h, 1.05146, 1.10557));
    case "star6": return one(star(6, A("adj", 28868, 0, 50000) * 2, w, h, 1, 1));
    case "star7": return one(star(7, A("adj", 34601, 0, 50000) * 2, w, h, 1.02572, 1.0521));
    case "star8": return one(star(8, A("adj", 37500, 0, 50000) * 2, w, h, 1, 1));
    case "star10": return one(star(10, A("adj", 42533, 0, 50000) * 2, w, h, 1.05146, 1));
    case "star12": return one(star(12, A("adj", 37500, 0, 50000) * 2, w, h, 1, 1));
    case "star16": return one(star(16, A("adj", 37500, 0, 50000) * 2, w, h, 1, 1));
    case "star24": return one(star(24, A("adj", 37500, 0, 50000) * 2, w, h, 1, 1));
    case "star32": return one(star(32, A("adj", 37500, 0, 50000) * 2, w, h, 1, 1));
    case "irregularSeal1": return one(U(w, h, [10800, 5800, 8352, 2295, 7312, 6320, 370, 2295, 4627, 7617, 0, 8615, 3722, 11775, 135, 14587, 5667, 13937, 4762, 17617, 7715, 15627, 8485, 21600, 10532, 14935, 13247, 19737, 14020, 14457, 18145, 18095, 16837, 12942, 21600, 13290, 17607, 10475, 21097, 8137, 16702, 7315, 18380, 4457, 14155, 5325, 14522, 0]));
    case "irregularSeal2": return one(U(w, h, [11462, 4342, 9722, 1887, 8550, 6382, 4502, 3625, 5372, 7955, 1172, 8550, 3075, 11487, 0, 12975, 3375, 14420, 1980, 19107, 6022, 17157, 6750, 21600, 8550, 16880, 11017, 18842, 12038, 16095, 14700, 18607, 14423, 15187, 17850, 15832, 16032, 13110, 19980, 12225, 17370, 10162, 21600, 7575, 17107, 6202, 18255, 1125, 13987, 4425, 12810, 2235]));
    case "wave": { const a = h * A("adj1", 12500, 0, 20000); return one(`M0,${r(a)} C${r(w / 3)},${r(-a)} ${r(w * 2 / 3)},${r(3 * a)} ${r(w)},${r(a)} L${r(w)},${r(h - a)} C${r(w * 2 / 3)},${r(h - 3 * a)} ${r(w / 3)},${r(h + a)} 0,${r(h - a)} Z`); }
    case "doubleWave": {
      const a = h * A("adj1", 6250, 0, 12500);
      return one(`M0,${r(a)} C${r(w / 6)},${r(-a)} ${r(w / 3)},${r(3 * a)} ${r(w / 2)},${r(a)} C${r(w * 2 / 3)},${r(-a)} ${r(w * 5 / 6)},${r(3 * a)} ${r(w)},${r(a)} L${r(w)},${r(h - a)} C${r(w * 5 / 6)},${r(h - 3 * a)} ${r(w * 2 / 3)},${r(h + a)} ${r(w / 2)},${r(h - a)} C${r(w / 3)},${r(h - 3 * a)} ${r(w / 6)},${r(h + a)} 0,${r(h - a)} Z`);
    }

    // Callouts (the pointer below, toward the left, as Word puts it).
    case "wedgeRectCallout": return one(rectCallout(w, h, 0, calloutTip(w, h)));
    case "wedgeRoundRectCallout": return one(rectCallout(w, h, m * A("adj3", 16667, 0, 50000), calloutTip(w, h)));
    case "wedgeEllipseCallout": {
      // The ellipse, open where the pointer leaves it (toward the pointer's tip).
      const tip = calloutTip(w, h);
      const toward = Math.atan2((tip[1] - h / 2) / h, (tip[0] - w / 2) / w);
      const [a1, a2] = [toward - 0.2, toward + 0.2];
      const p1: Pt = [w / 2 + Math.cos(a1) * w / 2, h / 2 + Math.sin(a1) * h / 2];
      const p2: Pt = [w / 2 + Math.cos(a2) * w / 2, h / 2 + Math.sin(a2) * h / 2];
      const inside = Math.hypot((tip[0] - w / 2) / (w / 2), (tip[1] - h / 2) / (h / 2)) <= 1;
      return one(inside ? E(w / 2, h / 2, w / 2, h / 2) : `M${r(p2[0])},${r(p2[1])} A${r(w / 2)},${r(h / 2)} 0 1,1 ${r(p1[0])},${r(p1[1])} L${r(tip[0])},${r(tip[1])} Z`);
    }
    case "cloudCallout": {
      // The cloud, and bubbles from its edge toward the tip, smaller as they go.
      const tip = calloutTip(w, h);
      const toward = Math.atan2((tip[1] - h / 2) / h, (tip[0] - w / 2) / w);
      const edge: Pt = [w / 2 + Math.cos(toward) * w / 2 * 0.95, h / 2 + Math.sin(toward) * h / 2 * 0.95];
      const bubble = (k: number, size: number) => E(edge[0] + (tip[0] - edge[0]) * k, edge[1] + (tip[1] - edge[1]) * k, m * size, m * size);
      const far = Math.hypot((tip[0] - w / 2) / (w / 2), (tip[1] - h / 2) / (h / 2)) > 1;
      return one(`${cloud(w / 2, h / 2, w / 2 * 0.86, h / 2 * 0.86)}${far ? ` ${bubble(0.35, 0.06)} ${bubble(0.7, 0.04)} ${bubble(1, 0.025)}` : ""}`);
    }

    // Equation shapes.
    case "mathPlus": { const t = m * 0.1176, a = m * 0.37; const cx = w / 2, cy = h / 2; return one(P([[cx - t, cy - a], [cx + t, cy - a], [cx + t, cy - t], [cx + a, cy - t], [cx + a, cy + t], [cx + t, cy + t], [cx + t, cy + a], [cx - t, cy + a], [cx - t, cy + t], [cx - a, cy + t], [cx - a, cy - t], [cx - t, cy - t]])); }
    case "mathMinus": { const t = h * 0.1176, a = w * 0.367; return one(P([[w / 2 - a, h / 2 - t], [w / 2 + a, h / 2 - t], [w / 2 + a, h / 2 + t], [w / 2 - a, h / 2 + t]])); }
    case "mathMultiply": {
      // A plus turned a quarter turn.
      const t = m * 0.083, a = m * 0.36, c = Math.SQRT1_2;
      const plus: Pt[] = [[-t, -a], [t, -a], [t, -t], [a, -t], [a, t], [t, t], [t, a], [-t, a], [-t, t], [-a, t], [-a, -t], [-t, -t]];
      return one(P(plus.map(([x, y]) => [w / 2 + (x - y) * c, h / 2 + (x + y) * c] as Pt)));
    }
    case "mathDivide": { const t = h * 0.0588, a = w * 0.367, d = m * 0.0882; return one(`${P([[w / 2 - a, h / 2 - t], [w / 2 + a, h / 2 - t], [w / 2 + a, h / 2 + t], [w / 2 - a, h / 2 + t]])} ${E(w / 2, h / 2 - t - d * 2.2, d, d)} ${E(w / 2, h / 2 + t + d * 2.2, d, d)}`); }
    case "mathEqual": { const t = h * 0.1176, g = h * 0.0588, a = w * 0.367; return one(`${P([[w / 2 - a, h / 2 - g - 2 * t], [w / 2 + a, h / 2 - g - 2 * t], [w / 2 + a, h / 2 - g], [w / 2 - a, h / 2 - g]])} ${P([[w / 2 - a, h / 2 + g], [w / 2 + a, h / 2 + g], [w / 2 + a, h / 2 + g + 2 * t], [w / 2 - a, h / 2 + g + 2 * t]])}`); }
    case "mathNotEqual": {
      const t = h * 0.1176, g = h * 0.0588, a = w * 0.367;
      const bars = `${P([[w / 2 - a, h / 2 - g - 2 * t], [w / 2 + a, h / 2 - g - 2 * t], [w / 2 + a, h / 2 - g], [w / 2 - a, h / 2 - g]])} ${P([[w / 2 - a, h / 2 + g], [w / 2 + a, h / 2 + g], [w / 2 + a, h / 2 + g + 2 * t], [w / 2 - a, h / 2 + g + 2 * t]])}`;
      const s = t * 0.8;
      return one(`${bars} ${P([[w * 0.58 - s, h * 0.1], [w * 0.58 + s, h * 0.1], [w * 0.42 + s, h * 0.9], [w * 0.42 - s, h * 0.9]])}`);
    }
    default: return one(P([[0, 0], [w, 0], [w, h], [0, h]]));
  }
}

/** The outline of Word's preset geometry `prst` in a box of `w`×`h` points (its main part). */
export function shapePath(prst: string, w: number, h: number, adj?: Adjust): string {
  return parts(prst, w, h, adj)[0].d;
}

/** A dash pattern for SVG / PDF (in pt) from Word's dash name and the line width; undefined for solid. */
export function dashArray(dash: string | undefined, width: number): number[] | undefined {
  const k = Math.max(width, 0.5);
  const units: Record<string, number[]> = {
    dash: [4, 3], sysDash: [3, 1], sysDot: [1, 1], dot: [1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
    dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
  };
  const u = dash ? units[dash] : undefined;
  return u?.map((n) => r(n * k));
}

/** An arrow head (or other line end) at `tip`, pointing away from `from`, for a line `width` thick. */
function lineEnd(kind: LineEnd, tip: Pt, from: Pt, width: number): ShapePart {
  const [dx, dy] = [tip[0] - from[0], tip[1] - from[1]];
  const len = Math.hypot(dx, dy) || 1;
  const [ux, uy] = [dx / len, dy / len]; // along the line, toward the tip
  const [nx, ny] = [-uy, ux];
  const size = Math.max(width * 3, 6.5);
  const L = size, W = size / 2;
  const at = (a: number, b: number): Pt => [tip[0] - ux * a + nx * b, tip[1] - uy * a + ny * b];
  switch (kind) {
    case "arrow": {
      const [l, rr] = [at(L, W), at(L, -W)];
      return { d: `M${r(l[0])},${r(l[1])} L${r(tip[0])},${r(tip[1])} L${r(rr[0])},${r(rr[1])}`, fill: "none", stroke: true, head: true };
    }
    case "stealth": return { d: P([tip, at(L, W), at(L * 0.6, 0), at(L, -W)]), fill: "line", stroke: false, head: true };
    case "diamond": return { d: P([tip, at(L / 2, W), at(L, 0), at(L / 2, -W)].map(([x, y]) => [x + ux * L / 2, y + uy * L / 2] as Pt)), fill: "line", stroke: false, head: true };
    case "oval": { const c = at(0, 0); return { d: E(c[0], c[1], W, W), fill: "line", stroke: false, head: true }; }
    default: return { d: P([tip, at(L, W), at(L, -W)]), fill: "line", stroke: false, head: true };
  }
}

/**
 * A shape's whole drawing in a w×h box: its outline and details, and a line's ends (arrow heads). Flips
 * and rotation are not applied here (the screen and the PDF turn the whole drawing).
 */
export function shapeDrawing(style: ShapeStyle, w: number, h: number): ShapePart[] {
  const out = parts(style.geometry, w, h, style.adj);
  const line: ShapeLine | undefined = style.line;
  if (line && (style.geometry === "line" || style.geometry === "straightConnector1")) {
    if (line.head) out.push(lineEnd(line.head, [0, 0], [w, h], line.width));
    if (line.tail) out.push(lineEnd(line.tail, [w, h], [0, 0], line.width));
  }
  return out;
}

/** A fill shaded from the shape's: darker (a cube's side) or lighter (its top). */
export function shadeFill(fill: string, how: "darken" | "lighten"): string {
  const [rr, g, b] = [0, 2, 4].map((i) => parseInt(fill.slice(i, i + 2), 16));
  const f = (v: number) => Math.round(how === "darken" ? v * 0.8 : v + (255 - v) * 0.4).toString(16).padStart(2, "0");
  return (f(rr) + f(g) + f(b)).toUpperCase();
}

/**
 * Where text sits in a shape, as shares of its width and height taken off each side (Word's text
 * rectangle for the geometry): inside an oval, below a triangle's point, within an arrow's shaft.
 */
export function textArea(prst: string, w: number, h: number, adj?: Adjust): { left: number; top: number; right: number; bottom: number } {
  const m = Math.min(w, h);
  const A = adjust(adj);
  const all = (k: number) => ({ left: k, top: k, right: k, bottom: k });
  const sides = (left: number, top: number, right: number, bottom: number) => ({ left, top, right, bottom });
  switch (prst) {
    case "ellipse": case "flowChartConnector": case "flowChartSummingJunction": case "flowChartOr": case "wedgeEllipseCallout": case "smileyFace": return all(0.146);
    case "donut": case "noSmoking": return all(0.2);
    case "triangle": case "flowChartExtract": return sides(0.25, 0.5, 0.25, 0);
    case "flowChartMerge": return sides(0.25, 0, 0.25, 0.5);
    case "rtTriangle": return sides(0.083, 0.583, 0.417, 0.083);
    case "diamond": case "flowChartDecision": case "flowChartSort": return all(0.25);
    case "parallelogram": case "trapezoid": case "hexagon": { const k = (m * 0.25) / w; return sides(k * 0.67, 0, k * 0.67, 0); }
    case "flowChartInputOutput": case "flowChartPreparation": return sides(0.2, 0, 0.2, 0);
    case "pentagon": return sides(0.19, 0.36, 0.19, 0);
    case "octagon": { const k = (m * 0.2929) / 2; return sides(k / w, k / h, k / w, k / h); }
    case "heptagon": case "decagon": case "dodecagon": return all(0.12);
    case "star4": return all(0.3);
    case "star5": return sides(0.31, 0.38, 0.31, 0.2);
    case "star6": case "star7": case "star8": case "star10": case "star12": case "star16": case "star24": case "star32": return all(0.2);
    case "irregularSeal1": case "irregularSeal2": return all(0.28);
    case "heart": return sides(0.2, 0.25, 0.2, 0.3);
    case "cloud": case "cloudCallout": return sides(0.15, 0.2, 0.15, 0.2);
    case "plus": { const k = (m * 0.25); return sides(0, k / h, 0, k / h); }
    case "frame": case "bevel": { const k = m * 0.125; return sides(k / w, k / h, k / w, k / h); }
    case "can": case "flowChartMagneticDisk": return sides(0, (Math.min(h * 0.25, m * 0.25)) / h, 0, 0);
    case "flowChartMagneticDrum": return sides(0, 0, (Math.min(w * 0.25, m * 0.25)) / w, 0);
    case "cube": { const k = m * 0.25; return sides(0, k / h, k / w, 0); }
    case "foldedCorner": return sides(0, 0, 0, 0.167);
    case "flowChartDocument": return sides(0, 0, 0, 0.2);
    case "flowChartMultidocument": return sides(0, 0.15, 0.18, 0.25);
    case "flowChartTerminator": return sides(Math.min(h / 2, w / 2) * 0.3 / w, 0, Math.min(h / 2, w / 2) * 0.3 / w, 0);
    case "flowChartDelay": return sides(0, 0.146, 0.146, 0.146);
    case "flowChartOnlineStorage": case "flowChartDisplay": return sides(0.167, 0, 0.167, 0);
    case "flowChartPredefinedProcess": return sides(0.125, 0, 0.125, 0);
    case "flowChartInternalStorage": return sides(0.125, 0.125, 0, 0);
    case "flowChartManualInput": return sides(0, 0.2, 0, 0);
    case "flowChartManualOperation": return sides(0.2, 0, 0.2, 0);
    case "flowChartOffpageConnector": return sides(0, 0, 0, 0.2);
    case "flowChartPunchedTape": return sides(0, 0.2, 0, 0.2);
    case "flowChartPunchedCard": return sides(0, 0.2, 0, 0);
    case "wave": return sides(0, 0.2, 0, 0.2);
    case "doubleWave": return sides(0, 0.12, 0, 0.12);
    case "rightArrow": case "leftArrow": case "notchedRightArrow": case "stripedRightArrow": case "leftRightArrow": {
      const shaft = (1 - A("adj1", 50000, 0, 100000)) / 2;
      const heads = prst === "leftRightArrow" ? 2 : 1;
      const hd = Math.min(w / heads, m * A("adj2", 50000, 0, (100000 * w) / (heads * m))) / w;
      return sides(prst === "leftArrow" || prst === "leftRightArrow" ? hd : prst === "notchedRightArrow" ? 0.12 : 0, shaft, prst === "leftArrow" ? 0 : hd, shaft);
    }
    case "upArrow": case "downArrow": case "upDownArrow": {
      const shaft = (1 - A("adj1", 50000, 0, 100000)) / 2;
      const heads = prst === "upDownArrow" ? 2 : 1;
      const hd = Math.min(h / heads, m * A("adj2", 50000, 0, (100000 * h) / (heads * m))) / h;
      return sides(shaft, prst === "downArrow" ? 0 : hd, shaft, prst === "upArrow" ? 0 : hd);
    }
    case "homePlate": return sides(0, 0, (m * 0.5) / w, 0);
    case "chevron": return sides((m * 0.5) / w, 0, (m * 0.5) / w, 0);
    case "roundRect": case "flowChartAlternateProcess": case "wedgeRoundRectCallout": { const k = (m * 0.1667 * 0.29); return sides(k / w, k / h, k / w, k / h); }
    case "sun": return all(0.3);
    case "moon": return sides(0.5, 0.25, 0, 0.25);
    case "teardrop": return all(0.146);
    case "mathPlus": case "mathMinus": case "mathMultiply": case "mathDivide": case "mathEqual": case "mathNotEqual": return all(0.12);
    default: return all(0);
  }
}

/**
 * The yellow handles that change a shape's adjustments (Word's): for each, where it is in the w×h box
 * and what the adjustments become when it is dragged to a point.
 */
interface Handle { at(w: number, h: number, A: ReturnType<typeof adjust>): Pt; to(x: number, y: number, w: number, h: number): Adjust }
const k100 = (v: number) => Math.round(v * 100000);
const clamp = (v: number, lo: number, hi: number) => Math.min(Math.max(v, lo), hi);
const HANDLES: Record<string, Handle[]> = (() => {
  const ss = (w: number, h: number) => Math.min(w, h);
  /** A handle along the top edge, `k` from the left (or from the right). */
  const fromLeft = (name: string, def: number, max: (w: number, h: number) => number): Handle => ({
    at: (w, h, A) => [ss(w, h) * A(name, def, 0, max(w, h)), 0],
    to: (x, _y, w, h) => ({ [name]: k100(clamp(x / ss(w, h), 0, max(w, h) / 100000)) }),
  });
  const fromRight = (name: string, def: number, max: (w: number, h: number) => number, y: (h: number) => number = () => 0): Handle => ({
    at: (w, h, A) => [w - ss(w, h) * A(name, def, 0, max(w, h)), y(h)],
    to: (x, _y, w, h) => ({ [name]: k100(clamp((w - x) / ss(w, h), 0, max(w, h) / 100000)) }),
  });
  const half = () => 50000;
  const star = (def: number): Handle[] => [{
    at: (w, h, A) => [w / 2, h / 2 - (h / 2) * A("adj", def, 0, 50000) * 2],
    to: (_x, y, _w, h) => ({ adj: k100(clamp((h / 2 - y) / (h / 2) / 2, 0, 0.5)) }),
  }];
  /** A block arrow's two handles: shaft thickness (at the head's base) and head length. */
  const arrow = (dir: "r" | "l" | "u" | "d", heads = 1): Handle[] => {
    const horizontal = dir === "r" || dir === "l";
    const size = (w: number, h: number, A: ReturnType<typeof adjust>) => {
      const [across, along] = horizontal ? [h, w] : [w, h];
      const m = ss(w, h);
      return { across, along, m, shaft: across * A("adj1", 50000, 0, 100000), hd: Math.min(along / heads, m * A("adj2", 50000, 0, (100000 * along) / (heads * m))) };
    };
    const base = (along: number, hd: number) => (dir === "r" || dir === "d" ? along - hd : hd);
    return [
      {
        at: (w, h, A) => { const s = size(w, h, A); const b = base(s.along, s.hd); const t = (s.across - s.shaft) / 2; return horizontal ? [b, t] : [t, b]; },
        to: (x, y, w, h) => { const across = horizontal ? h : w; const v = horizontal ? y : x; return { adj1: k100(clamp(1 - (2 * v) / across, 0, 1)) }; },
      },
      {
        at: (w, h, A) => { const s = size(w, h, A); const b = base(s.along, s.hd); return horizontal ? [b, 0] : [0, b]; },
        to: (x, y, w, h) => {
          const along = horizontal ? w : h;
          const v = horizontal ? x : y;
          const hd = dir === "r" || dir === "d" ? along - v : v;
          return { adj2: k100(clamp(hd / ss(w, h), 0, along / heads / ss(w, h))) };
        },
      },
    ];
  };
  const callout: Handle = {
    at: (w, h, A) => [w / 2 + w * A("adj1", -20833, -10000000, 10000000), h / 2 + h * A("adj2", 62500, -10000000, 10000000)],
    to: (x, y, w, h) => ({ adj1: k100((x - w / 2) / w), adj2: k100((y - h / 2) / h) }),
  };
  return {
    roundRect: [fromLeft("adj", 16667, half)],
    snip1Rect: [fromRight("adj", 16667, half)],
    round1Rect: [fromRight("adj", 16667, half)],
    snip2SameRect: [fromLeft("adj1", 16667, half), { at: (w, h, A) => [ss(w, h) * A("adj2", 0, 0, 50000), h], to: (x, _y, w, h) => ({ adj2: k100(clamp(x / ss(w, h), 0, 0.5)) }) }],
    round2SameRect: [fromRight("adj1", 16667, half), { at: (w, h, A) => [ss(w, h) * A("adj2", 0, 0, 50000), h], to: (x, _y, w, h) => ({ adj2: k100(clamp(x / ss(w, h), 0, 0.5)) }) }],
    snipRoundRect: [fromLeft("adj1", 16667, half), fromRight("adj2", 16667, half)],
    triangle: [{ at: (w, _h, A) => [w * A("adj", 50000, 0, 100000), 0], to: (x, _y, w) => ({ adj: k100(clamp(x / w, 0, 1)) }) }],
    parallelogram: [fromLeft("adj", 25000, (w, h) => (100000 * w) / ss(w, h))],
    trapezoid: [fromLeft("adj", 25000, (w, h) => (50000 * w) / ss(w, h))],
    hexagon: [fromLeft("adj", 25000, (w, h) => (50000 * w) / ss(w, h))],
    octagon: [fromLeft("adj", 29289, half)],
    plus: [fromLeft("adj", 25000, half)],
    frame: [fromLeft("adj1", 12500, half)],
    bevel: [fromLeft("adj", 12500, half)],
    donut: [{ at: (w, h, A) => [ss(w, h) * A("adj", 25000, 0, 50000), h / 2], to: (x, _y, w, h) => ({ adj: k100(clamp(x / ss(w, h), 0, 0.5)) }) }],
    noSmoking: [{ at: (w, h, A) => [ss(w, h) * A("adj", 18750, 0, 50000), h / 2], to: (x, _y, w, h) => ({ adj: k100(clamp(x / ss(w, h), 0, 0.5)) }) }],
    can: [{ at: (w, h, A) => [w / 2, ss(w, h) * A("adj", 25000, 0, (50000 * h) / ss(w, h))], to: (_x, y, w, h) => ({ adj: k100(clamp(y / ss(w, h), 0, h / ss(w, h) / 2)) }) }],
    cube: [{ at: (w, h, A) => [0, ss(w, h) * A("adj", 25000, 0, 100000)], to: (_x, y, w, h) => ({ adj: k100(clamp(y / ss(w, h), 0, 1)) }) }],
    foldedCorner: [fromRight("adj", 16667, half, (h) => h)],
    homePlate: [fromRight("adj", 50000, (w, h) => (100000 * w) / ss(w, h))],
    chevron: [fromRight("adj", 50000, (w, h) => (100000 * w) / ss(w, h))],
    moon: [{ at: (w, h, A) => [w * A("adj", 50000, 0, 87500), h / 2], to: (x, _y, w) => ({ adj: k100(clamp(x / w, 0, 0.875)) }) }],
    smileyFace: [{ at: (w, h, A) => [w / 2, h * 0.7 + (h * A("adj", 4653, -4653, 4653) * 3.44) / 2], to: (_x, y, _w, h) => ({ adj: k100(clamp(((y - h * 0.7) * 2) / (h * 3.44), -0.04653, 0.04653)) }) }],
    wave: [{ at: (_w, h, A) => [0, h * A("adj1", 12500, 0, 20000)], to: (_x, y, _w, h) => ({ adj1: k100(clamp(y / h, 0, 0.2)) }) }],
    doubleWave: [{ at: (_w, h, A) => [0, h * A("adj1", 6250, 0, 12500)], to: (_x, y, _w, h) => ({ adj1: k100(clamp(y / h, 0, 0.125)) }) }],
    leftBracket: [{ at: (w, h, A) => [w, Math.min(h / 2, ss(w, h) * A("adj", 8333, 0, (50000 * h) / ss(w, h)))], to: (_x, y, w, h) => ({ adj: k100(clamp(y / ss(w, h), 0, h / 2 / ss(w, h))) }) }],
    rightBracket: [{ at: (w, h, A) => [0, Math.min(h / 2, ss(w, h) * A("adj", 8333, 0, (50000 * h) / ss(w, h)))], to: (_x, y, w, h) => ({ adj: k100(clamp(y / ss(w, h), 0, h / 2 / ss(w, h))) }) }],
    bracketPair: [fromLeft("adj", 16667, half)],
    leftBrace: [{ at: (_w, h, A) => [0, h * A("adj2", 50000, 0, 100000)], to: (_x, y, _w, h) => ({ adj2: k100(clamp(y / h, 0, 1)) }) }],
    rightBrace: [{ at: (w, h, A) => [w, h * A("adj2", 50000, 0, 100000)], to: (_x, y, _w, h) => ({ adj2: k100(clamp(y / h, 0, 1)) }) }],
    rightArrow: arrow("r"), leftArrow: arrow("l"), upArrow: arrow("u"), downArrow: arrow("d"),
    notchedRightArrow: arrow("r"), stripedRightArrow: arrow("r"),
    leftRightArrow: arrow("l", 2), upDownArrow: arrow("u", 2),
    star4: star(12500), star5: star(19098), star6: star(28868), star7: star(34601), star8: star(37500),
    star10: star(42533), star12: star(37500), star16: star(37500), star24: star(37500), star32: star(37500),
    wedgeRectCallout: [callout], wedgeRoundRectCallout: [callout], wedgeEllipseCallout: [callout], cloudCallout: [callout],
  };
})();

/** Where a shape's adjustment handles are, in its w×h box (none for most flowchart and fixed shapes). */
export function shapeHandles(style: ShapeStyle, w: number, h: number): { x: number; y: number }[] {
  const A = adjust(style.adj);
  return (HANDLES[style.geometry] ?? []).map((hd) => { const [x, y] = hd.at(w, h, A); return { x, y }; });
}

/** The shape's adjustments with handle `index` dragged to (x, y) in its w×h box. */
export function adjustShape(style: ShapeStyle, index: number, x: number, y: number, w: number, h: number): Adjust {
  const hd = HANDLES[style.geometry]?.[index];
  return hd ? { ...style.adj, ...hd.to(x, y, w, h) } : { ...style.adj };
}
