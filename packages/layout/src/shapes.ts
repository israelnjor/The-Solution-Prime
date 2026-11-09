/**
 * Shapes' outlines as SVG paths in a w×h box (origin top-left, y down): one drawing for the screen
 * (SVG) and the PDF (pdf-lib's drawSvgPath). Word's common preset geometries; others as rectangles.
 */

/** The outline of Word's preset geometry `prst` in a box of `w`×`h` points. */
export function shapePath(prst: string, w: number, h: number): string {
  const r = (n: number) => Math.round(n * 100) / 100;
  const P = (pts: [number, number][]) => `M${pts.map(([x, y]) => `${r(x)},${r(y)}`).join(" L")} Z`;
  const m = Math.min(w, h);
  switch (prst) {
    case "ellipse": case "flowChartConnector":
      return `M0,${r(h / 2)} A${r(w / 2)},${r(h / 2)} 0 1,0 ${r(w)},${r(h / 2)} A${r(w / 2)},${r(h / 2)} 0 1,0 0,${r(h / 2)} Z`;
    case "roundRect": case "flowChartAlternateProcess": {
      const k = m * 0.1667;
      return `M${r(k)},0 L${r(w - k)},0 Q${r(w)},0 ${r(w)},${r(k)} L${r(w)},${r(h - k)} Q${r(w)},${r(h)} ${r(w - k)},${r(h)} L${r(k)},${r(h)} Q0,${r(h)} 0,${r(h - k)} L0,${r(k)} Q0,0 ${r(k)},0 Z`;
    }
    case "triangle": case "flowChartExtract": return P([[w / 2, 0], [w, h], [0, h]]);
    case "rtTriangle": return P([[0, 0], [w, h], [0, h]]);
    case "diamond": case "flowChartDecision": return P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
    case "parallelogram": case "flowChartInputOutput": { const k = m * 0.25; return P([[k, 0], [w, 0], [w - k, h], [0, h]]); }
    case "trapezoid": { const k = m * 0.25; return P([[k, 0], [w - k, 0], [w, h], [0, h]]); }
    case "pentagon": case "homePlate": { const k = m * 0.5; return P([[0, 0], [w - k, 0], [w, h / 2], [w - k, h], [0, h]]); }
    case "hexagon": { const k = m * 0.25; return P([[k, 0], [w - k, 0], [w, h / 2], [w - k, h], [k, h], [0, h / 2]]); }
    case "octagon": { const k = m * 0.2929; return P([[k, 0], [w - k, 0], [w, k], [w, h - k], [w - k, h], [k, h], [0, h - k], [0, k]]); }
    case "rightArrow": { const hd = Math.min(w, h * 1); const t = h * 0.25; return P([[0, t], [w - hd / 2, t], [w - hd / 2, 0], [w, h / 2], [w - hd / 2, h], [w - hd / 2, h - t], [0, h - t]]); }
    case "leftArrow": { const hd = Math.min(w, h); const t = h * 0.25; return P([[w, t], [hd / 2, t], [hd / 2, 0], [0, h / 2], [hd / 2, h], [hd / 2, h - t], [w, h - t]]); }
    case "upArrow": { const hd = Math.min(h, w); const t = w * 0.25; return P([[t, h], [t, hd / 2], [0, hd / 2], [w / 2, 0], [w, hd / 2], [w - t, hd / 2], [w - t, h]]); }
    case "downArrow": { const hd = Math.min(h, w); const t = w * 0.25; return P([[t, 0], [w - t, 0], [w - t, h - hd / 2], [w, h - hd / 2], [w / 2, h], [0, h - hd / 2], [t, h - hd / 2]]); }
    case "star5": {
      const pts: [number, number][] = [];
      for (let i = 0; i < 10; i++) {
        const a = -Math.PI / 2 + (i * Math.PI) / 5;
        const k = i % 2 ? 0.38 : 1;
        pts.push([w / 2 + (Math.cos(a) * w * k) / 2, h / 2 + (Math.sin(a) * h * k) / 2 + h * 0.05]);
      }
      return P(pts);
    }
    case "line": case "straightConnector1": return `M0,0 L${r(w)},${r(h)}`;
    default: return P([[0, 0], [w, 0], [w, h], [0, h]]);
  }
}

/** A line, not a closed outline (nothing to fill). */
export const isLineShape = (prst: string) => prst === "line" || prst === "straightConnector1";
