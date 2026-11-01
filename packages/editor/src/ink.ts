/** Ink (pen strokes) as an SVG element sized to its box: strokes are stored relative to the box. */
import type { FloatInfo } from "@prime/pdm";

const SVG = "http://www.w3.org/2000/svg";

export function inkSvg(ink: NonNullable<FloatInfo["ink"]>, width: number, height: number): SVGSVGElement {
  const svg = document.createElementNS(SVG, "svg");
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.cssText = "width:100%;height:100%;overflow:visible";
  for (const s of ink.strokes) {
    const path = document.createElementNS(SVG, "path");
    let d = "";
    for (let i = 0; i < s.length; i += 2) d += `${i ? "L" : "M"}${s[i] * width} ${s[i + 1] * height}`;
    if (s.length === 2) d += `L${s[0] * width + 0.01} ${s[1] * height}`; // a dot
    path.setAttribute("d", d);
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", `#${ink.color}`);
    path.setAttribute("stroke-width", String(ink.width));
    path.setAttribute("stroke-linecap", "round");
    path.setAttribute("stroke-linejoin", "round");
    path.setAttribute("vector-effect", "non-scaling-stroke");
    svg.appendChild(path);
  }
  return svg;
}
