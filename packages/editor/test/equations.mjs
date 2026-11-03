// Equations: LaTeX → OMML → LaTeX for the notation the editor covers; then a document for Word.
// Usage: node equations.mjs [out.docx]
import { writeFileSync } from "node:fs";
import { writeDocx } from "@prime/docx";
import { blankDocument, latexToOmml, ommlToLatex } from "@prime/editor";

const cases = [
  String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`,
  String.raw`E = mc^2`,
  String.raw`\sum_{i=1}^{n} x_i^2`,
  String.raw`\int_0^1 f(x) dx`,
  String.raw`\lim_{x \to 0} \frac{\sin x}{x} = 1`,
  String.raw`\left[ \alpha + \beta \right] \leq \pi r^2`,
  String.raw`\sqrt[3]{27} = 3`,
  String.raw`\hat{y} = \bar{x} + \vec{v}`,
  String.raw`a_{ij} \times b_{jk}`,
  String.raw`\text{Area} = \pi r^2`,
];
let ok = true;
const blocks = [];
for (const c of cases) {
  const xml = latexToOmml(c);
  const back = ommlToLatex(xml);
  const again = back && latexToOmml(back.latex);
  const same = again === xml;
  ok &&= same;
  console.log(same ? "ok  " : "FAIL", c, "→", back?.latex);
  blocks.push({ type: "paragraph", props: {}, content: [{ type: "text", text: `${c}:  ` }, { type: "preserved", preserved: { name: "m:oMath", xml: xml.replace(/ xmlns:m="[^"]+"/, "") } }] });
}
for (const bad of [String.raw`\frac{1}`, String.raw`\unknown x`, String.raw`\left( x`]) {
  try { latexToOmml(bad); console.log("FAIL (accepted)", bad); ok = false; } catch (e) { console.log("ok   rejected", bad, "—", e.message); }
}
const display = latexToOmml(String.raw`\sum_{k=0}^{\infty} \frac{1}{k!} = e`, true);
console.log("display:", ommlToLatex(display));
blocks.push({ type: "paragraph", props: {}, content: [{ type: "preserved", preserved: { name: "m:oMathPara", xml: display.replace(/ xmlns:m="[^"]+"/, "") } }] });
console.log(ok ? "all passed" : "FAILED");
if (process.argv[2]) { const d = blankDocument(); d.sections[0].blocks = blocks; writeFileSync(process.argv[2], writeDocx(d)); }
