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
  String.raw`\int_1^4 x^3 + 7x + 12 = ?? express the integrand at x = 4`,
  String.raw`f(x) = sin x + cos x`,
  String.raw`y = x^2 for all real x`,
  "\\int_0^1 x dx\ny = 2x + 1\nwhere x is the input",
  String.raw`A = \begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}`,
  String.raw`\det \begin{vmatrix} a & b \\ c & d \end{vmatrix} = ad - bc`,
  String.raw`\begin{bmatrix} x \\ y \\ z \end{bmatrix}`,
  String.raw`|x| = \begin{cases} x & x \geq 0 \\ -x & x < 0 \end{cases}`,
  String.raw`\binom{n}{k} = \frac{n!}{k!(n-k)!}`,
  String.raw`\log_{2} 8 = 3`,
  String.raw`\sin^2 x + \cos^2 x = 1`,
  "M = \\begin{pmatrix}\n1 & 0 \\\\\n0 & 1\n\\end{pmatrix}",
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
  const name = xml.startsWith("<m:oMathPara") ? "m:oMathPara" : "m:oMath";
  blocks.push({ type: "paragraph", props: {}, content: [{ type: "text", text: `${c.replace(/\n/g, " / ")}:  ` }, { type: "preserved", preserved: { name, xml: xml.replace(/ xmlns:m="[^"]+"/, "") } }] });
}
for (const bad of [String.raw`\frac{1}`, String.raw`\unknown x`, String.raw`\left( x`]) {
  try { latexToOmml(bad); console.log("FAIL (accepted)", bad); ok = false; } catch (e) { console.log("ok   rejected", bad, "—", e.message); }
}
const display = latexToOmml(String.raw`\sum_{k=0}^{\infty} \frac{1}{k!} = e`, true);
console.log("display:", ommlToLatex(display));
blocks.push({ type: "paragraph", props: {}, content: [{ type: "preserved", preserved: { name: "m:oMathPara", xml: display.replace(/ xmlns:m="[^"]+"/, "") } }] });
console.log(ok ? "all passed" : "FAILED");
if (process.argv[2]) { const d = blankDocument(); d.sections[0].blocks = blocks; writeFileSync(process.argv[2], writeDocx(d)); }
