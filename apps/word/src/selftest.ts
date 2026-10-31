/**
 * Development self-tests that need a real browser (run with ?selftest=paste in the dev build).
 * Results are logged as "[selftest] ..." lines for the test scripts to read.
 */
import { fromProseMirror } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import { TextSelection } from "prosemirror-state";
import * as pmView from "prosemirror-view";
import type { EditorView } from "prosemirror-view";
import type { Slice } from "prosemirror-model";

// ProseMirror's clipboard serializer (exported, though not in this version's type definitions).
const serializeForClipboard = (pmView as unknown as {
  __serializeForClipboard(view: EditorView, slice: Slice): { dom: HTMLElement };
}).__serializeForClipboard;

const SAMPLE = `
<h1>Pasted heading</h1>
<p style="text-align:center"><b>bold <i>bold-italic</i></b> and <span style="color:#c00000;font-size:16pt;font-family:'Georgia',serif">red Georgia 16</span></p>
<ul><li>bullet one</li><li>bullet two<ul><li>nested bullet</li></ul></li></ul>
<ol><li>number one</li><li>number two</li></ol>
<p class=MsoListParagraph style='margin-left:36pt;text-indent:-18pt;mso-list:l0 level1 lfo1'><span style='mso-list:Ignore'>·<span>&nbsp;&nbsp;</span></span>word list item</p>
<table border="1"><tr><td>A1</td><td style="background-color:#ffff00">B1</td></tr><tr><td colspan="2">wide</td></tr></table>
<p>see <a href="https://example.com/pasted">the link</a> and <u>underlined</u> <s>struck</s> H<sub>2</sub>O</p>
<p><img alt="dot" width="20" height="10" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="></p>
`;

function summary(doc: PrimeDocument) {
  return doc.sections.flatMap((s) => s.blocks).map((b) => {
    if (b.type === "table") return { table: b.rows.map((r) => r.cells.map((c) => ({ span: c.props.span, fill: c.props.shading?.fill }))), grid: b.grid.length, borders: !!b.props.extra?.length };
    if (b.type !== "paragraph") return { block: b.type };
    return {
      style: b.props.style, align: b.props.align, list: b.props.list,
      runs: b.content.map((i) => (i.type === "text" ? { t: i.text, ...i.props }
        : i.type === "link" ? { link: i.href, t: i.content.map((c) => (c.type === "text" ? c.text : c.type)).join("") }
        : i.type === "image" ? { image: i.imageId, w: Math.round(i.width) } : { type: i.type })),
    };
  });
}

export function pasteSelfTest(view: EditorView, base: () => PrimeDocument) {
  // 1. External HTML (web / Word) into an empty document.
  view.dispatch(view.state.tr.delete(0, view.state.doc.content.size));
  view.pasteHTML(SAMPLE);
  const pasted = fromProseMirror(view.state.doc, base());
  console.log("[selftest] external", JSON.stringify(summary(pasted)));
  console.log("[selftest] lists", base().lists.definitions.length, "images", Object.keys(base().images).length);
}

/** Copies the whole document through the real clipboard serializer and pastes it back: must be identical. */
export function copySelfTest(view: EditorView, base: () => PrimeDocument) {
  const before = JSON.stringify(fromProseMirror(view.state.doc, base()).sections.map((s) => s.blocks));
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 0, view.state.doc.content.size)));
  const { dom } = serializeForClipboard(view, view.state.selection.content());
  view.dispatch(view.state.tr.delete(0, view.state.doc.content.size));
  view.pasteHTML(dom.innerHTML);
  const after = JSON.stringify(fromProseMirror(view.state.doc, base()).sections.map((s) => s.blocks));
  console.log("[selftest] internal copy identical:", before === after, "blocks:", JSON.parse(before).flat().length, "->", JSON.parse(after).flat().length);
  if (before !== after) {
    const a = JSON.parse(before).flat(), b = JSON.parse(after).flat();
    const i = a.findIndex((x: unknown, k: number) => JSON.stringify(x) !== JSON.stringify(b[k]));
    console.log("[selftest] first difference at block", i, JSON.stringify(a[i]).slice(0, 300), "VS", JSON.stringify(b[i]).slice(0, 300));
  }
}

/** Types misspelled words, waits for the checker, then applies a suggestion the way the menu does. */
export function spellSelfTest(view: EditorView) {
  view.dispatch(view.state.tr.insertText(" The colour of teh centre is beautifull. ", 1));
  const marked = () => [...view.dom.querySelectorAll(".pw-misspelled")].map((e) => e.textContent);
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  void (async () => {
    for (let i = 0; i < 40 && !marked().includes("teh"); i++) await wait(250);
    console.log("[selftest] underlined:", JSON.stringify(marked()));
    const spans = [...view.dom.querySelectorAll(".pw-misspelled")].filter((e) => e.textContent === "teh");
    const from = spans.length ? view.posAtDOM(spans[0], 0) : -1;
    if (from >= 0) view.dispatch(view.state.tr.insertText("the", from, from + 3));
    await wait(1200);
    console.log("[selftest] after fix:", JSON.stringify(marked()), "spellcheck attr:", view.dom.getAttribute("spellcheck"));
  })();
}
