/**
 * Development self-tests that need a real browser (run with ?selftest=paste in the dev build).
 * Results are logged as "[selftest] ..." lines for the test scripts to read.
 */
import { fromProseMirror, timed } from "@prime/editor";
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

/**
 * Typing performance: types `count` characters one at a time (as fast as the editor accepts them), then
 * waits for the work that follows a pause (pagination, spelling). Reports keystroke times and every
 * main-thread task over 50 ms (a "long task" — the moments typing would feel stuck).
 */
export function perfSelfTest(view: EditorView, count = 40) {
  const long: number[] = [];
  const observer = new PerformanceObserver((list) => { for (const e of list.getEntries()) long.push(Math.round(e.duration)); });
  try { observer.observe({ type: "longtask", buffered: false }); } catch { /* not supported */ }
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  void (async () => {
    await wait(3000); // let the initial layout and spelling settle
    long.length = 0;
    // Somewhere in the middle of the document, like an author editing an existing chapter.
    let target = Math.floor(view.state.doc.content.size / 2);
    view.state.doc.nodesBetween(target, view.state.doc.content.size, (n, pos) => {
      if (n.isTextblock && n.content.size > 0 && target === Math.floor(view.state.doc.content.size / 2)) { target = pos + 1; return false; }
      return true;
    });
    const keys: number[] = [];
    const typingStart = performance.now();
    for (let i = 0; i < count; i++) {
      const t0 = performance.now();
      timed("dispatch", () => view.dispatch(view.state.tr.insertText(i % 6 === 5 ? " " : "x", target + i)));
      await wait(0); // the browser paints and React re-renders here
      keys.push(performance.now() - t0);
      await wait(60); // a fast typist: ~16 characters a second
    }
    const typingLong = [...long];
    long.length = 0;
    const t1 = performance.now();
    await wait(4000);
    const sorted = [...keys].sort((a, b) => a - b);
    // Where the time went: total and count per timed step ("pw:" performance entries).
    const steps: Record<string, [number, number]> = {};
    for (const e of performance.getEntriesByType("measure")) {
      if (!e.name.startsWith("pw:") || e.startTime < typingStart) continue;
      const s = (steps[e.name.slice(3)] ??= [0, 0]);
      s[0] += Math.round(e.duration); s[1]++;
    }
    console.log("[selftest] steps (ms, count)", JSON.stringify(steps));
    console.log("[selftest] perf", JSON.stringify({
      blocks: view.state.doc.childCount,
      keyMedianMs: Math.round(sorted[Math.floor(sorted.length / 2)]),
      keyMaxMs: Math.round(sorted.at(-1)!),
      longTasksWhileTyping: typingLong,
      longTasksAfterPause: long,
      settleWindowMs: Math.round(performance.now() - t1),
    }));
    console.log("[selftest] align", JSON.stringify(alignmentReport()));
    observer.disconnect();
  })();
}

/**
 * Page alignment on screen: after each page-gap spacer, content must start at the top of its sheet's
 * body area (as the first page's content does). Reports the largest error in CSS pixels.
 */
export function alignmentReport() {
  const sheets = Array.from(document.querySelectorAll(".pw-sheet"));
  const flow = document.querySelector(".pw-flow");
  if (!sheets.length || !flow) return { checked: 0 };
  const bodyOffset = flow.getBoundingClientRect().top - sheets[0].getBoundingClientRect().top;
  let checked = 0;
  let worst = 0;
  let off = 0;
  const bad: [number, number][] = [];
  for (const gap of Array.from(document.querySelectorAll<HTMLElement>(".pw-page-gap"))) {
    const sheet = sheets[Number(gap.dataset.page) - 1];
    if (!sheet) continue;
    const error = Math.abs(gap.getBoundingClientRect().bottom - (sheet.getBoundingClientRect().top + bodyOffset));
    checked++;
    worst = Math.max(worst, error);
    if (error > 1) { off++; bad.push([Number(gap.dataset.page), Math.round(gap.getBoundingClientRect().bottom - (sheet.getBoundingClientRect().top + bodyOffset))]); }
  }
  const corrections = performance.getEntriesByName("pw:correction").length;
  const rounds = performance.getEntriesByType("measure").filter((e) => e.name.startsWith("pw:round")).map((e) => Number(e.name.slice(9)));
  return { corrections, rounds, pages: sheets.length, checked, worstPx: Math.round(worst * 10) / 10, offByMoreThan1px: off, firstOff: bad.slice(0, 8), masked: Array.from(document.querySelectorAll('.pw-mask')).length, unmasked: sheets.length - 1 - document.querySelectorAll(".pw-mask").length };
}

export function alignSelfTest() {
  setTimeout(() => console.log("[selftest] align", JSON.stringify(alignmentReport())), 3000);
}
