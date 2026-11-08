/**
 * Development self-tests that need a real browser (run with ?selftest=paste in the dev build).
 * Results are logged as "[selftest] ..." lines for the test scripts to read.
 */
import { fromProseMirror, timed } from "@prime/editor";
import { layoutDocument, type Measurer } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { NodeSelection, TextSelection } from "prosemirror-state";
import * as pmView from "prosemirror-view";
import type { EditorView } from "prosemirror-view";
import type { Slice } from "prosemirror-model";
import { sheetTop } from "./PageSheets";

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
  view.dispatch(view.state.tr.insertText(" The colour of teh centre is beautifull, see https://github.com/Dept-of-Comp/seg26-edutrack or www.youtubee.com. ", 1));
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
export function perfSelfTest(first: EditorView, count = 40) {
  // The current editor (a long document may be rebuilt once its fonts arrive, after this starts).
  const views = { get current() { return (globalThis as { primeView?: EditorView }).primeView ?? first; } };
  const long: number[] = [];
  const observer = new PerformanceObserver((list) => { for (const e of list.getEntries()) long.push(Math.round(e.duration)); });
  try { observer.observe({ type: "longtask", buffered: false }); } catch { /* not supported */ }
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  void (async () => {
    // Let the initial layout and spelling settle (?settle=ms for very long documents).
    await wait(Number(new URLSearchParams(location.search).get("settle")) || 3000);
    const view = views.current;
    // What opening the document cost (layout, fitting, drawing…), before any typing.
    const opening: Record<string, [number, number]> = {};
    for (const e of performance.getEntriesByType("measure")) {
      if (!e.name.startsWith("pw:") || e.name.startsWith("pw:correction") || e.name.startsWith("pw:round")) continue;
      const o = (opening[e.name.slice(3)] ??= [0, 0]);
      o[0] += Math.round(e.duration); o[1]++;
    }
    console.log("[selftest] opening (ms, count)", JSON.stringify({ blocks: view.state.doc.childCount, sinceStart: Math.round(performance.now()), ...opening }));
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
    const fits = performance.getEntriesByName("pw:fit").filter((e) => e.startTime >= typingStart).map((e) => e.duration).sort((a, b) => a - b);
    console.log("[selftest] fit ms (median, max)", JSON.stringify([Math.round(fits[fits.length >> 1] ?? 0), Math.round(fits.at(-1) ?? 0)]));
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
  // Sheets are drawn only near the screen: their places are worked out (sheetTop) for every page.
  const box = document.querySelector<HTMLElement>(".pw-sheets");
  const flow = document.querySelector(".pw-flow");
  if (!box || !flow) return { checked: 0 };
  const pages = Number(box.dataset.count);
  const bodyOffset = flow.getBoundingClientRect().top - sheetTop(1)!;
  let checked = 0;
  let worst = 0;
  let off = 0;
  const bad: [number, number][] = [];
  for (const gap of Array.from(document.querySelectorAll<HTMLElement>(".pw-page-gap"))) {
    const page = Number(gap.dataset.page);
    if (!(page >= 1 && page <= pages)) continue;
    const top = sheetTop(page)!;
    const error = Math.abs(gap.getBoundingClientRect().bottom - (top + bodyOffset));
    checked++;
    worst = Math.max(worst, error);
    if (error > 1) { off++; bad.push([page, Math.round(gap.getBoundingClientRect().bottom - (top + bodyOffset))]); }
  }
  const corrected = performance.getEntriesByType("measure").filter((e) => e.name.startsWith("pw:correction")).map((e) => e.name.slice(14));
  const corrections = corrected.length;
  console.log("[selftest] corrected", corrected.length, JSON.stringify(corrected.slice(-8)));
  // How far each page's content ends from its body bottom (the top of the mask below it), in px:
  // positive = the screen ran longer than the layout there. (Masks are drawn near the screen only.)
  const ends = Array.from(document.querySelectorAll<HTMLElement>(".pw-page-gap")).map((g) => {
    const mask = document.querySelector(`.pw-mask[data-page="${g.dataset.page}"]`);
    return mask ? [Number(g.dataset.page) - 1, Math.round(g.getBoundingClientRect().top - mask.getBoundingClientRect().top)] : null;
  }).filter(Boolean);
  console.log("[selftest] page ends", JSON.stringify(ends));
  const rounds = performance.getEntriesByType("measure").filter((e) => e.name.startsWith("pw:round")).map((e) => Number(e.name.slice(9)));
  const masked = Number(box.dataset.masked);
  return { corrections, rounds, pages, checked, worstPx: Math.round(worst * 10) / 10, offByMoreThan1px: off, firstOff: bad.slice(0, 8), masked, unmasked: pages - 1 - masked };
}

/**
 * Page alignment as the reader sees it: scrolls to every `step`-th page (?step=, default 5), lets the
 * browser draw it (long documents skip drawing what is off screen) and refit, and checks the breaks
 * on screen. Reports the worst error and the pages off by more than 1 px.
 */
export function alignScanSelfTest() {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const params = new URLSearchParams(location.search);
  void (async () => {
    await wait(Number(params.get("settle")) || 3000);
    const box = document.querySelector<HTMLElement>(".pw-sheets");
    const flow = document.querySelector(".pw-flow");
    const scroller = document.querySelector(".canvas");
    if (!box || !flow || !scroller) return console.log("[selftest] align", JSON.stringify({ error: "no pages" }));
    const pages = Number(box.dataset.count);
    const step = Number(params.get("step")) || 5;
    const bodyOffset = flow.getBoundingClientRect().top - sheetTop(1)!;
    let checked = 0;
    let worst = 0;
    const off: [number, number][] = [];
    const t0 = performance.now();
    for (let n = 2; n <= pages; n += step) {
      scroller.scrollTop += sheetTop(n)! - scroller.getBoundingClientRect().top - 40;
      await wait(400); // drawn, and refitted (a block coming into view is fitted again after 150 ms)
      const view = scroller.getBoundingClientRect();
      for (const gap of Array.from(document.querySelectorAll<HTMLElement>(".pw-page-gap"))) {
        const r = gap.getBoundingClientRect();
        if (r.bottom < view.top || r.top > view.bottom) continue;
        const page = Number(gap.dataset.page);
        const error = r.bottom - (sheetTop(page)! + bodyOffset);
        checked++;
        worst = Math.max(worst, Math.abs(error));
        if (Math.abs(error) > 1) off.push([page, Math.round(error)]);
      }
    }
    console.log("[selftest] align", JSON.stringify({ pages, step, checked, worstPx: Math.round(worst * 10) / 10, offByMoreThan1px: off.length, firstOff: off.slice(0, 10), ms: Math.round(performance.now() - t0) }));
  })();
}

export function alignSelfTest() {
  setTimeout(() => console.log("[selftest] align", JSON.stringify(alignmentReport())), 3000);
}

/** Scrolls page `n` (from ?page=n) to the top of the window, for screenshots. */
export function scrollSelfTest() {
  const n = Number(new URLSearchParams(location.search).get("page") ?? 1);
  setTimeout(() => {
    // ?at=<text>: the first paragraph containing it instead of a page (also in Web Layout).
    const at = new URLSearchParams(location.search).get("at");
    const target = at ? [...document.querySelectorAll(".ProseMirror p")].find((p) => p.textContent?.includes(at)) : undefined;
    if (target) target.scrollIntoView({ block: "start" });
    else {
      const scroller = document.querySelector(".canvas");
      const y = sheetTop(n);
      if (scroller && y !== undefined) scroller.scrollTop += y - scroller.getBoundingClientRect().top;
    }
    // Debug aid: what sits around the page (spacers and wrap room), relative to the page's sheet.
    const at0 = sheetTop(n);
    const sheet = at0 === undefined ? undefined : { top: at0 };
    if (sheet) {
      const rel = (el: Element) => Math.round(el.getBoundingClientRect().top - sheet.top);
      const near = (el: Element) => Math.abs(rel(el)) < 2400;
      const items = [...document.querySelectorAll(".pw-page-gap, .pw-wrap")].filter(near).map((el) =>
        el.classList.contains("pw-wrap")
          ? `wrap@${rel(el)} in "${el.closest("p,h1,h2,h3,h4,div")?.textContent?.slice(0, 25)}" boxes=${[...el.children].map((c) => (c as HTMLElement).style.cssText).join(" | ")}`
          : `gap p${(el as HTMLElement).dataset.page}@${rel(el)} h=${Math.round(el.getBoundingClientRect().height)}`);
      console.log("[selftest] near", JSON.stringify(items));
      const paras = [...document.querySelectorAll(".ProseMirror > *")].filter(near).filter((el) => rel(el) > 0 && rel(el) < 1500)
        .map((el) => `${rel(el)}-${Math.round(el.getBoundingClientRect().bottom - sheet.top)}:${(el.textContent ?? "").slice(0, 14)}`);
      console.log("[selftest] paras", JSON.stringify(paras));
    }
    const toc = document.querySelector(".ProseMirror p[data-style^=TOC]");
    console.log("[selftest] align", JSON.stringify({ scrolled: n, righttab: document.querySelectorAll("p.pw-righttab").length,
      toc: toc?.outerHTML.slice(0, 300), display: toc ? getComputedStyle(toc).display : null }));
  }, 2500);
}

/**
 * Line counts per paragraph: the layout engine's (with the browser's measurer) against the lines the
 * browser actually drew. Lists the first paragraphs where they differ.
 */
export function linesSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  setTimeout(() => {
    const layout = layoutDocument(fromProseMirror(view.state.doc, base()), measurer);
    // Lines the browser drew per paragraph (the engine's are counted per paragraph laid out alone).
    const engine = new Map<number, number>();
    const diffs: string[] = [];
    view.state.doc.forEach((node, offset, index) => {
      if (node.type.name !== "paragraph" || !node.textContent.trim()) return;
      const dom = view.nodeDOM(offset) as HTMLElement | null;
      if (!dom) return;
      // Lines of text only (not list labels or other widgets): distinct baselines of text runs.
      const ys = new Set<number>();
      const walker = document.createTreeWalker(dom, NodeFilter.SHOW_TEXT);
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        if ((t.parentElement as HTMLElement | null)?.closest("[contenteditable=false]")) continue;
        const range = document.createRange();
        range.selectNodeContents(t);
        for (const r of range.getClientRects()) if (r.width > 0) ys.add(Math.round(r.bottom / 4));
      }
      engine.set(index, ys.size);
    });
    const one = (index: number) => {
      const node = view.state.doc.child(index);
      const doc = { ...base(), sections: [{ ...base().sections[0], blocks: [fromProseMirror(view.state.doc.type.schema.node("doc", null, [node]), base()).sections[0].blocks[0]] }] };
      const l = layoutDocument(doc, measurer);
      return l.pages.reduce((n, p) => n + p.lines.filter((x) => !x.story).length, 0);
    };
    for (const [index, domLines] of engine) {
      const ours = one(index);
      if (ours !== domLines && diffs.length < 12) diffs.push(`#${index} engine ${ours} screen ${domLines}: ${view.state.doc.child(index).textContent.slice(0, 50)}`);
    }
    console.log("[selftest] lines", JSON.stringify(diffs));
    // Table rows: heights on screen against the engine's (each table laid out alone), in pt.
    const rows: string[] = [];
    view.state.doc.forEach((node, offset, index) => {
      if (node.type.name !== "table") return;
      const dom = view.nodeDOM(offset) as HTMLElement | null;
      const screen = dom ? [...dom.querySelectorAll("tr")].map((tr) => Math.round(tr.getBoundingClientRect().height * 0.75 * 10) / 10) : [];
      const doc = { ...base(), sections: [{ ...base().sections[0], blocks: [fromProseMirror(view.state.doc.type.schema.node("doc", null, [node]), base()).sections[0].blocks[0]] }] };
      const l = layoutDocument(doc, measurer);
      const ys = [...new Set(l.pages.flatMap((p) => p.rects.filter((r) => r.sides).map((r) => `${p.number}:${Math.round(r.y * 10) / 10}:${Math.round(r.height * 10) / 10}`)))];
      const engine = ys.map((k) => Number(k.split(":")[2]));
      rows.push(`#${index} screen ${JSON.stringify(screen)} engine ${JSON.stringify(engine)}`);
    });
    console.log("[selftest] rows", JSON.stringify(rows));
    // Every paragraph's height (first line top to last line bottom), screen against engine, in pt.
    const heights: [number, string][] = [];
    view.state.doc.forEach((node, offset, index) => {
      if (node.type.name !== "paragraph") return;
      const dom = view.nodeDOM(offset) as HTMLElement | null;
      if (!dom) return;
      const cs = getComputedStyle(dom);
      const screen = (dom.getBoundingClientRect().height - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom)) * 0.75;
      const doc = { ...base(), sections: [{ ...base().sections[0], blocks: [fromProseMirror(view.state.doc.type.schema.node("doc", null, [node]), base()).sections[0].blocks[0]] }] };
      const lines = layoutDocument(doc, measurer).pages.flatMap((p) => p.lines.filter((x) => !x.story));
      const engine = lines.length ? Math.max(...lines.map((l) => l.y + l.height)) - Math.min(...lines.map((l) => l.y)) : 0;
      heights.push([Math.round((screen - engine) * 10) / 10, `#${index} screen ${screen.toFixed(1)} engine ${engine.toFixed(1)}: ${node.textContent.slice(0, 30)}`]);
    });
    heights.sort((a, b) => b[0] - a[0]);
    console.log("[selftest] heights", JSON.stringify(heights.slice(0, 8).map((h) => h[1])));
    // Paragraph tops on their page, screen against engine (pt from the page top), where they drift.
    const sheets = Array.from({ length: Number(document.querySelector<HTMLElement>(".pw-sheets")?.dataset.count ?? 0) }, (_, i) => sheetTop(i + 1)!);
    const drift: string[] = [];
    let last = 0;
    let n = 0;
    view.state.doc.forEach((node, offset, index) => {
      if (node.type.name !== "paragraph" || !node.textContent.trim()) return;
      // The engine lists table-cell paragraphs too: find this one by its text, in order.
      const key = node.textContent.slice(0, 20);
      let j = n;
      while (j < layout.paragraphs.length && layout.paragraphs[j].text.slice(0, 20) !== key) j++;
      if (j >= layout.paragraphs.length) return;
      const want = layout.paragraphs[j];
      n = j + 1;
      const dom = view.nodeDOM(offset) as HTMLElement | null;
      if (!dom) return;
      const top = dom.getBoundingClientRect().top;
      const page = sheets.filter((t) => t <= top + 1).length;
      const y = (top - sheets[page - 1]) * 0.75;
      const d = page === want.page ? Math.round((y - want.y) * 10) / 10 : 999;
      if (Math.abs(d - last) > 0.8 && drift.length < 14) drift.push(`#${index} p${want.page} ${want.y.toFixed(1)} -> p${page} ${y.toFixed(1)} (${d}) ${node.textContent.slice(0, 25)}`);
      last = d;
    });
    console.log("[selftest] drift", JSON.stringify(drift));
    console.log("[selftest] align", JSON.stringify({ done: true }));
  }, 3000);
}

/** Drags the first floating picture 60px right and 30px down, then a corner 40px out; logs its box. */
export function floatDragSelfTest() {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const box = (el: HTMLElement) => ["x", "y", "w", "h"].map((k) => Math.round(Number(el.dataset[k]))).join(",");
  const mouse = (type: string, target: EventTarget, x: number, y: number) =>
    target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: y }));
  void (async () => {
    await wait(2500);
    const first = () => document.querySelector<HTMLElement>('.pw-float[data-kind="picture"]');
    let el = first();
    if (!el) return console.log("[selftest] align", JSON.stringify({ error: "no floating picture" }));
    el.scrollIntoView();
    const before = box(el);
    const r = el.getBoundingClientRect();
    mouse("mousedown", el, r.left + 20, r.top + 20);
    mouse("mousemove", window, r.left + 80, r.top + 50);
    mouse("mouseup", window, r.left + 80, r.top + 50);
    await wait(1500);
    el = first()!;
    const moved = box(el);
    const r2 = el.getBoundingClientRect();
    mouse("mousedown", el.querySelector(".pw-se")!, r2.right, r2.bottom);
    mouse("mousemove", window, r2.right + 40, r2.bottom + 40);
    mouse("mouseup", window, r2.right + 40, r2.bottom + 40);
    await wait(1500);
    console.log("[selftest] align", JSON.stringify({ before, moved, resized: box(first()!) }));
  })();
}

/** Review UI: opens the Review tab and the comments pane, then adds a comment on the first words. */
export function reviewSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string) => [...document.querySelectorAll("button")].find((b) => b.textContent?.trim() === label)?.click();
  void (async () => {
    await wait(1500);
    click("Review");
    await wait(200);
    click("Comments");
    await wait(300);
    let at = 0;
    view.state.doc.descendants((n, pos) => { if (!at && n.isTextblock && n.textContent.length > 10) at = pos + 1; return !at; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at, at + 8)));
    click("New Comment");
    await wait(300);
    const box = document.querySelector<HTMLTextAreaElement>(".comments-pane textarea");
    if (box) {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")!.set!;
      setter.call(box, "Added by the review self-test.");
      box.dispatchEvent(new Event("input", { bubbles: true }));
      await wait(100);
      click("Post");
    }
    await wait(500);
    const items = [...document.querySelectorAll(".comments-pane .comment:not(.new)")].map((c) => c.textContent?.slice(0, 60));
    console.log("[selftest] align", JSON.stringify({ comments: items, highlighted: document.querySelectorAll(".pw-commented").length }));
  })();
}

/** Keyboard: Alt+R opens Review with focus on its first control; F6 cycles ribbon → document → status bar. */
export function keyboardSelfTest() {
  const key = (k: string, opts: KeyboardEventInit = {}) => window.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true, ...opts }));
  const where = () => {
    const a = document.activeElement as HTMLElement | null;
    return a?.closest(".toolbar") ? `ribbon:${a.textContent?.trim().slice(0, 14)}` : a?.closest(".status") ? "status" : a?.getAttribute("aria-label") ?? a?.tagName ?? "none";
  };
  setTimeout(async () => {
    const steps: string[] = [];
    key("r", { altKey: true });
    await new Promise((r) => setTimeout(r, 50));
    steps.push(`Alt+R tab=${document.querySelector('[role="tab"][aria-selected="true"]')?.textContent} focus=${where()}`);
    for (let i = 0; i < 3; i++) { key("F6"); steps.push(`F6 -> ${where()}`); }
    key("F6", { shiftKey: true });
    steps.push(`Shift+F6 -> ${where()}`);
    const unnamed = [...document.querySelectorAll("button")].filter((b) => !(b.getAttribute("aria-label") || b.textContent?.trim().length && b.textContent.trim().length > 1));
    steps.push(`buttons without a readable name: ${unnamed.map((b) => b.outerHTML.slice(0, 60)).join(" | ") || "none"}`);
    console.log("[selftest] align", JSON.stringify(steps));
  }, 1500);
}

/** Picture tab: select the first picture, choose Square wrap, drag it 80px right; log what happened. */
export function pictureSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  void (async () => {
    await wait(1500);
    let pos = -1;
    view.state.doc.descendants((n, p) => { if (pos < 0 && n.type.name === "image") pos = p; return pos < 0; });
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
    await wait(300);
    const tab = [...document.querySelectorAll(".tabs button")].find((b) => b.textContent === "Picture")?.className;
    const select = document.querySelector<HTMLSelectElement>('select[aria-label="Wrap text"]');
    if (select) {
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(select, "square");
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
    await wait(2500);
    const el = document.querySelector<HTMLElement>('.pw-float[data-kind="picture"]');
    const before = el && ["x", "y", "w", "h"].map((k) => Math.round(Number(el.dataset[k]))).join(",");
    if (el) {
      const r = el.getBoundingClientRect();
      const fire = (type: string, target: EventTarget, x: number, y: number) => target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, clientX: x, clientY: y }));
      fire("mousedown", el, r.left + 10, r.top + 10);
      fire("mousemove", window, r.left + 90, r.top + 10);
      fire("mouseup", window, r.left + 90, r.top + 10);
    }
    await wait(2500);
    const after = document.querySelector<HTMLElement>('.pw-float[data-kind="picture"]');
    console.log("[selftest] align", JSON.stringify({ tab, wrap: select?.value, before, after: after && ["x", "y", "w", "h"].map((k) => Math.round(Number(after.dataset[k]))).join(",") }));
  })();
}

/** Crop & Rotate: rotate the first picture a quarter turn; its width and height should swap. */
export function rotateSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string) => [...document.querySelectorAll("button")].find((b) => b.textContent?.trim().startsWith(label))?.click();
  void (async () => {
    await wait(1500);
    let pos = -1;
    view.state.doc.descendants((n, p) => { if (pos < 0 && n.type.name === "image") pos = p; return pos < 0; });
    const size = (p: number) => { const n = view.state.doc.nodeAt(p)!; return `${Math.round(n.attrs.width)}x${Math.round(n.attrs.height)} ${n.attrs.imageId}`; };
    const before = size(pos);
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
    await wait(300);
    click("Crop & Rotate");
    await wait(800);
    click("⟳ Rotate right");
    await wait(300);
    click("Apply");
    await wait(500);
    console.log("[selftest] align", JSON.stringify({ before, after: size(pos) }));
  })();
}

/** Drag handle: hover the 2nd paragraph, drag its handle and drop below the 5th; log the order. */
export function dragBlockSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const order = () => { const out: string[] = []; view.state.doc.forEach((n) => out.push(n.textContent.slice(0, 14))); return out.slice(0, 6); };
  void (async () => {
    await wait(2500);
    const before = order();
    const blocks = [...view.dom.children].filter((el) => el.tagName === "P" || el.tagName === "H1") as HTMLElement[];
    const source = blocks[1].getBoundingClientRect();
    view.dom.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: source.left + 20, clientY: source.top + 4 }));
    await wait(100);
    const handle = document.querySelector<HTMLElement>(".pw-block-handle");
    const shown = handle?.style.display;
    const dt = new DataTransfer();
    handle?.dispatchEvent(new DragEvent("dragstart", { bubbles: true, dataTransfer: dt }));
    const target = blocks[4].getBoundingClientRect();
    const at = { clientX: target.left + 4, clientY: target.bottom - 2 };
    view.dom.dispatchEvent(new DragEvent("dragover", { bubbles: true, cancelable: true, dataTransfer: dt, ...at }));
    view.dom.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt, ...at }));
    await wait(300);
    console.log("[selftest] align", JSON.stringify({ shown, before, after: order() }));
  })();
}

/** Letterhead: fill the designer and save it as a template (?save=0 leaves it open, for a screenshot). */
export function letterheadSelfTest() {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string) => [...document.querySelectorAll("button")].find((b) => b.textContent?.trim().startsWith(label))?.click();
  const type = (el: HTMLInputElement | HTMLTextAreaElement | null, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  void (async () => {
    await wait(1200);
    click("Insert");
    await wait(200);
    click("Letterhead");
    await wait(400);
    const form = document.querySelector(".letterhead-form")!;
    const [name, tagline] = [...form.querySelectorAll("input")];
    const [address, contacts] = [...form.querySelectorAll("textarea")];
    type(name, "Accra Community Library");
    type(tagline, "Knowledge for everyone");
    type(address, "12 Independence Avenue\nAccra, Ghana");
    type(contacts, "+233 30 222 0000\ninfo@acl.org.gh");
    await wait(800);
    if (new URLSearchParams(location.search).get("save") === "0") return console.log("[selftest] align", JSON.stringify({ filled: true }));
    click("Save as template");
    await wait(800);
    const { listEntries } = await import("./library");
    const list = await listEntries("templates");
    console.log("[selftest] align", JSON.stringify({ templates: list.map((t) => t.name), preview: !!document.querySelector(".page-thumb") }));
  })();
}

/**
 * Signature: a synthetic phone photo (lighting falling off to one side, a shadow, blue pen) is cleaned,
 * saved, and inserted from the Signature dialog as a floating picture in front of the text.
 */
export function signatureSelfTest(view: { state: { doc: { descendants(f: (n: { type: { name: string }; attrs: Record<string, unknown> }) => void): void } } }) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string) => [...document.querySelectorAll("button")].find((b) => b.textContent?.trim().startsWith(label))?.click();
  void (async () => {
    await wait(1200);
    const { cleanSignature, toCanvas, pngOf } = await import("./signature");
    const photo = document.createElement("canvas");
    photo.width = 1200;
    photo.height = 800;
    const ctx = photo.getContext("2d")!;
    const light = ctx.createLinearGradient(0, 0, 1200, 0);
    light.addColorStop(0, "#f2efe8");
    light.addColorStop(1, "#9c968c"); // paper in shadow at the right
    ctx.fillStyle = light;
    ctx.fillRect(0, 0, 1200, 800);
    ctx.fillStyle = "rgba(0,0,0,0.18)"; // a hand's shadow
    ctx.beginPath();
    ctx.ellipse(950, 650, 260, 160, 0, 0, Math.PI * 2);
    ctx.fill();
    ctx.strokeStyle = "#1b2f7a";
    ctx.lineWidth = 7;
    ctx.lineCap = "round";
    ctx.beginPath();
    ctx.moveTo(250, 450);
    for (let x = 250; x <= 1000; x += 10) ctx.lineTo(x, 420 + Math.sin(x / 35) * 60 - (x - 250) * 0.1);
    ctx.stroke();
    const { canvas, coverage } = cleanSignature(toCanvas(photo), { sensitivity: 0, colour: "original" });
    const d = canvas.getContext("2d")!.getImageData(0, 0, canvas.width, canvas.height).data;
    let opaque = 0, partial = 0;
    for (let i = 3; i < d.length; i += 4) { if (d[i] > 200) opaque++; else if (d[i] > 0) partial++; }
    const corner = d[3]; // the cropped corner is paper: transparent
    const png = pngOf(canvas);
    const { saveEntry, listEntries, deleteEntry } = await import("./library");
    for (const e of await listEntries("signatures")) if (e.name === "Selftest signature") await deleteEntry("signatures", e.id);
    await saveEntry("signatures", { id: "selftest-signature", name: "Selftest signature", created: Date.now(), data: { png: png.data, width: png.width, height: png.height } });
    click("Insert");
    await wait(200);
    click("Signature");
    await wait(600);
    [...document.querySelectorAll<HTMLButtonElement>(".sig-open")].find((b) => b.title === "Insert Selftest signature")?.click();
    await wait(400);
    const pictures: unknown[] = [];
    view.state.doc.descendants((n) => { if (n.type.name === "image") pictures.push({ wrap: (n.attrs.float as { wrap?: string } | null)?.wrap, width: n.attrs.width }); });
    await deleteEntry("signatures", "selftest-signature");
    console.log("[selftest] align", JSON.stringify({ size: [canvas.width, canvas.height], coverage: +coverage.toFixed(3), opaque, partial, corner, pictures }));
  })();
}

/**
 * Blocks: make a stamp and an address in the Blocks dialog, then insert the address centred at the
 * cursor and the stamp into the header, on the right (?save=0 stops at the stamp maker, for a screenshot).
 */
export function blocksSelfTest(view: { state: { doc: { attrs: Record<string, unknown>; forEach(f: (n: { type: { name: string }; attrs: Record<string, unknown>; textContent: string }) => void): void } } }) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll("button")].find((b) => b.textContent?.trim().startsWith(label))?.click();
  const type = (el: HTMLInputElement | HTMLTextAreaElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const dialog = () => document.querySelector(".hf-dialog.blocks")!;
  const open = async () => { click("Insert"); await wait(200); click("Blocks"); await wait(500); };
  void (async () => {
    await wait(1200);
    const { listEntries, deleteEntry } = await import("./library");
    for (const e of await listEntries("blocks")) if (e.name.startsWith("Selftest")) await deleteEntry("blocks", e.id);
    await open();
    click("Stamp", dialog());
    await wait(300);
    type(dialog().querySelector<HTMLInputElement>("input:not([type])"), "Accra Community Library");
    type(dialog().querySelector<HTMLInputElement>(".sig-name input"), "Selftest stamp");
    await wait(400);
    if (new URLSearchParams(location.search).get("save") === "0") return console.log("[selftest] align", JSON.stringify({ stamp: true }));
    click("Save block", dialog());
    await wait(600);
    click("Address", dialog());
    await wait(300);
    type(dialog().querySelector("textarea"), "Accra Community Library\n12 Independence Avenue\nAccra, Ghana");
    type(dialog().querySelector<HTMLInputElement>(".sig-name input"), "Selftest address");
    await wait(300);
    click("Save block", dialog());
    await wait(800);
    const cards = () => [...dialog().querySelectorAll<HTMLButtonElement>(".mine-open")];
    const listed = cards().map((b) => b.title);
    click("Centre", dialog());
    await wait(200);
    cards().find((b) => b.title === "Insert Selftest address")?.click();
    await wait(400);
    await open();
    choose(dialog().querySelector("select"), "header");
    click("Right", dialog());
    await wait(200);
    cards().find((b) => b.title === "Insert Selftest stamp")?.click();
    await wait(600);
    const centred: string[] = [];
    view.state.doc.forEach((n) => { if (centred.length < 4 && (n.attrs.props as { align?: string })?.align === "center") centred.push(n.textContent); });
    type Story = { blocks: { props: { align?: string }; content: { type: string; float?: unknown }[] }[] };
    const h = (view.state.doc.attrs.section as { headers?: { default?: Story } } | null)?.headers?.default?.blocks;
    const header = h?.map((b) => `${b.props.align}:${b.content.map((i) => (i.float ? "float" : i.type)).join(",")}`);
    for (const e of await listEntries("blocks")) if (e.name.startsWith("Selftest")) await deleteEntry("blocks", e.id);
    console.log("[selftest] align", JSON.stringify({ listed, centred, header }));
  })();
}

/**
 * AutoCorrect as you type, in the editor: typing goes through the editor's own text-input path, keys
 * are real key events. Checks corrections, Backspace straight after a change, Enter, and lists.
 */
export function autoCorrectSelfTest(view: import("prosemirror-view").EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const type = (s: string) => {
    for (const ch of s) {
      const { from, to } = view.state.selection;
      if (!view.someProp("handleTextInput", (f) => f(view, from, to, ch))) view.dispatch(view.state.tr.insertText(ch, from, to));
    }
  };
  const key = (k: string, mods: KeyboardEventInit = {}) => view.dom.dispatchEvent(new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...mods }));
  const para = () => view.state.selection.$from.parent;
  void (async () => {
    await wait(1200);
    view.focus();
    const { TextSelection } = await import("prosemirror-state");
    view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
    key("Enter");
    type("teh report is \"done\" -- (c) 2026 ");
    const typed = para().textContent;
    type("adn ");
    key("Backspace"); // straight after a correction: takes it back
    const backspaced = para().textContent;
    type("more teh");
    key("Enter"); // Enter ends the word: corrected
    const entered = view.state.doc.resolve(view.state.selection.from - 2).parent.textContent;
    type("1. first");
    const list = !!para().attrs.props.list;
    const listText = para().textContent;
    key("z", { ctrlKey: true });
    await wait(50);
    const undone = { list: !!para().attrs.props.list, text: para().textContent };
    key("z", { ctrlKey: true }); // the second undo takes back the list itself
    await wait(50);
    const undone2 = { list: !!para().attrs.props.list, text: para().textContent };
    console.log("[selftest] align", JSON.stringify({ typed, backspaced, entered, list, listText, undone, undone2 }));
  })();
}

/**
 * Wave 1, through the UI: grow font, change case, clear formatting, format painter, right-click menu →
 * Font dialog, Symbol, Date & Time (as a field), Show ¶, the ruler (adds a tab stop), tabs at their stop,
 * Page X of Y, Paste Options (keep text only), the Paragraph dialog's line and page breaks.
 */
export function wave1SelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => (b.textContent?.trim() ?? "").startsWith(label) || b.getAttribute("aria-label") === label)?.click();
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const type = (el: HTMLInputElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    const { toggleRun } = await import("@prime/editor");
    await wait(1500);
    // The first top-level paragraph after the title with some text: select its first word.
    let at = -1;
    view.state.doc.forEach((n, pos) => { if (at < 0 && n.type.name === "paragraph" && n.textContent.length > 20 && pos > 0) at = pos; });
    const para = () => view.state.doc.nodeAt(at)!;
    const runOf = (n: import("prosemirror-model").Node | null | undefined) => n?.marks.find((m) => m.type.name === "run")?.attrs.props ?? {};
    const word = para().textContent.indexOf(" ");
    const selectWord = () => view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + 1, at + 1 + word)));
    selectWord();
    view.focus();
    click("Grow Font"); await wait(100);
    out.grow = runOf(view.state.doc.nodeAt(at + 1)).size;
    choose(document.querySelector<HTMLSelectElement>(".case-select"), "upper"); await wait(100);
    out.upper = view.state.doc.textBetween(at + 1, at + 1 + word);
    click("Clear All Formatting"); await wait(100);
    out.cleared = JSON.stringify(runOf(view.state.doc.nodeAt(at + 1)));

    // Format painter: bold the first word, copy it, brush it onto the next words.
    selectWord();
    toggleRun("bold")(view.state, view.dispatch);
    click("🖌 Format Painter"); await wait(100);
    const p2 = para().textContent.indexOf(" ", word + 1);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + 2 + word, at + 1 + p2)));
    view.dom.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    await wait(100);
    out.painted = !!runOf(view.state.doc.nodeAt(at + 2 + word)).bold;
    out.painterOff = !document.querySelector(".canvas.painting");

    // Right-click → Font… → 16 pt.
    const target = view.nodeDOM(at) as HTMLElement;
    const r = target.getBoundingClientRect();
    selectWord();
    target.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: r.left + 5, clientY: r.top + 5 }));
    await wait(200);
    out.menu = [...document.querySelectorAll(".context-menu button span")].map((s) => s.textContent).slice(0, 9);
    click("Font…", document.querySelector(".context-menu")!); await wait(200);
    type(document.querySelector<HTMLInputElement>(".font-dialog input[type=number]"), "16");
    await wait(50);
    click("OK", document.querySelector(".font-dialog")!); await wait(150);
    out.fontDialog = runOf(view.state.doc.nodeAt(at + 1)).size;

    // Symbol: ₵ at the end of the paragraph.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + para().nodeSize - 1)));
    click("Insert"); await wait(150);
    click("Symbol"); await wait(200);
    click("Currency", document.querySelector(".symbol-dialog")!); await wait(100);
    [...document.querySelectorAll<HTMLButtonElement>(".symbol-dialog .symbol-grid button")].find((b) => b.textContent === "₵")?.click();
    click("Close", document.querySelector(".symbol-dialog")!); await wait(100);
    out.symbol = para().textContent.endsWith("₵");

    // Date & Time, as a field.
    click("Date & Time"); await wait(200);
    const auto = document.querySelector<HTMLInputElement>(".datetime-dialog input[type=checkbox]");
    if (auto && !auto.checked) auto.click();
    await wait(50);
    click("OK", document.querySelector(".datetime-dialog")!); await wait(150);
    let field = "";
    para().forEach((n) => { const f = n.marks.find((m) => m.type.name === "field"); if (f) field = `${f.attrs.instr.trim()} = ${n.text}`; });
    out.dateField = field;

    // Show ¶.
    click("Home"); await wait(100);
    click("Show formatting marks"); await wait(400);
    out.marks = { on: view.dom.classList.contains("pw-show-marks"), spaces: document.querySelectorAll(".pw-space").length };
    click("Show formatting marks"); await wait(100);

    // Ruler: turn it on, click at 5 cm: a left tab stop for this paragraph.
    click("View"); await wait(150);
    const rulerBox = [...document.querySelectorAll<HTMLInputElement>(".toolbar input[type=checkbox]")][0];
    if (rulerBox && !rulerBox.checked) rulerBox.click();
    await wait(300);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + 1)));
    await wait(100);
    const text = document.querySelector<HTMLElement>(".ruler-text");
    const column = view.dom.getBoundingClientRect();
    const section = view.state.doc.attrs.section?.props ?? {};
    const scale = column.width / ((section.pageWidth ?? 612) - (section.margins?.left ?? 72) - (section.margins?.right ?? 72));
    const cm = 72 / 2.54;
    text?.dispatchEvent(new MouseEvent("click", { bubbles: true, clientX: column.left + 5 * cm * scale, clientY: text.getBoundingClientRect().top + 5 }));
    await wait(150);
    out.rulerTab = (para().attrs.props.extra ?? []).find((e: { name: string }) => e.name === "w:tabs")?.xml;
    out.rulerMarkers = document.querySelectorAll(".ruler-marker").length;

    // A tab at the start of the paragraph now ends at 5 cm on screen.
    view.dispatch(view.state.tr.insert(at + 1, view.state.schema.nodes.tab.create()));
    await wait(200);
    const tabEl = (view.nodeDOM(at) as HTMLElement).querySelector<HTMLElement>(".pw-tab");
    out.tabEndsAtCm = tabEl ? +((tabEl.getBoundingClientRect().right - column.left) / scale / cm).toFixed(2) : null;

    // Page X of Y.
    out.status = [...document.querySelectorAll(".status span")].map((s) => s.textContent).find((t) => t?.startsWith("Page"));

    // Paste formatted text, then Keep Text Only from the Paste Options button.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + 2)));
    view.focus();
    view.pasteHTML("<p><b style='color:red'>Pasted bold</b> words</p>");
    await wait(200);
    out.pasteButton = !!document.querySelector(".paste-options");
    click("Paste Options"); await wait(100);
    click("Keep Text Only", document.querySelector(".paste-menu") ?? document); await wait(150);
    let pastedBold = false;
    view.state.doc.nodesBetween(at, at + 40, (n) => { if (n.isText && n.text?.includes("Pasted") && runOf(n).bold) pastedBold = true; });
    out.pasteTextOnly = !pastedBold && para().textContent.includes("Pasted bold words");

    // Paragraph dialog: Keep with next.
    click("Home"); await wait(100);
    click("Paragraph settings"); await wait(200);
    [...document.querySelectorAll<HTMLLabelElement>(".para-breaks label")].find((l) => l.textContent?.includes("Keep with next"))?.querySelector("input")?.click();
    await wait(50);
    click("OK", document.querySelector(".para-dialog")!); await wait(150);
    out.keepNext = !!para().attrs.props.keepNext;
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** File tab and My templates: the File tab's buttons, the Home tab on one row, then the rail's My templates page. */
export function fileTabSelfTest() {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent?.trim().startsWith(label))?.click();
  void (async () => {
    await wait(1500);
    const toolbar = () => document.querySelector<HTMLElement>(".toolbar")!;
    const homeHeight = toolbar().getBoundingClientRect().height;
    click("File", document.querySelector("nav")!); await wait(200);
    const file = [...toolbar().querySelectorAll("button")].map((b) => b.textContent?.trim());
    click("Home screen", toolbar()); await wait(400);
    click("My templates", document.querySelector(".home-rail")!); await wait(400);
    const heading = document.querySelector(".home-main h1")?.textContent;
    const shot = new URLSearchParams(location.search).get("shot");
    if (shot !== "templates") { click("Home", document.querySelector(".home-rail")!); await wait(300); }
    console.log("[selftest] align", JSON.stringify({ file, homeToolbarPx: Math.round(homeHeight), templatesPage: heading, back: document.querySelector(".home-main h1")?.textContent }));
  })();
}

/**
 * References through the UI: a table of contents at the top (page numbers filled in after layout),
 * a caption under the first picture, a cross-reference to a heading, F9, and Ctrl+Click on an entry.
 */
export function referencesSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => (b.textContent?.trim() ?? "").startsWith(label))?.click();
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    await wait(1500);
    const entries = () => {
      const list: string[] = [];
      view.state.doc.descendants((n) => { if (/^TOC\d/.test(n.attrs.props?.style ?? "")) list.push(n.textContent); return n.type.name !== "paragraph"; });
      return list;
    };
    // Contents at the very start.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
    click("References", document.querySelector("nav")!); await wait(150);
    choose(document.querySelector<HTMLSelectElement>(".toc-select"), "2");
    await wait(2500); // the layout runs, then the page numbers are filled in
    out.toc = entries().slice(0, 6);
    out.tocCount = entries().length;

    // A caption under the first picture.
    let pic = -1;
    view.state.doc.descendants((n, pos) => { if (pic < 0 && n.type.name === "image") pic = pos; return pic < 0; });
    if (pic >= 0) {
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pic)));
      click("Insert Caption"); await wait(200);
      const input = document.querySelector<HTMLInputElement>(".caption-dialog input");
      if (input) { Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(input, "Quarterly chart"); input.dispatchEvent(new Event("input", { bubbles: true })); }
      await wait(50);
      out.captionPreview = document.querySelector(".caption-preview")?.textContent;
      click("OK", document.querySelector(".caption-dialog")!); await wait(200);
      let caption = "";
      view.state.doc.descendants((n) => { if (n.attrs.props?.style === "Caption" && !caption) caption = n.textContent; return n.type.name !== "paragraph"; });
      out.caption = caption;
    }

    // A cross-reference to the second heading, at the end of the document.
    view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
    click("Cross-reference"); await wait(200);
    const options = [...document.querySelectorAll<HTMLButtonElement>(".crossref-list button")];
    options[1]?.click(); await wait(50);
    out.refTarget = options[1]?.textContent;
    click("Insert", document.querySelector(".crossref-dialog")!); await wait(200);
    out.refText = view.state.doc.lastChild?.textContent.slice(-40);

    // F9, then Ctrl+Click the third contents entry: the cursor goes to that heading.
    view.dom.dispatchEvent(new KeyboardEvent("keydown", { key: "F9", bubbles: true, cancelable: true }));
    await wait(300);
    let third = -1;
    let k = 0;
    view.state.doc.descendants((n, pos) => { if (/^TOC\d/.test(n.attrs.props?.style ?? "") && k++ === 2) third = pos; return third < 0 && n.type.name !== "paragraph"; });
    const el = third >= 0 ? (view.nodeDOM(third) as HTMLElement) : null;
    if (el) {
      const r = el.getBoundingClientRect();
      el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, ctrlKey: true, button: 0, clientX: r.left + 20, clientY: r.top + r.height / 2 }));
      await wait(200);
      out.jumpedTo = view.state.selection.$from.parent.textContent.slice(0, 40);
      out.entryClicked = el.textContent?.slice(0, 40);
    }
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/**
 * Footnotes through the UI: Insert Footnote, type its text in the Notes pane, see the number in the text
 * and the note at the foot of the page; a second one earlier in the text renumbers the first.
 */
export function notesSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => (b.textContent?.trim() ?? "").startsWith(label))?.click();
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    await wait(1500);
    const numbers = () => [...view.dom.querySelectorAll(".pw-noteref")].map((el) => getComputedStyle(el, "::after").content.replace(/"/g, ""));
    // At the end of the third paragraph.
    let at = -1;
    let k = 0;
    view.state.doc.forEach((n, pos) => { if (n.type.name === "paragraph" && n.textContent.length > 20 && k++ === 2) at = pos + n.nodeSize - 1; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
    view.focus();
    click("References", document.querySelector("nav")!); await wait(150);
    click("Insert Footnote"); await wait(400);
    out.paneOpen = !!document.querySelector(".notes-pane");
    out.focusedInNote = !!document.activeElement?.closest(".note-editor");
    document.execCommand("insertText", false, "Source: Ghana Statistical Service, 2024.");
    await wait(2500); // the layout runs
    out.numbersAfterFirst = numbers();
    out.footAreaText = [...document.querySelectorAll(".pw-chrome-note")].map((e) => e.textContent).join(" | ");
    out.rule = document.querySelectorAll(".pw-chrome-rule").length;
    // A second footnote earlier in the text: it becomes 1, the first becomes 2.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 3)));
    view.focus();
    click("Insert Footnote"); await wait(400);
    document.execCommand("insertText", false, "An earlier note.");
    await wait(2500);
    out.numbersAfterSecond = numbers();
    out.paneNumbers = [...document.querySelectorAll(".notes-pane .note-head strong")].map((e) => e.textContent);
    out.footArea = [...document.querySelectorAll(".pw-chrome-note")].map((e) => e.textContent);
    // Undo in the document takes the second note away again.
    view.focus();
    view.dom.dispatchEvent(new KeyboardEvent("keydown", { key: "z", ctrlKey: true, bubbles: true, cancelable: true }));
    await wait(200);
    out.afterUndo = { refs: view.dom.querySelectorAll(".pw-noteref").length, pane: [...document.querySelectorAll(".notes-pane .note-editor")].map((e) => e.textContent), foot: [...document.querySelectorAll(".pw-chrome-note")].map((e) => e.textContent) };
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Citations through the UI: a new source, cited with a page; the reference list; then IEEE style. */
export function citationsSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => (b.textContent?.trim() ?? "").startsWith(label))?.click();
  const setValue = (el: HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(Object.getPrototypeOf(el), "value")!.set!.call(el, value);
    el.dispatchEvent(new Event(el.tagName === "SELECT" ? "change" : "input", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    await wait(1500);
    let at = -1;
    view.state.doc.forEach((n, pos) => { if (at < 0 && n.type.name === "paragraph" && n.textContent.length > 40) at = pos + n.nodeSize - 1; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
    click("References", document.querySelector("nav")!); await wait(150);
    setValue(document.querySelector<HTMLSelectElement>(".cite-select"), "new"); await wait(300);
    const dialog = () => document.querySelector(".sources-dialog")!;
    setValue(dialog().querySelector("textarea"), "Mensah, Ama\nOwusu, Kwame");
    const inputs = () => [...dialog().querySelectorAll<HTMLInputElement>("input")];
    setValue(inputs()[0], "Climate and Farming in Ghana");
    setValue(inputs()[1], "2024");
    setValue(inputs()[2], "Sub-Saharan Publishers");
    await wait(100);
    click("Save source", dialog()); await wait(200);
    out.listed = [...dialog().querySelectorAll(".source strong")].map((e) => e.textContent);
    setValue(dialog().querySelector<HTMLInputElement>(".sig-name input"), "12");
    await wait(50);
    click("Cite", dialog()); await wait(200);
    const para = () => view.state.doc.resolve(at).parent.textContent;
    out.citedApa = para().slice(-32);
    // The reference list at the end, then IEEE.
    view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
    click("Bibliography"); await wait(200);
    const last = () => { const out: string[] = []; view.state.doc.forEach((n) => out.push(n.textContent)); return out.slice(-2); };
    out.bibliographyApa = last();
    setValue(document.querySelector<HTMLSelectElement>('select[aria-label="Citation style"]'), "IEEE"); await wait(300);
    out.citedIeee = para().slice(-12);
    out.bibliographyIeee = last();
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Equations through the UI: type one with a live preview, insert it, double-click it, change it. */
export function equationSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => (b.textContent?.trim() ?? "").startsWith(label))?.click();
  const type = (el: HTMLTextAreaElement | null, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
  };
  const equations = () => {
    const out: string[] = [];
    view.state.doc.descendants((n) => { if (n.type.name === "inline_preserved" && /^m:oMath/.test(n.attrs.preserved?.name ?? "")) out.push(n.attrs.preserved.xml); return true; });
    return out;
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    await wait(1500);
    const before = equations().length;
    let at = -1;
    view.state.doc.forEach((n, pos) => { if (at < 0 && n.type.name === "paragraph" && n.textContent.length > 40) at = pos + n.nodeSize - 1; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
    click("Insert", document.querySelector("nav")!); await wait(150);
    click("Equation"); await wait(300);
    type(document.querySelector(".equation-dialog textarea"), String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`);
    await wait(300);
    out.preview = !!document.querySelector(".eq-preview svg");
    type(document.querySelector(".equation-dialog textarea"), String.raw`x = \frac{1}{`);
    await wait(200);
    out.error = document.querySelector(".eq-error")?.textContent;
    type(document.querySelector(".equation-dialog textarea"), String.raw`x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}`);
    await wait(200);
    click("Insert", document.querySelector(".equation-dialog")!); await wait(400);
    out.inserted = equations().length - before;
    out.onScreen = view.dom.querySelectorAll(".pw-math").length;
    // Double-click it: the dialog shows its LaTeX; change and update.
    const math = [...view.dom.querySelectorAll<HTMLElement>(".pw-math")].pop();
    math?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true, cancelable: true }));
    await wait(300);
    out.reopened = document.querySelector<HTMLTextAreaElement>(".equation-dialog textarea")?.value;
    type(document.querySelector(".equation-dialog textarea"), String.raw`E = mc^2`);
    await wait(200);
    click("Update", document.querySelector(".equation-dialog")!); await wait(400);
    out.after = equations().length - before;
    out.updated = equations().some((x) => x.includes("<m:t>E=mc</m:t>") || (x.includes(">E=") && x.includes("m:sSup")));
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/**
 * Tables through the UI: insert, fill, merge the title row, a style with a total row, borders,
 * shading, a column dragged wider, sorting, and text to table (?shot=1 stops before sorting, for a screenshot).
 */
export function tablesSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const click = (label: string, scope: ParentNode = document) => [...scope.querySelectorAll<HTMLButtonElement>("button")].find((b) => (b.textContent?.trim() ?? "").startsWith(label))?.click();
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    const { insertTable } = await import("@prime/editor");
    await wait(1500);
    // An empty paragraph at the end, with a table in it.
    view.dispatch(view.state.tr.insert(view.state.doc.content.size, view.state.schema.nodes.paragraph.create()));
    view.dispatch(view.state.tr.setSelection(TextSelection.atEnd(view.state.doc)));
    insertTable(() => ({ sections: [{ props: {}, blocks: [] }] }) as never, 5, 3)(view.state, view.dispatch);
    let tablePos = -1;
    view.state.doc.forEach((n, pos) => { if (n.type.name === "table") tablePos = pos; });
    const table = () => view.state.doc.nodeAt(tablePos)!;
    const cellPos = (r: number, k: number) => { let p = tablePos + 1; for (let i = 0; i < r; i++) p += table().child(i).nodeSize; p += 1; for (let i = 0; i < k; i++) p += table().child(r).child(i).nodeSize; return p; };
    const data = [["Results", "", ""], ["Region", "Rainfall", "Yield"], ["Accra", "810", "1.9"], ["Ashanti", "1420", "2.6"], ["Total", "2230", "4.5"]];
    for (let r = data.length - 1; r >= 0; r--) for (let k = 2; k >= 0; k--) if (data[r][k]) view.dispatch(view.state.tr.insertText(data[r][k], cellPos(r, k) + 2));
    const select = (r1: number, k1: number, r2: number, k2: number) => view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, cellPos(r1, k1) + 2, cellPos(r2, k2) + 2)));
    view.focus();
    select(0, 0, 0, 2);
    await wait(100);
    out.tabs = [...document.querySelectorAll("nav button.contextual")].map((b) => b.textContent);
    click("Table Layout", document.querySelector("nav")!); await wait(100);
    click("Merge Cells"); await wait(150);
    out.titleCells = table().child(0).childCount;
    select(2, 1, 2, 1);
    click("Table Design", document.querySelector("nav")!); await wait(100);
    choose(document.querySelector<HTMLSelectElement>(".table-style-select"), "GridTable4-Accent1"); await wait(150);
    [...document.querySelectorAll<HTMLLabelElement>(".toolbar label.choice")].find((l) => l.textContent?.includes("Total Row"))?.querySelector("input")?.click();
    await wait(150);
    out.style = table().attrs.props.style;
    out.look = (table().attrs.props.extra ?? []).find((x: { name: string }) => x.name === "w:tblLook")?.xml.match(/w:lastRow="(\d)"/)?.[1];
    select(2, 1, 3, 2);
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Borders"]'), "outside"); await wait(150);
    click("Shading", document.querySelector('[aria-label="Shading"]')!); await wait(150);
    out.shaded = table().child(2).child(1).attrs.props.shading?.fill;
    out.borders = (table().child(2).child(1).attrs.props.extra ?? []).find((x: { name: string }) => x.name === "w:tcBorders")?.xml.match(/<w:(top|left)\b/g);
    // Drag the border between columns 1 and 2 to the right by 60px.
    await wait(400);
    const td = (view.nodeDOM(cellPos(2, 0)) as HTMLElement);
    const r = td.getBoundingClientRect();
    const before = (table().attrs.grid as number[]).map((w) => Math.round(w));
    td.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: r.right - 1, clientY: r.top + 5 }));
    td.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, button: 0, clientX: r.right - 1, clientY: r.top + 5 }));
    window.dispatchEvent(new MouseEvent("mousemove", { clientX: r.right + 59, clientY: r.top + 5 }));
    window.dispatchEvent(new MouseEvent("mouseup", { clientX: r.right + 59, clientY: r.top + 5 }));
    await wait(200);
    out.columns = { before, after: (table().attrs.grid as number[]).map((w) => Math.round(w)) };
    if (new URLSearchParams(location.search).get("shot")) return console.log("[selftest] align", JSON.stringify(out));
    // Sort by rainfall, largest first (title and heading rows stay).
    select(2, 1, 2, 1);
    click("Table Layout", document.querySelector("nav")!); await wait(100);
    select(0, 0, 1, 0); click("Repeat Header Rows"); await wait(100);
    select(2, 1, 2, 1);
    click("Sort"); await wait(200);
    choose(document.querySelectorAll<HTMLSelectElement>(".sort-dialog select")[0], "1");
    choose(document.querySelectorAll<HTMLSelectElement>(".sort-dialog select")[1], "number");
    choose(document.querySelectorAll<HTMLSelectElement>(".sort-dialog select")[2], "d");
    await wait(50);
    click("OK", document.querySelector(".sort-dialog")!); await wait(200);
    out.sorted = Array.from({ length: table().childCount }, (_, i) => table().child(i).firstChild!.textContent);
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/**
 * Diagnosis (Web Layout): for consecutive top-level paragraphs on the same page of the layout, the
 * screen's distance between them against the layout's. Lists the paragraphs where the two differ most.
 */
export function driftSelfTest(view: EditorView, base: () => import("@prime/pdm").PrimeDocument, measurer: import("@prime/layout").Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  void (async () => {
    const { fromProseMirror } = await import("@prime/editor");
    const { layoutDocument } = await import("@prime/layout");
    await wait(Number(new URLSearchParams(location.search).get("settle")) || 4000);
    const v = (globalThis as { primeView?: EditorView }).primeView ?? view;
    if (new URLSearchParams(location.search).get("resize")) { document.fonts.dispatchEvent(new Event("loadingdone")); await wait(500); }
    const result = layoutDocument(fromProseMirror(v.state.doc, base()), measurer);
    // Layout: each top-level block's page and top (from pageStarts and lines is hard; use paragraphs in order).
    const tops: { page: number; y: number }[] = result.paragraphs.map((p) => ({ page: p.page, y: p.y }));
    const box = v.dom.getBoundingClientRect();
    const zoom = box.width / parseFloat(getComputedStyle(v.dom).width);
    const blocks: { node: import("prosemirror-model").Node; el: HTMLElement }[] = [];
    v.state.doc.forEach((n, pos) => { const el = v.nodeDOM(pos); if (el instanceof HTMLElement && n.type.name === "paragraph") blocks.push({ node: n, el }); });
    // Paragraphs in tables are in result.paragraphs too: keep only top-level ones (match by order of top-level paragraphs).
    let k = 0;
    const layoutOf: { page: number; y: number }[] = [];
    v.state.doc.forEach((n) => {
      if (n.type.name === "paragraph") layoutOf.push(tops[k]);
      let count = 0;
      n.descendants((c) => { if (c.type.name === "paragraph") count++; return true; });
      k += n.type.name === "paragraph" ? 1 : count;
    });
    const diffs: { i: number; screen: number; layout: number; text: string; props: string; html: string }[] = [];
    for (let i = 0; i + 1 < blocks.length; i++) {
      const a = layoutOf[i], b = layoutOf[i + 1];
      if (!a || !b || a.page !== b.page) continue;
      // Only neighbours on screen too (no page gap between them).
      if (blocks[i].el.nextElementSibling !== blocks[i + 1].el) continue;
      const screen = ((blocks[i + 1].el.getBoundingClientRect().top - blocks[i].el.getBoundingClientRect().top) / zoom) * 0.75;
      const layout = b.y - a.y;
      if (Math.abs(screen - layout) > 0.5) {
        diffs.push({ i, screen: +screen.toFixed(1), layout: +layout.toFixed(1), text: blocks[i].node.textContent.slice(0, 60),
          props: JSON.stringify(blocks[i].node.attrs.props).slice(0, 160), html: blocks[i].el.outerHTML.slice(0, 300) });
      }
    }
    const total = diffs.reduce((s, d) => s + (d.screen - d.layout), 0);
    const buckets: Record<string, number> = {};
    for (const d of diffs) { const k = String(Math.round(d.screen - d.layout)); buckets[k] = (buckets[k] ?? 0) + 1; }
    const small = diffs.filter((d) => Math.abs(d.screen - d.layout) < 3).slice(0, 4).map(({ html: _h, ...d }) => d);
    console.log("[selftest] buckets", JSON.stringify({ buckets, small }));
    diffs.sort((x, y) => Math.abs(y.screen - y.layout) - Math.abs(x.screen - x.layout));
    console.log("[selftest] drift", JSON.stringify({ compared: blocks.length, differing: diffs.length, totalPt: +total.toFixed(1), worst: diffs.slice(0, 6).map(({ html: _h, ...d }) => d) }));
    // The worst paragraphs line by line: the words on each screen line, and the layout's lines.
    for (const d of diffs.slice(0, 3)) {
      const el = blocks[d.i].el;
      const lines: string[] = [];
      let lastTop = NaN;
      const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      for (let t = walker.nextNode() as Text | null; t; t = walker.nextNode() as Text | null) {
        for (const m of t.data.matchAll(/\S+/g)) {
          const r = document.createRange();
          r.setStart(t, m.index!);
          r.setEnd(t, m.index! + m[0].length);
          const top = Math.round(r.getBoundingClientRect().top);
          if (top !== lastTop) { lines.push(""); lastTop = top; }
          lines[lines.length - 1] += m[0] + " ";
        }
      }
      const a2 = layoutOf[d.i];
      const b2 = layoutOf[d.i + 1];
      const page = result.pages[a2.page - 1];
      const layoutLines = page.lines.filter((l) => !l.story && l.y >= a2.y - 0.5 && l.y < b2.y - 0.5).map((l) => l.segments.map((g) => g.text).join("").trim());
      console.log("[selftest] lines", JSON.stringify({ i: d.i, screen: lines.map((l) => l.trim()), layout: layoutLines }));
      // Where each word of the first line starts (pt from the column's left), on screen and in the layout.
      const col = v.dom.getBoundingClientRect().left;
      const xs: string[] = [];
      const w2 = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      let firstTop = NaN;
      let lineNo = 0;
      const want = Number(new URLSearchParams(location.search).get("line")) || 1;
      for (let t = w2.nextNode() as Text | null; t; t = w2.nextNode() as Text | null) {
        for (const m of t.data.matchAll(/\S+/g)) {
          const r = document.createRange();
          r.setStart(t, m.index!);
          r.setEnd(t, m.index! + m[0].length);
          const rr = r.getBoundingClientRect();
          if (Number.isNaN(firstTop) || Math.abs(rr.top - firstTop) > 2) { firstTop = rr.top; lineNo++; }
          if (lineNo !== want) continue;
          xs.push(`${m[0]}@${(((rr.left - col) / zoom) * 0.75).toFixed(1)}`);
        }
      }
      const first = page.lines.filter((l) => !l.story && l.y >= a2.y - 0.5)[want - 1];
      const lx = (first?.segments ?? []).filter((g) => g.text.trim()).map((g) => `${g.text}@${(g.x - page.bodyLeft).toFixed(1)}`);
      const tabs = [...el.querySelectorAll<HTMLElement>(".pw-tab")].map((t) => t.style.width || "(unsized)");
      console.log("[selftest] xs", JSON.stringify({ i: d.i, screen: xs.slice(0, 30), layout: lx.slice(0, 30), tabs }));
    }
    console.log("[selftest] align", JSON.stringify({ done: true }));
  })();
}

/**
 * Section and column breaks (Wave 4): inserted from Layout → Breaks; checks the sections, how each
 * starts, the pages they produce (an odd-page section starts on an odd page, a blank page between when
 * needed) and that a saved file keeps them.
 */
export function sectionsSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    const { writeDocx, readDocx } = await import("@prime/docx");
    await wait(1500);
    // The n-th top-level paragraph with text.
    const para = (n: number) => { let k = 0; let at = -1; view.state.doc.forEach((node, pos) => { if (at < 0 && node.type.name === "paragraph" && node.textContent.length > 20 && k++ === n) at = pos; }); return at; };
    const breaks = () => document.querySelector<HTMLSelectElement>('select[aria-label="Breaks"]');
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Layout")?.click();
    await wait(150);
    out.menu = breaks() ? [...breaks()!.options].map((o) => o.value).filter(Boolean) : "missing";
    const pagesBefore = layoutDocument(fromProseMirror(view.state.doc, base()), measurer).pages.length;
    // A continuous break in the middle of a paragraph: it is split there.
    const p1 = para(2);
    const text1 = view.state.doc.nodeAt(p1)!.textContent;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, p1 + 1 + 10)));
    choose(breaks(), "continuous"); await wait(150);
    out.splitAt = [view.state.doc.nodeAt(p1)!.textContent, text1.slice(0, 10)];
    // A continuous section starts no new page.
    out.continuousAddsNoPage = layoutDocument(fromProseMirror(view.state.doc, base()), measurer).pages.length === pagesBefore;
    // An odd-page section a few paragraphs later; a column break after it.
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, para(5) + 1)));
    choose(breaks(), "oddPage"); await wait(150);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, para(8) + 1 + 5)));
    choose(breaks(), "column"); await wait(150);
    const model = fromProseMirror(view.state.doc, base());
    out.sections = model.sections.map((s) => s.props.breakType ?? "nextPage");
    out.headersInherited = model.sections.slice(1).every((s) => !Object.keys(s.headers).length);
    const layout = layoutDocument(model, measurer);
    out.pages = [pagesBefore, layout.pages.length];
    const firstOf = (si: number) => layout.pages.find((p) => p.section === si)?.number;
    out.oddSectionStartsOn = firstOf(2);
    // Saved and read again.
    const again = readDocx(writeDocx(model)).doc;
    out.saved = again.sections.map((s) => s.props.breakType ?? "nextPage");
    out.columnBreakSaved = JSON.stringify(again.sections).includes('"kind":"column"');
    // The last section landscape (Layout → Landscape applies to the section at the cursor).
    const lastPara = para(10);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, lastPara + 1)));
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Landscape")?.click();
    await wait(3000); // laid out and fitted
    const laidOut = layoutDocument(fromProseMirror(view.state.doc, base()), measurer);
    out.pageSizes = [...new Set(laidOut.pages.map((p) => `s${p.section}:${Math.round(p.width)}x${Math.round(p.height)}`))];
    // On screen: the landscape section's text starts at its own page's left margin.
    const el = view.nodeDOM(lastPara) as HTMLElement;
    el.scrollIntoView({ block: "center" });
    await wait(600);
    const page = laidOut.pages.find((p) => p.section === laidOut.pages.at(-1)!.section)!;
    const sheet = document.querySelector<HTMLElement>(`.pw-sheet[data-page="${page.number}"]`);
    const scale = sheet ? sheet.getBoundingClientRect().width / page.width : 0;
    out.landscapeTextLeftPx = sheet ? Math.round((el.getBoundingClientRect().left - sheet.getBoundingClientRect().left - page.bodyLeft * scale) * 10) / 10 : "no sheet";
    out.landscapeTextWidthPx = sheet ? Math.round((el.getBoundingClientRect().width - (page.bodyRight - page.bodyLeft) * scale) * 10) / 10 : "no sheet";
    out.alignment = alignmentReport();
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/**
 * Design tab (Wave 4): watermark, page colour and page borders from the ribbon; checks the document,
 * the laid-out pages, the screen and a saved file.
 */
export function designSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const choose = (el: HTMLSelectElement | HTMLInputElement | null | undefined, value: string, event = "change") => {
    if (!el) return;
    const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event(event, { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { writeDocx, readDocx } = await import("@prime/docx");
    const { watermarkShape } = await import("@prime/layout");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Design")?.click();
    await wait(150);
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Watermark"]'), "CONFIDENTIAL"); await wait(150);
    // A colour input reports "input" as the user drags and "change" at the end; React listens to both.
    choose(document.querySelector<HTMLInputElement>('input[aria-label="Page Color"]'), "#fff8e1", "input"); await wait(150);
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Page Borders"]'), "double"); await wait(150);
    choose(document.querySelector<HTMLInputElement>('input[aria-label="Page border colour"]'), "#1f4e79", "input"); await wait(150);
    const model = fromProseMirror(view.state.doc, base());
    out.background = model.background?.xml;
    out.borders = model.sections.map((s) => s.props.extra?.find((x) => x.name === "w:pgBorders")?.xml.slice(0, 120));
    out.watermarkInHeaders = model.sections.map((s) => Object.keys(s.headers).filter((k) => JSON.stringify(s.headers[k as "default"]).includes("PowerPlusWaterMarkObject")));
    const layout = layoutDocument(model, measurer);
    const decor = layout.pages[0].decor;
    out.decor = { background: decor?.background, borders: decor?.borders?.length, borderStyle: decor?.borders?.[0]?.style, watermark: decor?.watermark && { text: decor.watermark.text, size: Math.round(decor.watermark.size), rotation: decor.watermark.rotation } };
    out.everyPage = layout.pages.every((p) => p.decor?.watermark && p.decor.borders);
    await wait(2500); // laid out on screen
    out.screen = {
      watermark: document.querySelector(".pw-watermark")?.textContent,
      sheetColour: getComputedStyle(document.querySelector(".pw-sheet")!).backgroundColor,
      borderLines: document.querySelector(".pw-page-borders")?.querySelectorAll("line").length,
    };
    const again = readDocx(writeDocx(model)).doc;
    const runs = (again.sections[0].headers.default?.blocks ?? []).flatMap((b) => (b.type === "paragraph" ? b.content : []));
    const shape = runs.map((i) => ("preserved" in i ? watermarkShape(i.preserved.xml) : undefined)).find(Boolean);
    out.saved = { background: again.background?.xml, borders: !!again.sections[0].props.extra?.some((x) => x.name === "w:pgBorders"), watermark: shape?.text };
    // Removing them again.
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Watermark"]'), "remove"); await wait(150);
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "No Color")?.click(); await wait(150);
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Page Borders"]'), "none"); await wait(150);
    const cleared = fromProseMirror(view.state.doc, base());
    out.removed = !cleared.background && !JSON.stringify(cleared.sections).includes("PowerPlusWaterMarkObject") && !JSON.stringify(cleared.sections).includes("w:pgBorders");
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Line numbers (Wave 4): Layout → Line Numbers; checks the section, the layout, the screen and suppression. */
export function lineNumbersSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Layout")?.click();
    await wait(150);
    const menu = () => document.querySelector<HTMLSelectElement>('select[aria-label="Line Numbers"]');
    choose(menu(), "continuous"); await wait(150);
    const numbers = () => layoutDocument(fromProseMirror(view.state.doc, base()), measurer).pages.flatMap((p) => p.lines.filter((l) => l.story === "lineNumber").map((l) => Number(l.segments[0].text)));
    out.xml = fromProseMirror(view.state.doc, base()).sections[0].props.extra?.find((x) => x.name === "w:lnNumType")?.xml;
    out.menuShows = menu()?.value;
    const all = numbers();
    out.count = all.length;
    out.continuous = all.every((n, i) => n === i + 1);
    await wait(2500);
    out.onScreen = [...document.querySelectorAll(".pw-chrome-text")].filter((e) => /^\d+$/.test(e.textContent ?? "")).length > 5;
    // Suppress the first long paragraph: its lines lose their numbers.
    let at = -1;
    view.state.doc.forEach((n, pos) => { if (at < 0 && n.type.name === "paragraph" && n.textContent.length > 200) at = pos; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + 1)));
    choose(menu(), "suppress"); await wait(150);
    out.suppressedFewer = all.length - numbers().length;
    choose(menu(), "newPage"); await wait(150);
    const pages = layoutDocument(fromProseMirror(view.state.doc, base()), measurer).pages;
    out.restartEachPage = pages.filter((p) => p.lines.some((l) => l.story === "lineNumber")).every((p) => p.lines.find((l) => l.story === "lineNumber")!.segments[0].text === "1");
    choose(menu(), "none"); await wait(150);
    out.off = numbers().length === 0;
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Hyphenation (Wave 4): Layout → Hyphenation → Automatic; checks the setting, the layout, the screen and a saved file. */
export function hyphenationSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const choose = (el: HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { writeDocx, readDocx } = await import("@prime/docx");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Layout")?.click();
    await wait(150);
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Hyphenation"]'), "auto");
    await wait(3500); // laid out and fitted
    const model = fromProseMirror(view.state.doc, base());
    out.setting = model.hyphenation;
    out.hyphenated = layoutDocument(model, measurer).hyphens.length;
    out.softHyphens = document.querySelectorAll(".pw-shy").length;
    // Each soft hyphen ends a line on screen: the text after it starts lower.
    out.breaksOnScreen = [...document.querySelectorAll(".pw-shy")].filter((el) => {
      const range = document.createRange();
      const next = el.nextSibling ?? el.parentElement?.nextSibling;
      if (!next) return false;
      range.setStartAfter(el); range.setEnd(range.startContainer, range.startOffset);
      const before = el.previousSibling ? (() => { const r = document.createRange(); r.selectNodeContents(el.previousSibling!); const rs = r.getClientRects(); return rs[rs.length - 1]; })() : undefined;
      const after = (() => { const r = document.createRange(); r.selectNodeContents(next); return r.getClientRects()[0]; })();
      return before && after && after.top > before.top + 2;
    }).length;
    out.alignment = alignmentReport();
    out.saved = readDocx(writeDocx(model)).doc.hyphenation;
    choose(document.querySelector<HTMLSelectElement>('select[aria-label="Hyphenation"]'), "none"); await wait(2500);
    out.offAgain = { setting: fromProseMirror(view.state.doc, base()).hyphenation ?? null, softHyphens: document.querySelectorAll(".pw-shy").length };
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Styles (Wave 4): modify Heading 1 and create a style from a paragraph, from Home → Styles…. */
export function stylesSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const setValue = (el: HTMLInputElement | HTMLSelectElement | null | undefined, value: string, event = "change") => {
    if (!el) return;
    const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event(event === "change" && el instanceof HTMLInputElement ? "input" : event, { bubbles: true }));
  };
  const field = (label: string) => [...document.querySelectorAll<HTMLLabelElement>(".styles-dialog label")].find((l) => l.textContent?.trim().startsWith(label))?.querySelector<HTMLInputElement | HTMLSelectElement>("input, select");
  const click = (text: string) => [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent?.trim() === text)?.click();
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    const { undo } = await import("prosemirror-history");
    const { writeDocx, readDocx } = await import("@prime/docx");
    await wait(1500);
    const heading = () => { let at = -1; view.state.doc.forEach((n, pos) => { if (at < 0 && n.attrs.props?.style === "Heading1") at = pos; }); return at; };
    const h = heading();
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, h + 1)));
    const sizeOnScreen = () => parseFloat(getComputedStyle((view.nodeDOM(heading()) as HTMLElement).querySelector("span") ?? (view.nodeDOM(heading()) as HTMLElement)).fontSize);
    const before = sizeOnScreen();
    click("Styles…"); await wait(200);
    out.opensOn = (document.querySelector<HTMLSelectElement>('select[aria-label="Style to modify"]'))?.value;
    setValue(field("Size") as HTMLInputElement, "24"); await wait(50);
    setValue(field("Space before (pt)") as HTMLInputElement, "30"); await wait(50);
    click("Apply"); await wait(600);
    const style = () => fromProseMirror(view.state.doc, base()).styles.styles.find((s) => s.id === "Heading1")!;
    out.style = { size: style().run?.size, before: style().paragraph?.spacing?.before };
    out.screenSize = [before, sizeOnScreen()];
    const laid = layoutDocument(fromProseMirror(view.state.doc, base()), measurer);
    out.laidOutSize = laid.pages.flatMap((p) => p.lines).find((l) => l.segments.some((s) => s.text.includes("summary")))?.segments.find((s) => s.text.includes("summary"))?.size;
    undo(view.state, view.dispatch); await wait(300);
    out.undone = style().run?.size !== 24 && sizeOnScreen() === before;
    // A new style from a body paragraph.
    let body = -1;
    view.state.doc.forEach((n, pos) => { if (body < 0 && n.type.name === "paragraph" && n.textContent.length > 100) body = pos; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, body + 1)));
    click("Styles…"); await wait(200);
    setValue(document.querySelector<HTMLInputElement>('input[aria-label="New style name"]'), "Report Body"); await wait(50);
    click("Create Style"); await wait(300);
    const doc = fromProseMirror(view.state.doc, base());
    out.created = { paragraphStyle: view.state.doc.nodeAt(body)?.attrs.props?.style, inSheet: doc.styles.styles.find((s) => s.name === "Report Body")?.id };
    const again = readDocx(writeDocx(doc)).doc;
    out.saved = !!again.styles.styles.find((s) => s.name === "Report Body");
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Design → Style Set, Fonts and Colors (Wave 4): each restyles the body, title and headings in one undoable step. */
export function styleSetsSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const choose = (label: string, value: string) => {
    const el = document.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`);
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { undo } = await import("prosemirror-history");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Design")?.click();
    await wait(150);
    const sheet = () => fromProseMirror(view.state.doc, base()).styles.styles;
    const look = (id: string) => { const s = sheet().find((x) => x.id === id); return s && { font: s.run?.font?.ascii, size: s.run?.size, color: s.run?.color, bold: s.run?.bold }; };
    const pagesBefore = layoutDocument(fromProseMirror(view.state.doc, base()), measurer).pages.length;
    choose("Style Set", "Formal Report"); await wait(200);
    out.formal = { normal: look(sheet().find((s) => s.isDefault && s.type === "paragraph")!.id), h1: look("Heading1") };
    out.pages = [pagesBefore, layoutDocument(fromProseMirror(view.state.doc, base()), measurer).pages.length];
    choose("Theme Fonts", "Georgia"); await wait(200);
    out.fonts = { h1: look("Heading1")?.font, h2: look("Heading2")?.font };
    choose("Theme Colors", "Green"); await wait(200);
    out.colors = { h1: look("Heading1")?.color, h2: look("Heading2")?.color };
    undo(view.state, view.dispatch); await wait(200);
    out.undoColors = look("Heading1")?.color;
    await wait(3000);
    out.alignment = alignmentReport();
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Drop cap (Wave 4): Insert → Drop Cap on a paragraph; checks the frame, the layout, the screen, spelling and undoing it with "None". */
export function dropCapSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const choose = (value: string) => {
    const el = document.querySelector<HTMLSelectElement>('select[aria-label="Drop Cap"]');
    if (!el) return;
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event("change", { bubbles: true }));
  };
  const out: Record<string, unknown> = {};
  void (async () => {
    const { TextSelection } = await import("prosemirror-state");
    await wait(1500);
    let at = -1;
    view.state.doc.forEach((n, pos) => { if (at < 0 && n.type.name === "paragraph" && n.textContent.length > 200) at = pos; });
    const text = view.state.doc.nodeAt(at)!.textContent;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + 5)));
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Insert")?.click();
    await wait(150);
    choose("drop"); await wait(300);
    const cap = view.state.doc.nodeAt(at)!;
    out.cap = { letter: cap.textContent, frame: JSON.stringify(cap.attrs.props.extra).includes('w:dropCap=\\"drop\\"'), size: cap.firstChild?.marks.find((m) => m.type.name === "run")?.attrs.props.size };
    out.rest = view.state.doc.nodeAt(at + cap.nodeSize)!.textContent.slice(0, 12);
    const laid = layoutDocument(fromProseMirror(view.state.doc, base()), measurer);
    out.laidOut = laid.pages.flatMap((p) => p.lines).some((l) => l.segments.some((s) => s.text === cap.textContent && s.size > 30));
    await wait(3000); // laid out, fitted and spell-checked
    // On screen: the first three lines of the paragraph start right of the letter; the fourth does not.
    const capEl = view.nodeDOM(at) as HTMLElement;
    const para = capEl.nextElementSibling as HTMLElement;
    const r = document.createRange();
    const lefts: number[] = [];
    const tops: number[] = [];
    const walker = document.createTreeWalker(para, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      r.selectNodeContents(n);
      for (const b of r.getClientRects()) if (!tops.some((t) => Math.abs(t - b.top) < 3)) { tops.push(b.top); lefts.push(b.left); }
    }
    const paraLeft = para.getBoundingClientRect().left;
    out.narrowedLines = lefts.filter((l) => l > paraLeft + 10).length;
    out.floated = getComputedStyle(capEl).float;
    out.misspelledPrime = [...document.querySelectorAll(".pw-misspelled")].some((e) => /^rime/.test(e.textContent ?? ""));
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at + cap.nodeSize + 3)));
    choose("none"); await wait(300);
    out.restored = view.state.doc.nodeAt(at)!.textContent === text;
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Text columns (Wave 4): Layout → Columns; checks the section, the layout, the screen (text split between columns) and a saved file. */
export function columnsSelfTest(view: EditorView, base: () => PrimeDocument, measurer: Measurer) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const set = (el: HTMLInputElement | HTMLSelectElement | null | undefined, value: string) => {
    if (!el) return;
    const proto = el instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, "value")!.set!.call(el, value);
    el.dispatchEvent(new Event(el instanceof HTMLSelectElement ? "change" : "input", { bubbles: true }));
  };
  const menu = () => document.querySelector<HTMLSelectElement>('select[aria-label="Columns"]');
  const out: Record<string, unknown> = {};
  void (async () => {
    const { writeDocx, readDocx } = await import("@prime/docx");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Layout")?.click();
    await wait(150);
    set(menu(), "two"); await wait(200);
    const model = () => fromProseMirror(view.state.doc, base());
    out.xml = model().sections.map((s) => s.props.extra?.find((x) => x.name === "w:cols")?.xml);
    const laid = layoutDocument(model(), measurer);
    out.pages = laid.pages.length;
    out.columnStarts = laid.columnStarts.length;
    await wait(3500); // laid out, fitted
    const root = view.dom;
    out.screen = { columnsClass: root.classList.contains("pw-columns"), count: getComputedStyle(root).columnCount, breaks: root.querySelectorAll(".pw-col-break").length };
    // Some paragraph has lines in both columns: its text rectangles start at two different x positions.
    const lefts = new Set<number>();
    let split = false;
    for (const p of [...root.querySelectorAll(":scope > p")]) {
      const r = document.createRange();
      r.selectNodeContents(p);
      const xs = new Set([...r.getClientRects()].map((b) => Math.round(b.left / 20)));
      if (xs.size > 1) { split = true; break; }
      for (const x of xs) lefts.add(x);
    }
    out.paragraphSplitBetweenColumns = split;
    out.alignment = alignmentReport();
    // More Columns…: three, with a line between.
    set(menu(), "more"); await wait(150);
    out.dialogOpen = !!document.querySelector('input[aria-label="Number of columns"]');
    set(document.querySelector<HTMLInputElement>('input[aria-label="Number of columns"]'), "3"); await wait(50);
    [...document.querySelectorAll<HTMLLabelElement>(".popover label.choice")].find((l) => l.textContent?.includes("Line between"))?.querySelector("input")?.click();
    await wait(50);
    [...document.querySelectorAll<HTMLButtonElement>(".popover button")].find((b) => b.textContent === "Apply")?.click();
    await wait(3500);
    out.three = { xml: model().sections[0].props.extra?.find((x) => x.name === "w:cols")?.xml, count: getComputedStyle(root).columnCount, rule: getComputedStyle(root).columnRuleStyle };
    out.saved = readDocx(writeDocx(model())).doc.sections[0].props.extra?.find((x) => x.name === "w:cols")?.xml;
    set(menu(), "one"); await wait(3000);
    out.one = { pages: layoutDocument(model(), measurer).pages.length, columnsClass: root.classList.contains("pw-columns") };
    // Columns for selected paragraphs only: two body paragraphs in the middle of the text.
    const { TextSelection } = await import("prosemirror-state");
    const bodies: number[] = [];
    view.state.doc.forEach((n, pos) => { if (n.type.name === "paragraph" && n.textContent.length > 200) bodies.push(pos); });
    const a = bodies[2];
    const b = bodies[3] + view.state.doc.nodeAt(bodies[3])!.nodeSize - 2;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, a + 5, b)));
    set(menu(), "two"); await wait(3500);
    const m = model();
    out.selection = {
      sections: m.sections.map((x) => `${x.props.breakType ?? "nextPage"}:${/w:num="(\d+)"/.exec(x.props.extra?.find((e) => e.name === "w:cols")?.xml ?? "")?.[1] ?? 1}:${x.blocks.length}`),
      spanned: root.querySelectorAll(":scope > .pw-span").length,
      inColumns: [...root.children].filter((el) => el.matches("p, table") && !el.classList.contains("pw-span")).length,
      alignment: alignmentReport(),
    };
    // A selection in a section already in columns: the text around it goes back to one column.
    const counts = () => model().sections.map((x) => Number(/w:num="(\d+)"/.exec(x.props.extra?.find((e) => e.name === "w:cols")?.xml ?? "")?.[1] ?? 1));
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1)));
    set(menu(), "three"); await wait(3000);
    const allThree = counts();
    const later: number[] = [];
    view.state.doc.forEach((n, pos) => { if (n.type.name === "paragraph" && n.textContent.length > 200) later.push(pos); });
    const c = later[later.length - 4];
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, c + 5, c + view.state.doc.nodeAt(c)!.nodeSize - 2)));
    set(menu(), "two"); await wait(3500);
    out.inColumnsSection = { before: allThree, after: counts(), alignment: alignmentReport() };
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Track changes: typing, deleting, replacing and Enter are recorded; own insertions delete outright; resizing a table column records nothing; saved and read back; Accept All. */
export function trackChangesSelfTest(view: EditorView, base: () => PrimeDocument) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const out: Record<string, unknown> = {};
  const key = (k: string, mods: KeyboardEventInit = {}) => {
    const e = new KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...mods });
    view.someProp("handleKeyDown", (f) => f(view, e));
  };
  const type = (text: string) => {
    for (const ch of text) {
      const { from, to } = view.state.selection;
      if (!view.someProp("handleTextInput", (f) => f(view, from, to, ch))) view.dispatch(view.state.tr.insertText(ch, from, to));
    }
  };
  // Character deletion is the browser's own (ProseMirror reads it back as a deletion).
  const backspace = () => { const h = view.state.selection.head; view.dispatch(view.state.tr.delete(h - 1, h)); };
  const del = () => { const h = view.state.selection.head; view.dispatch(view.state.tr.delete(h, h + 1)); };
  void (async () => {
    const ed = await import("@prime/editor");
    const { writeDocx, readDocx } = await import("@prime/docx");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Review")?.click();
    await wait(150);
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Track Changes")?.click();
    await wait(100);
    out.on = ed.trackingOn(view.state);
    let para = -1;
    view.state.doc.forEach((n, pos) => { if (para < 0 && n.type.name === "paragraph" && n.textContent.length > 80 && !/[^\x20-\x7e’‘“”–—]/.test(n.textContent.slice(0, 40))) para = pos; });
    const text = () => view.state.doc.nodeAt(para)!.textContent;
    const original = text();
    const list = () => ed.changes(view.state.doc).map((c) => `${c.kind}:${c.paragraph !== undefined ? "¶" : view.state.doc.textBetween(c.from, c.to)}`);
    // Typing in the middle of a word boundary.
    const space = original.indexOf(" ", 20);
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, para + 1 + space)));
    type(" new");
    out.typed = list();
    out.author = ed.changes(view.state.doc)[0]?.author;
    // Backspace on your own insertion removes it.
    backspace();
    out.afterOwnBackspace = list();
    // Backspace into the original text: kept, marked deleted, the cursor moves before it.
    const at = para + 1 + space;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, at)));
    backspace(); backspace();
    out.deleted = list();
    out.cursorMoved = at - view.state.selection.head;
    out.textKept = text().startsWith(original.slice(0, space));
    // Delete (forward) marks the next character, the cursor after it.
    const fwd = para + 1 + original.indexOf(" ", space + 5) + 1;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, fwd)));
    del();
    out.forward = { changes: list().length, cursor: view.state.selection.head - fwd };
    // Typing over a selected word: the word is deleted, the new text inserted after it.
    const w = para + 1 + original.lastIndexOf(" ", 70) + 1;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, w, w + 4)));
    type("ZZ");
    out.replaced = list().slice(-2);
    // Enter in the middle: the new paragraph mark is an insertion.
    const mid = para + 1 + 10;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, mid)));
    key("Enter");
    out.enter = list().filter((c) => c.endsWith("¶"));
    // Backspace at the start of a paragraph marks the paragraph mark before it deleted.
    let other = -1;
    view.state.doc.forEach((n, pos) => { if (other < 0 && pos > para + 200 && n.type.name === "paragraph" && n.textContent.length > 20 && view.state.doc.resolve(pos).index() > 0 && view.state.doc.child(view.state.doc.resolve(pos).index() - 1).type.name === "paragraph") other = pos; });
    const blocks = view.state.doc.childCount;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, other + 1)));
    key("Backspace");
    out.joinKept = { blocks: view.state.doc.childCount === blocks, paragraphMarks: list().filter((c) => c === "del:¶").length };
    // A table column resized: nothing recorded.
    const before = list().length;
    ed.insertTable(base, 2, 2)(view.state, view.dispatch);
    const afterInsert = list().length;
    let tpos = -1;
    view.state.doc.forEach((n, pos) => { if (n.type.name === "table") tpos = pos; });
    const grid = [...view.state.doc.nodeAt(tpos)!.attrs.grid] as number[];
    if (grid.length > 1) { grid[0] += 20; grid[1] -= 20; }
    ed.setColumnWidths(tpos, grid)(view.state, view.dispatch);
    out.table = { insertRecorded: afterInsert > before, resizeRecorded: list().length - afterInsert, resized: (view.state.doc.nodeAt(tpos)!.attrs.grid as number[])[0] === grid[0] };
    // Formatting: a word made bold, a paragraph centred, a second word bolded then un-bolded (no change left).
    let fp = -1;
    view.state.doc.forEach((n, pos) => { if (fp < 0 && pos > other + 400 && n.type.name === "paragraph" && n.textContent.length > 80) fp = pos; });
    const fw = fp + 1 + view.state.doc.nodeAt(fp)!.textContent.indexOf(" ") + 1;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, fw, fw + 4)));
    ed.toggleRun("bold")(view.state, view.dispatch);
    ed.setParagraph("align", "center")(view.state, view.dispatch);
    const fw2 = fw + 20;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, fw2, fw2 + 3)));
    ed.toggleRun("italic")(view.state, view.dispatch);
    ed.toggleRun("italic")(view.state, view.dispatch);
    const formats = () => ed.changes(view.state.doc).filter((c) => c.kind === "format").map((c) => (c.paragraph !== undefined ? "¶" : view.state.doc.textBetween(c.from, c.to)));
    await wait(300);
    out.formatted = { changes: formats(), boldStays: ed.isActive(view.state.apply(view.state.tr.setSelection(TextSelection.create(view.state.doc, fw + 1))), "bold"), marked: view.dom.querySelectorAll(".pw-fmt").length, bar: view.dom.querySelectorAll(".pw-fmt-para").length };
    // Reject the paragraph's change: left aligned again; the bold stays (its own change).
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, fp + 3)));
    const beforeReject = view.state.doc.nodeAt(fp)!.attrs.props.align;
    ed.acceptOrReject(false)(view.state, view.dispatch);
    out.rejectParagraph = { before: beforeReject, after: view.state.doc.nodeAt(fp)!.attrs.props.align ?? "none", left: formats() };
    ed.setParagraph("align", "right")(view.state, view.dispatch);
    // Saved and read back (Word's own markup and setting).
    const docx = writeDocx(fromProseMirror(view.state.doc, base()));
    const back = readDocx(docx).doc;
    const files = await import("fflate").then((f) => f.unzipSync(docx)).catch(() => null);
    const dec = (n: string) => (files?.[n] ? new TextDecoder().decode(files[n]) : "");
    const body = dec("word/document.xml");
    const ids = [...body.matchAll(/<w:(?:ins|del) w:id="(\d+)"/g)].map((m) => m[1]);
    out.saved = { rPrChange: (body.match(/<w:rPrChange /g) ?? []).length, pPrChange: (body.match(/<w:pPrChange /g) ?? []).length, trackRevisions: !!back.trackRevisions, setting: /<w:trackRevisions\/>/.test(dec("word/settings.xml")), ins: (body.match(/<w:ins /g) ?? []).length, del: (body.match(/<w:del /g) ?? []).length, delText: (body.match(/<w:delText/g) ?? []).length, uniqueIds: new Set(ids).size === ids.length, paraMarkFirst: !/<w:rPr>(?:(?!<\/w:rPr>).)*?<w:(?:rFonts|sz|b)\b[^>]*\/>(?:(?!<\/w:rPr>).)*?<w:(?:ins|del) /.test(body.replace(/<w:r>.*?<\/w:r>/g, "")) };
    // Accept All: the edits stay, the marks go; not itself recorded.
    ed.acceptOrReject(true, true)(view.state, view.dispatch);
    out.accepted = { left: list().length, typedKept: view.state.doc.textContent.includes("ZZ") };
    // Off: typing is not recorded.
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Track Changes")?.click();
    type("q");
    out.off = { on: ed.trackingOn(view.state), recorded: list().length };
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Navigation pane: headings listed and followed; collapsing; search results; dragging a section to move it. */
export function navPaneSelfTest(view: EditorView, base: () => PrimeDocument) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const out: Record<string, unknown> = {};
  void (async () => {
    const ed = await import("@prime/editor");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "View")?.click();
    await wait(150);
    const box = [...document.querySelectorAll<HTMLLabelElement>("label.choice")].find((l) => l.textContent?.includes("Navigation Pane"))?.querySelector("input");
    if (box && !box.checked) box.click();
    await wait(600);
    const pane = document.querySelector(".nav-pane");
    const items = ed.outline(view.state, base());
    const rows = () => [...document.querySelectorAll<HTMLLIElement>(".nav-outline li")];
    out.open = !!pane;
    out.headings = { outline: items.length, rows: rows().length, first: rows()[0]?.textContent };
    // Click the 5th heading: the cursor goes there and the heading is near the top of the window.
    const target = Math.min(4, items.length - 1);
    rows()[target]?.querySelector<HTMLButtonElement>(".heading")?.click();
    await wait(400);
    const canvas = view.dom.closest(".canvas") as HTMLElement;
    const top = view.coordsAtPos(items[target].pos + 1).top - canvas.getBoundingClientRect().top;
    out.jump = { cursorIn: view.state.selection.from === items[target].pos + 1, topPx: Math.round(top), current: rows()[target]?.classList.contains("current") };
    // Collapse a heading with subheadings.
    const parent = rows().findIndex((r) => !r.querySelector<HTMLButtonElement>(".twisty")?.disabled);
    const before = rows().length;
    if (parent >= 0) rows()[parent].querySelector<HTMLButtonElement>(".twisty")!.click();
    await wait(100);
    out.collapse = { parent, before, after: rows().length };
    if (parent >= 0) rows()[parent].querySelector<HTMLButtonElement>(".twisty")!.click();
    // Search.
    const input = document.querySelector<HTMLInputElement>('input[aria-label="Search document"]')!;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(input, "trust");
    input.dispatchEvent(new Event("input", { bubbles: true }));
    await wait(400);
    const all = ed.findAll(view.state.doc, { text: "trust" });
    out.search = { matches: all.length, count: document.querySelector(".nav-count")?.textContent, listed: document.querySelectorAll(".nav-results li").length, highlighted: view.dom.querySelectorAll(".pw-find").length > 0 };
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
    await wait(200);
    out.enterSelects = view.state.doc.textBetween(view.state.selection.from, view.state.selection.to);
    [...document.querySelectorAll<HTMLButtonElement>(".nav-results li button")][2]?.click();
    await wait(300);
    out.resultClick = view.state.selection.from === all[2]?.from;
    // Drag the first heading of a level below the next of the same level.
    [...document.querySelectorAll<HTMLButtonElement>(".nav-tabs button")].find((b) => b.textContent === "Headings")?.click();
    await wait(100);
    const fresh = ed.outline(view.state, base());
    const a = fresh.findIndex((h, i) => fresh.slice(i + 1).some((x) => x.level === h.level));
    const b = fresh.findIndex((h, i) => i > a && h.level === fresh[a].level);
    const names = () => ed.outline(view.state, base()).filter((h) => h.level === fresh[a].level).map((h) => h.text.slice(0, 25));
    const order = names();
    const dt = new DataTransfer();
    const ra = rows()[a];
    const rb = rows()[b];
    ra.dispatchEvent(new DragEvent("dragstart", { bubbles: true, dataTransfer: dt }));
    const rect = rb.getBoundingClientRect();
    rb.dispatchEvent(new DragEvent("dragover", { bubbles: true, cancelable: true, dataTransfer: dt, clientY: rect.bottom - 2 }));
    await wait(50);
    rb.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: dt, clientY: rect.bottom - 2 }));
    await wait(300);
    const moved = names();
    out.drag = { a, b, before: order.slice(0, 3), after: moved.slice(0, 3), blocks: view.state.doc.childCount };
    const { undo } = await import("prosemirror-history");
    undo(view.state, view.dispatch);
    out.undone = names().join("|") === order.join("|");
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Print dialog (pages chosen, preview, the frame handed to the printer) and Read Mode (pages turned, one or two). */
export function printReadSelfTest() {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const until = async (f: () => unknown, ms = 60000) => { const t = performance.now(); while (!f() && performance.now() - t < ms) await wait(100); return f(); };
  const out: Record<string, unknown> = {};
  const inked = (c: HTMLCanvasElement | null) => {
    if (!c || !c.width) return 0;
    const d = c.getContext("2d")!.getImageData(0, 0, c.width, c.height).data;
    let dark = 0;
    for (let i = 0; i < d.length; i += 16) if (d[i] < 128) dark++;
    return dark;
  };
  void (async () => {
    const { parsePages } = await import("./PrintDialog");
    out.parse = [parsePages("1-3, 5", 10), parsePages("8-", 10), parsePages("2,2,1", 10), parsePages("12", 10), parsePages("x", 10)];
    await wait(2500);
    const t0 = performance.now();
    window.dispatchEvent(new Event("prime-word:print"));
    await until(() => document.querySelector(".print-dialog canvas[aria-label]"));
    out.printOpenMs = Math.round(performance.now() - t0);
    await wait(1200);
    out.previewInk = inked(document.querySelector(".print-dialog canvas"));
    const custom = document.querySelector<HTMLInputElement>('input[aria-label="Pages to print"]')!;
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")!.set!.call(custom, "2-3");
    custom.dispatchEvent(new Event("input", { bubbles: true }));
    await wait(300);
    out.printButton = document.querySelector<HTMLButtonElement>(".print-dialog button.primary")?.textContent;
    // The printer window cannot open in a test: the frame's print is caught instead.
    let printed: { sheets: number; sizes: string[] } | undefined;
    (globalThis as { primePrintHook?: (d: Document) => void }).primePrintHook = (d) => {
      printed = { sheets: d.querySelectorAll(".sheet img").length, sizes: [...d.querySelectorAll("style")].map((x) => x.textContent ?? "").join("").match(/size: [^;]+/g) ?? [] };
    };
    document.querySelector<HTMLButtonElement>(".print-dialog button.primary")!.click();
    await until(() => printed, 30000);
    out.printed = printed;
    out.dialogClosed = await until(() => !document.querySelector(".print-dialog"), 5000);
    // Read Mode.
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "View")?.click();
    await wait(150);
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Read Mode")?.click();
    await until(() => document.querySelectorAll(".read-mode canvas.read-page").length);
    await wait(1500);
    const pagesShown = () => document.querySelector(".read-pages")?.textContent;
    const canv = () => [...document.querySelectorAll<HTMLCanvasElement>(".read-mode canvas.read-page")];
    out.read = { shown: pagesShown(), canvases: canv().length, ink: canv().map(inked), sideBySide: canv().length === 2 && canv()[1].getBoundingClientRect().left > canv()[0].getBoundingClientRect().right };
    const mode = document.querySelector<HTMLElement>(".read-mode")!;
    mode.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    await wait(300);
    out.afterRight = pagesShown();
    [...document.querySelectorAll<HTMLButtonElement>(".read-mode button")].find((b) => b.textContent === "One Page")?.click();
    await wait(300);
    out.onePage = { shown: pagesShown(), canvases: canv().length };
    mode.dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true }));
    await wait(300);
    out.end = pagesShown();
    mode.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
    await wait(200);
    out.closed = !document.querySelector(".read-mode");
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Outline view: levels and symbols, Show Level, folding a heading, promote / demote, back to Print Layout. */
export function outlineSelfTest(view: () => EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const out: Record<string, unknown> = {};
  void (async () => {
    const ed = await import("@prime/editor");
    await wait(1500);
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "View")?.click();
    await wait(150);
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Outline")?.click();
    await wait(1500);
    const v = view();
    const root = v.dom;
    const visible = () => [...root.querySelectorAll<HTMLElement>(":scope > .pw-ol")].filter((e) => e.offsetParent !== null);
    out.outlineTab = document.querySelector("nav button.contextual.on")?.textContent;
    out.symbols = { plus: root.querySelectorAll(".pw-ol-plus").length, minus: root.querySelectorAll(".pw-ol-minus").length, body: root.querySelectorAll(".pw-ol-body").length };
    const h2 = root.querySelector<HTMLElement>(".pw-ol-plus ~ .pw-ol-plus");
    out.indents = [...root.querySelectorAll<HTMLElement>(":scope > .pw-ol")].slice(0, 12).map((e) => Math.round(e.getBoundingClientRect().left));
    out.all = visible().length;
    const show = document.querySelector<HTMLSelectElement>('select[aria-label="Show level"]')!;
    const pick = (sel: HTMLSelectElement, val: string) => { Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, "value")!.set!.call(sel, val); sel.dispatchEvent(new Event("change", { bubbles: true })); };
    pick(show, "1");
    await wait(300);
    out.level1 = { shown: visible().length, texts: visible().slice(0, 3).map((e) => e.textContent?.slice(0, 20)) };
    pick(show, "10");
    await wait(300);
    // Fold the first heading with text under it: a click on its ⊕ (left of the paragraph).
    const first = root.querySelector<HTMLElement>(".pw-ol-plus")!;
    const r = first.getBoundingClientRect();
    first.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, clientX: r.left - 12, clientY: r.top + 6 }));
    await wait(300);
    out.folded = { shown: visible().length, foldedClass: first.classList.contains("pw-ol-folded") || !!root.querySelector(".pw-ol-folded") };
    first.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, clientX: r.left - 12, clientY: r.top + 6 }));
    await wait(300);
    out.unfolded = visible().length;
    // Demote the first Heading 1 to Heading 2, then promote it back; a body paragraph promoted.
    let hpos = -1;
    v.state.doc.forEach((n, pos) => { if (hpos < 0 && ed.outlineLevelAt(v.state.apply(v.state.tr.setSelection(TextSelection.create(v.state.doc, pos + 1))), (globalThis as { primeBase?: () => PrimeDocument }).primeBase!()) === 1) hpos = pos; });
    v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, hpos + 1)));
    const level = () => (document.querySelector<HTMLSelectElement>('select[aria-label="Outline level"]')?.value);
    await wait(200);
    const before = level();
    document.querySelector<HTMLButtonElement>('button[aria-label="Demote"]')!.click();
    await wait(200);
    const demoted = { level: level(), style: view().state.doc.nodeAt(hpos)!.attrs.props.style };
    document.querySelector<HTMLButtonElement>('button[aria-label="Promote"]')!.click();
    await wait(200);
    out.demotePromote = { before, demoted, back: level() };
    void h2;
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Close Outline View")?.click();
    await wait(2500);
    out.backToPrint = !!document.querySelector(".pw-editor.paged") && !document.querySelector(".pw-outline");
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Read Aloud with a stand-in voice (a test browser cannot speak): paragraphs read in turn from the cursor, each word highlighted. */
export function readAloudSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const out: Record<string, unknown> = { spoken: [] as string[], highlighted: [] as string[] };
  // The stand-in: says each utterance's words (boundary events), then ends it.
  const fake = {
    speaking: false, paused: false, pending: false,
    getVoices: () => [], addEventListener() {}, removeEventListener() {}, cancel() { this.cancelled++; }, cancelled: 0,
    pause() {}, resume() {},
    speak(u: SpeechSynthesisUtterance) {
      (out.spoken as string[]).push(u.text.slice(0, 30));
      void (async () => {
        for (const m of u.text.slice(0, 40).matchAll(/\S+/g)) {
          await wait(30);
          u.onboundary?.({ name: "word", charIndex: m.index!, charLength: m[0].length } as unknown as SpeechSynthesisEvent);
          const hl = view.dom.querySelector(".pw-speaking")?.textContent;
          if (hl) (out.highlighted as string[]).push(hl);
        }
        if ((out.spoken as string[]).length < 3) u.onend?.({} as SpeechSynthesisEvent);
      })();
    },
  };
  Object.defineProperty(window, "speechSynthesis", { value: fake, configurable: true });
  void (async () => {
    await wait(1500);
    let para = -1;
    view.state.doc.forEach((n, pos) => { if (para < 0 && n.type.name === "paragraph" && n.textContent.length > 80) para = pos; });
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, para + 1)));
    [...document.querySelectorAll<HTMLButtonElement>("nav button")].find((b) => b.textContent === "Review")?.click();
    await wait(150);
    [...document.querySelectorAll<HTMLButtonElement>("button")].find((b) => b.textContent === "Read Aloud")?.click();
    await wait(2500);
    out.player = !!document.querySelector(".read-aloud");
    out.firstParagraph = view.state.doc.nodeAt(para)!.textContent.slice(0, 30);
    out.highlighted = (out.highlighted as string[]).slice(0, 8);
    document.querySelector<HTMLButtonElement>('.read-aloud button[aria-label="Stop"]')?.click();
    await wait(200);
    out.stopped = { highlight: !!view.dom.querySelector(".pw-speaking"), cancelled: fake.cancelled > 0 };
    document.querySelector<HTMLButtonElement>('button[aria-label="Close read aloud"]')?.click();
    await wait(200);
    out.closed = !document.querySelector(".read-aloud");
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/** Grammar: issues typed into a paragraph are underlined; the right-click menu explains and fixes one. */
export function grammarSelfTest(view: EditorView) {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const out: Record<string, unknown> = {};
  void (async () => {
    await wait(3000);
    let para = -1;
    view.state.doc.forEach((n, pos) => { if (para < 0 && n.type.name === "paragraph" && n.textContent.length > 80) para = pos; });
    const at = para + view.state.doc.nodeAt(para)!.nodeSize - 1;
    view.dispatch(view.state.tr.insertText(" We saw a elephant in the the park. then i left.", at));
    await wait(1500);
    const marks = [...view.dom.querySelectorAll(".pw-grammar")].map((e) => e.textContent);
    out.underlined = marks;
    const el = [...view.dom.querySelectorAll<HTMLElement>(".pw-grammar")].find((e) => e.textContent === "a");
    if (el) {
      const r = el.getBoundingClientRect();
      el.dispatchEvent(new MouseEvent("contextmenu", { bubbles: true, cancelable: true, clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 }));
      await wait(300);
      out.menu = { message: document.querySelector(".spell-menu .grammar-message")?.textContent, options: [...document.querySelectorAll(".spell-menu .suggestion")].map((b) => b.textContent) };
      document.querySelector<HTMLButtonElement>(".spell-menu .suggestion")?.click();
      await wait(1200);
      out.fixed = view.state.doc.nodeAt(para)!.textContent.includes("saw an elephant");
    }
    out.status = document.querySelector("footer .link")?.textContent;
    console.log("[selftest] align", JSON.stringify(out));
  })();
}

/**
 * Compare: an "original" made from the open document (a word changed, one deleted, a paragraph
 * removed, one added), compared with it. Accepting every change must give the revised text exactly;
 * rejecting every change, the original's.
 */
export function compareSelfTest() {
  const wait = (ms: number) => new Promise((r) => setTimeout(r, ms));
  const out: Record<string, unknown> = {};
  const g = globalThis as unknown as { primeView: EditorView; primeBase: () => PrimeDocument; primeCompare: (f: { bytes: Uint8Array; name: string }) => Promise<void> };
  void (async () => {
    const ed = await import("@prime/editor");
    const { writeDocx } = await import("@prime/docx");
    await wait(2500);
    const v = g.primeView;
    const revisedText = (d: import("prosemirror-model").Node) => { const t: string[] = []; d.forEach((n) => t.push(n.textContent)); return t.join("\n"); };
    const revised = revisedText(v.state.doc);
    // The original: the open document, edited.
    let tr = v.state.tr;
    const paras: number[] = [];
    v.state.doc.forEach((n, pos) => { if (n.type.name === "paragraph" && n.textContent.length > 60) paras.push(pos); });
    const p1 = paras[3];
    const t1 = v.state.doc.nodeAt(p1)!.textContent;
    const w = t1.indexOf("preserve");
    tr = tr.insertText("keep", p1 + 1 + w, p1 + 1 + w + "preserve".length);            // a word changed
    const p2 = tr.mapping.map(paras[6]);
    tr = tr.delete(p2 + 1, p2 + 1 + 6);                                                    // a word deleted (the original lacks it: an insertion)
    const p3 = tr.mapping.map(paras[9]);
    tr = tr.delete(p3, p3 + tr.doc.nodeAt(p3)!.nodeSize);                                  // a paragraph the original lacks
    const p4 = tr.mapping.map(paras[12]);
    tr = tr.insert(p4, v.state.schema.nodes.paragraph.create({}, v.state.schema.text("A paragraph only the original had.")));
    const original = revisedText(tr.doc);
    const bytes = writeDocx(ed.fromProseMirror(tr.doc, g.primeBase()));
    await g.primeCompare({ bytes, name: "original.docx" });
    await wait(4000);
    const v2 = (globalThis as unknown as { primeView: EditorView }).primeView;
    const list = ed.changes(v2.state.doc).map((c) => `${c.kind}:${c.paragraph !== undefined ? "¶" : v2.state.doc.textBetween(c.from, c.to).slice(0, 40)}`);
    out.changes = list;
    const all = (accept: boolean) => { let st = v2.state; ed.acceptOrReject(accept, true)(st, (t) => { st = st.apply(t); }); return revisedText(st.doc); };
    const accepted = all(true);
    const rejected = all(false);
    out.acceptAllIsRevised = accepted === revised;
    out.rejectAllIsOriginal = rejected === original;
    if (!out.rejectAllIsOriginal) {
      const a = rejected.split("\n"); const b = original.split("\n");
      out.firstDiff = a.findIndex((x, i) => x !== b[i]);
      out.lines = [a[out.firstDiff as number]?.slice(0, 80), b[out.firstDiff as number]?.slice(0, 80)];
    }
    out.name = document.querySelector("footer.status span")?.textContent;
    console.log("[selftest] align", JSON.stringify(out));
  })();
}
