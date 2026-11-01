/**
 * Development self-tests that need a real browser (run with ?selftest=paste in the dev build).
 * Results are logged as "[selftest] ..." lines for the test scripts to read.
 */
import { fromProseMirror, timed } from "@prime/editor";
import { layoutDocument, type Measurer } from "@prime/layout";
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
  const corrected = performance.getEntriesByType("measure").filter((e) => e.name.startsWith("pw:correction")).map((e) => e.name.slice(14));
  const corrections = corrected.length;
  console.log("[selftest] corrected", corrected.length, JSON.stringify(corrected.slice(-8)));
  // How far each page's content ends from its body bottom (the top of the mask below it), in px:
  // positive = the screen ran longer than the layout there.
  const masks = Array.from(document.querySelectorAll(".pw-mask"));
  const ends = Array.from(document.querySelectorAll<HTMLElement>(".pw-page-gap")).map((g) => {
    const top = g.getBoundingClientRect().top;
    const mask = masks.map((m) => m.getBoundingClientRect().top).filter((t) => t <= top + 200).pop();
    return mask === undefined ? null : [Number(g.dataset.page) - 1, Math.round(top - mask)];
  });
  console.log("[selftest] page ends", JSON.stringify(ends));
  const rounds = performance.getEntriesByType("measure").filter((e) => e.name.startsWith("pw:round")).map((e) => Number(e.name.slice(9)));
  return { corrections, rounds, pages: sheets.length, checked, worstPx: Math.round(worst * 10) / 10, offByMoreThan1px: off, firstOff: bad.slice(0, 8), masked: Array.from(document.querySelectorAll('.pw-mask')).length, unmasked: sheets.length - 1 - document.querySelectorAll(".pw-mask").length };
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
    else document.querySelectorAll(".pw-sheet")[n - 1]?.scrollIntoView({ block: "start" });
    // Debug aid: what sits around the page (spacers and wrap room), relative to the page's sheet.
    const sheet = document.querySelectorAll(".pw-sheet")[n - 1]?.getBoundingClientRect();
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
    console.log("[selftest] align", JSON.stringify({ scrolled: n }));
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
    const sheets = [...document.querySelectorAll(".pw-sheet")].map((e) => e.getBoundingClientRect().top);
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
