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
