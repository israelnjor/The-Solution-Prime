/**
 * Paginated editing (Phase 4, ADR-007). The layout engine decides where pages start; the editor keeps
 * one flowing document and receives page-gap spacers at those points. After each redraw the spacers
 * are sized from the real screen positions, so every page's content starts exactly at the top of
 * its sheet's body area — any small difference between engine and browser corrects itself.
 */
import { type LayoutResult, layoutDocument, type Measurer } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { fromProseMirror } from "./convert.js";
import { type ResolvedStart, resolvePageStarts } from "./pages.js";

/** Whether `timed` records (self-tests turn it on; entries would otherwise pile up over a session). */
const timing = () => (globalThis as { primeTiming?: boolean }).primeTiming === true;

/** Records an event on the browser's performance timeline ("pw:" entries; see the perf self-test). */
export function mark(name: string) {
  if (timing()) performance.measure(`pw:${name}`, { start: performance.now(), end: performance.now() });
}

/** Times a step for the browser's performance timeline (when timing is on). */
export function timed<T>(name: string, fn: () => T): T {
  if (!timing()) return fn();
  const t0 = performance.now();
  try { return fn(); } finally { performance.measure(`pw:${name}`, { start: t0, end: performance.now() }); }
}

/** Gap between page sheets on screen, in pt. */
export const PAGE_GAP = 16;
const PX = 96 / 72;

export interface PagedOptions {
  /** The document the editor was opened from (styles, lists, package parts). */
  base: () => PrimeDocument;
  measurer: Measurer;
  /** Called with every new layout (to draw page sheets, headers and footers). */
  onLayout?: (layout: LayoutResult) => void;
  /**
   * Called after spacers are sized, with the pages whose preceding content overflowed into the
   * margin (the screen ran longer than the layout). Margin masks must not hide those.
   */
  onFit?: (overflow: Set<number>) => void;
  /** Quiet time after typing before re-paginating, in ms. */
  delay?: number;
}

interface PagedState {
  decorations: DecorationSet;
}

export const paginationKey = new PluginKey<PagedState>("pagination");

function gap(page: number, inline: boolean): () => HTMLElement {
  return () => {
    const el = document.createElement(inline ? "span" : "div");
    el.className = "pw-page-gap";
    el.dataset.page = String(page);
    el.contentEditable = "false";
    return el;
  };
}

function decorations(doc: Node, starts: ResolvedStart[]): DecorationSet {
  const out: Decoration[] = [];
  for (const s of starts) {
    if (s.kind === "row") {
      const row = doc.nodeAt(s.pos);
      if (row) out.push(Decoration.node(s.pos, s.pos + row.nodeSize, { class: "pw-page-row", "data-page": String(s.page) }));
    } else {
      out.push(Decoration.widget(s.pos, gap(s.page, s.kind === "inline"), { side: -1, key: `gap-${s.page}-${s.kind}` }));
    }
  }
  return DecorationSet.create(doc, out);
}

/**
 * Where the break for a page should move when the screen's content overflows that page: before the
 * first row or top-level block (after the page's own start) that crosses the bottom of the body.
 */
function earlierBreak(view: EditorView, limit: number, after: number, before: number,
  top: (el: Element) => number, bottom: (el: Element) => number): { kind: "row" | "block"; pos: number; top: number } | undefined {
  // Only the blocks on this page (found by binary search: blocks are in screen order), and the rows of
  // its tables — scanning the whole document for every page made long documents slow.
  const blocks = Array.from(view.dom.children).filter((el) => !el.classList.contains("pw-page-gap"));
  let lo = 0;
  let hi = blocks.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (top(blocks[mid]) <= after + 1) lo = mid + 1;
    else hi = mid;
  }
  const near: Element[] = [];
  // From the block that starts before the page (a table continuing onto it) to the page's end.
  for (let i = Math.max(0, lo - 1); i < blocks.length && top(blocks[i]) < before; i++) near.push(blocks[i]);
  const rows = near.flatMap((el) => Array.from(el.querySelectorAll("tr")));
  const candidates = [...rows, ...near]
    .filter((el) => top(el) > after + 1 && top(el) < before && bottom(el) > limit)
    .sort((a, b) => top(a) - top(b));
  for (const el of candidates) {
    const $pos = view.state.doc.resolve(view.posAtDOM(el, 0));
    for (let d = $pos.depth; d > 0; d--) {
      const name = $pos.node(d).type.name;
      if (el.tagName === "TR" && name === "table_row") return { kind: "row", pos: $pos.before(d), top: top(el) };
      if (el.tagName !== "TR" && d === 1) return { kind: "block", pos: $pos.before(1), top: top(el) };
    }
  }
  return undefined;
}

export function pagination(opts: PagedOptions): Plugin {
  let layout: LayoutResult | undefined;
  /** The page starts currently drawn (from the layout, possibly moved earlier by `fit`). */
  let starts: ResolvedStart[] = [];
  /** Rounds of corrections (each moves every break that needs it) since the last layout: a loop guard. */
  let rounds = 0;
  /** Pages whose overflow has no earlier break (e.g. one block taller than the page), until the next layout. */
  let unfixable = new Set<number>();

  /** Spacer sizes currently in the stylesheet, by spacer ("gap-3", "row-3"). */
  let applied = new Map<string, number>();

  /**
   * Sizes each spacer so the next content lands at the top of its page's body area. Sizes go into a
   * stylesheet outside the editor: changing the editor's own DOM would make ProseMirror redraw it.
   * If the screen's content overflows a page (it ran longer than the layout predicted), that page's
   * break moves earlier, so a page never overflows; the screen has the last word.
   *
   * Positions are read in one pass (each read after a style change would re-measure the whole
   * document — in a long document, that is what made typing slow): growing a spacer by d moves
   * everything after it down by d, so the later positions follow from the earlier sizes. One
   * check afterwards catches the rare case where that does not hold exactly (text reflowing
   * around an inline spacer), with a second pass. Breaks that must move earlier are found with the
   * same arithmetic, all in one pass and one redraw (one redraw per page made a long document's
   * pages jump around for seconds after each pause in typing).
   */
  const fit = (view: EditorView, sheet: HTMLStyleElement, pass = 0) => {
    if (!layout?.pages.length) return;
    const first = layout.pages[0];
    const pitch = (first.height + PAGE_GAP) * PX;
    const box = timed("fit.reflow", () => view.dom.getBoundingClientRect());
    // Screen measurements include any zoom; sizes we set are in unzoomed CSS pixels.
    const zoom = view.dom.offsetWidth ? box.width / view.dom.offsetWidth : 1;
    const top = (el: Element) => (el.getBoundingClientRect().top - box.top) / zoom;
    const bottom = (el: Element) => (el.getBoundingClientRect().bottom - box.top) / zoom;
    const bodyTopAt = (page: number) => (page - 1) * pitch + ((layout!.pages[page - 1]?.bodyTop ?? first.bodyTop) - first.bodyTop) * PX;
    const overflow = new Set<number>();
    const items = Array.from(view.dom.querySelectorAll<HTMLElement>(".pw-page-gap, .pw-page-row"));
    const tops = items.map(top); // the one measuring pass
    const sizes = new Map<string, number>();
    let shift = 0; // how far the planned sizes move content, so far
    let off = false;
    let moves = 0;
    let previousEnd = 0; // where the previous page's content starts, after the planned sizes
    // When the previous page's break moved earlier, content between the new break and the old spacer
    // (measured above `at`) never had the old spacer's size, which `shift` has already taken away.
    let lastMove: { at: number; size: number } | undefined;
    const extra = (measured: number) => (lastMove && measured < lastMove.at ? lastMove.size : 0);
    // Positions as they will be once the planned sizes apply (valid up to the next spacer).
    const plannedTop = (el: Element) => { const t = top(el); return t + shift + extra(t); };
    const plannedBottom = (el: Element) => bottom(el) + shift + extra(top(el));
    for (let i = 0; i < items.length; i++) {
      const el = items[i];
      const page = Number(el.dataset.page);
      const key = `${el.classList.contains("pw-page-row") ? "row" : "gap"}-${page}`;
      const current = applied.get(key) ?? 0;
      // Where the spacer would sit with no size of its own, after the sizes planned before it.
      const natural = tops[i] + shift;
      const size = Math.max(0, bodyTopAt(page) - natural);
      // Content before this spacer should end inside the previous page's body area.
      const previous = layout.pages[page - 2];
      const limit = previous ? (page - 2) * pitch + (previous.bodyBottom - first.bodyTop) * PX + 1 : Infinity;
      if (natural > limit) {
        if (rounds < 10 && !unfixable.has(page)) {
          // Only content after the previous page's start may move to this page.
          const moved = timed("fit.earlierBreak", () => earlierBreak(view, limit, Math.max(previousEnd, bodyTopAt(page - 1)), natural, plannedTop, plannedBottom));
          if (moved) {
            moves++;
            mark("correction");
            starts = starts.map((s) => (s.page === page ? { page, kind: moved.kind, pos: moved.pos } : s));
            // The spacer moves up to the new break: content from there moves by its new size, and
            // content after the old break no longer has the old spacer's size.
            const movedSize = Math.max(0, bodyTopAt(page) - moved.top);
            sizes.delete(key);
            sizes.set(`${moved.kind === "row" ? "row" : "gap"}-${page}`, movedSize);
            shift += movedSize - current;
            previousEnd = moved.top + movedSize;
            lastMove = { at: tops[i], size: current };
            continue;
          }
          unfixable.add(page);
        }
        overflow.add(page);
      }
      sizes.set(key, size);
      shift += size - current;
      previousEnd = natural + size;
      lastMove = undefined;
      if (Math.abs(size - current) > 0.5) off = true;
    }
    if (moves) {
      rounds++;
      mark(`round ${moves}`);
      // Sizes first, then the breaks: the redraw that follows refits (and checks) everything.
      write(sheet, sizes);
      view.dispatch(view.state.tr.setMeta(paginationKey, starts).setMeta("addToHistory", false));
      return;
    }
    if (off || sizes.size !== applied.size) timed("fit.write", () => write(sheet, sizes));
    // Verify: every spacer should now end at its page's body top.
    if (off && pass < 2) {
      const wrong = items.some((el) => Math.abs(top(el) + (sizes.get(`${el.classList.contains("pw-page-row") ? "row" : "gap"}-${el.dataset.page}`) ?? 0) - bodyTopAt(Number(el.dataset.page))) > 1);
      if (wrong) return fit(view, sheet, pass + 1);
    }
    opts.onFit?.(overflow);
  };

  const write = (sheet: HTMLStyleElement, sizes: Map<string, number>) => {
    const rules: string[] = [];
    for (const [key, size] of sizes) {
      const [kind, page] = key.split("-");
      rules.push(kind === "row"
        ? `.pw-editor tr.pw-page-row[data-page="${page}"] > td { padding-top: ${size}px; }`
        : `.pw-editor .pw-page-gap[data-page="${page}"] { height: ${size}px; }`);
    }
    sheet.textContent = rules.join("\n");
    applied = new Map(sizes);
  };

  return new Plugin<PagedState>({
    key: paginationKey,
    state: {
      init: () => ({ decorations: DecorationSet.empty }),
      apply(tr, old) {
        const starts = tr.getMeta(paginationKey) as ResolvedStart[] | undefined;
        if (starts) return { decorations: decorations(tr.doc, starts) };
        if (!tr.docChanged) return old;
        return { decorations: old.decorations.map(tr.mapping, tr.doc) };
      },
    },
    props: {
      decorations: (state) => paginationKey.getState(state)?.decorations,
    },
    view(view) {
      let timer: ReturnType<typeof setTimeout> | undefined;
      let fitTimer: ReturnType<typeof setTimeout> | undefined;
      const sheet = document.createElement("style");
      sheet.dataset.primePages = "";
      document.head.appendChild(sheet);
      // A timeout, not requestAnimationFrame: frames do not run in hidden or headless windows, and
      // reading positions forces layout anyway.
      const refit = () => {
        clearTimeout(fitTimer);
        fitTimer = setTimeout(() => timed("fit", () => fit(view, sheet)), 0);
      };
      const relayout = () => {
        const doc = timed("toModel", () => fromProseMirror(view.state.doc, opts.base()));
        const result = timed("layout", () => layoutDocument(doc, opts.measurer));
        layout = result;
        timed("onLayout", () => opts.onLayout?.(result));
        starts = timed("pageStarts", () => resolvePageStarts(view.state.doc, result.pageStarts));
        rounds = 0;
        unfixable = new Set();
        timed("drawBreaks", () => view.dispatch(view.state.tr.setMeta(paginationKey, starts).setMeta("addToHistory", false)));
        refit();
      };
      timer = setTimeout(relayout, 0);
      const onResize = () => refit();
      window.addEventListener("resize", onResize);
      document.fonts?.addEventListener?.("loadingdone", onResize);
      return {
        update(v, prev) {
          if (v.state.doc !== prev.doc) {
            clearTimeout(timer);
            timer = setTimeout(relayout, opts.delay ?? 250);
          }
          if (v.state.doc !== prev.doc || paginationKey.getState(v.state) !== paginationKey.getState(prev)) refit();
        },
        destroy() {
          clearTimeout(timer);
          clearTimeout(fitTimer);
          sheet.remove();
          window.removeEventListener("resize", onResize);
          document.fonts?.removeEventListener?.("loadingdone", onResize);
        },
      };
    },
  });
}
