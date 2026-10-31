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
  top: (el: Element) => number, bottom: (el: Element) => number): { kind: "row" | "block"; pos: number } | undefined {
  const rows = Array.from(view.dom.querySelectorAll("tr"));
  const blocks = Array.from(view.dom.children).filter((el) => !el.classList.contains("pw-page-gap"));
  const candidates = [...rows, ...blocks]
    .filter((el) => top(el) > after + 1 && top(el) < before && bottom(el) > limit)
    .sort((a, b) => top(a) - top(b));
  for (const el of candidates) {
    const $pos = view.state.doc.resolve(view.posAtDOM(el, 0));
    for (let d = $pos.depth; d > 0; d--) {
      const name = $pos.node(d).type.name;
      if (el.tagName === "TR" && name === "table_row") return { kind: "row", pos: $pos.before(d) };
      if (el.tagName !== "TR" && d === 1) return { kind: "block", pos: $pos.before(1) };
    }
  }
  return undefined;
}

export function pagination(opts: PagedOptions): Plugin {
  let layout: LayoutResult | undefined;
  /** The page starts currently drawn (from the layout, possibly moved earlier by `fit`). */
  let starts: ResolvedStart[] = [];
  /** Corrections made since the last layout, to stop a run-away loop. */
  let corrections = 0;

  /**
   * Sizes each spacer so the next content lands at the top of its page's body area. Sizes go into a
   * stylesheet outside the editor: changing the editor's own DOM would make ProseMirror redraw it.
   * If the screen's content overflows a page (it ran longer than the layout predicted), that page's
   * break moves earlier, so a page never overflows; the screen has the last word.
   */
  const fit = (view: EditorView, sheet: HTMLStyleElement) => {
    if (!layout?.pages.length) return;
    const first = layout.pages[0];
    const pitch = (first.height + PAGE_GAP) * PX;
    const rules: string[] = [];
    sheet.textContent = "";
    const box = view.dom.getBoundingClientRect();
    // Screen measurements include any zoom; sizes we set are in unzoomed CSS pixels.
    const zoom = view.dom.offsetWidth ? box.width / view.dom.offsetWidth : 1;
    const top = (el: Element) => (el.getBoundingClientRect().top - box.top) / zoom;
    const bottom = (el: Element) => (el.getBoundingClientRect().bottom - box.top) / zoom;
    const bodyTopAt = (page: number) => (page - 1) * pitch + ((layout!.pages[page - 1]?.bodyTop ?? first.bodyTop) - first.bodyTop) * PX;
    const overflow = new Set<number>();
    const items = view.dom.querySelectorAll<HTMLElement>(".pw-page-gap, .pw-page-row");
    for (const el of Array.from(items)) {
      // In document order: each rule moves everything after it, so measure after applying the previous one.
      const page = Number(el.dataset.page);
      const target = bodyTopAt(page);
      const size = Math.max(0, target - top(el));
      // Content before this spacer should end inside the previous page's body area.
      const previous = layout.pages[page - 2];
      const limit = previous ? (page - 2) * pitch + (previous.bodyBottom - first.bodyTop) * PX + 1 : Infinity;
      if (top(el) > limit) {
        const moved = corrections < 50 ? earlierBreak(view, limit, bodyTopAt(page - 1), top(el), top, bottom) : undefined;
        if (moved) {
          corrections++;
          starts = starts.map((s) => (s.page === page ? { page, ...moved } : s));
          view.dispatch(view.state.tr.setMeta(paginationKey, starts).setMeta("addToHistory", false));
          return; // the redraw triggers another fit
        }
        overflow.add(page);
      }
      rules.push(el.classList.contains("pw-page-row")
        ? `.pw-editor tr.pw-page-row[data-page="${page}"] > td { padding-top: ${size}px; }`
        : `.pw-editor .pw-page-gap[data-page="${page}"] { height: ${size}px; }`);
      sheet.textContent = rules.join("\n");
    }
    opts.onFit?.(overflow);
  };

  return new Plugin<PagedState>({
    key: paginationKey,
    state: {
      init: () => ({ decorations: DecorationSet.empty }),
      apply(tr, old) {
        const starts = tr.getMeta(paginationKey) as ResolvedStart[] | undefined;
        if (starts) return { decorations: decorations(tr.doc, starts) };
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
        fitTimer = setTimeout(() => fit(view, sheet), 0);
      };
      const relayout = () => {
        const doc = fromProseMirror(view.state.doc, opts.base());
        layout = layoutDocument(doc, opts.measurer);
        opts.onLayout?.(layout);
        starts = resolvePageStarts(view.state.doc, layout.pageStarts);
        corrections = 0;
        view.dispatch(view.state.tr.setMeta(paginationKey, starts).setMeta("addToHistory", false));
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
          refit();
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
