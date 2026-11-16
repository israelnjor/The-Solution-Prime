/**
 * Clicking anywhere on a page puts the cursor where text can be typed, as in Word: on the nearest line
 * of the page clicked, at the nearest point along it. Below the last line, in a margin, between pages,
 * on the grey beside a page, or on the space a page break leaves: the browser does nothing there by
 * itself (it is not the editor), and the editor lost the focus, so no cursor showed at all.
 * Dragging from there selects, and Shift+click extends the selection, as from the text.
 */
import { Plugin, Selection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/** Places a click on the page leaves to what is there (controls, notes, pictures, text boxes). */
const OWN_CLICK = "button, input, select, textarea, a, label, [role=dialog], .pw-chrome-note, .pw-float, .conversion, .paste-options";
/** Spaces in the editor with no text to put the cursor in (between pages, above a block kept at the bottom). */
const NO_TEXT = ".pw-page-gap, .pw-pin-gap";

/** The pages on screen (the sheets in Print Layout; the one page otherwise). */
function pagesOf(view: EditorView): DOMRect[] {
  const sheets = view.dom.closest(".pw-sheets")?.querySelectorAll<HTMLElement>(".pw-sheet");
  if (sheets?.length) return Array.from(sheets, (s) => s.getBoundingClientRect());
  const page = view.dom.closest<HTMLElement>(".page") ?? view.dom;
  return [page.getBoundingClientRect()];
}

/**
 * Whether the cursor at `pos` is on the line at height y. The box of the character after it is looked
 * at first: the cursor's own box, at the start of the first line on a page, is the space between the
 * pages before it (the browser measures it against that), which would take a click in the margin for
 * a click on the line. Before a character it cannot have (the end of a line), the box before it.
 */
function onLine(view: EditorView, pos: number, y: number): boolean {
  const spans = (side: number, tall: number) => {
    try {
      const c = view.coordsAtPos(pos, side);
      return c.bottom - c.top > 2 && c.bottom - c.top < tall && y >= c.top - 1 && y <= c.bottom + 1;
    } catch {
      return false;
    }
  };
  return spans(1, Infinity) || spans(-1, 200);
}

/**
 * The text position nearest a point on a page: the nearest line of that page (searching up and down
 * from the point, a few pixels at a time), at the point's distance along the line (kept inside the
 * text column). Undefined when there is no text at all.
 */
export function caretNear(view: EditorView, x: number, y: number): number | undefined {
  const text = view.dom.getBoundingClientRect();
  const left = text.left + 1;
  const right = Math.max(left, text.right - 1);
  const pages = pagesOf(view);
  // The page clicked, or the nearest one (a click between pages, or beside them).
  const distance = (r: DOMRect) => (y < r.top ? r.top - y : y > r.bottom ? y - r.bottom : 0);
  const page = pages.reduce((best, r) => (distance(r) < distance(best) ? r : best), pages[0]);
  const top = Math.max(page.top, text.top);
  const bottom = Math.min(page.bottom, text.bottom);
  const at = Math.min(Math.max(y, top), bottom);
  const cx = Math.min(Math.max(x, left), right);
  /** The position at a point, if it is on a line of text there (moved into the text from a cell's padding). */
  const textAt = (px: number, yy: number): number | undefined => {
    const found = view.posAtCoords({ left: px, top: yy });
    // Not inside any block: the space between paragraphs, or between pages.
    if (!found || found.inside < 0) return undefined;
    const $pos = view.state.doc.resolve(found.pos);
    if ($pos.parent.inlineContent) return onLine(view, found.pos, yy) ? found.pos : undefined;
    // Between blocks inside a table (a cell's padding, beside a row): the text on this line either side.
    return [Selection.findFrom($pos, 1, true), Selection.findFrom($pos, -1, true)].map((sel) => sel?.head)
      .find((pos) => pos !== undefined && onLine(view, pos, yy));
  };
  const hit = (yy: number): number | undefined => {
    if (yy < top || yy > bottom) return undefined;
    if (document.elementFromPoint(cx, yy)?.closest(NO_TEXT)) return undefined;
    let pos = textAt(cx, yy);
    if (pos === undefined) {
      // Beside a block narrower than the text column (a table, an indented paragraph): its nearest edge.
      const block = [left, (left + right) / 2, right].map((px) => document.elementFromPoint(px, yy))
        .find((el) => el && el !== view.dom && view.dom.contains(el) && !el.closest(NO_TEXT));
      const box = block?.closest("tr, li, p, h1, h2, h3, h4, h5, h6, pre, blockquote")?.getBoundingClientRect();
      if (box) pos = textAt(Math.min(Math.max(cx, box.left + 2), box.right - 2), yy);
    }
    return pos;
  };
  /** The line found: asked again at its middle (at its very edge, the point can read as the paragraph's start). */
  const middle = (pos: number): number => {
    try {
      const c = view.coordsAtPos(pos, 1);
      const again = c.bottom - c.top > 2 ? textAt(cx, (c.top + c.bottom) / 2) : undefined;
      return again ?? pos;
    } catch {
      return pos;
    }
  };
  for (let d = 0; d <= bottom - top; d += 3) {
    const pos = hit(at + d) ?? hit(at - d);
    if (pos !== undefined) return middle(pos);
  }
  return view.posAtCoords({ left: cx, top: at })?.pos;
}

/** Whether a mouse-down is one this plugin places: on the page but not on its text or on a control. */
function placed(view: EditorView, e: MouseEvent): boolean {
  // Drawing a shape (drawing.ts) takes the mouse.
  if (e.button !== 0 || e.defaultPrevented || document.body.classList.contains("pw-drawing")) return false;
  const t = e.target as HTMLElement | null;
  if (!t?.closest || !view.editable) return false;
  if (t.closest(OWN_CLICK)) return false;
  if (view.dom.contains(t)) return t === view.dom || !!t.closest(NO_TEXT);
  // On a page (its sheet, margins, header and footer), or on the grey around the pages: an element
  // the editor is in, within the document's scrolling area (not the ribbon, a pane or a banner).
  const onPage = !!t.closest(".pw-sheets, .page");
  const around = t.contains(view.dom) && !!t.closest(".canvas");
  if (!onPage && !around) return false;
  // The scroll bar (it belongs to the scrolling element, which is not zoomed): dragging it must scroll.
  if (getComputedStyle(t).overflowY !== "visible" && !t.closest(".zoomed")) {
    const box = t.getBoundingClientRect();
    if (e.clientX >= box.left + t.clientLeft + t.clientWidth || e.clientY >= box.top + t.clientTop + t.clientHeight) return false;
  }
  // A picture behind the text, clicked in the margin, is selected by the floats plugin.
  return !Array.from(view.dom.querySelectorAll<HTMLElement>(".pw-float-behind")).some((el) => {
    const r = el.getBoundingClientRect();
    return e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
  });
}

export function pageClicks(): Plugin {
  return new Plugin({
    view(view) {
      const down = (e: MouseEvent) => {
        if (!placed(view, e)) return;
        const pos = caretNear(view, e.clientX, e.clientY);
        if (pos === undefined) return;
        // Not the browser's: it would take the focus from the editor (and select nothing). Not the
        // editor's either, on its own background or a page gap: on mouse-up it would put the cursor at
        // its own guess for the point (the start of the paragraph after the gap).
        e.preventDefault();
        if (view.dom.contains(e.target as Node)) e.stopPropagation();
        const { doc, selection } = view.state;
        const anchor = e.shiftKey ? selection.anchor : pos;
        view.dispatch(view.state.tr.setSelection(TextSelection.between(doc.resolve(anchor), doc.resolve(pos))));
        view.focus();
        // Dragging selects from there.
        let frame = 0;
        const move = (m: MouseEvent) => {
          if (!(m.buttons & 1)) return up();
          cancelAnimationFrame(frame);
          frame = requestAnimationFrame(() => {
            const head = caretNear(view, m.clientX, m.clientY);
            if (head === undefined || view.isDestroyed) return;
            const d = view.state.doc;
            const a = Math.min(anchor, d.content.size);
            if (view.state.selection.head !== head || view.state.selection.anchor !== a) {
              view.dispatch(view.state.tr.setSelection(TextSelection.between(d.resolve(a), d.resolve(head))));
            }
          });
        };
        const up = () => {
          cancelAnimationFrame(frame);
          window.removeEventListener("mousemove", move);
          window.removeEventListener("mouseup", up);
        };
        window.addEventListener("mousemove", move);
        window.addEventListener("mouseup", up);
      };
      // Captured: before the editor's own handling of a mouse-down on its root or a page gap.
      document.addEventListener("mousedown", down, true);
      return { destroy: () => document.removeEventListener("mousedown", down, true) };
    },
  });
}
