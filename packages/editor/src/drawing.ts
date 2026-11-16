/**
 * Drawing a shape on the page, as in Word: pick a shape, then drag where it goes (Shift: a square,
 * circle or straight line at 45° steps). A click without dragging puts it there at Word's default
 * size. Esc, or a click outside the pages, stops drawing. The shape is anchored to the paragraph
 * beside where it was drawn, and stays where it was drawn.
 */
import type { EditorView } from "prosemirror-view";
import { caretNear } from "./pageclick.js";
import { insertShape, isLine, type ShapeChoice, type ShapePlace } from "./shapes.js";

/** Screen pixels per point on the page (with the zoom). */
function pxPerPt(view: EditorView): number {
  const page = view.dom.closest<HTMLElement>(".pw-sheets, .page");
  const zoom = page && page.offsetWidth ? page.getBoundingClientRect().width / page.offsetWidth : 1;
  return (96 / 72) * zoom;
}

/** Where a box drawn on the screen (px) goes in the document: its paragraph and offsets from it (pt). */
export function placeOf(view: EditorView, box: { left: number; top: number; width: number; height: number }, flip: { flipH?: boolean; flipV?: boolean } = {}): ShapePlace | undefined {
  const k = pxPerPt(view);
  const pos = caretNear(view, box.left, box.top);
  if (pos === undefined) return undefined;
  const doc = view.state.doc;
  // A paragraph of the body (not inside a table): the nearest one at or after the place.
  const $pos = doc.resolve(pos);
  let index = $pos.index(0);
  const paragraphAt = (i: number) => (i >= 0 && i < doc.childCount && doc.child(i).type.name === "paragraph" ? i : undefined);
  let found = paragraphAt(index);
  for (let d = 1; found === undefined && d < doc.childCount; d++) found = paragraphAt(index + d) ?? paragraphAt(index - d);
  if (found === undefined) return undefined;
  index = found;
  const posOf = (i: number) => { let p = 0; for (let j = 0; j < i; j++) p += doc.child(j).nodeSize; return p; };
  let paraPos = posOf(index);
  let top = (view.nodeDOM(paraPos) as HTMLElement | null)?.getBoundingClientRect().top ?? box.top;
  // A paragraph that starts on an earlier page: the next one, starting on this page, if there is one
  // (an offset across a page break would put the shape on the page the paragraph starts on).
  const sheet = Array.from(view.dom.closest(".pw-sheets")?.querySelectorAll<HTMLElement>(".pw-sheet") ?? [])
    .map((s) => s.getBoundingClientRect()).find((r) => box.top >= r.top && box.top <= r.bottom);
  if (sheet && top < sheet.top) {
    for (let i = index + 1; i < doc.childCount; i++) {
      if (doc.child(i).type.name !== "paragraph") continue;
      const t = (view.nodeDOM(posOf(i)) as HTMLElement | null)?.getBoundingClientRect().top;
      if (t !== undefined && t >= sheet.top && t <= sheet.bottom) { paraPos = posOf(i); top = t; }
      break;
    }
  }
  const left = view.dom.getBoundingClientRect().left;
  return {
    paraPos,
    x: (box.left - left) / k, y: (box.top - top) / k,
    width: box.width / k, height: box.height / k,
    ...flip,
  };
}

/**
 * Starts drawing `choice` on the page (a shape, or a text box). Returns a function that stops it.
 * `onEnd` is called when drawing ends: with true when a shape was drawn.
 */
export function drawShape(view: EditorView, kind: "textbox" | "shape", choice: ShapeChoice, onEnd?: (drawn: boolean) => void): () => void {
  const line = kind === "shape" && isLine(choice.geometry);
  document.body.classList.add("pw-drawing");
  let band: HTMLElement | undefined;
  let ended = false;
  const stop = (drawn = false) => {
    if (ended) return;
    ended = true;
    document.body.classList.remove("pw-drawing");
    document.removeEventListener("mousedown", down, true);
    window.removeEventListener("keydown", key, true);
    band?.remove();
    onEnd?.(drawn);
  };
  const key = (e: KeyboardEvent) => { if (e.key === "Escape") { e.preventDefault(); stop(); } };
  const onPage = (t: HTMLElement) => !!t.closest(".pw-sheets, .page") || (t.contains(view.dom) && !!t.closest(".canvas"));
  const down = (e: MouseEvent) => {
    const t = e.target as HTMLElement;
    if (e.button !== 0 || !t.closest || !onPage(t) || t.closest("button, input, select, textarea")) { stop(); return; }
    e.preventDefault();
    e.stopPropagation();
    const start = { x: e.clientX, y: e.clientY };
    band = document.createElement("div");
    band.className = line ? "pw-draw-band pw-draw-line" : "pw-draw-band";
    document.body.appendChild(band);
    let end = start;
    let shift = false;
    /** The drawn box, with Shift: as wide as high (a line: along 0°, 45° or 90°). */
    const shaped = () => {
      let dx = end.x - start.x;
      let dy = end.y - start.y;
      if (shift && line) {
        const a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4);
        const len = Math.hypot(dx, dy);
        [dx, dy] = [Math.round(Math.cos(a) * len), Math.round(Math.sin(a) * len)];
      } else if (shift) {
        const s = Math.max(Math.abs(dx), Math.abs(dy));
        [dx, dy] = [Math.sign(dx || 1) * s, Math.sign(dy || 1) * s];
      }
      return { dx, dy };
    };
    const draw = () => {
      const { dx, dy } = shaped();
      if (!band) return;
      if (line) {
        Object.assign(band.style, { left: `${start.x}px`, top: `${start.y}px`, width: `${Math.hypot(dx, dy)}px`, rotate: `${Math.atan2(dy, dx)}rad` });
      } else {
        Object.assign(band.style, { left: `${Math.min(start.x, start.x + dx)}px`, top: `${Math.min(start.y, start.y + dy)}px`, width: `${Math.abs(dx)}px`, height: `${Math.abs(dy)}px` });
      }
    };
    const move = (m: MouseEvent) => { end = { x: m.clientX, y: m.clientY }; shift = m.shiftKey; draw(); };
    const up = (m: MouseEvent) => {
      window.removeEventListener("mousemove", move, true);
      window.removeEventListener("mouseup", up, true);
      end = { x: m.clientX, y: m.clientY };
      shift = m.shiftKey;
      let { dx, dy } = shaped();
      const k = pxPerPt(view);
      // A click: Word's default size there (a line an inch long, across).
      if (Math.abs(dx) < 4 && Math.abs(dy) < 4) {
        [dx, dy] = line ? [144 * k, 0] : kind === "textbox" ? [180 * k, 72 * k] : [72 * k, 72 * k];
      }
      const box = { left: Math.min(start.x, start.x + dx), top: Math.min(start.y, start.y + dy), width: Math.abs(dx), height: Math.abs(dy) };
      // A line runs from where the drag started: mirrored when it goes left or up.
      const flip = line ? { ...(dx < 0 ? { flipH: true } : {}), ...(dy < 0 ? { flipV: true } : {}) } : {};
      const place = placeOf(view, box, flip);
      if (!place) { stop(); return; }
      insertShape(kind, choice, place)(view.state, view.dispatch);
      view.focus();
      stop(true);
    };
    window.addEventListener("mousemove", move, true);
    window.addEventListener("mouseup", up, true);
    draw();
  };
  document.addEventListener("mousedown", down, true);
  window.addEventListener("keydown", key, true);
  return () => stop();
}
