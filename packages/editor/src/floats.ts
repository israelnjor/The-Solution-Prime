/**
 * Moving and resizing floating pictures and text boxes. The object stays the XML Word wrote; only its
 * position (wp:positionH / wp:positionV) and size (wp:extent and the shape's a:ext) are rewritten.
 */
import type { FloatInfo, PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, NodeSelection, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { adjustShape, shapeHandles } from "@prime/layout";
import { shapeSvg } from "./paged.js";
import { setShapeAdjust, setShapeTurn } from "./shapes.js";

const EMU = 12700;
const emu = (pt: number) => String(Math.round(pt * EMU));

/** A floating object's box: x from the column's left edge, y from its paragraph's first line (pt). */
export interface FloatBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** The `index`-th floating object in the paragraph at `paraPos` (document order). */
export function floatAt(doc: Node, paraPos: number, index: number): { pos: number; node: Node } | undefined {
  const para = doc.nodeAt(paraPos);
  if (!para) return undefined;
  let n = 0;
  let found: { pos: number; node: Node } | undefined;
  para.descendants((node, pos) => {
    if (found) return false;
    // Kept floating objects (from Word files) and floating pictures, in document order (as the layout counts them).
    if ((node.type.name === "run_preserved" || node.type.name === "image") && node.attrs.float && n++ === index) found = { pos: paraPos + 1 + pos, node };
    return true;
  });
  return found;
}

/**
 * The object's XML with a new box. Positions become offsets from the column and the paragraph (what
 * Word writes after a drag); the old VML fallback, which would keep the old size, is dropped.
 */
export function withBox(xml: string, box: FloatBox, keepHeight: boolean): string {
  const height = (cy: string) => (keepHeight ? cy : emu(box.height));
  return xml
    .replace(/<mc:Fallback>[\s\S]*?<\/mc:Fallback>/, "")
    .replace(/(<wp:anchor\b[^>]*?)simplePos="1"/, '$1simplePos="0"')
    .replace(/<wp:positionH\b[^>]*>[\s\S]*?<\/wp:positionH>/, `<wp:positionH relativeFrom="column"><wp:posOffset>${emu(box.x)}</wp:posOffset></wp:positionH>`)
    .replace(/<wp:positionV\b[^>]*>[\s\S]*?<\/wp:positionV>/, `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${emu(box.y)}</wp:posOffset></wp:positionV>`)
    .replace(/<wp:extent cx="\d+" cy="(\d+)"\s*\/>/, (_, cy) => `<wp:extent cx="${emu(box.width)}" cy="${height(cy)}"/>`)
    .replace(/(<a:xfrm\b[^>]*>\s*<a:off\b[^>]*\/>\s*)<a:ext cx="\d+" cy="(\d+)"\s*\/>/, (_, head, cy) => `${head}<a:ext cx="${emu(box.width)}" cy="${height(cy)}"/>`);
}

/** Moves and/or resizes the floating object to `box`. */
export function setFloatBox(paraPos: number, index: number, box: FloatBox): Command {
  return (state, dispatch) => {
    const target = floatAt(state.doc, paraPos, index);
    if (!target) return false;
    if (dispatch) {
      const float = target.node.attrs.float as FloatInfo;
      if (target.node.type.name === "image") {
        // A picture of our own: its position and size are attributes.
        const tr = state.tr.setNodeMarkup(target.pos, undefined, {
          ...target.node.attrs, width: box.width, height: box.height,
          float: { ...float, x: { from: "column", offset: box.x }, y: { from: "paragraph", offset: box.y } },
        });
        if (state.selection instanceof NodeSelection && state.selection.from === target.pos) tr.setSelection(NodeSelection.create(tr.doc, target.pos));
        dispatch(tr);
        return true;
      }
      // A text box that grows with its text keeps doing so: only its width is set.
      const keepHeight = float.kind === "textbox" && (float.autoHeight || float.height < 1);
      const preserved = target.node.attrs.preserved as { name: string; xml: string };
      const tr = state.tr.setNodeMarkup(target.pos, undefined, {
        ...target.node.attrs,
        preserved: { ...preserved, xml: withBox(preserved.xml, box, keepHeight) },
        float: {
          ...float, x: { from: "column", offset: box.x }, y: { from: "paragraph", offset: box.y },
          width: box.width, ...(keepHeight ? {} : { height: box.height }),
        },
      });
      // Still selected after a move or resize (the next nudge, or its format, applies to it).
      if (state.selection instanceof NodeSelection && state.selection.from === target.pos) tr.setSelection(NodeSelection.create(tr.doc, target.pos));
      dispatch(tr);
    }
    return true;
  };
}

/** Removes the floating object. */
export function deleteFloat(paraPos: number, index: number): Command {
  return (state, dispatch) => {
    const target = floatAt(state.doc, paraPos, index);
    if (!target) return false;
    if (dispatch) dispatch(state.tr.delete(target.pos, target.pos + target.node.nodeSize));
    return true;
  };
}

/** The paragraph position and float index of a drawn float (see floatLayer in paged.ts). */
function locate(view: EditorView, el: HTMLElement): { paraPos: number; index: number } | undefined {
  const para = el.parentElement?.closest("p");
  if (!para) return undefined;
  const inside = view.posAtDOM(para, 0);
  const $pos = view.state.doc.resolve(inside);
  return { paraPos: $pos.before($pos.depth), index: Number(el.dataset.index) };
}

/** The floating object (its node and position) a layer on the page shows. */
export function floatNodeOf(view: EditorView, el: HTMLElement): { pos: number; node: Node } | undefined {
  const where = locate(view, el);
  return where ? floatAt(view.state.doc, where.paraPos, where.index) : undefined;
}

/** A key typed with a shape or text box selected: { pos, text } (the app opens it for typing). */
export const SHAPE_TYPING = "prime-word:shape-typing";

/** Screen pixels per point for a drawn object: its untransformed width (turning ignored) and the page's zoom. */
function scaleOf(el: HTMLElement, width: number): number {
  const page = el.closest<HTMLElement>(".pw-sheets, .page");
  const zoom = page && page.offsetWidth ? page.getBoundingClientRect().width / page.offsetWidth : 1;
  return ((el.offsetWidth || width * (96 / 72)) / width) * zoom || 96 / 72;
}

/** The drawn layer of the floating object at `pos`, if it is on screen. */
export function floatElementAt(view: EditorView, pos: number): HTMLElement | undefined {
  const $pos = view.state.doc.resolve(pos);
  const paraPos = $pos.before($pos.depth);
  let index = -1;
  for (let i = 0; ; i++) {
    const f = floatAt(view.state.doc, paraPos, i);
    if (!f) break;
    if (f.pos === pos) { index = i; break; }
  }
  const para = view.nodeDOM(paraPos) as HTMLElement | null;
  if (index < 0 || !para?.querySelectorAll) return undefined;
  return Array.from(para.querySelectorAll<HTMLElement>(".pw-float")).find((el) => el.dataset.index === String(index) && el.parentElement?.closest("p") === para);
}

/** The topmost picture behind the text under a point, if any. */
function behindAt(view: EditorView, x: number, y: number): HTMLElement | null {
  const all = Array.from(view.dom.querySelectorAll<HTMLElement>(".pw-float-behind")).reverse();
  return all.find((el) => { const r = el.getBoundingClientRect(); return x >= r.left && x <= r.right && y >= r.top && y <= r.bottom; }) ?? null;
}

/** Whether a point is on a character of the text (its letter's box). */
function overText(x: number, y: number): boolean {
  const range = (document as Document & { caretRangeFromPoint?(x: number, y: number): Range | null }).caretRangeFromPoint?.(x, y);
  const node = range?.startContainer;
  if (!range || !node || node.nodeType !== 3) return false; // a text node
  const length = node.textContent?.length ?? 0;
  for (const at of [range.startOffset, range.startOffset - 1]) {
    if (at < 0 || at >= length) continue;
    const letter = document.createRange();
    letter.setStart(node, at);
    letter.setEnd(node, at + 1);
    const b = letter.getBoundingClientRect();
    if (x >= b.left - 2 && x <= b.right + 2 && y >= b.top - 2 && y <= b.bottom + 2) return true;
  }
  return false;
}

/**
 * Click a floating object to select it; drag it to move it; drag a corner handle to resize it
 * (pictures keep their proportions); Delete removes it, Escape deselects.
 */
export function floatEditing(): Plugin {
  let selected: HTMLElement | undefined;
  const select = (el: HTMLElement | undefined) => {
    selected?.classList.remove("pw-float-selected");
    selected = el;
    el?.classList.add("pw-float-selected");
  };
  function onMouseDown(view: EditorView, event: MouseEvent): boolean {
    const target = event.target as HTMLElement;
    // A picture behind the text lets the mouse through (styles.css): a click on text places the
    // cursor; a click where no text is selects it (as Word does).
    const el = (target.closest?.(".pw-float") as HTMLElement | null)
      ?? (overText(event.clientX, event.clientY) ? null : behindAt(view, event.clientX, event.clientY));
    if (!el) { select(undefined); return false; }
    event.preventDefault();
    select(el);
    // Also select it in the document, so the ribbon shows the picture's options.
    const where = locate(view, el);
    const node = where && floatAt(view.state.doc, where.paraPos, where.index);
    if (node) view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, node.pos)));
    const handle = target.dataset.handle; // "nw" | "ne" | "sw" | "se" when resizing, "rot" when turning
    const start = { x: event.clientX, y: event.clientY };
    const box: FloatBox = { x: Number(el.dataset.x), y: Number(el.dataset.y), width: Number(el.dataset.w), height: Number(el.dataset.h) };
    const turned = Number(el.dataset.rotation ?? 0);
    // Screen pixels per point, including any zoom (measured unturned: a turned box's bounds are larger).
    const scale = scaleOf(el, box.width);
    const picture = el.dataset.kind === "picture";
    const line = el.dataset.line === "1";
    if (handle === "rot") {
      // Turning: about the centre, following the mouse (0 degrees straight up); Shift in steps of 15.
      const r = el.getBoundingClientRect();
      const [cx, cy] = [r.left + r.width / 2, r.top + r.height / 2];
      let angle = turned;
      const turn = (e: MouseEvent) => {
        angle = (Math.atan2(e.clientX - cx, -(e.clientY - cy)) * 180) / Math.PI;
        if (e.shiftKey) angle = Math.round(angle / 15) * 15;
        angle = ((angle % 360) + 360) % 360;
        el.style.rotate = `${angle}deg`;
      };
      const done = () => {
        window.removeEventListener("mousemove", turn);
        window.removeEventListener("mouseup", done);
        if (Math.abs(angle - turned) > 0.05) setShapeTurn({ rotation: Math.round(angle * 10) / 10 })(view.state, view.dispatch);
      };
      window.addEventListener("mousemove", turn);
      window.addEventListener("mouseup", done);
      return true;
    }
    if (handle?.startsWith("adj:")) {
      // Reshaping: the handle follows the mouse (in the shape's own, unturned and unmirrored box); the
      // outline is drawn again as it goes, and the adjustments are kept on release.
      const f = node?.node.attrs.float as FloatInfo | undefined;
      const style = f?.shape;
      if (!f || !style) return true;
      const index = Number(handle.slice(4));
      const r = el.getBoundingClientRect();
      const [cx, cy] = [r.left + r.width / 2, r.top + r.height / 2];
      const t = (turned * Math.PI) / 180;
      let adj = style.adj ?? {};
      const point = (e: MouseEvent) => {
        const [dx, dy] = [(e.clientX - cx) / scale, (e.clientY - cy) / scale];
        let x = dx * Math.cos(t) + dy * Math.sin(t) + box.width / 2;
        let y = -dx * Math.sin(t) + dy * Math.cos(t) + box.height / 2;
        if (f.flipH) x = box.width - x;
        if (f.flipV) y = box.height - y;
        return [x, y];
      };
      const reshape = (e: MouseEvent) => {
        const [x, y] = point(e);
        adj = adjustShape({ ...style, adj }, index, x, y, box.width, box.height);
        const next = { ...style, adj };
        el.querySelector("svg.pw-shape")?.replaceWith(shapeSvg(next, box.width, box.height, f));
        const at = shapeHandles(next, box.width, box.height)[index];
        if (at) Object.assign(target.style, { left: `${f.flipH ? box.width - at.x : at.x}pt`, top: `${f.flipV ? box.height - at.y : at.y}pt` });
      };
      const done = () => {
        window.removeEventListener("mousemove", reshape);
        window.removeEventListener("mouseup", done);
        if (JSON.stringify(adj) !== JSON.stringify(style.adj ?? {})) setShapeAdjust(adj)(view.state, view.dispatch);
      };
      window.addEventListener("mousemove", reshape);
      window.addEventListener("mouseup", done);
      return true;
    }
    let next = box;
    const a = (turned * Math.PI) / 180;
    const [cos, sin] = [Math.cos(a), Math.sin(a)];
    const move = (e: MouseEvent) => {
      const dx = (e.clientX - start.x) / scale;
      const dy = (e.clientY - start.y) / scale;
      if (!handle) {
        next = { ...box, x: box.x + dx, y: box.y + dy };
        el.style.transform = `translate(${dx * scale}px, ${dy * scale}px)`;
        return;
      }
      // Resizing: the mouse's move along the object's own (turned) sides; the opposite corner stays.
      const lx = dx * cos + dy * sin;
      const ly = -dx * sin + dy * cos;
      const west = handle.includes("w");
      const north = handle.includes("n");
      const least = line ? 0.01 : 12;
      let width = Math.max(least, box.width + (west ? -lx : lx));
      let height = Math.max(least, box.height + (north ? -ly : ly));
      // Pictures keep their proportions; shapes too with Shift (a circle stays a circle).
      if ((picture || e.shiftKey) && !line) {
        const k = Math.max(width / box.width, height / box.height);
        width = box.width * k;
        height = box.height * k;
      }
      const [dw, dh] = [width - box.width, height - box.height];
      const [ox, oy] = [(west ? -dw : dw) / 2, (north ? -dh : dh) / 2];
      const [cx, cy] = [box.x + box.width / 2 + ox * cos - oy * sin, box.y + box.height / 2 + ox * sin + oy * cos];
      next = { x: cx - width / 2, y: cy - height / 2, width, height };
      el.style.transform = `translate(${(next.x - box.x) * scale}px, ${(next.y - box.y) * scale}px)`;
      el.style.width = `${width}pt`;
      el.style.height = `${height}pt`;
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      const changed = Math.abs(next.x - box.x) + Math.abs(next.y - box.y) + Math.abs(next.width - box.width) + Math.abs(next.height - box.height) > 0.5;
      const at = locate(view, el);
      if (changed && at) setFloatBox(at.paraPos, at.index, next)(view.state, view.dispatch);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    return true;
  }
  /** The selected object shown as selected (outline, handles), also after it is drawn again. */
  const mark = (view: EditorView) => {
    const sel = view.state.selection;
    const el = sel instanceof NodeSelection && sel.node.attrs.float ? floatElementAt(view, sel.from) : undefined;
    if (el !== selected) select(el);
  };
  return new Plugin({
    // The page margins are outside the editor: a click there, on a picture behind the text, selects it too.
    view: (editorView) => {
      const onPage = (e: MouseEvent) => {
        const t = e.target as HTMLElement;
        if (e.button !== 0 || editorView.dom.contains(t) || !t.closest?.(".pw-flow, .pw-sheets, .canvas")) return;
        if (behindAt(editorView, e.clientX, e.clientY) && !overText(e.clientX, e.clientY)) onMouseDown(editorView, e);
      };
      document.addEventListener("mousedown", onPage);
      return {
        update: (view) => mark(view),
        destroy: () => document.removeEventListener("mousedown", onPage),
      };
    },
    props: {
      handleDOMEvents: {
        mousedown: (view, event) => onMouseDown(view, event),
      },
      handleKeyDown(view, event) {
        // The selected object as drawn now (it may have been drawn again since it was selected).
        mark(view);
        if (!selected || !selected.isConnected) { selected = undefined; return false; }
        if (event.key === "Escape") { select(undefined); return true; }
        if (event.key === "Delete" || event.key === "Backspace") {
          const at = locate(view, selected);
          select(undefined);
          return at ? deleteFloat(at.paraPos, at.index)(view.state, view.dispatch) : false;
        }
        // Typing with a shape or text box selected types in it (as in Word), rather than replacing it.
        const kind = selected.dataset.kind;
        if ((kind === "shape" || kind === "textbox") && selected.dataset.line !== "1" && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
          const sel = view.state.selection;
          if (sel instanceof NodeSelection) window.dispatchEvent(new CustomEvent(SHAPE_TYPING, { detail: { pos: sel.from, text: event.key } }));
          return true;
        }
        // Arrow keys nudge it (Ctrl: a little).
        const step = event.ctrlKey || event.metaKey ? 1 : 6;
        const by = ({ ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] } as Record<string, number[]>)[event.key];
        if (by && !event.shiftKey && !event.altKey) {
          const at = locate(view, selected);
          const d = selected.dataset;
          // From where it is now (as stored: the screen may not be drawn again yet after the last nudge).
          const node = at && floatAt(view.state.doc, at.paraPos, at.index)?.node;
          const f = node?.attrs.float as FloatInfo | undefined;
          const stored = f && f.x.from === "column" && f.y.from === "paragraph" && f.x.offset !== undefined && f.y.offset !== undefined;
          const box = stored ? { x: f.x.offset!, y: f.y.offset!, width: f.width, height: f.height } : { x: Number(d.x), y: Number(d.y), width: Number(d.w), height: Number(d.h) };
          return at ? setFloatBox(at.paraPos, at.index, { ...box, x: box.x + by[0], y: box.y + by[1] })(view.state, view.dispatch) : false;
        }
        return false;
      },
    },
  });
}

// --- pictures: wrapping, replacing, alt text ---------------------------------------------------

/** How a picture sits with the text (Word's Wrap Text). */
export type PictureWrap = "inline" | "square" | "tight" | "topAndBottom" | "behind" | "inFront";

/** The selected picture (an ordinary picture, or a floating object kept from a Word file), if any. */
export function pictureAt(state: EditorState): { pos: number; node: Node; kept: boolean } | undefined {
  const sel = state.selection;
  if (!(sel instanceof NodeSelection)) return undefined;
  const node = sel.node;
  if (node.type.name === "image") return { pos: sel.from, node, kept: false };
  if (node.type.name === "run_preserved" && node.attrs.float) return { pos: sel.from, node, kept: true };
  return undefined;
}

/** A picture's current wrap. */
export function pictureWrap(node: Node): PictureWrap {
  const f = node.attrs.float as FloatInfo | null;
  if (!f) return "inline";
  if (f.wrap === "none") return f.behind ? "behind" : "inFront";
  return f.wrap === "through" ? "tight" : f.wrap;
}

const WRAP_ELEMENT = /<wp:wrap(None|Square|Tight|Through|TopAndBottom)\b[^>]*?(\/>|>[\s\S]*?<\/wp:wrap\1>)/;

/** Wrap XML for a kept floating object (its wp:wrap* element and behindDoc flag). */
function withWrap(xml: string, wrap: PictureWrap): string {
  const polygon = '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>';
  const element = wrap === "square" ? '<wp:wrapSquare wrapText="bothSides"/>'
    : wrap === "tight" ? `<wp:wrapTight wrapText="bothSides">${polygon}</wp:wrapTight>`
    : wrap === "topAndBottom" ? "<wp:wrapTopAndBottom/>" : "<wp:wrapNone/>";
  return xml.replace(WRAP_ELEMENT, element).replace(/(<wp:anchor\b[^>]*?)behindDoc="[01]"/, `$1behindDoc="${wrap === "behind" ? 1 : 0}"`);
}

/**
 * Changes how a picture sits with the text. In line → floating: it floats where it was (at the left of
 * its column, just below its paragraph's first line), and can then be dragged anywhere. Floating → in
 * line works for pictures of our own; kept objects from Word files stay floating.
 */
export function setPictureWrap(wrap: PictureWrap): Command {
  return (state, dispatch) => {
    const pic = pictureAt(state);
    if (!pic || (pic.kept && wrap === "inline")) return false;
    if (!dispatch) return true;
    const old = pic.node.attrs.float as FloatInfo | null;
    const float: FloatInfo | null = wrap === "inline" ? null : {
      kind: "picture",
      width: pic.node.attrs.width ?? old?.width ?? 0,
      height: pic.node.attrs.height ?? old?.height ?? 0,
      x: old?.x ?? { from: "column", offset: 0 },
      y: old?.y ?? { from: "paragraph", offset: 0 },
      wrap: wrap === "behind" || wrap === "inFront" ? "none" : wrap,
      side: "bothSides",
      ...(wrap === "behind" ? { behind: true } : {}),
      // Word's default distance from the text: 0.125" at the sides.
      dist: old?.dist ?? { top: 0, bottom: 0, left: 9, right: 9 },
      ...(old?.imagePart ? { imagePart: old.imagePart } : {}),
    };
    const attrs = pic.kept
      ? { ...pic.node.attrs, float: { ...old, ...float }, preserved: { ...pic.node.attrs.preserved, xml: withWrap(pic.node.attrs.preserved.xml, wrap) } }
      : { ...pic.node.attrs, float };
    const tr = state.tr.setNodeMarkup(pic.pos, undefined, attrs);
    dispatch(tr.setSelection(NodeSelection.create(tr.doc, pic.pos)));
    return true;
  };
}

/**
 * Replaces a picture's image (after cropping or rotating it): a new image resource, at the given size
 * in pt. The old image stays in the document only if something else still uses it.
 */
export function replacePicture(base: () => PrimeDocument, image: { data: string; contentType: string; width: number; height: number }): Command {
  return (state, dispatch) => {
    const pic = pictureAt(state);
    if (!pic || pic.kept) return false;
    if (!dispatch) return true;
    const doc = base();
    const ext = image.contentType === "image/jpeg" ? "jpeg" : "png";
    let n = 1;
    while (doc.images[`word/media/prime-edit${n}.${ext}`]) n++;
    const imageId = `word/media/prime-edit${n}.${ext}`;
    doc.images[imageId] = { contentType: image.contentType, extension: ext, data: image.data };
    const float = pic.node.attrs.float as FloatInfo | null;
    const tr = state.tr.setNodeMarkup(pic.pos, undefined, {
      ...pic.node.attrs, imageId, width: image.width, height: image.height,
      float: float ? { ...float, width: image.width, height: image.height } : null,
    });
    dispatch(tr.setSelection(NodeSelection.create(tr.doc, pic.pos)));
    return true;
  };
}

/** Sets a picture's alternative text (read by screen readers; Word's Alt Text). */
export function setAltText(text: string): Command {
  return (state, dispatch) => {
    const pic = pictureAt(state);
    if (!pic || pic.kept) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup(pic.pos, undefined, { ...pic.node.attrs, description: text || null }));
    return true;
  };
}

/**
 * Resize handles on a selected picture in line with the text (Word's eight): a corner keeps the
 * picture's proportions, a side stretches it one way. The new size is set when the drag ends (one
 * undo step); the outline shows the size while dragging.
 */
export function inlinePictureResizing(): Plugin {
  return new Plugin({
    view(view) {
      const box = document.createElement("div");
      box.className = "pw-resize-box";
      box.style.display = "none";
      const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];
      for (const h of HANDLES) {
        const el = document.createElement("span");
        el.className = `pw-resize-handle h-${h}`;
        el.dataset.handle = h;
        box.append(el);
      }
      document.body.append(box);
      // A picture of ours, or one kept from Word as it was (its drawing XML, sized by wp:extent).
      const target = () => {
        const sel = view.state.selection;
        if (!(sel instanceof NodeSelection) || sel.node.attrs.float) return undefined;
        const kept = sel.node.type.name === "run_preserved" && /<wp:inline[ >][\s\S]*?<wp:extent /.test(sel.node.attrs.preserved?.xml ?? "");
        if (sel.node.type.name !== "image" && !kept) return undefined;
        const dom = view.nodeDOM(sel.from) as HTMLElement | null;
        const img = dom?.tagName === "IMG" ? dom : dom?.querySelector?.("img");
        return img ? { pos: sel.from, node: sel.node, img: img as HTMLImageElement, kept } : undefined;
      };
      const place = () => {
        const t = target();
        if (!t) { if (!box.dataset.dragging) box.style.display = "none"; return; }
        const r = t.img.getBoundingClientRect();
        if (!r.width) { box.style.display = "none"; return; }
        Object.assign(box.style, { display: "block", left: `${r.left}px`, top: `${r.top}px`, width: `${r.width}px`, height: `${r.height}px` });
      };
      const down = (e: PointerEvent) => {
        const h = (e.target as HTMLElement).dataset.handle;
        const t = target();
        if (!h || !t) return;
        e.preventDefault();
        e.stopPropagation();
        const start = t.img.getBoundingClientRect();
        const ptWidth = t.kept ? Number(/<wp:extent cx="(\d+)"/.exec(t.node.attrs.preserved.xml)?.[1] ?? 0) / 12700 : t.node.attrs.width;
        const px = start.width / (ptWidth || start.width); // screen pixels per point (zoom included)
        const ratio = start.width / start.height;
        const x0 = e.clientX;
        const y0 = e.clientY;
        let w = start.width;
        let hh = start.height;
        box.dataset.dragging = "1";
        const move = (ev: PointerEvent) => {
          const dx = (ev.clientX - x0) * (h.includes("w") ? -1 : 1);
          const dy = (ev.clientY - y0) * (h.includes("n") ? -1 : 1);
          w = start.width;
          hh = start.height;
          if (h.length === 2) {
            // A corner: proportions kept, by whichever way the pointer went further.
            const k = Math.max((start.width + dx) / start.width, (start.height + dy) / start.height);
            w = start.width * k;
            hh = w / ratio;
          } else if (h === "e" || h === "w") w = start.width + dx;
          else hh = start.height + dy;
          w = Math.max(8, w);
          hh = Math.max(8, hh);
          Object.assign(box.style, {
            width: `${w}px`, height: `${hh}px`,
            left: `${h.includes("w") ? start.right - w : start.left}px`, top: `${h.includes("n") ? start.bottom - hh : start.top}px`,
          });
        };
        const up = () => {
          window.removeEventListener("pointermove", move);
          window.removeEventListener("pointerup", up);
          delete box.dataset.dragging;
          const now = target();
          if (now && (Math.abs(w - start.width) > 1 || Math.abs(hh - start.height) > 1)) {
            const width = Math.round((w / px) * 10) / 10;
            const height = Math.round((hh / px) * 10) / 10;
            let attrs: Record<string, unknown> = { ...now.node.attrs, width, height };
            if (now.kept) {
              // A drawing kept from Word: its size in EMU, in the inline extent and the picture's own transform.
              const cx = Math.round(width * 12700);
              const cy = Math.round(height * 12700);
              const xml = (now.node.attrs.preserved.xml as string)
                .replace(/(<wp:extent )cx="\d+" cy="\d+"/, `$1cx="${cx}" cy="${cy}"`)
                .replace(/(<a:xfrm\b[^>]*>[\s\S]*?<a:ext )cx="\d+" cy="\d+"/, `$1cx="${cx}" cy="${cy}"`);
              attrs = { ...now.node.attrs, preserved: { ...now.node.attrs.preserved, xml } };
            }
            const tr = view.state.tr.setNodeMarkup(now.pos, undefined, attrs);
            view.dispatch(tr.setSelection(NodeSelection.create(tr.doc, now.pos)));
          }
          view.focus();
          requestAnimationFrame(place);
        };
        window.addEventListener("pointermove", move);
        window.addEventListener("pointerup", up);
      };
      box.addEventListener("pointerdown", down);
      const onScroll = () => requestAnimationFrame(place);
      window.addEventListener("scroll", onScroll, true);
      window.addEventListener("resize", onScroll);
      return {
        update: () => requestAnimationFrame(place),
        destroy() {
          window.removeEventListener("scroll", onScroll, true);
          window.removeEventListener("resize", onScroll);
          box.remove();
        },
      };
    },
  });
}
