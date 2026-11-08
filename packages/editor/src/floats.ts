/**
 * Moving and resizing floating pictures and text boxes. The object stays the XML Word wrote; only its
 * position (wp:positionH / wp:positionV) and size (wp:extent and the shape's a:ext) are rewritten.
 */
import type { FloatInfo, PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { type Command, type EditorState, NodeSelection, Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

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
        dispatch(state.tr.setNodeMarkup(target.pos, undefined, {
          ...target.node.attrs, width: box.width, height: box.height,
          float: { ...float, x: { from: "column", offset: box.x }, y: { from: "paragraph", offset: box.y } },
        }));
        return true;
      }
      // A text box that grows with its text keeps doing so: only its width is set.
      const keepHeight = float.kind === "textbox" && (float.autoHeight || float.height < 1);
      const preserved = target.node.attrs.preserved as { name: string; xml: string };
      dispatch(state.tr.setNodeMarkup(target.pos, undefined, {
        ...target.node.attrs,
        preserved: { ...preserved, xml: withBox(preserved.xml, box, keepHeight) },
        float: {
          ...float, x: { from: "column", offset: box.x }, y: { from: "paragraph", offset: box.y },
          width: box.width, ...(keepHeight ? {} : { height: box.height }),
        },
      }));
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
  return new Plugin({
    props: {
      handleDOMEvents: {
        mousedown(view, event) {
          const target = event.target as HTMLElement;
          const el = target.closest?.(".pw-float") as HTMLElement | null;
          if (!el) { select(undefined); return false; }
          event.preventDefault();
          select(el);
          // Also select it in the document, so the ribbon shows the picture's options.
          const where = locate(view, el);
          const node = where && floatAt(view.state.doc, where.paraPos, where.index);
          if (node) view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, node.pos)));
          const handle = target.dataset.handle; // "nw" | "ne" | "sw" | "se" when resizing
          const start = { x: event.clientX, y: event.clientY };
          const box: FloatBox = { x: Number(el.dataset.x), y: Number(el.dataset.y), width: Number(el.dataset.w), height: Number(el.dataset.h) };
          // Screen pixels per point, including any zoom.
          const scale = el.getBoundingClientRect().width / box.width || 96 / 72;
          const picture = el.dataset.kind === "picture";
          let next = box;
          const move = (e: MouseEvent) => {
            const dx = (e.clientX - start.x) / scale;
            const dy = (e.clientY - start.y) / scale;
            if (!handle) {
              next = { ...box, x: box.x + dx, y: box.y + dy };
              el.style.transform = `translate(${dx * scale}px, ${dy * scale}px)`;
              return;
            }
            const west = handle.includes("w");
            const north = handle.includes("n");
            let width = Math.max(12, box.width + (west ? -dx : dx));
            let height = Math.max(12, box.height + (north ? -dy : dy));
            if (picture) {
              const k = Math.max(width / box.width, height / box.height);
              width = box.width * k;
              height = box.height * k;
            }
            next = { x: west ? box.x + box.width - width : box.x, y: north ? box.y + box.height - height : box.y, width, height };
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
        },
      },
      handleKeyDown(view, event) {
        if (!selected || !selected.isConnected) { selected = undefined; return false; }
        if (event.key === "Escape") { select(undefined); return true; }
        if (event.key === "Delete" || event.key === "Backspace") {
          const at = locate(view, selected);
          select(undefined);
          return at ? deleteFloat(at.paraPos, at.index)(view.state, view.dispatch) : false;
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
