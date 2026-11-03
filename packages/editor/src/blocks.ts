/**
 * Moving blocks (paragraphs, tables): Alt+Shift+↑/↓ as in Word, and a drag handle beside the block
 * under the pointer. A block moves within its container (the document body, or a table cell).
 */
import type { Node, ResolvedPos } from "prosemirror-model";
import { type Command, NodeSelection, Plugin, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import * as pmView from "prosemirror-view";

/**
 * The depth of the block containing `$pos`: a child of the document or a content control; in a table,
 * its row (Word's Alt+Shift+↑/↓ move table rows).
 */
function blockDepth($pos: ResolvedPos, rows = true): number {
  if (rows) for (let d = $pos.depth; d > 0; d--) if ($pos.node(d).type.name === "table_row") return d;
  for (let d = $pos.depth; d > 0; d--) {
    const parent = $pos.node(d - 1).type.name;
    if (parent === "doc" || parent === "table_cell" || parent === "content_control") return d;
  }
  return 1;
}

/** Moves the block(s) the selection is in up (-1) or down (1), past the neighbouring block. */
export function moveBlock(dir: -1 | 1): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    const depth = blockDepth($from);
    const container = $from.node(depth - 1);
    const first = $from.index(depth - 1);
    // The selection's last block, if it ends in the same container.
    const last = $to.depth >= depth && $to.node(depth - 1) === container ? $to.index(depth - 1) : first;
    const neighbour = dir < 0 ? first - 1 : last + 1;
    if (neighbour < 0 || neighbour >= container.childCount) return false;
    if (dispatch) {
      const start = $from.before(depth);
      const end = $to.depth >= depth && $to.node(depth - 1) === container ? $to.after(depth) : $from.after(depth);
      const other = container.child(neighbour);
      const moving = state.doc.slice(start, end).content;
      const tr = state.tr;
      if (dir < 0) {
        // [other][moving] → [moving][other]
        const otherStart = start - other.nodeSize;
        tr.replaceWith(otherStart, end, moving.append(pmFragment(other)));
        const shift = -other.nodeSize;
        tr.setSelection(TextSelection.create(tr.doc, state.selection.from + shift, state.selection.to + shift));
      } else {
        // [moving][other] → [other][moving]
        const otherEnd = end + other.nodeSize;
        tr.replaceWith(start, otherEnd, pmFragment(other).append(moving));
        const shift = other.nodeSize;
        tr.setSelection(TextSelection.create(tr.doc, state.selection.from + shift, state.selection.to + shift));
      }
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

function pmFragment(node: Node) {
  return node.type.schema.nodes.doc.create(null, node).content;
}

// ProseMirror's clipboard serializer (exported, though not in this version's type definitions).
const serializeForClipboard = (pmView as unknown as {
  __serializeForClipboard(view: EditorView, slice: import("prosemirror-model").Slice): { dom: HTMLElement; text: string };
}).__serializeForClipboard;

/**
 * A drag handle (⠿) beside the block under the pointer. Dragging it moves the whole block — a
 * paragraph, a table, a paragraph with a picture — to wherever it is dropped.
 */
export function blockHandle(): Plugin {
  return new Plugin({
    view(view) {
      const handle = document.createElement("div");
      handle.className = "pw-block-handle";
      handle.textContent = "⠿";
      handle.draggable = true;
      handle.title = "Drag to move this paragraph (Alt+Shift+↑/↓ also move it)";
      handle.contentEditable = "false";
      let target: { pos: number; dom: HTMLElement } | undefined;
      const host = () => view.dom.parentElement;
      const place = (dom: HTMLElement) => {
        const parent = host();
        if (!parent) return;
        if (handle.parentElement !== parent) parent.appendChild(handle);
        const a = dom.getBoundingClientRect();
        const b = parent.getBoundingClientRect();
        const zoom = parent.offsetWidth ? b.width / parent.offsetWidth : 1;
        handle.style.top = `${(a.top - b.top) / zoom}px`;
        handle.style.left = `${(a.left - b.left) / zoom - 22}px`;
        handle.style.display = "block";
      };
      /** The top-level block (or block in a table cell) under the pointer. */
      const blockAt = (x: number, y: number) => {
        const hit = view.posAtCoords({ left: x, top: y });
        if (!hit) return undefined;
        // The position in the text under the pointer (hit.inside is the start of a node: one level up).
        const $pos = view.state.doc.resolve(hit.pos);
        if ($pos.depth === 0) {
          const node = view.state.doc.nodeAt(hit.pos);
          const dom = node && (view.nodeDOM(hit.pos) as HTMLElement | null);
          return dom instanceof HTMLElement ? { pos: hit.pos, dom } : undefined;
        }
        // The handle moves whole blocks: a table, not one of its rows.
        let depth = blockDepth($pos, false);
        while (depth > 1 && $pos.node(depth - 1).type.name !== "doc" && $pos.node(depth - 1).type.name !== "content_control") depth--;
        const pos = $pos.before(depth);
        const dom = view.nodeDOM(pos);
        return dom instanceof HTMLElement ? { pos, dom } : undefined;
      };
      const onMove = (e: MouseEvent) => {
        if (e.buttons) return;
        const found = blockAt(Math.min(e.clientX, view.dom.getBoundingClientRect().right - 4), e.clientY);
        if (!found) return;
        target = found;
        place(found.dom);
      };
      const onLeave = (e: MouseEvent) => {
        if (e.relatedTarget === handle) return;
        handle.style.display = "none";
      };
      handle.addEventListener("dragstart", (e) => {
        if (!target || !e.dataTransfer) return;
        const sel = NodeSelection.create(view.state.doc, target.pos);
        view.dispatch(view.state.tr.setSelection(sel));
        const slice = sel.content();
        const { dom, text } = serializeForClipboard(view, slice);
        e.dataTransfer.clearData();
        e.dataTransfer.setData("text/html", dom.innerHTML);
        e.dataTransfer.setData("text/plain", text);
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setDragImage(target.dom, 0, 0);
        // ProseMirror moves (not copies) what it is dragging when this is set.
        (view as unknown as { dragging: unknown }).dragging = { slice, move: true };
      });
      handle.addEventListener("dragend", () => { handle.style.display = "none"; });
      view.dom.addEventListener("mousemove", onMove);
      view.dom.addEventListener("mouseleave", onLeave);
      return {
        destroy() {
          view.dom.removeEventListener("mousemove", onMove);
          view.dom.removeEventListener("mouseleave", onLeave);
          handle.remove();
        },
      };
    },
  });
}
