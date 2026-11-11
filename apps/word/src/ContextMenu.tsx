import {
  addColumn, addRow, deleteColumn, deleteRow, deleteTable, inTable, linkAt, pasteAs, type PasteMode, pictureAt, selectedPages,
} from "@prime/editor";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useLayoutEffect, useRef, useState } from "react";

export interface ContextActions {
  onFont(): void;
  onParagraph(): void;
  onLink(): void;
  onComment(): void;
  onEditPicture(): void;
  onAltText(): void;
  onSaveBlock(): void;
  onNotice(message: string): void;
  /** Deletes the page(s) the cursor or selection is on. */
  onDeletePage(): void;
}

/** Pastes from the system clipboard (formatted when it has HTML), then applies a paste option. */
async function pasteFromClipboard(view: EditorView, mode: PasteMode, notice: (m: string) => void) {
  try {
    let html = "";
    let text = "";
    if (navigator.clipboard.read) {
      for (const item of await navigator.clipboard.read()) {
        if (item.types.includes("text/html")) html = await (await item.getType("text/html")).text();
        if (item.types.includes("text/plain")) text = await (await item.getType("text/plain")).text();
      }
    } else text = await navigator.clipboard.readText();
    view.focus();
    if (html && mode !== "text") view.pasteHTML(html);
    else if (text) view.pasteText(text);
    else return;
    if (mode !== "keep") pasteAs(mode)(view.state, view.dispatch);
  } catch {
    notice("Prime Word could not read the clipboard here: press Ctrl+V to paste");
  }
}

/** Removes the link around the cursor (or across the selection). */
const removeLink: Command = (state, dispatch) => {
  const link = state.schema.marks.link;
  let { from, to } = state.selection;
  if (from === to) {
    // The whole link the cursor is in: the run of text nodes around it that carry the same link.
    const $p = state.selection.$from;
    const mark = $p.marks().find((m) => m.type === link) ?? $p.nodeBefore?.marks.find((m) => m.type === link);
    if (!mark) return false;
    // Stretches of text with this link, in order; the one containing the cursor.
    const runs: [number, number][] = [];
    $p.parent.forEach((n, offset) => {
      if (!mark.isInSet(n.marks)) return;
      const a = $p.start() + offset;
      const last = runs[runs.length - 1];
      if (last && last[1] === a) last[1] = a + n.nodeSize;
      else runs.push([a, a + n.nodeSize]);
    });
    const hit = runs.find(([a, b]) => a <= $p.pos && $p.pos <= b);
    if (!hit) return false;
    [from, to] = hit;
  }
  if (dispatch) dispatch(state.tr.removeMark(from, to, link));
  return true;
};

type Item = { label: string; keys?: string; run(): void; disabled?: boolean } | "-";

/** The right-click menu: what fits where you clicked (text, a link, a table, a picture). */
export function ContextMenu({ view, x, y, actions, onClose }: { view: EditorView; x: number; y: number; actions: ContextActions; onClose(): void }) {
  const menu = useRef<HTMLDivElement>(null);
  const [at, setAt] = useState({ x, y });
  const state = view.state;
  const cmd = (c: Command) => () => { c(view.state, view.dispatch, view); view.focus(); };
  const has = !state.selection.empty;
  const picture = pictureAt(state);
  const link = linkAt(state);

  const items: Item[] = [
    { label: "Cut", keys: "Ctrl+X", disabled: !has, run: () => { view.focus(); document.execCommand("cut"); } },
    { label: "Copy", keys: "Ctrl+C", disabled: !has, run: () => { view.focus(); document.execCommand("copy"); } },
    { label: "Paste", keys: "Ctrl+V", run: () => void pasteFromClipboard(view, "keep", actions.onNotice) },
    { label: "Paste: Merge Formatting", run: () => void pasteFromClipboard(view, "merge", actions.onNotice) },
    { label: "Paste: Keep Text Only", run: () => void pasteFromClipboard(view, "text", actions.onNotice) },
    "-",
  ];
  if (picture) {
    items.push(
      { label: "Crop & Rotate…", run: actions.onEditPicture },
      { label: "Alt Text…", run: actions.onAltText },
      { label: "Delete Picture", run: () => { view.dispatch(view.state.tr.deleteSelection()); view.focus(); } },
      "-",
    );
  }
  if (inTable(state)) {
    items.push(
      { label: "Insert Row Above", run: cmd(addRow(-1)) }, { label: "Insert Row Below", run: cmd(addRow(1)) },
      { label: "Insert Column Left", run: cmd(addColumn(-1)) }, { label: "Insert Column Right", run: cmd(addColumn(1)) },
      { label: "Delete Row", run: cmd(deleteRow) }, { label: "Delete Column", run: cmd(deleteColumn) }, { label: "Delete Table", run: cmd(deleteTable) },
      "-",
    );
  }
  items.push(
    { label: "Font…", keys: "Ctrl+D", run: actions.onFont },
    { label: "Paragraph…", run: actions.onParagraph },
    "-",
    link
      ? { label: "Edit Link…", run: actions.onLink }
      : { label: "Link…", keys: "Ctrl+K", run: actions.onLink },
  );
  if (link) items.push({ label: "Remove Link", run: cmd(removeLink) });
  items.push(
    { label: "New Comment", disabled: !has, run: actions.onComment },
    { label: "Save Selection as Block…", disabled: !has, run: actions.onSaveBlock },
  );
  // Print Layout: the page(s) the cursor or selection is on.
  const pages = selectedPages(state);
  if (pages) {
    items.push("-", {
      label: pages.first === pages.last ? `Delete Page ${pages.first}` : `Delete Pages ${pages.first}–${pages.last}`,
      run: actions.onDeletePage,
    });
  }

  // Keep the menu on screen.
  useLayoutEffect(() => {
    const r = menu.current?.getBoundingClientRect();
    if (!r) return;
    setAt({ x: Math.max(4, Math.min(x, innerWidth - r.width - 4)), y: Math.max(4, Math.min(y, innerHeight - r.height - 4)) });
  }, [x, y]);
  useEffect(() => {
    const close = (e: Event) => { if (!menu.current?.contains(e.target as Node)) onClose(); };
    const key = (e: KeyboardEvent) => { if (e.key === "Escape") { onClose(); view.focus(); } };
    window.addEventListener("mousedown", close, true);
    window.addEventListener("keydown", key, true);
    window.addEventListener("blur", onClose);
    menu.current?.querySelector<HTMLElement>("button:not(:disabled)")?.focus();
    return () => { window.removeEventListener("mousedown", close, true); window.removeEventListener("keydown", key, true); window.removeEventListener("blur", onClose); };
  }, [onClose, view]);
  // Arrow keys move through the items.
  const move = (e: React.KeyboardEvent) => {
    if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
    e.preventDefault();
    const buttons = [...(menu.current?.querySelectorAll<HTMLButtonElement>("button:not(:disabled)") ?? [])];
    const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
    buttons[(i + (e.key === "ArrowDown" ? 1 : buttons.length - 1)) % buttons.length]?.focus();
  };

  return (
    <div ref={menu} className="context-menu" role="menu" style={{ left: at.x, top: at.y }} onKeyDown={move}>
      {items.map((it, i) => it === "-"
        ? <div key={i} className="context-sep" role="separator" />
        : (
          <button key={it.label} role="menuitem" disabled={it.disabled} onClick={() => { onClose(); it.run(); }}>
            <span>{it.label}</span>{it.keys && <kbd>{it.keys}</kbd>}
          </button>
        ))}
    </div>
  );
}
