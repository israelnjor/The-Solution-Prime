import { findNext, findState, moveSection, outline, outlineIndexAt, type OutlineItem, setFindQuery } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";

/** Scrolls the document so `pos` is near the top of the window (a heading) or in its middle (a search result). */
export function scrollToPos(view: EditorView, pos: number, align: "top" | "center") {
  const canvas = view.dom.closest(".canvas") as HTMLElement | null;
  if (!canvas) return;
  let coords: { top: number };
  try { coords = view.coordsAtPos(pos); } catch { return; }
  const box = canvas.getBoundingClientRect();
  canvas.scrollTop += coords.top - box.top - (align === "top" ? 24 : box.height / 2);
}

const NONE: { from: number; to: number }[] = [];

/** A result's line: the words around the match, the match itself bold. */
function snippet(doc: Node, from: number, to: number): [string, string, string] {
  const $from = doc.resolve(from);
  const start = $from.start();
  const end = $from.end();
  const before = doc.textBetween(Math.max(start, from - 40), from, " ", " ");
  const after = doc.textBetween(to, Math.min(end, to + 60), " ", " ");
  return [(from - 40 > start ? "…" : "") + before.replace(/^\S*\s/, (m) => (from - 40 > start ? "" : m)), doc.textBetween(from, to), after];
}

/**
 * View → Navigation Pane (Ctrl+F): the document's headings, as an outline to jump around and to
 * reorder sections by dragging; and a search box listing every match in the document.
 */
export function NavPane({ view, base, doc, searchFocus, onAdvanced, onClose }: {
  view: EditorView; base: () => PrimeDocument; doc: Node; searchFocus: number; onAdvanced(): void; onClose(): void;
}) {
  const [tab, setTab] = useState<"headings" | "results">("headings");
  const [query, setQuery] = useState("");
  const [items, setItems] = useState<OutlineItem[]>(() => outline(view.state, base()));
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());
  const [drop, setDrop] = useState<{ index: number; after: boolean } | null>(null);
  const dragging = useRef(-1);
  const search = useRef<HTMLInputElement>(null);
  const pane = useRef<HTMLElement>(null);
  // Level with the document area's top (the ribbon's height differs from tab to tab): measured when the
  // document area changes size, not on every render (measuring then lays out the whole document).
  useLayoutEffect(() => {
    const canvas = view.dom.closest(".canvas") as HTMLElement | null;
    if (!canvas) return;
    const place = () => { if (pane.current) pane.current.style.top = `${Math.round(canvas.getBoundingClientRect().top)}px`; };
    place();
    const watch = new ResizeObserver(place);
    watch.observe(canvas);
    return () => watch.disconnect();
  }, [view]);

  // The outline follows the document (a moment after typing stops in a long one).
  useEffect(() => {
    const t = setTimeout(() => setItems(outline(view.state, base())), doc.content.size > 300_000 ? 600 : 150);
    return () => clearTimeout(t);
  }, [doc, view, base]);
  useEffect(() => {
    if (!searchFocus) return;
    setTab("results");
    search.current?.focus();
    search.current?.select();
  }, [searchFocus]);
  // Search: matches highlighted in the document while the pane is open.
  // (Not on an editor already replaced: it is rebuilt when fonts arrive.)
  useEffect(() => {
    if (!view.isDestroyed) setFindQuery({ text: query })(view.state, view.dispatch);
  }, [query, view]);
  useEffect(() => () => { if (!view.isDestroyed) setFindQuery({ text: "" })(view.state, view.dispatch); }, [view]);

  const fs = findState(view.state);
  const matches = query ? fs?.matches ?? NONE : NONE;
  const current = outlineIndexAt(items, view.state.selection.from);
  // Headings whose section holds a match (marked while searching, as Word does).
  const withResults = useMemo(() => {
    const s = new Set<number>();
    for (const m of matches) { const i = outlineIndexAt(items, m.from); if (i >= 0) s.add(i); }
    return s;
  }, [matches, items]);

  const key = (h: OutlineItem) => `${h.level}:${h.text}`;
  // Rows shown: those not under a collapsed heading.
  const rows = useMemo(() => {
    const out: { item: OutlineItem; index: number; parent: boolean }[] = [];
    let hideBelow = Infinity;
    items.forEach((item, index) => {
      if (item.level <= hideBelow) hideBelow = Infinity;
      const parent = items[index + 1]?.level > item.level;
      if (hideBelow === Infinity) {
        out.push({ item, index, parent });
        if (parent && collapsed.has(key(item))) hideBelow = item.level;
      }
    });
    return out;
  }, [items, collapsed]);

  const go = (index: number) => {
    // Positions from the latest document (the outline may be a moment behind).
    const fresh = outline(view.state, base());
    const h = fresh[index] ?? items[index];
    if (!h) return;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, h.pos + 1)));
    view.focus();
    requestAnimationFrame(() => scrollToPos(view, h.pos + 1, "top"));
  };
  const goMatch = (from: number, to: number) => {
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
    requestAnimationFrame(() => scrollToPos(view, from, "center"));
  };
  const step = (dir: 1 | -1) => {
    if (findNext(dir)(view.state, view.dispatch)) requestAnimationFrame(() => scrollToPos(view, view.state.selection.from, "center"));
  };
  const toggle = (h: OutlineItem) => setCollapsed((old) => {
    const next = new Set(old);
    if (next.has(key(h))) next.delete(key(h)); else next.add(key(h));
    return next;
  });
  const dropOn = () => {
    const from = dragging.current;
    dragging.current = -1;
    const d = drop;
    setDrop(null);
    if (from < 0 || !d) return;
    const fresh = outline(view.state, base());
    const target = d.after ? fresh[d.index]?.end : fresh[d.index]?.pos;
    if (target === undefined) return;
    if (moveSection(fresh, from, target)(view.state, view.dispatch)) {
      setItems(outline(view.state, base()));
      requestAnimationFrame(() => scrollToPos(view, view.state.selection.from, "top"));
    }
  };

  // The outline is drawn again only when it, the current heading, the search or a drop marker changes (not on every keystroke).
  const outlineView = useMemo(() => (
        items.length ? (
          <ul className="nav-outline" onDragLeave={(e) => { if (e.currentTarget === e.target) setDrop(null); }}>
            {rows.map(({ item, index, parent }) => (
              <li key={`${index}-${item.pos}`} style={{ paddingLeft: `${(item.level - 1) * 14 + 4}px` }}
                className={[index === current ? "current" : "", withResults.has(index) ? "has-results" : "",
                  drop?.index === index ? (drop.after ? "drop-after" : "drop-before") : ""].join(" ")}
                draggable onDragStart={(e) => { dragging.current = index; e.dataTransfer.effectAllowed = "move"; e.dataTransfer.setData("text/plain", item.text); }}
                onDragOver={(e) => {
                  if (dragging.current < 0) return;
                  e.preventDefault();
                  const r = e.currentTarget.getBoundingClientRect();
                  const after = e.clientY > r.top + r.height / 2;
                  if (drop?.index !== index || drop.after !== after) setDrop({ index, after });
                }}
                onDrop={(e) => { e.preventDefault(); dropOn(); }}
                onDragEnd={() => { dragging.current = -1; setDrop(null); }}>
                <button className="twisty" aria-label={collapsed.has(key(item)) ? "Expand" : "Collapse"} disabled={!parent}
                  onClick={() => toggle(item)}>{parent ? (collapsed.has(key(item)) ? "▸" : "▾") : ""}</button>
                <button className="heading" onClick={() => go(index)} title={item.text}>{item.text}</button>
              </li>
            ))}
          </ul>
        ) : (
          <p className="nav-empty">This document has no headings. Use the Heading styles (Home → Styles) for the titles of sections, and they appear here.</p>
        )
  ), [view, items, rows, current, withResults, drop, collapsed]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <aside ref={pane} className="nav-pane" aria-label="Navigation">
      <header>
        <strong>Navigation</strong>
        <button className="link" onClick={onClose} title="Close the navigation pane" aria-label="Close navigation">✕</button>
      </header>
      <div className="nav-search">
        <input ref={search} value={query} placeholder="Search document" aria-label="Search document"
          onChange={(e) => { setQuery(e.target.value); setTab("results"); }}
          onKeyDown={(e) => {
            if (e.key === "Enter") { e.preventDefault(); step(e.shiftKey ? -1 : 1); }
            if (e.key === "Escape") { setQuery(""); view.focus(); }
          }} />
        <button onClick={() => step(-1)} disabled={!matches.length} title="Previous result (Shift+Enter)" aria-label="Previous result">▲</button>
        <button onClick={() => step(1)} disabled={!matches.length} title="Next result (Enter)" aria-label="Next result">▼</button>
      </div>
      <nav className="nav-tabs">
        <button className={tab === "headings" ? "on" : ""} onClick={() => setTab("headings")}>Headings</button>
        <button className={tab === "results" ? "on" : ""} onClick={() => setTab("results")}>Results</button>
        <button className="link" onClick={onAdvanced} title="Find with options (match case, whole words) and replace">Advanced…</button>
      </nav>
      {tab === "headings" && outlineView}
      {tab === "results" && (
        <div className="nav-results">
          <p className="nav-count">{query ? `${matches.length} result${matches.length === 1 ? "" : "s"}` : "Type in the box above to search the document."}</p>
          <ul>
            {matches.slice(0, 300).map((m) => {
              const [before, hit, after] = snippet(view.state.doc, m.from, m.to);
              return (
                <li key={m.from} className={view.state.selection.from === m.from ? "current" : ""}>
                  <button onClick={() => goMatch(m.from, m.to)}>{before}<strong>{hit}</strong>{after}</button>
                </li>
              );
            })}
          </ul>
          {matches.length > 300 && <p className="nav-count">Showing the first 300. Use ▼ to go through the rest.</p>}
        </div>
      )}
    </aside>
  );
}
