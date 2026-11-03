import { setIndent, setTabStops } from "@prime/editor";
import { Resolver, type TabStop } from "@prime/layout";
import type { Paragraph, PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useEffect, useReducer, useRef, useState } from "react";

const CM = 72 / 2.54;
type Kind = TabStop["align"];
const KINDS: Kind[] = ["left", "center", "right", "decimal", "bar"];
const KIND_LABEL: Record<Kind, string> = { left: "Left tab", center: "Centre tab", right: "Right tab", decimal: "Decimal tab", bar: "Bar tab" };
/** Word's tab marks. */
const KIND_MARK: Record<Kind, string> = { left: "⌊", center: "⊥", right: "⌋", decimal: "⊥.", bar: "|" };
/** Dragging snaps to 1/8 cm, as Word's ruler does in centimetres. */
const snap = (pt: number) => Math.round(pt / (CM / 8)) * (CM / 8);

type Drag =
  | { what: "first" | "left" | "hanging" | "right"; startX: number; start: { left: number; right: number; firstLine: number } }
  | { what: "tab"; index: number; startX: number; pos: number; away: boolean };

/**
 * The horizontal ruler over the text column: margins, the paragraph's indents (first line ▽, hanging
 * △, left □, right △) and its tab stops. Drag the indent marks; click the ruler to add a tab stop of
 * the kind shown at the left (click that to change it); drag a tab stop to move it, or off the ruler
 * to remove it.
 */
export function Ruler({ view, doc, pageWidth, margins }: { view: EditorView; doc: PrimeDocument; pageWidth: number; margins: { left: number; right: number } }) {
  const bar = useRef<HTMLDivElement>(null);
  const [, redraw] = useReducer((n: number) => n + 1, 0);
  const [kind, setKind] = useState<Kind>("left");
  const [drag, setDrag] = useState<Drag>();
  const [preview, setPreview] = useState<{ left: number; right: number; firstLine: number; tabs?: TabStop[] }>();

  // Follow the page as the window scrolls sideways or resizes.
  useEffect(() => {
    window.addEventListener("resize", redraw);
    window.addEventListener("scroll", redraw, true);
    return () => { window.removeEventListener("resize", redraw); window.removeEventListener("scroll", redraw, true); };
  }, []);

  // The editor is rebuilt when fonts arrive: a scroll in between must not ask the old one.
  if (view.isDestroyed) return null;
  const para = view.state.selection.$from.parent;
  const props = para.type.name === "paragraph" ? para.attrs.props : {};
  const rp = new Resolver(doc).paragraph({ type: "paragraph", props, content: [] } as Paragraph, undefined, false);
  // The column the paragraph is in (the page's text column, or a table cell).
  const dom = view.domAtPos(view.state.selection.from).node;
  const el = (dom instanceof HTMLElement ? dom : dom.parentElement)?.closest("td, .ProseMirror") as HTMLElement | null;
  const column = (el ?? view.dom).getBoundingClientRect();
  const outer = bar.current?.getBoundingClientRect();
  const textWidth = pageWidth - margins.left - margins.right;
  const scale = view.dom.getBoundingClientRect().width / textWidth; // screen px per pt
  const origin = column.left - (outer?.left ?? 0); // where 0 (the column's left edge) is on the ruler
  const width = column.width / scale;
  const x = (pt: number) => origin + pt * scale;
  // The page's text column (margins and the numbered scale), whatever column the cursor is in.
  const pageOrigin = view.dom.getBoundingClientRect().left - (outer?.left ?? 0);
  const px = (pt: number) => pageOrigin + pt * scale;
  const ind = preview ?? { left: rp.left, right: rp.right, firstLine: rp.firstLine };
  const tabs = preview?.tabs ?? rp.tabs;

  const ptAt = (clientX: number) => (clientX - (outer?.left ?? 0) - origin) / scale;

  const begin = (what: Drag["what"], e: React.PointerEvent, index = -1) => {
    e.preventDefault();
    e.stopPropagation();
    (e.target as HTMLElement).setPointerCapture(e.pointerId);
    if (what === "tab") setDrag({ what, index, startX: e.clientX, pos: tabs[index].pos, away: false });
    else setDrag({ what, startX: e.clientX, start: { left: rp.left, right: rp.right, firstLine: rp.firstLine } });
  };
  const move = (e: React.PointerEvent) => {
    if (!drag) return;
    const d = (e.clientX - drag.startX) / scale;
    if (drag.what === "tab") {
      const away = Math.abs(e.clientY - (outer?.top ?? 0) - 12) > 30;
      const next = tabs.map((t, i) => (i === drag.index ? { ...t, pos: Math.max(0, snap(drag.pos + d)) } : t));
      setDrag({ ...drag, away });
      setPreview({ ...ind, tabs: away ? next.filter((_, i) => i !== drag.index) : next });
      return;
    }
    const s = drag.start;
    const firstAbs = s.left + s.firstLine; // where the first line starts
    if (drag.what === "first") setPreview({ ...s, firstLine: snap(firstAbs + d) - s.left });
    else if (drag.what === "hanging") {
      const left = snap(s.left + d); // the other lines move; the first line stays where it is
      setPreview({ ...s, left, firstLine: firstAbs - left });
    } else if (drag.what === "left") setPreview({ ...s, left: snap(s.left + d) });
    else setPreview({ ...s, right: snap(s.right - d) });
  };
  const end = () => {
    if (drag && preview) {
      if (drag.what === "tab") setTabStops(() => doc, preview.tabs ?? [])(view.state, view.dispatch);
      else setIndent({ left: preview.left, right: preview.right, firstLine: preview.firstLine })(view.state, view.dispatch);
    }
    setDrag(undefined);
    setPreview(undefined);
    view.focus();
  };
  const addTab = (e: React.MouseEvent) => {
    if (drag) return;
    const pos = snap(ptAt(e.clientX));
    if (pos <= 0 || pos >= width) return;
    setTabStops(() => doc, [...rp.tabs, { pos, align: kind }].sort((a, b) => a.pos - b.pos))(view.state, view.dispatch);
    view.focus();
  };

  // Ticks every quarter centimetre, numbered every centimetre (from the column's left edge, both ways).
  const ticks: JSX.Element[] = [];
  const from = Math.ceil(-margins.left / (CM / 4));
  const to = Math.floor((pageWidth - margins.left) / (CM / 4));
  for (let i = from; i <= to; i++) {
    const at = px((i * CM) / 4);
    if (i % 4 === 0) { if (i !== 0) ticks.push(<span key={i} className="ruler-num" style={{ left: at }}>{Math.abs(i / 4)}</span>); }
    else ticks.push(<span key={i} className={`ruler-tick ${i % 2 ? "" : "half"}`} style={{ left: at }} />);
  }

  return (
    <div className="ruler" ref={bar} onPointerMove={move} onPointerUp={end} onPointerCancel={end} aria-label="Ruler">
      <button className="ruler-kind" title={`${KIND_LABEL[kind]}: click to change; click the ruler to add one`} aria-label={KIND_LABEL[kind]}
        onClick={() => setKind(KINDS[(KINDS.indexOf(kind) + 1) % KINDS.length])}>{KIND_MARK[kind]}</button>
      <div className="ruler-margin" style={{ left: px(-margins.left), width: margins.left * scale }} />
      <div className="ruler-text" style={{ left: px(0), width: textWidth * scale }} onClick={addTab} title="Click to add a tab stop" />
      <div className="ruler-margin" style={{ left: px(textWidth), width: margins.right * scale }} />
      {ticks}
      {tabs.map((t, i) => (
        <span key={`${t.pos}-${i}`} className={`ruler-tab ${drag?.what === "tab" && drag.index === i && drag.away ? "away" : ""}`} style={{ left: x(t.pos) }}
          title={`${KIND_LABEL[t.align]} at ${(t.pos / CM).toFixed(2)} cm — drag to move, drag off the ruler to remove`}
          onPointerDown={(e) => begin("tab", e, i)}>{KIND_MARK[t.align]}</span>
      ))}
      <span className="ruler-marker first" style={{ left: x(ind.left + ind.firstLine) }} title="First Line Indent" onPointerDown={(e) => begin("first", e)} />
      <span className="ruler-marker hanging" style={{ left: x(ind.left) }} title="Hanging Indent" onPointerDown={(e) => begin("hanging", e)} />
      <span className="ruler-marker left" style={{ left: x(ind.left) }} title="Left Indent" onPointerDown={(e) => begin("left", e)} />
      <span className="ruler-marker right" style={{ left: x(width - ind.right) }} title="Right Indent" onPointerDown={(e) => begin("right", e)} />
      {preview && drag?.what !== "tab" && <span className="ruler-guide" style={{ left: x(drag?.what === "first" ? ind.left + ind.firstLine : drag?.what === "right" ? width - ind.right : ind.left) }} />}
    </div>
  );
}
