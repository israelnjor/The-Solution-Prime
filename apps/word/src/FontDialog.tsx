import { FONT_SIZES, runAtSelection, runCss, updateRun } from "@prime/editor";
import { Resolver } from "@prime/layout";
import type { PrimeDocument, RunProps } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useMemo, useState } from "react";

const UNDERLINES: [string, string][] = [
  ["none", "(none)"], ["single", "Single"], ["words", "Words only"], ["double", "Double"], ["thick", "Thick"],
  ["dotted", "Dotted"], ["dash", "Dashed"], ["dotDash", "Dot dash"], ["wave", "Wave"],
];

/** "font-size:12pt;color:#f00" as a React style object. */
const cssObject = (css: string): React.CSSProperties =>
  Object.fromEntries(css.split(";").filter(Boolean).map((d) => {
    const i = d.indexOf(":");
    return [d.slice(0, i).trim().replace(/-([a-z])/g, (_, c: string) => c.toUpperCase()), d.slice(i + 1).trim()];
  }));

type Effect = "strike" | "doubleStrike" | "superscript" | "subscript" | "caps" | "hidden";
const EFFECTS: [Effect, string][] = [
  ["strike", "Strikethrough"], ["doubleStrike", "Double strikethrough"], ["superscript", "Superscript"],
  ["subscript", "Subscript"], ["caps", "All caps"], ["hidden", "Hidden"],
];

/**
 * Word's Font dialog (Ctrl+D): font, style, size, colour, underline, effects and character spacing,
 * with a preview. Only what is changed here is applied, so a selection with mixed fonts keeps them
 * unless the font is changed.
 */
export function FontDialog({ view, doc, fonts, onClose }: { view: EditorView; doc: PrimeDocument; fonts: string[]; onClose(): void }) {
  const state = view.state;
  const direct = runAtSelection(state);
  const styleId = state.selection.$from.parent.attrs.props?.style ?? "";
  const r = useMemo(() => new Resolver(doc).run(styleId, direct), [doc, styleId, direct]);
  const initial = {
    font: r.font, size: r.size, bold: r.bold, italic: r.italic, color: r.color ?? "000000",
    underline: direct.underline?.val ?? (r.underline ? "single" : "none"),
    strike: !!direct.strike, doubleStrike: !!direct.doubleStrike, superscript: r.vertAlign === "superscript",
    subscript: r.vertAlign === "subscript", caps: r.caps, hidden: r.hidden, spacing: direct.charSpacing ?? 0,
  };
  const [f, setF] = useState(initial);
  const [touched, setTouched] = useState(new Set<keyof typeof initial>());
  const set = <K extends keyof typeof initial>(k: K, v: (typeof initial)[K]) => {
    setF((s) => {
      const next = { ...s, [k]: v };
      // Superscript and subscript, and the two strikethroughs, exclude each other.
      if (k === "superscript" && v) next.subscript = false;
      if (k === "subscript" && v) next.superscript = false;
      if (k === "strike" && v) next.doubleStrike = false;
      if (k === "doubleStrike" && v) next.strike = false;
      return next;
    });
    setTouched((t) => new Set(t).add(k).add(k === "superscript" ? "subscript" : k === "subscript" ? "superscript" : k === "strike" ? "doubleStrike" : k === "doubleStrike" ? "strike" : k));
  };
  const style = f.bold ? (f.italic ? "Bold Italic" : "Bold") : f.italic ? "Italic" : "Regular";

  const apply = () => {
    const t = touched;
    updateRun((p) => {
      const n: RunProps = { ...p };
      if (t.has("font")) {
        const { asciiTheme: _a, hAnsiTheme: _h, csTheme: _c, ...rest } = n.font ?? {};
        n.font = { ...rest, ascii: f.font, hAnsi: f.font, cs: f.font };
      }
      if (t.has("size")) n.size = f.size;
      if (t.has("bold")) n.bold = f.bold;
      if (t.has("italic")) n.italic = f.italic;
      if (t.has("color")) n.color = f.color.toUpperCase();
      if (t.has("underline")) { if (f.underline === "none") n.underline = { val: "none" }; else n.underline = { val: f.underline }; }
      if (t.has("strike")) n.strike = f.strike || undefined;
      if (t.has("doubleStrike")) n.doubleStrike = f.doubleStrike || undefined;
      if (t.has("superscript") || t.has("subscript")) n.vertAlign = f.superscript ? "superscript" : f.subscript ? "subscript" : undefined;
      if (t.has("caps")) n.caps = f.caps || undefined;
      if (t.has("hidden")) n.hidden = f.hidden || undefined;
      if (t.has("spacing")) n.charSpacing = f.spacing || undefined;
      return n;
    })(view.state, view.dispatch, view);
    onClose();
    view.focus();
  };

  const sample = state.selection.empty ? f.font : state.doc.textBetween(state.selection.from, state.selection.to, " ").slice(0, 60) || f.font;
  const preview = runCss({
    font: f.font, size: Math.min(f.size, 36), bold: f.bold, italic: f.italic, color: f.color, underline: f.underline !== "none",
    underlineStyle: f.underline === "single" || f.underline === "words" ? undefined : f.underline,
    strike: f.strike || f.doubleStrike, caps: f.caps, charSpacing: f.spacing, vertAlign: f.superscript ? "superscript" : f.subscript ? "subscript" : undefined,
  });

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); if (e.key === "Enter" && (e.target as HTMLElement).tagName !== "BUTTON") apply(); }}>
      <div className="hf-dialog font-dialog" role="dialog" aria-label="Font">
        <h3>Font</h3>
        <div className="font-grid">
          <label>Font<input list="pw-font-list" value={f.font} onChange={(e) => set("font", e.target.value)} autoFocus />
            <datalist id="pw-font-list">{fonts.map((x) => <option key={x} value={x} />)}</datalist></label>
          <label>Font style<select value={style} onChange={(e) => { set("bold", e.target.value.includes("Bold")); set("italic", e.target.value.includes("Italic")); }}>
            {["Regular", "Italic", "Bold", "Bold Italic"].map((s) => <option key={s}>{s}</option>)}
          </select></label>
          <label>Size<input list="pw-size-list" type="number" min={1} max={1638} step={0.5} value={f.size} onChange={(e) => set("size", Number(e.target.value) || f.size)} />
            <datalist id="pw-size-list">{FONT_SIZES.map((s) => <option key={s} value={s} />)}</datalist></label>
          <label>Font colour<input type="color" value={`#${f.color.toLowerCase()}`} onChange={(e) => set("color", e.target.value.slice(1))} /></label>
          <label>Underline style<select value={f.underline} onChange={(e) => set("underline", e.target.value)}>
            {UNDERLINES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select></label>
        </div>
        <fieldset className="font-effects"><legend>Effects</legend>
          {EFFECTS.map(([k, l]) => (
            <label key={k} className="choice"><input type="checkbox" checked={f[k]} onChange={(e) => set(k, e.target.checked)} /> {l}</label>
          ))}
        </fieldset>
        <fieldset className="font-spacing"><legend>Character spacing</legend>
          <label>Spacing <select value={f.spacing > 0 ? "expanded" : f.spacing < 0 ? "condensed" : "normal"} onChange={(e) => {
            const v = e.target.value;
            set("spacing", v === "normal" ? 0 : v === "expanded" ? Math.abs(f.spacing) || 1 : -(Math.abs(f.spacing) || 1));
          }}>
            <option value="normal">Normal</option><option value="expanded">Expanded</option><option value="condensed">Condensed</option>
          </select></label>
          <label>By <input type="number" min={0} max={100} step={0.1} value={Math.abs(f.spacing)} disabled={!f.spacing}
            onChange={(e) => set("spacing", (f.spacing < 0 ? -1 : 1) * Math.abs(Number(e.target.value)))} /> pt</label>
        </fieldset>
        <div className="font-preview" aria-label="Preview"><span style={cssObject(preview)}>{sample}</span></div>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" onClick={apply}>OK</button>
        </div>
      </div>
    </div>
  );
}
