import { createStyle, modifyStyle, type StyleChange, stylesOf } from "@prime/editor";
import { Resolver } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useMemo, useState } from "react";

type Form = Required<Omit<StyleChange, "basedOn" | "next">> & { basedOn: string; next: string };

const SPACINGS: [string, number][] = [["Single", 1], ["1.15", 1.15], ["1.5 lines", 1.5], ["Double", 2]];

/**
 * Home → Styles: modify a paragraph style (Word's Modify Style) — every paragraph of the style
 * changes — or create one from the paragraph at the cursor (Word's Create a Style).
 */
export function StylesDialog({ view, base, fonts, onClose }: { view: EditorView; base: PrimeDocument; fonts: string[]; onClose(): void }) {
  const sheet = stylesOf(view.state, base);
  const doc = useMemo(() => ({ ...base, styles: sheet }), [base, sheet]);
  const paragraphStyles = sheet.styles.filter((s) => s.type === "paragraph" && !s.name?.toLowerCase().startsWith("toc"));
  const label = (id: string) => sheet.styles.find((s) => s.id === id)?.name ?? id;
  const atCursor = (view.state.selection.$from.parent.attrs.props?.style as string | undefined)
    ?? paragraphStyles.find((s) => s.isDefault)?.id ?? paragraphStyles[0]?.id ?? "";
  const [id, setId] = useState(atCursor);
  const [newName, setNewName] = useState("");
  const [error, setError] = useState("");

  // The style's values as they apply (with what it inherits), to start the form from.
  const current = useMemo((): Form => {
    const rp = new Resolver(doc).paragraph({ type: "paragraph", props: { style: id }, content: [] });
    const style = sheet.styles.find((s) => s.id === id);
    const m = rp.mark;
    return {
      name: style?.name ?? id, basedOn: style?.basedOn ?? "", next: style?.next ?? "",
      font: m.font, size: m.size, bold: m.bold, italic: m.italic, underline: m.underline, color: m.color ?? "000000",
      align: (["left", "center", "right", "both"].includes(rp.align) ? rp.align : "left") as Form["align"],
      spaceBefore: Math.round(rp.spaceBefore * 10) / 10, spaceAfter: Math.round(rp.spaceAfter * 10) / 10,
      lineSpacing: rp.lineRule === "auto" ? Math.round((rp.line / 240) * 100) / 100 : 1,
      indentLeft: Math.round(rp.left * 10) / 10, firstLine: Math.round(rp.firstLine * 10) / 10,
    };
  }, [doc, id, sheet]);
  const [form, setForm] = useState<Form>(current);
  const [shownFor, setShownFor] = useState(id);
  if (shownFor !== id) { setShownFor(id); setForm(current); }
  const set = <K extends keyof Form>(k: K, v: Form[K]) => setForm((f) => ({ ...f, [k]: v }));

  const apply = () => {
    // Only what was changed: a style keeps inheriting everything else from the one it is based on.
    const change: StyleChange = {};
    for (const k of Object.keys(form) as (keyof Form)[]) {
      if (form[k] === current[k]) continue;
      if (k === "basedOn" || k === "next") change[k] = form[k] || null;
      else (change as Record<string, unknown>)[k] = form[k];
    }
    if (Object.keys(change).length) modifyStyle(base, id, change)(view.state, view.dispatch);
    onClose();
    view.focus();
  };
  const create = () => {
    const name = newName.trim();
    if (!name) { setError("Give the new style a name."); return; }
    if (sheet.styles.some((s) => s.name?.toLowerCase() === name.toLowerCase())) { setError(`There is already a style called "${name}".`); return; }
    if (!createStyle(base, name)(view.state, view.dispatch)) { setError("Put the cursor in a paragraph first."); return; }
    onClose();
    view.focus();
  };
  const num = (k: "size" | "spaceBefore" | "spaceAfter" | "indentLeft" | "firstLine", step = 1) => (
    <input type="number" step={step} value={form[k]} onChange={(e) => set(k, Number(e.target.value))} />
  );

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog styles-dialog" role="dialog" aria-label="Styles">
        <h3>Styles</h3>
        <label>Style<select value={id} onChange={(e) => setId(e.target.value)} aria-label="Style to modify">
          {paragraphStyles.map((s) => <option key={s.id} value={s.id}>{s.name ?? s.id}</option>)}
        </select></label>
        <fieldset>
          <legend>Modify “{label(id)}” — every paragraph in this style changes</legend>
          <label>Name<input value={form.name} onChange={(e) => set("name", e.target.value)} /></label>
          <label>Based on<select value={form.basedOn} onChange={(e) => set("basedOn", e.target.value)}>
            <option value="">(no style)</option>
            {paragraphStyles.filter((s) => s.id !== id).map((s) => <option key={s.id} value={s.id}>{s.name ?? s.id}</option>)}
          </select></label>
          <label>Next paragraph<select value={form.next} onChange={(e) => set("next", e.target.value)}>
            <option value="">(same style)</option>
            {paragraphStyles.map((s) => <option key={s.id} value={s.id}>{s.name ?? s.id}</option>)}
          </select></label>
          <div className="row">
            <label>Font<input list="pw-style-fonts" value={form.font} onChange={(e) => set("font", e.target.value)} aria-label="Font" />
              <datalist id="pw-style-fonts">{fonts.map((x) => <option key={x} value={x} />)}</datalist></label>
            <label>Size {num("size", 0.5)}</label>
            <label>Colour<input type="color" value={`#${form.color}`} onChange={(e) => set("color", e.target.value.slice(1).toUpperCase())} /></label>
          </div>
          <div className="row">
            <label className="choice"><input type="checkbox" checked={form.bold} onChange={(e) => set("bold", e.target.checked)} /> Bold</label>
            <label className="choice"><input type="checkbox" checked={form.italic} onChange={(e) => set("italic", e.target.checked)} /> Italic</label>
            <label className="choice"><input type="checkbox" checked={form.underline} onChange={(e) => set("underline", e.target.checked)} /> Underline</label>
          </div>
          <div className="row">
            <label>Alignment<select value={form.align} onChange={(e) => set("align", e.target.value as Form["align"])}>
              <option value="left">Left</option><option value="center">Centre</option><option value="right">Right</option><option value="both">Justify</option>
            </select></label>
            <label>Line spacing<select value={form.lineSpacing} onChange={(e) => set("lineSpacing", Number(e.target.value))}>
              {SPACINGS.map(([n, v]) => <option key={n} value={v}>{n}</option>)}
              {!SPACINGS.some(([, v]) => v === form.lineSpacing) && <option value={form.lineSpacing}>{form.lineSpacing}</option>}
            </select></label>
          </div>
          <div className="row">
            <label>Space before (pt) {num("spaceBefore")}</label>
            <label>Space after (pt) {num("spaceAfter")}</label>
          </div>
          <div className="row">
            <label>Left indent (pt) {num("indentLeft")}</label>
            <label>First line (pt; negative: hanging) {num("firstLine")}</label>
          </div>
        </fieldset>
        <fieldset>
          <legend>New style from the paragraph at the cursor</legend>
          <div className="row">
            <label>Name<input value={newName} onChange={(e) => { setNewName(e.target.value); setError(""); }} aria-label="New style name" /></label>
            <button onClick={create}>Create Style</button>
          </div>
        </fieldset>
        {error && <p className="eq-error">{error}</p>}
        <div className="hf-actions">
          <button className="primary" onClick={apply}>Apply</button>
          <button onClick={onClose}>Cancel</button>
        </div>
      </div>
    </div>
  );
}
