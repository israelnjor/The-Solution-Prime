import { letterheadDocument, letterheadStories, type LetterheadSpec, setStory } from "@prime/editor";
import type { PrimeDocument } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useMemo, useState } from "react";
import { PageThumb, saveTemplate } from "./MyTemplates";

const FONTS = ["Calibri", "Cambria", "Arial", "Georgia", "Times New Roman", "Segoe UI", "Garamond", "Trebuchet MS"];
const LAYOUTS: [LetterheadSpec["layout"], string][] = [["logoLeft", "Logo left"], ["centre", "Centred"], ["logoRight", "Logo right"]];
const STORE = "prime-word:letterhead";

/** The last letterhead designed on this computer, to start from next time. */
function remembered(): LetterheadSpec {
  try {
    const saved = localStorage.getItem(STORE);
    if (saved) return JSON.parse(saved);
  } catch { /* storage unavailable */ }
  return { name: "", tagline: "", address: [], contacts: [], layout: "logoLeft", color: "1F3864", font: "Calibri", rule: true, footer: false };
}

/** A picture file as image data (base64) with its pixel size. */
function readLogo(): Promise<LetterheadSpec["logo"]> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg";
    input.onchange = async () => {
      const f = input.files?.[0];
      if (!f) return resolve(undefined);
      const url = URL.createObjectURL(f);
      const img = new Image();
      img.onload = async () => {
        const bytes = new Uint8Array(await f.arrayBuffer());
        let s = "";
        for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
        resolve({ data: btoa(s), contentType: f.type === "image/jpeg" ? "image/jpeg" : "image/png", width: img.naturalWidth, height: img.naturalHeight });
        URL.revokeObjectURL(url);
      };
      img.src = url;
    };
    input.click();
  });
}

/**
 * Design a letterhead (name, tagline, address, contacts, logo, layout, colour, font), with a live
 * preview. Save it as a template, start a letter with it, or put it on the open document.
 */
export function LetterheadDesigner({ view, base, onNewDocument, onSaved, onClose }: {
  view?: EditorView; base: () => PrimeDocument; onNewDocument(doc: PrimeDocument, name: string): void; onSaved(): void; onClose(): void;
}) {
  const [spec, setSpec] = useState<LetterheadSpec>(remembered);
  const set = <K extends keyof LetterheadSpec>(k: K, v: LetterheadSpec[K]) => setSpec((s) => ({ ...s, [k]: v }));
  const lines = (text: string) => text.split("\n").map((l) => l.trim()).filter(Boolean);
  const preview = useMemo(() => {
    const doc = letterheadDocument({ ...spec, name: spec.name || "Your organisation" });
    doc.sections[0].blocks = [{ type: "paragraph", props: {}, content: [{ type: "text", text: "Dear …," }] }];
    return doc;
  }, [spec]);
  const keep = () => { try { localStorage.setItem(STORE, JSON.stringify(spec)); } catch { /* storage unavailable */ } };
  const ready = !!spec.name.trim();

  const apply = () => {
    if (!view) return;
    const doc = base();
    const section = doc.sections[0]?.props ?? {};
    const width = (section.pageWidth ?? 612) - (section.margins?.left ?? 72) - (section.margins?.right ?? 72);
    const { header, footer } = letterheadStories(doc, spec, width);
    setStory("header", header)(view.state, view.dispatch, view);
    if (footer) setStory("footer", footer)(view.state, view.dispatch, view);
    keep();
    onClose();
  };

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog letterhead" role="dialog" aria-label="Design a letterhead">
        <h3>Letterhead</h3>
        <div className="letterhead-grid">
          <form className="letterhead-form" onSubmit={(e) => e.preventDefault()}>
            <label>Organisation name<input value={spec.name} onChange={(e) => set("name", e.target.value)} placeholder="e.g. Accra Community Library" autoFocus /></label>
            <label>Tagline (optional)<input value={spec.tagline ?? ""} onChange={(e) => set("tagline", e.target.value)} placeholder="e.g. Knowledge for everyone" /></label>
            <label>Address (one line each)<textarea rows={2} value={spec.address.join("\n")} onChange={(e) => set("address", lines(e.target.value))} /></label>
            <label>Phone, e-mail, website (one per line)<textarea rows={3} value={spec.contacts.join("\n")} onChange={(e) => set("contacts", lines(e.target.value))} /></label>
            <div className="row">
              <button type="button" onClick={() => void readLogo().then((logo) => logo && set("logo", logo))}>{spec.logo ? "Change logo…" : "Add logo…"}</button>
              {spec.logo && <button type="button" className="link" onClick={() => set("logo", undefined)}>Remove logo</button>}
            </div>
            <fieldset className="row"><legend>Layout</legend>
              {LAYOUTS.map(([id, label]) => (
                <label key={id} className="choice"><input type="radio" name="layout" checked={spec.layout === id} onChange={() => set("layout", id)} /> {label}</label>
              ))}
            </fieldset>
            <div className="row">
              <label className="choice">Colour <input type="color" value={`#${spec.color}`} onChange={(e) => set("color", e.target.value.slice(1).toUpperCase())} /></label>
              <label className="choice">Font <select value={spec.font} onChange={(e) => set("font", e.target.value)}>{FONTS.map((f) => <option key={f}>{f}</option>)}</select></label>
            </div>
            <label className="choice"><input type="checkbox" checked={spec.rule} onChange={(e) => set("rule", e.target.checked)} /> Line under the letterhead</label>
            <label className="choice"><input type="checkbox" checked={spec.footer} onChange={(e) => set("footer", e.target.checked)} /> Address and contacts in the footer</label>
          </form>
          <figure className="letterhead-preview"><PageThumb doc={preview} width={300} /><figcaption>Preview</figcaption></figure>
        </div>
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          {view && <button disabled={!ready} onClick={apply} title="Put this letterhead on the open document">Apply to this document</button>}
          <button disabled={!ready} onClick={() => { keep(); void saveTemplate(letterheadDocument(spec), `${spec.name} letterhead`, "letterhead").then(onSaved); }}>Save as template</button>
          <button className="primary" disabled={!ready} onClick={() => { keep(); onNewDocument(letterheadDocument(spec), `${spec.name} letter.docx`); }}>Start a letter</button>
        </div>
      </div>
    </div>
  );
}
