import {
  applicationOf, type ApplyOptions, commitApplication, defaultOptions, fromProseMirror, PAPER, readAnyFormat, storyAt,
  storyShows, templateLook, TEMPLATES, withAdditions,
} from "@prime/editor";
import type { PrimeDocument, SectionProps } from "@prime/pdm";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useState } from "react";
import { loadTemplates, PageThumb } from "./MyTemplates";
import { openDocument } from "./platform";

/** A template to choose from: a saved one, a built-in one with a letterhead, or one from a file. */
interface Source { id: string; name: string; kind: string; doc: PrimeDocument }

/** "A4 · margins 2.5 cm" — a page setup in words. */
function pageWords(p: SectionProps): string {
  const w = p.pageWidth ?? 612, h = p.pageHeight ?? 792;
  const near = (a: number, b: number) => Math.abs(a - b) < 2;
  const paper = Object.entries(PAPER).find(([, s]) => (near(s.width, w) && near(s.height, h)) || (near(s.width, h) && near(s.height, w)))?.[0]
    ?? `${(w / 72 * 2.54).toFixed(1)} × ${(h / 72 * 2.54).toFixed(1)} cm`;
  const m = p.margins ?? {};
  const cm = (pt: number | undefined) => ((pt ?? 72) / 72 * 2.54).toFixed(1).replace(/\.0$/, "");
  const sides = [cm(m.top), cm(m.right), cm(m.bottom), cm(m.left)];
  const margins = sides.every((s) => s === sides[0]) ? `${sides[0]} cm` : `${sides.join(" / ")} cm`;
  return `${paper}${w > h ? " landscape" : ""}, margins ${margins}`;
}

const DESCRIBE = {
  every: "This letterhead is on every page.",
  first: "This letterhead is on the first page only.",
  firstAndRest: "This letterhead is on the first page; the following pages have a smaller header of their own.",
  none: "This template has no header, footer or page artwork to bring over. You can still use its page size, margins and styles.",
} as const;

/**
 * Apply a template or letterhead to the open document: its header and footer (on the first page only,
 * or on every page), artwork pinned to its first page, and, if wanted, its page setup and styles. The
 * document's text stays as it is; the preview shows its first two pages as they will be.
 */
export function ApplyTemplate({ view, base, onApplied, onClose }: {
  view: EditorView; base: () => PrimeDocument; onApplied(name: string): void; onClose(): void;
}) {
  const [sources, setSources] = useState<Source[]>();
  const [picked, setPicked] = useState<Source>();
  const [opts, setOpts] = useState<ApplyOptions>();
  const [error, setError] = useState("");
  const [noPage2, setNoPage2] = useState(false);

  useEffect(() => {
    void loadTemplates().catch(() => []).then((saved) => {
      const mine: Source[] = saved.map(({ entry, doc }) => ({ id: entry.id, name: entry.name, kind: entry.data.kind === "letterhead" ? "Letterhead" : "My template", doc }));
      // Built-in templates that have a letterhead (a header or footer) of their own.
      const builtIn: Source[] = TEMPLATES.filter((t) => t.id !== "blank").flatMap((t) => {
        const doc = t.create();
        return templateLook(doc).pattern === "none" ? [] : [{ id: `builtin:${t.id}`, name: t.name, kind: "Built-in", doc }];
      });
      setSources([...mine, ...builtIn]);
    });
  }, []);

  const look = useMemo(() => (picked ? templateLook(picked.doc) : undefined), [picked]);
  const ours = view.state;
  const ourPage = sectionPage(ours);
  const ourHas = !!(storyAt(ours, "header") || storyAt(ours, "footer")) && (storyShows(storyAt(ours, "header")) || storyShows(storyAt(ours, "footer")));
  const pageDiffers = !!look && pageWords(look.page) !== pageWords(ourPage);

  const pick = (s: Source) => {
    const l = templateLook(s.doc);
    setPicked(s);
    setError("");
    setNoPage2(false);
    // A letterhead made for another paper size looks wrong on this one: its page setup comes too, unless unticked.
    setOpts({ ...defaultOptions(l), pageSetup: pageWords(l.page).split(",")[0] !== pageWords(ourPage).split(",")[0] });
  };

  const fromFile = async () => {
    const f = await openDocument();
    if (!f) return;
    try {
      const doc = readAnyFormat(f.bytes, f.name).doc;
      const s: Source = { id: `file:${f.name}`, name: f.name.replace(/\.(docx|dotx|doc|rtf|odt)$/i, ""), kind: "From a file", doc };
      setSources((list) => [s, ...(list ?? []).filter((x) => x.id !== s.id)]);
      pick(s);
    } catch (e) {
      setError(`Prime Word could not read ${f.name}: ${(e as Error).message}`);
    }
  };

  // The document as it will be: its first two pages, laid out with the template applied.
  const preview = useMemo(() => {
    if (!picked || !opts) return undefined;
    try {
      const doc = base();
      const app = applicationOf(picked.doc, doc, view.state, opts);
      const after = view.state.apply(app.tr);
      const d = fromProseMirror(after.doc, withAdditions(doc, app));
      // Only the start matters for two thumbnails: the first section's opening blocks.
      const first = d.sections[0];
      const short = d.sections.length > 1 || first.blocks.length > 80;
      return short ? { ...d, sections: [{ ...first, blocks: first.blocks.slice(0, 80) }] } : d;
    } catch {
      return undefined;
    }
  }, [picked, opts, base, view]);
  useEffect(() => setNoPage2(false), [preview]);

  const apply = () => {
    if (!picked || !opts) return;
    const doc = base();
    const app = applicationOf(picked.doc, doc, view.state, opts);
    commitApplication(doc, app);
    view.dispatch(app.tr.scrollIntoView());
    view.focus();
    onApplied(picked.name);
  };

  const set = <K extends keyof ApplyOptions>(k: K, v: ApplyOptions[K]) => setOpts((o) => (o ? { ...o, [k]: v } : o));
  const restOffered = look?.pattern === "firstAndRest";
  const askFollowing = !!look && look.pattern !== "none" && opts?.pages === "first" && (restOffered || ourHas);

  return (
    <div className="hf-overlay" onKeyDown={(e) => { if (e.key === "Escape") onClose(); }}>
      <div className="hf-dialog apply-template" role="dialog" aria-label="Apply a template">
        <h3>Apply a template or letterhead</h3>
        <p className="hint">Brings a template's letterhead (header, footer, page artwork) to this document. Your text stays as it is.</p>
        <div className="apply-grid">
          <div className="apply-sources">
            <button type="button" onClick={() => void fromFile()} title="Use a Word document or template from your computer as the template">From a file…</button>
            {!sources && <p className="recent-empty">Loading…</p>}
            {sources && !sources.length && <p className="recent-empty">Letterheads and templates you save appear here. You can also use one from a file.</p>}
            <div className="apply-list" role="listbox" aria-label="Templates">
              {sources?.map((s) => (
                <button key={s.id} role="option" aria-selected={picked?.id === s.id} className={`apply-source ${picked?.id === s.id ? "on" : ""}`} onClick={() => pick(s)} title={s.name}>
                  <PageThumb doc={s.doc} width={64} />
                  <span><strong>{s.name}</strong><small>{s.kind}</small></span>
                </button>
              ))}
            </div>
          </div>

          {picked && look && opts ? (
            <div className="apply-options">
              <p className="apply-describe">{DESCRIBE[look.pattern]}{look.artwork ? " Its page artwork goes into the first page's header, so it stays in place as you type." : ""}</p>
              {look.pattern !== "none" && (
                <fieldset><legend>Put the letterhead on</legend>
                  <label className="choice"><input type="radio" name="apply-pages" checked={opts.pages === "first"} onChange={() => set("pages", "first")} /> The first page only</label>
                  <label className="choice"><input type="radio" name="apply-pages" checked={opts.pages === "every"} onChange={() => set("pages", "every")} /> Every page</label>
                </fieldset>
              )}
              {askFollowing && (
                <fieldset><legend>On the following pages</legend>
                  {restOffered && <label className="choice"><input type="radio" name="apply-rest" checked={opts.following === "template"} onChange={() => set("following", "template")} /> The template's own follow-on header</label>}
                  {ourHas && <label className="choice"><input type="radio" name="apply-rest" checked={opts.following === "keep"} onChange={() => set("following", "keep")} /> This document's header and footer, as now</label>}
                  <label className="choice"><input type="radio" name="apply-rest" checked={opts.following === "none" || (!ourHas && opts.following === "keep")} onChange={() => set("following", "none")} /> Nothing (blank)</label>
                </fieldset>
              )}
              {pageDiffers && (
                <label className="choice"><input type="checkbox" checked={opts.pageSetup} onChange={(e) => set("pageSetup", e.target.checked)} />
                  <span>Use the template's page: {pageWords(look.page)}<small>This document: {pageWords(ourPage)}</small></span></label>
              )}
              <label className="choice"><input type="checkbox" checked={opts.styles} onChange={(e) => set("styles", e.target.checked)} />
                <span>Use the template's styles too<small>Fonts, headings and colours of this document change to the template's</small></span></label>
              {look.pattern === "none" && !opts.styles && !opts.pageSetup && <p className="hint">Tick an option above, or choose another template.</p>}
            </div>
          ) : (
            <div className="apply-options"><p className="recent-empty">Choose a template or letterhead on the left.</p></div>
          )}

          <div className="apply-preview">
            {preview ? (
              <>
                <figure className="letterhead-preview"><PageThumb doc={preview} width={150} /><figcaption>Page 1</figcaption></figure>
                {!noPage2 && <figure className="letterhead-preview"><PageThumb doc={preview} width={150} page={1} onMissing={() => setNoPage2(true)} /><figcaption>Page 2</figcaption></figure>}
              </>
            ) : <p className="recent-empty">The preview of your document appears here.</p>}
          </div>
        </div>
        {error && <p className="hint error">{error}</p>}
        <div className="hf-actions">
          <button onClick={onClose}>Cancel</button>
          <button className="primary" disabled={!picked || !opts || (look?.pattern === "none" && !opts.styles && !opts.pageSetup)} onClick={apply}>Apply</button>
        </div>
      </div>
    </div>
  );
}

/** The page setup of the document's first section. */
function sectionPage(state: EditorView["state"]): SectionProps {
  let props: SectionProps | undefined;
  state.doc.forEach((n) => { if (!props && n.type.name === "section_break") props = (n.attrs.section as { props: SectionProps }).props; });
  return props ?? ((state.doc.attrs.section as { props: SectionProps } | null)?.props ?? {});
}
