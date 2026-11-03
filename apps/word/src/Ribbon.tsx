import {
  addColumn, addRow, changeLevel, deleteColumn, deleteRow, deleteTable, inTable, insertImage, insertPageBreak,
  insertTable, isActive, linkAt, listKind, PAPER, pageSetup, paragraphProp, runAtSelection, setColor, setFont, setFontSize,
  setHighlight, setLink, setPageSetup, setParagraph, toggleList, toggleRun, bordersAt, type BorderKind, setBorders,
  setLineSpacing, acceptOrReject, gotoChange, pictureAt, pictureWrap, type PictureWrap, setPictureWrap, setAltText,
} from "@prime/editor";
import { Resolver } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { redo, undo } from "prosemirror-history";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useState } from "react";
import { RecentMenu } from "./RecentList";

const FONTS = ["Aptos", "Arial", "Book Antiqua", "Calibri", "Calibri Light", "Cambria", "Candara", "Century Gothic",
  "Consolas", "Constantia", "Corbel", "Courier New", "Garamond", "Georgia", "Palatino Linotype", "Segoe UI",
  "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana"];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Raised by Ctrl+K in the editor; the ribbon opens its link box. */
export const LINK_EVENT = "prime-word:link";

const HIGHLIGHTS: [string, string][] = [["yellow", "#ffff00"], ["green", "#00ff00"], ["cyan", "#00ffff"], ["magenta", "#ff00ff"],
  ["red", "#ff0000"], ["blue", "#0000ff"], ["lightGray", "#c0c0c0"]];

type Tab = "home" | "insert" | "layout" | "review" | "table" | "picture";
/** Word's keys for the tabs: Alt+H, Alt+N, Alt+P, Alt+R. */
const TAB_KEYS: Record<string, string> = { home: "H", insert: "N", layout: "P", review: "R" };
export type Markup = "all" | "none" | "original";

const CM = 72 / 2.54;
/** Word's margin presets (in cm). */
const MARGINS: [string, { top: number; bottom: number; left: number; right: number }][] = [
  ["Normal", { top: 2.54, bottom: 2.54, left: 2.54, right: 2.54 }],
  ["Narrow", { top: 1.27, bottom: 1.27, left: 1.27, right: 1.27 }],
  ["Moderate", { top: 2.54, bottom: 2.54, left: 1.91, right: 1.91 }],
  ["Wide", { top: 2.54, bottom: 2.54, left: 5.08, right: 5.08 }],
];
const inCm = (pt: number | undefined) => Math.round(((pt ?? 72) / CM) * 100) / 100;

export interface RibbonProps {
  view: EditorView | null;
  doc: PrimeDocument;
  onNew(): void;
  onOpen(): void;
  onOpenRecent(path: string): void;
  onSave(saveAs?: boolean): void;
  onExportPdf(): void;
  onPrint(): void;
  onFind(replace: boolean): void;
  onHeaderFooter(): void;
  onParagraph(): void;
  mode: "print" | "web";
  onMode(mode: "print" | "web"): void;
  markup: Markup;
  onMarkup(markup: Markup): void;
  onComments(): void;
  onNewComment(): void;
  onHome(): void;
  onConvertPdf(): void;
  onEditPicture(): void;
  onLetterhead(): void;
  onSaveTemplate(): void;
  onSignature(): void;
  onBlocks(): void;
}

/** "heading 1" -> "Heading 1" (Word stores built-in style names in lower case). */
const displayName = (name: string) => name.charAt(0).toUpperCase() + name.slice(1);

function readImage(): Promise<{ data: string; contentType: string; width: number; height: number; name: string } | undefined> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/gif,image/bmp";
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) return resolve(undefined);
      const bytes = new Uint8Array(await file.arrayBuffer());
      const bitmap = await createImageBitmap(new Blob([bytes], { type: file.type }));
      let binary = "";
      for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
      resolve({ data: btoa(binary), contentType: file.type, width: bitmap.width, height: bitmap.height, name: file.name });
    };
    input.click();
  });
}

export function Ribbon({ view: v, doc, onNew, onOpen, onOpenRecent, onSave, onExportPdf, onPrint, onFind, onHeaderFooter, onParagraph, mode, onMode, markup, onMarkup, onComments, onNewComment, onHome, onConvertPdf, onEditPicture, onLetterhead, onSaveTemplate, onSignature, onBlocks }: RibbonProps) {
  const [tab, setTab] = useState<Tab>("home");
  // Alt+H / N / P / R: open that tab and put the keyboard on its first control.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (!e.altKey || e.ctrlKey || e.metaKey || e.shiftKey) return;
      const t = (Object.keys(TAB_KEYS) as Tab[]).find((k) => TAB_KEYS[k] === e.key.toUpperCase());
      if (!t) return;
      e.preventDefault();
      setTab(t);
      setTimeout(() => document.querySelector<HTMLElement>(".toolbar button:not(:disabled), .toolbar select")?.focus(), 0);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);
  const [popover, setPopover] = useState<"link" | "table" | "margins" | null>(null);
  const [custom, setCustom] = useState({ top: 2.54, bottom: 2.54, left: 2.54, right: 2.54 });
  const [linkUrl, setLinkUrl] = useState("");
  const [tableSize, setTableSize] = useState({ rows: 3, cols: 3 });
  const resolver = useMemo(() => new Resolver(doc), [doc]);
  const docRef = () => doc;

  const run = (cmd: Command) => () => { if (v) { cmd(v.state, v.dispatch, v); v.focus(); } };
  const state = v?.state;
  const styles = doc.styles.styles.filter((s) => s.type === "paragraph" && !s.name?.toLowerCase().startsWith("toc"));
  const defaultStyle = styles.find((s) => s.isDefault)?.id ?? "";
  const currentStyle = (state && paragraphProp(state, "style")) ?? defaultStyle;
  const align = (state && paragraphProp(state, "align")) ?? "left";
  const list = state ? listKind(doc, paragraphProp(state, "list")?.numId) : undefined;
  const effective = state ? resolver.run(currentStyle, runAtSelection(state)) : undefined;
  const fonts = [...new Set([...FONTS, ...(effective ? [effective.font] : [])])].sort();
  const tableActive = !!state && inTable(state);
  const picture = state ? pictureAt(state) : undefined;
  const activeTab = (tab === "table" && !tableActive) || (tab === "picture" && !picture) ? "home" : tab;
  // Selecting a picture shows its tab (as Word does); leaving it returns to where you were.
  const pictureKey = picture ? `${picture.pos}` : "";
  useEffect(() => {
    if (pictureKey) setTab((t) => (t === "picture" ? t : "picture"));
  }, [pictureKey]);

  const openLink = () => {
    setLinkUrl((state && linkAt(state)) ?? "https://");
    setTab("insert");
    setPopover("link");
  };
  // Ctrl+K in the editor opens the link box.
  useEffect(() => {
    const onLink = () => openLink();
    window.addEventListener(LINK_EVENT, onLink);
    return () => window.removeEventListener(LINK_EVENT, onLink);
  });

  return (
    <header className="ribbon">
      <nav className="tabs" role="tablist" aria-label="Ribbon tabs">
        <button className="brand" onClick={onHome} title="Home: new, open, recent documents">Prime Word</button>
        {([["home", "Home"], ["insert", "Insert"], ["layout", "Layout"], ["review", "Review"]] as const).map(([t, label]) => (
          <button key={t} role="tab" aria-selected={activeTab === t} className={activeTab === t ? "on" : ""} onClick={() => setTab(t)}
            title={`${label} (Alt+${TAB_KEYS[t]})`}>{label}</button>
        ))}
        {tableActive && <button className={`contextual ${activeTab === "table" ? "on" : ""}`} onClick={() => setTab("table")}>Table</button>}
        {picture && <button className={`contextual ${activeTab === "picture" ? "on" : ""}`} onClick={() => setTab("picture")}>Picture</button>}
        <span className="spacer" />
        <button className={mode === "print" ? "on" : ""} onClick={() => onMode("print")} title="Pages as they will print">Print Layout</button>
        <button className={mode === "web" ? "on" : ""} onClick={() => onMode("web")} title="One continuous column">Web Layout</button>
      </nav>

      <div className="toolbar" role="toolbar" aria-label="Ribbon">
        {activeTab === "home" && (
          <>
            <div className="group">
              <button onClick={onNew} title="New document">New</button>
              <button onClick={onOpen} title="Open a .docx (Ctrl+O)">Open</button>
              <RecentMenu onOpen={onOpenRecent} />
              <button onClick={onConvertPdf} title="Turn a PDF into an editable Word document">Convert PDF</button>
              <button onClick={() => onSave()} title="Save (Ctrl+S)">Save</button>
              <button onClick={() => onSave(true)} title="Save as a new file">Save As</button>
              <button onClick={onSaveTemplate} title="Save this document as a template, to start new documents from">Save as Template</button>
              <button onClick={onExportPdf} title="Export as PDF">Export PDF</button>
              <button onClick={onPrint} title="Print (Ctrl+P)">Print</button>
            </div>
            <div className="group">
              <button onClick={run(undo)} title="Undo (Ctrl+Z)" aria-label="Undo">↶</button>
              <button onClick={run(redo)} title="Redo (Ctrl+Y)" aria-label="Redo">↷</button>
            </div>
            <div className="group">
              <select value={currentStyle} onChange={(e) => run(setParagraph("style", e.target.value))()} title="Paragraph style" className="style-select">
                {!styles.some((s) => s.id === currentStyle) && <option value={currentStyle}>{currentStyle}</option>}
                {styles.map((s) => <option key={s.id} value={s.id}>{displayName(s.name ?? s.id)}</option>)}
              </select>
            </div>
            <div className="group">
              <select value={effective?.font ?? ""} onChange={(e) => run(setFont(e.target.value))()} title="Font" className="font-select">
                {fonts.map((f) => <option key={f} value={f} style={{ fontFamily: `"${f}"` }}>{f}</option>)}
              </select>
              <select value={effective?.size ?? 11} onChange={(e) => run(setFontSize(Number(e.target.value)))()} title="Font size" className="size-select">
                {[...new Set([...SIZES, effective?.size ?? 11])].sort((a, b) => a - b).map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
              {([["bold", "B", "Bold (Ctrl+B)"], ["italic", "I", "Italic (Ctrl+I)"], ["underline", "U", "Underline (Ctrl+U)"], ["strike", "S", "Strikethrough"]] as const)
                .map(([key, label, title]) => (
                  <button key={key} className={`fmt-${key} ${state && isActive(state, key) ? "on" : ""}`} onClick={run(toggleRun(key))} title={title}
                    aria-label={title} aria-pressed={!!(state && isActive(state, key))}>{label}</button>
                ))}
              <label className="color" title="Font colour" aria-label="Font colour">
                <span style={{ borderBottom: `3px solid #${effective?.color ?? "000000"}` }}>A</span>
                <input type="color" value={`#${(effective?.color ?? "000000").toLowerCase()}`} onChange={(e) => run(setColor(e.target.value.slice(1)))()} />
              </label>
              <select value={effective?.highlight ?? ""} onChange={(e) => run(setHighlight(e.target.value || undefined))()} title="Highlight" className="highlight-select">
                <option value="">No highlight</option>
                {HIGHLIGHTS.map(([name, css]) => <option key={name} value={name} style={{ background: css }}>{name}</option>)}
              </select>
            </div>
            <div className="group">
              <button className={list === "bullet" ? "on" : ""} onClick={run(toggleList(docRef, "bullet"))} title="Bullets">• List</button>
              <button className={list === "number" ? "on" : ""} onClick={run(toggleList(docRef, "number"))} title="Numbering">1. List</button>
              <button onClick={run(changeLevel(-1))} title="Decrease indent (Shift+Tab)" aria-label="Decrease list level" disabled={!list}>⇤</button>
              <button onClick={run(changeLevel(1))} title="Increase indent (Tab)" aria-label="Increase list level" disabled={!list}>⇥</button>
            </div>
            <div className="group">
              {([["left", "Left"], ["center", "Center"], ["right", "Right"], ["both", "Justify"]] as const).map(([value, label]) => (
                <button key={value} className={align === value || (value === "both" && align === "justify") ? "on" : ""}
                  onClick={run(setParagraph("align", value === "left" ? undefined : value))} title={`Align: ${label}`}>{label}</button>
              ))}
            </div>
            <div className="group">
              <select value="" onChange={(e) => {
                const v = e.target.value;
                if (v === "dialog") onParagraph();
                else if (v) run(setLineSpacing("multiple", Number(v)))();
              }} title="Line spacing" className="spacing-select">
                <option value="">Spacing</option>
                {["1", "1.15", "1.5", "2", "2.5", "3"].map((x) => <option key={x} value={x}>{x}</option>)}
                <option value="dialog">Paragraph settings…</option>
              </select>
              <select value={(state && bordersAt(state)) ?? ""} onChange={(e) => { if (e.target.value) run(setBorders(e.target.value as BorderKind))(); }} title="Borders" className="border-select">
                <option value="">Borders</option>
                <option value="none">No border</option>
                <option value="bottom">Bottom border</option>
                <option value="top">Top border</option>
                <option value="box">Box</option>
              </select>
              <button onClick={onParagraph} title="Paragraph settings: indents and spacing" aria-label="Paragraph settings">¶</button>
            </div>
            <div className="group">
              <button onClick={() => onFind(false)} title="Find (Ctrl+F)">Find</button>
              <button onClick={() => onFind(true)} title="Replace (Ctrl+H)">Replace</button>
            </div>
          </>
        )}

        {activeTab === "insert" && (
          <>
            <div className="group">
              <button onClick={run(insertPageBreak)} title="Page break (Ctrl+Enter)">Page Break</button>
            </div>
            <div className="group popover-anchor">
              <button onClick={() => setPopover(popover === "table" ? null : "table")} title="Insert a table">Table</button>
              {popover === "table" && (
                <div className="popover">
                  <label>Rows <input type="number" min={1} max={100} value={tableSize.rows} onChange={(e) => setTableSize((s) => ({ ...s, rows: Number(e.target.value) }))} /></label>
                  <label>Columns <input type="number" min={1} max={20} value={tableSize.cols} onChange={(e) => setTableSize((s) => ({ ...s, cols: Number(e.target.value) }))} /></label>
                  <button className="primary" onClick={() => {
                    run(insertTable(docRef, Math.max(1, tableSize.rows), Math.max(1, Math.min(20, tableSize.cols))))();
                    setPopover(null);
                  }}>Insert</button>
                </div>
              )}
            </div>
            <div className="group">
              <button onClick={async () => {
                const img = await readImage();
                if (img) run(insertImage(docRef, img))();
              }} title="Insert a picture from a file">Picture</button>
            </div>
            <div className="group">
              <button onClick={onHeaderFooter} title="Edit the header and footer, add page numbers">Header &amp; Footer</button>
              <button onClick={onLetterhead} title="Design a letterhead: logo, name, address, contacts">Letterhead…</button>
              <button onClick={onSignature} title="Insert your signature, or make one from a photo, by drawing or typing">Signature…</button>
              <button onClick={onBlocks} title="Signature blocks, addresses, stamps and saved text: insert left, centre or right">Blocks…</button>
            </div>
            <div className="group popover-anchor">
              <button onClick={openLink} title="Insert or edit a link (Ctrl+K)">Link</button>
              {popover === "link" && (
                <form className="popover" onSubmit={(e) => { e.preventDefault(); run(setLink(linkUrl.trim()))(); setPopover(null); }}>
                  <label>Address <input autoFocus value={linkUrl} onChange={(e) => setLinkUrl(e.target.value)} style={{ width: 240 }} /></label>
                  <button className="primary" type="submit">Apply</button>
                  <button type="button" onClick={() => { run(setLink(""))(); setPopover(null); }}>Remove link</button>
                </form>
              )}
            </div>
          </>
        )}


        {activeTab === "layout" && state && (() => {
          const setup = pageSetup(state);
          const m = setup.margins!;
          const landscape = setup.pageWidth > setup.pageHeight;
          const short = Math.min(setup.pageWidth, setup.pageHeight);
          const long = Math.max(setup.pageWidth, setup.pageHeight);
          const paper = Object.entries(PAPER).find(([, s]) => Math.abs(s.width - short) < 1 && Math.abs(s.height - long) < 1)?.[0] ?? "Custom";
          const preset = MARGINS.find(([, p]) => (["top", "bottom", "left", "right"] as const).every((k) => Math.abs(p[k] * CM - (m[k] ?? 72)) < 0.5))?.[0] ?? "Custom";
          return (
            <>
              <div className="group">
                <label className="field">Margins
                  <select value={preset} onChange={(e) => {
                    const p = MARGINS.find(([n]) => n === e.target.value)?.[1];
                    if (p) run(setPageSetup({ margins: { top: p.top * CM, bottom: p.bottom * CM, left: p.left * CM, right: p.right * CM } }))();
                  }}>
                    {MARGINS.map(([name, p]) => <option key={name} value={name}>{name} ({p.top} / {p.left} cm)</option>)}
                    {preset === "Custom" && <option value="Custom">Custom</option>}
                  </select>
                </label>
                <span className="popover-anchor">
                  <button onClick={() => {
                    setCustom({ top: inCm(m.top), bottom: inCm(m.bottom), left: inCm(m.left), right: inCm(m.right) });
                    setPopover(popover === "margins" ? null : "margins");
                  }}>Custom Margins…</button>
                  {popover === "margins" && (
                    <form className="popover" onSubmit={(e) => {
                      e.preventDefault();
                      run(setPageSetup({ margins: { top: custom.top * CM, bottom: custom.bottom * CM, left: custom.left * CM, right: custom.right * CM } }))();
                      setPopover(null);
                    }}>
                      {(["top", "bottom", "left", "right"] as const).map((k) => (
                        <label key={k}>{k[0].toUpperCase() + k.slice(1)} <input type="number" step={0.1} min={0} max={20} value={custom[k]}
                          onChange={(e) => setCustom((c) => ({ ...c, [k]: Number(e.target.value) }))} /> cm</label>
                      ))}
                      <button className="primary" type="submit">Apply</button>
                    </form>
                  )}
                </span>
              </div>
              <div className="group">
                <button className={!landscape ? "on" : ""} onClick={run(setPageSetup({ orientation: "portrait" }))}>Portrait</button>
                <button className={landscape ? "on" : ""} onClick={run(setPageSetup({ orientation: "landscape" }))}>Landscape</button>
              </div>
              <div className="group">
                <label className="field">Size
                  <select value={paper} onChange={(e) => { const p = PAPER[e.target.value]; if (p) run(setPageSetup({ paper: p }))(); }}>
                    {Object.entries(PAPER).map(([name, s]) => <option key={name} value={name}>{name} ({inCm(s.width)} × {inCm(s.height)} cm)</option>)}
                    {paper === "Custom" && <option value="Custom">Custom</option>}
                  </select>
                </label>
              </div>
            </>
          );
        })()}

        {activeTab === "review" && (
          <>
            <div className="group" aria-label="Comments">
              <button onClick={onNewComment} title="Comment on the selected text">New Comment</button>
              <button onClick={onComments} title="Show or hide the comments">Comments</button>
            </div>
            <div className="group" aria-label="Tracked changes">
              <button onClick={run(gotoChange(-1))} title="Previous change">◀ Previous</button>
              <button onClick={run(gotoChange(1))} title="Next change">Next ▶</button>
              <button onClick={run(acceptOrReject(true))} title="Accept the change at the cursor">Accept</button>
              <button onClick={run(acceptOrReject(false))} title="Reject the change at the cursor">Reject</button>
              <button onClick={run(acceptOrReject(true, true))} title="Accept every change in the document">Accept All</button>
              <button onClick={run(acceptOrReject(false, true))} title="Reject every change in the document">Reject All</button>
            </div>
            <div className="group">
              <select value={markup} onChange={(e) => onMarkup(e.target.value as Markup)} title="How tracked changes are shown" aria-label="Markup">
                <option value="all">All Markup</option>
                <option value="none">No Markup (final)</option>
                <option value="original">Original</option>
              </select>
            </div>
          </>
        )}

        {activeTab === "picture" && picture && (
          <>
            <div className="group">
              <label className="field">Wrap text{" "}
                <select value={pictureWrap(picture.node)} onChange={(e) => run(setPictureWrap(e.target.value as PictureWrap))()}
                  title="How the picture sits with the text" aria-label="Wrap text">
                  {!picture.kept && <option value="inline">In line with text</option>}
                  <option value="square">Square</option>
                  <option value="tight">Tight</option>
                  <option value="topAndBottom">Top and bottom</option>
                  <option value="behind">Behind text</option>
                  <option value="inFront">In front of text</option>
                </select>
              </label>
              {pictureWrap(picture.node) !== "inline" && <span className="hint">Drag the picture to move it; drag a corner to resize.</span>}
            </div>
            <div className="group">
              <button onClick={onEditPicture} disabled={picture.kept} title={picture.kept ? "Pictures with effects from Word cannot be cropped yet" : "Crop and rotate"}>Crop &amp; Rotate…</button>
              <button disabled={picture.kept} onClick={() => {
                const text = window.prompt("Alternative text (read aloud by screen readers):", picture.node.attrs.description ?? "");
                if (text !== null) run(setAltText(text))();
              }}>Alt Text…</button>
              <button onClick={() => v && (v.dispatch(v.state.tr.deleteSelection()), v.focus())}>Delete</button>
            </div>
          </>
        )}

        {activeTab === "table" && (
          <>
            <div className="group">
              <button onClick={run(addRow(-1))}>Row Above</button>
              <button onClick={run(addRow(1))}>Row Below</button>
            </div>
            <div className="group">
              <button onClick={run(addColumn(-1))}>Column Left</button>
              <button onClick={run(addColumn(1))}>Column Right</button>
            </div>
            <div className="group">
              <button onClick={run(deleteRow)}>Delete Row</button>
              <button onClick={run(deleteColumn)}>Delete Column</button>
              <button onClick={run(deleteTable)}>Delete Table</button>
            </div>
          </>
        )}
      </div>
    </header>
  );
}


