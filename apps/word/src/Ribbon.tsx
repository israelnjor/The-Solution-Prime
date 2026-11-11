import {
  addColumn, addRow, changeLevel, deleteColumn, deleteRow, deleteTable, inTable, insertImage, insertPageBreak,
  insertTable, isActive, linkAt, listKind, PAPER, pageSetup, paragraphProp, runAtSelection, setColor, setFont, setFontSize,
  setHighlight, setLink, setPageSetup, setParagraph, toggleList, toggleRun, bordersAt, type BorderKind, setBorders,
  setLineSpacing, acceptOrReject, gotoChange, pictureAt, pictureWrap, type PictureWrap, setPictureWrap, setAltText,
  changeCase, type CaseMode, clearFormatting, growFont, toggleVertAlign, type CitationStyle, insertCitation, readSources,
  autofitWindow, type BorderChoice, distributeColumns, isHeaderRow, mergeCells, setCellAlign, setCellBorders, setCellShading, setTableStyle,
  splitCell, TABLE_STYLES, tableLookAt, type TableLook, tableToText, textToTable, toggleHeaderRows,
  insertColumnBreak, insertSectionBreak, type SectionStart, type LineNumbering, lineNumbersOf, lineNumbersSuppressed, setLineNumbers,
  toggleSuppressLineNumbers, insertShape, SHAPES, setShapeFill, setShapeLine, shapeAt, BODY, changeOutlineLevel, moveBlock, outlineLevelAt, outlineShown, setOutlineLevel, setTracking, trackingOn, hyphenationOn, setHyphenation, columnPreset, columnsOf, setColumns,
  applyColors, applyFontPair, applyStyleSet, COLOR_SETS, FONT_PAIRS, STYLE_SETS, type DropCapKind, dropCapOf,
  pageBordersOf, pageColorOf, setPageBorders, setPageColor, setWatermark, watermarkOf, type WatermarkSpec,
} from "@prime/editor";
import { Resolver } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { redo, undo } from "prosemirror-history";
import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { useEffect, useMemo, useState } from "react";
import { RecentMenu } from "./RecentList";
import { ShareMenu } from "./ShareMenu";
import type { ShareFormat } from "./share";

export const FONTS = ["Aptos", "Arial", "Book Antiqua", "Calibri", "Calibri Light", "Cambria", "Candara", "Century Gothic",
  "Consolas", "Constantia", "Corbel", "Courier New", "Garamond", "Georgia", "Palatino Linotype", "Segoe UI",
  "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana"];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
/** Raised by Ctrl+K in the editor; the ribbon opens its link box. */
export const LINK_EVENT = "prime-word:link";

const HIGHLIGHTS: [string, string][] = [["yellow", "#ffff00"], ["green", "#00ff00"], ["cyan", "#00ffff"], ["magenta", "#ff00ff"],
  ["red", "#ff0000"], ["blue", "#0000ff"], ["lightGray", "#c0c0c0"]];

type Tab = "file" | "home" | "insert" | "design" | "layout" | "references" | "review" | "view" | "tableDesign" | "table" | "picture" | "outlining";
/** Word's keys for the tabs: Alt+F, Alt+H, Alt+N, Alt+P, Alt+S, Alt+R, Alt+W. */
const TAB_KEYS: Record<string, string> = { file: "F", home: "H", insert: "N", design: "G", layout: "P", references: "S", review: "R", view: "W" };
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
  /** Sends a copy of the document through the system's share panel. */
  onShare(format: ShareFormat): void;
  onPrint(): void;
  onFind(replace: boolean): void;
  onHeaderFooter(): void;
  onParagraph(): void;
  mode: "print" | "web" | "outline";
  onMode(mode: "print" | "web" | "outline"): void;
  markup: Markup;
  onMarkup(markup: Markup): void;
  onComments(): void;
  onNewComment(): void;
  onHome(): void;
  onConvertPdf(): void;
  /** Deletes the page(s) the cursor or selection is on (Print Layout). */
  onDeletePage(): void;
  onEditPicture(): void;
  onLetterhead(): void;
  onSaveTemplate(): void;
  onSignature(): void;
  onBlocks(): void;
  onAutoCorrect(): void;
  /** The format painter: on (once, or kept on with a double-click) or off. */
  painter: boolean;
  onPainter(sticky: boolean): void;
  onFontDialog(): void;
  onSymbol(): void;
  onEquation(): void;
  onDateTime(): void;
  showMarks: boolean;
  onShowMarks(): void;
  ruler: boolean;
  onRuler(): void;
  navPane: boolean;
  onNavPane(): void;
  onReadMode(): void;
  onInfo(): void;
  onThesaurus(): void;
  onChart(): void;
  onProtect(): void;
  onCompare(): void;
  onReadAloud(): void;
  onZoomFit(fit: "page" | "width"): void;
  /** References: a table of contents (levels), updating fields, captions, a table of figures, cross-references. */
  onToc(levels: number): void;
  onUpdateFields(): void;
  onCaption(): void;
  onTableOfFigures(): void;
  onCrossRef(): void;
  onNote(kind: "footnote" | "endnote"): void;
  onShowNotes(): void;
  /** Citations: the style (all citations follow it), the sources dialog, the reference list. */
  onCitationStyle(style: CitationStyle): void;
  onSources(startNew: boolean): void;
  onBibliography(): void;
  onSortTable(): void;
  /** Insert → Drop Cap: dropped, in the margin, or none, for the paragraph at the cursor. */
  onDropCap(kind: DropCapKind): void;
  /** Home → Styles…: modify or create paragraph styles. */
  onStyles(): void;
}

const BORDER_CHOICES: [BorderChoice, string][] = [
  ["all", "All Borders"], ["outside", "Outside Borders"], ["inside", "Inside Borders"], ["top", "Top Border"], ["bottom", "Bottom Border"],
  ["left", "Left Border"], ["right", "Right Border"], ["insideH", "Inside Horizontal"], ["insideV", "Inside Vertical"], ["none", "No Border"],
];
const LOOK_OPTIONS: [keyof TableLook, string][] = [
  ["firstRow", "Header Row"], ["lastRow", "Total Row"], ["bandedRows", "Banded Rows"], ["firstColumn", "First Column"], ["lastColumn", "Last Column"], ["bandedColumns", "Banded Columns"],
];

const CASES: [CaseMode, string][] = [["sentence", "Sentence case."], ["lower", "lowercase"], ["upper", "UPPERCASE"], ["title", "Capitalize Each Word"], ["toggle", "tOGGLE cASE"]];

/** "heading 1" -> "Heading 1" (Word stores built-in style names in lower case). */
/** Small line icons for the ribbon (drawn in the text colour). */
function Icon({ name }: { name: "left" | "center" | "right" | "both" | "bullets" | "numbers" | "find" | "replace" }) {
  const lines = (rows: [number, number][]) => rows.map(([x1, x2], i) => <line key={i} x1={x1} x2={x2} y1={3 + i * 3.5} y2={3 + i * 3.5} />);
  const body = {
    left: lines([[2, 14], [2, 10], [2, 14], [2, 9]]),
    center: lines([[2, 14], [4, 12], [2, 14], [5, 11]]),
    right: lines([[2, 14], [6, 14], [2, 14], [7, 14]]),
    both: lines([[2, 14], [2, 14], [2, 14], [2, 14]]),
    bullets: <>{[4, 8, 12].map((y) => <g key={y}><circle cx={3} cy={y} r={1.1} fill="currentColor" stroke="none" /><line x1={6} x2={14} y1={y} y2={y} /></g>)}</>,
    numbers: <>{[4, 8, 12].map((y, i) => <g key={y}><text x={1} y={y + 1.8} fontSize={5} fill="currentColor" stroke="none" fontFamily="Segoe UI, sans-serif">{i + 1}</text><line x1={6} x2={14} y1={y} y2={y} /></g>)}</>,
    find: <><circle cx={7} cy={7} r={4} /><line x1={10} y1={10} x2={14} y2={14} /></>,
    replace: <><path d="M2 6h10l-2.5-2.5M14 10H4l2.5 2.5" /></>,
  }[name];
  return <svg className="icon" viewBox="0 0 16 16" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{body}</svg>;
}

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

export function Ribbon({ view: v, doc, onNew, onOpen, onOpenRecent, onSave, onExportPdf, onShare, onDeletePage, onPrint, onFind, onHeaderFooter, onParagraph, mode, onMode, markup, onMarkup, onComments, onNewComment, onHome, onConvertPdf, onEditPicture, onLetterhead, onSaveTemplate, onSignature, onBlocks, onAutoCorrect, painter, onPainter, onFontDialog, onSymbol, onEquation, onDateTime, showMarks, onShowMarks, ruler, onRuler, navPane, onNavPane, onInfo, onThesaurus, onChart, onProtect, onCompare, onReadAloud, onReadMode, onZoomFit, onToc, onUpdateFields, onCaption, onTableOfFigures, onCrossRef, onNote, onShowNotes, onCitationStyle, onSources, onBibliography, onSortTable, onStyles, onDropCap }: RibbonProps) {
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
  const [popover, setPopover] = useState<"link" | "table" | "margins" | "watermark" | "columns" | null>(null);
  const [colForm, setColForm] = useState({ count: 2, space: 1.27, line: false });
  const [mark, setMark] = useState<Required<Omit<WatermarkSpec, "font">>>({ text: "DRAFT", color: "C0C0C0", layout: "diagonal", semitransparent: true });
  /** The pen for table borders (Table Design): width in pt and colour. */
  const [pen, setPen] = useState({ width: 0.5, color: "000000" });
  const [shade, setShade] = useState("D9E2F3");
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
  // A shape or text box selected: its own tab (Shape Format), in the picture tab's place.
  const shape = state ? shapeAt(state) : undefined;
  const activeTab = ((tab === "table" || tab === "tableDesign") && !tableActive) || (tab === "picture" && !picture) ? "home" : tab === "outlining" && mode !== "outline" ? "view" : tab;
  // Outline view opens on its Outlining tab.
  useEffect(() => { if (mode === "outline") setTab("outlining"); }, [mode]);
  const look = state && tableActive ? tableLookAt(state) : undefined;
  const currentTableStyle = tableActive && state ? (() => {
    for (let d = state.selection.$from.depth; d > 0; d--) if (state.selection.$from.node(d).type.name === "table") return state.selection.$from.node(d).attrs.props.style as string | undefined;
    return undefined;
  })() : undefined;
  const columnWidth = state ? (() => { const p = pageSetup(state); return (p.pageWidth ?? 612) - (p.margins?.left ?? 72) - (p.margins?.right ?? 72); })() : 468;
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
        {([["file", "File"], ["home", "Home"], ["insert", "Insert"], ["design", "Design"], ["layout", "Layout"], ["references", "References"], ["review", "Review"], ["view", "View"]] as const).map(([t, label]) => (
          <button key={t} role="tab" aria-selected={activeTab === t} className={activeTab === t ? "on" : ""} onClick={() => setTab(t)}
            title={`${label} (Alt+${TAB_KEYS[t]})`}>{label}</button>
        ))}
        {tableActive && <button className={`contextual ${activeTab === "tableDesign" ? "on" : ""}`} onClick={() => setTab("tableDesign")}>Table Design</button>}
        {tableActive && <button className={`contextual ${activeTab === "table" ? "on" : ""}`} onClick={() => setTab("table")}>Table Layout</button>}
        {picture && <button className={`contextual ${activeTab === "picture" ? "on" : ""}`} onClick={() => setTab("picture")}>{shape ? "Shape Format" : "Picture"}</button>}
        {mode === "outline" && <button className={`contextual ${activeTab === "outlining" ? "on" : ""}`} onClick={() => setTab("outlining")}>Outlining</button>}
        <span className="spacer" />
        <ShareMenu onShare={onShare} className="share-button" />
        <button className={`layout-mode ${mode === "print" ? "on" : ""}`} onClick={() => onMode("print")} title="Pages as they will print">Print Layout</button>
        <button className={`layout-mode ${mode === "web" ? "on" : ""}`} onClick={() => onMode("web")} title="One continuous column">Web Layout</button>
      </nav>

      <div className="toolbar" role="toolbar" aria-label="Ribbon">
        {activeTab === "file" && (
          <>
            <div className="group" aria-label="Document">
              <button onClick={onHome} title="The home screen: templates, recent documents, My templates">Home screen</button>
              <button onClick={onNew} title="New document (from a template)">New</button>
              <button onClick={onOpen} title="Open a document: Word (.docx, .doc), Rich Text (.rtf), OpenDocument (.odt) or plain text (.txt) (Ctrl+O)">Open</button>
              <RecentMenu onOpen={onOpenRecent} />
              <button onClick={onConvertPdf} title="Turn a PDF into an editable Word document">Convert PDF</button>
              <button onClick={onInfo} title="Document properties (title, author, tags) and statistics">Info</button>
              <button onClick={onProtect} title="Protect the document with a password (or remove it)">Protect</button>
            </div>
            <div className="group" aria-label="Save">
              <button onClick={() => onSave()} title="Save (Ctrl+S)">Save</button>
              <button onClick={() => onSave(true)} title="Save as a new file">Save As</button>
              <button onClick={onSaveTemplate} title="Save this document as a template, to start new documents from">Save as Template</button>
            </div>
            <div className="group" aria-label="Share">
              <ShareMenu onShare={onShare} label="Share ▾" />
              <button onClick={onExportPdf} title="Export as PDF">Export PDF</button>
              <button onClick={onPrint} title="Print (Ctrl+P)">Print</button>
            </div>
          </>
        )}

        {activeTab === "home" && (
          <>
            <div className="group">
              <button onClick={run(undo)} title="Undo (Ctrl+Z)" aria-label="Undo">↶</button>
              <button onClick={run(redo)} title="Redo (Ctrl+Y)" aria-label="Redo">↷</button>
              <button className={painter ? "on" : ""} aria-pressed={painter} onClick={() => onPainter(false)} onDoubleClick={() => onPainter(true)}
                title="Format Painter: copy this formatting, then select text to apply it (double-click to keep it on; Esc to stop)">🖌 <span className="label">Format Painter</span></button>
            </div>
            <div className="group">
              <select value={currentStyle} onChange={(e) => run(setParagraph("style", e.target.value))()} title="Paragraph style" className="style-select">
                {!styles.some((s) => s.id === currentStyle) && <option value={currentStyle}>{currentStyle}</option>}
                {styles.map((s) => <option key={s.id} value={s.id}>{displayName(s.name ?? s.id)}</option>)}
              </select>
              <button onClick={onStyles} title="Styles: modify a style, or create one from this paragraph" aria-label="Styles">Styles…</button>
            </div>
            <div className="group">
              <select value={effective?.font ?? ""} onChange={(e) => run(setFont(e.target.value))()} title="Font" className="font-select">
                {fonts.map((f) => <option key={f} value={f} style={{ fontFamily: `"${f}"` }}>{f}</option>)}
              </select>
              <select value={effective?.size ?? 11} onChange={(e) => run(setFontSize(Number(e.target.value)))()} title="Font size" className="size-select">
                {[...new Set([...SIZES, effective?.size ?? 11])].sort((a, b) => a - b).map((s) => <option key={s} value={s}>{s}</option>)}
              </select>
              <button onClick={run(growFont(docRef, 1))} title="Grow Font (Ctrl+Shift+>)" aria-label="Grow Font">A<sup>▲</sup></button>
              <button onClick={run(growFont(docRef, -1))} title="Shrink Font (Ctrl+Shift+<)" aria-label="Shrink Font">A<sup>▼</sup></button>
              <select value="" onChange={(e) => { if (e.target.value) run(changeCase(e.target.value as CaseMode))(); }} title="Change Case (Shift+F3)" aria-label="Change Case" className="case-select">
                <option value="">Aa</option>
                {CASES.map(([m, l]) => <option key={m} value={m}>{l}</option>)}
              </select>
              <button onClick={run(clearFormatting)} title="Clear All Formatting: back to plain text" aria-label="Clear All Formatting">A⌫</button>
              {([["bold", "B", "Bold (Ctrl+B)"], ["italic", "I", "Italic (Ctrl+I)"], ["underline", "U", "Underline (Ctrl+U)"], ["strike", "S", "Strikethrough"]] as const)
                .map(([key, label, title]) => (
                  <button key={key} className={`fmt-${key} ${state && isActive(state, key) ? "on" : ""}`} onClick={run(toggleRun(key))} title={title}
                    aria-label={title} aria-pressed={!!(state && isActive(state, key))}>{label}</button>
                ))}
              <button className={effective?.vertAlign === "subscript" ? "on" : ""} onClick={run(toggleVertAlign("subscript"))} title="Subscript (Ctrl+=)" aria-label="Subscript">x<sub>2</sub></button>
              <button className={effective?.vertAlign === "superscript" ? "on" : ""} onClick={run(toggleVertAlign("superscript"))} title="Superscript (Ctrl+Shift+=)" aria-label="Superscript">x<sup>2</sup></button>
              <label className="color" title="Font colour" aria-label="Font colour">
                <span style={{ borderBottom: `3px solid #${effective?.color ?? "000000"}` }}>A</span>
                <input type="color" value={`#${(effective?.color ?? "000000").toLowerCase()}`} onChange={(e) => run(setColor(e.target.value.slice(1)))()} />
              </label>
              <select value={effective?.highlight ?? ""} onChange={(e) => run(setHighlight(e.target.value || undefined))()} title="Highlight" className="highlight-select">
                <option value="">Highlight</option>
                {HIGHLIGHTS.map(([name, css]) => <option key={name} value={name} style={{ background: css }}>{name}</option>)}
              </select>
              <button onClick={onFontDialog} title="Font settings (Ctrl+D)" aria-label="Font settings" className="launcher">⋯</button>
            </div>
            <div className="group">
              <button className={list === "bullet" ? "on" : ""} onClick={run(toggleList(docRef, "bullet"))} title="Bullets" aria-label="Bullets"><Icon name="bullets" /></button>
              <button className={list === "number" ? "on" : ""} onClick={run(toggleList(docRef, "number"))} title="Numbering" aria-label="Numbering"><Icon name="numbers" /></button>
              <button onClick={run(changeLevel(-1))} title="Decrease indent (Shift+Tab)" aria-label="Decrease list level" disabled={!list}>⇤</button>
              <button onClick={run(changeLevel(1))} title="Increase indent (Tab)" aria-label="Increase list level" disabled={!list}>⇥</button>
              <span className="divider" aria-hidden="true" />
              {([["left", "Left"], ["center", "Center"], ["right", "Right"], ["both", "Justify"]] as const).map(([value, label]) => (
                <button key={value} className={align === value || (value === "both" && align === "justify") ? "on" : ""}
                  onClick={run(setParagraph("align", value === "left" ? undefined : value))} title={`Align: ${label}`} aria-label={`Align ${label}`}><Icon name={value} /></button>
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
              <button className={showMarks ? "on" : ""} aria-pressed={showMarks} onClick={onShowMarks} title="Show/Hide ¶: paragraph marks, spaces and tabs (Ctrl+Shift+8)" aria-label="Show formatting marks">¶</button>
              <button onClick={onParagraph} title="Paragraph settings: indents, spacing, line and page breaks, tabs" aria-label="Paragraph settings" className="launcher">⋯</button>
            </div>
            <div className="group">
              <button onClick={() => onFind(false)} title="Find in the navigation pane (Ctrl+F)" aria-label="Find"><Icon name="find" /><span className="label">Find</span></button>
              <button onClick={() => onFind(true)} title="Replace (Ctrl+H)" aria-label="Replace"><Icon name="replace" /><span className="label">Replace</span></button>
            </div>
          </>
        )}

        {activeTab === "insert" && (
          <>
            <div className="group">
              <button onClick={run(insertPageBreak)} title="Page break (Ctrl+Enter)">Page Break</button>
              <button onClick={onDeletePage} title="Delete the page the cursor is on (or every page the selection covers); Ctrl+Z brings it back">Delete Page</button>
            </div>
            <div className="group" aria-label="Text box and shapes">
              <button onClick={() => { run(insertShape("textbox"))(); setTab("picture"); }} title="A box of text you can place anywhere on the page">Text Box</button>
              <button onClick={onChart} title="A column, bar, line or pie chart from a small table of numbers (double-click a chart to change it)">Chart…</button>
              <select value="" aria-label="Shapes" title="Insert a shape"
                onChange={(e) => { if (e.target.value) { run(insertShape("shape", e.target.value))(); setTab("picture"); } }}>
                <option value="">Shapes…</option>
                {SHAPES.map((sh) => <option key={sh.geometry} value={sh.geometry}>{sh.label}</option>)}
              </select>
            </div>
            <div className="group" aria-label="Drop Cap">
              <select value={state ? dropCapOf(state) : "none"} aria-label="Drop Cap" title="A large first letter beside the paragraph's first lines"
                onChange={(e) => onDropCap(e.target.value as DropCapKind)}>
                <option value="none">Drop Cap: None</option>
                <option value="drop">Drop Cap: Dropped</option>
                <option value="margin">Drop Cap: In Margin</option>
              </select>
            </div>
            <div className="group popover-anchor">
              <button onClick={() => setPopover(popover === "table" ? null : "table")} title="Insert a table">Table</button>
              <button onClick={run(textToTable(docRef))} title="Turn the selected lines (separated by tabs) into a table">Text to Table</button>
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
            <div className="group">
              <button onClick={onEquation} title="Insert an equation (Alt+=): fractions, roots, sums, Greek…">Equation…</button>
              <button onClick={onSymbol} title="Insert a symbol: ©, €, ₵, ½, →, Greek letters…">Symbol…</button>
              <button onClick={onDateTime} title="Insert today's date or the time">Date &amp; Time…</button>
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


        {activeTab === "design" && state && (() => {
          const color = pageColorOf(state);
          const border = pageBordersOf(state);
          const current = watermarkOf(state);
          const borderChoice = !border ? "none" : border.style === "single" && border.width >= 2 ? "thick" : border.style;
          const borderOf = (choice: string, hex?: string) => (choice === "none" ? undefined : {
            style: choice === "thick" ? "single" : choice, width: choice === "thick" ? 3 : choice === "double" ? 0.75 : 0.5,
            ...(hex ? { color: hex } : border?.color ? { color: border.color } : {}),
          });
          return (
            <>
              <div className="group" aria-label="Document Formatting">
                <select value="" aria-label="Style Set" title="A coordinated look for the body text, title and headings"
                  onChange={(e) => { if (e.target.value) run(applyStyleSet(doc, e.target.value))(); }}>
                  <option value="">Style Set</option>
                  {STYLE_SETS.map((x) => <option key={x.name} value={x.name}>{x.name}</option>)}
                </select>
                <select value="" aria-label="Theme Fonts" title="Heading and body fonts"
                  onChange={(e) => { if (e.target.value) run(applyFontPair(doc, e.target.value))(); }}>
                  <option value="">Fonts</option>
                  {FONT_PAIRS.map((x) => <option key={x.name} value={x.name}>{x.name === x.heading ? x.name : `${x.name}: ${x.heading} / ${x.body}`}</option>)}
                </select>
                <select value="" aria-label="Theme Colors" title="Colours of the title and headings"
                  onChange={(e) => { if (e.target.value) run(applyColors(doc, e.target.value))(); }}>
                  <option value="">Colors</option>
                  {COLOR_SETS.map((x) => <option key={x.name} value={x.name}>{x.name}</option>)}
                </select>
              </div>
              <div className="group" aria-label="Watermark">
                <select value="" aria-label="Watermark" title="Text across every page, behind the text" onChange={(e) => {
                  const v2 = e.target.value;
                  if (v2 === "remove") run(setWatermark(undefined))();
                  else if (v2 === "custom") {
                    setMark({ text: current?.text ?? "DRAFT", color: current?.color ?? "C0C0C0", layout: current?.layout ?? "diagonal", semitransparent: current?.semitransparent ?? true });
                    setPopover("watermark");
                  } else if (v2) run(setWatermark({ text: v2 }))();
                }}>
                  <option value="">{current ? `Watermark: ${current.text}` : "Watermark"}</option>
                  {["CONFIDENTIAL", "DO NOT COPY", "DRAFT", "SAMPLE", "URGENT", "ASAP"].map((t) => <option key={t} value={t}>{t}</option>)}
                  <option value="custom">Custom Watermark…</option>
                  {current && <option value="remove">Remove Watermark</option>}
                </select>
                {popover === "watermark" && (
                  <span className="popover-anchor">
                    <form className="popover" onSubmit={(e) => {
                      e.preventDefault();
                      if (mark.text.trim()) run(setWatermark({ ...mark, text: mark.text.trim() }))();
                      setPopover(null);
                    }}>
                      <label>Text <input value={mark.text} onChange={(e) => setMark((m) => ({ ...m, text: e.target.value }))} autoFocus /></label>
                      <label>Colour <input type="color" value={`#${mark.color}`} onChange={(e) => setMark((m) => ({ ...m, color: e.target.value.slice(1).toUpperCase() }))} /></label>
                      <label>Layout <select value={mark.layout} onChange={(e) => setMark((m) => ({ ...m, layout: e.target.value as "diagonal" | "horizontal" }))}>
                        <option value="diagonal">Diagonal</option><option value="horizontal">Horizontal</option>
                      </select></label>
                      <label className="choice"><input type="checkbox" checked={mark.semitransparent} onChange={(e) => setMark((m) => ({ ...m, semitransparent: e.target.checked }))} /> Semitransparent</label>
                      <button className="primary" type="submit">Apply</button>
                    </form>
                  </span>
                )}
              </div>
              <div className="group" aria-label="Page Color">
                <label className="field">Page Color <input type="color" value={`#${color ?? "FFFFFF"}`} aria-label="Page Color"
                  onChange={(e) => run(setPageColor(e.target.value))()} /></label>
                <button onClick={run(setPageColor(undefined))} disabled={!color} title="White pages">No Color</button>
              </div>
              <div className="group" aria-label="Page Borders">
                <select value={borderChoice} aria-label="Page Borders" title="A box round every page" onChange={(e) => run(setPageBorders(borderOf(e.target.value)))()}>
                  <option value="none">No page border</option>
                  <option value="single">Box: single line</option>
                  <option value="double">Box: double line</option>
                  <option value="thick">Box: thick line</option>
                  <option value="dotted">Box: dotted</option>
                  <option value="dashed">Box: dashed</option>
                </select>
                <input type="color" value={`#${border?.color ?? "000000"}`} aria-label="Page border colour" title="Border colour" disabled={!border}
                  onChange={(e) => run(setPageBorders(borderOf(borderChoice, e.target.value.slice(1).toUpperCase())))()} />
              </div>
            </>
          );
        })()}

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
                        <label key={k}>{k[0].toUpperCase() + k.slice(1)} <input type="number" step={0.01} min={0} max={20} value={custom[k]}
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
              <div className="group popover-anchor" aria-label="Columns">
                {(() => {
                  const c = columnsOf(state);
                  const text = setup.pageWidth - (m.left ?? 72) - (m.right ?? 72);
                  const current = c.widths ? (c.widths[0] < c.widths[1] ? "left" : "right") : c.count <= 3 ? ["one", "two", "three"][c.count - 1] : "more";
                  return (
                    <select value={current} aria-label="Columns" title="Split the text into columns" onChange={(e) => {
                      const v2 = e.target.value;
                      if (v2 === "more") {
                        setColForm({ count: Math.max(2, c.count), space: Math.round((c.space / CM) * 100) / 100, line: !!c.line });
                        setPopover("columns");
                      } else run(setColumns(columnPreset(v2 as "one", text)))();
                    }}>
                      <option value="one">Columns: One</option>
                      <option value="two">Columns: Two</option>
                      <option value="three">Columns: Three</option>
                      <option value="left">Columns: Left</option>
                      <option value="right">Columns: Right</option>
                      <option value="more">{current === "more" ? `Columns: ${c.count}` : "More Columns…"}</option>
                    </select>
                  );
                })()}
                {popover === "columns" && (
                  <form className="popover" onSubmit={(e) => {
                    e.preventDefault();
                    run(setColumns({ count: colForm.count, space: colForm.space * CM, line: colForm.line }))();
                    setPopover(null);
                  }}>
                    <label>Number of columns <input type="number" min={1} max={12} value={colForm.count} aria-label="Number of columns"
                      onChange={(e) => setColForm((f) => ({ ...f, count: Number(e.target.value) }))} /></label>
                    <label>Spacing <input type="number" min={0} max={10} step={0.01} value={colForm.space} aria-label="Column spacing"
                      onChange={(e) => setColForm((f) => ({ ...f, space: Number(e.target.value) }))} /> cm</label>
                    <label className="choice"><input type="checkbox" checked={colForm.line} onChange={(e) => setColForm((f) => ({ ...f, line: e.target.checked }))} /> Line between</label>
                    <button className="primary" type="submit">Apply</button>
                  </form>
                )}
              </div>
              <div className="group" aria-label="Breaks">
                <select value="" title="Breaks" aria-label="Breaks" onChange={(e) => {
                  const v = e.target.value;
                  if (v === "page") run(insertPageBreak)();
                  else if (v === "column") run(insertColumnBreak)();
                  else if (v) run(insertSectionBreak(v as SectionStart))();
                }}>
                  <option value="">Breaks</option>
                  <optgroup label="Page breaks">
                    <option value="page">Page (Ctrl+Enter)</option>
                    <option value="column">Column (Ctrl+Shift+Enter)</option>
                  </optgroup>
                  <optgroup label="Section breaks">
                    <option value="nextPage">Next page</option>
                    <option value="continuous">Continuous</option>
                    <option value="evenPage">Even page</option>
                    <option value="oddPage">Odd page</option>
                  </optgroup>
                </select>
              </div>
              <div className="group" aria-label="Hyphenation">
                <select value={hyphenationOn(state) ? "auto" : "none"} aria-label="Hyphenation" title="Split words at the end of lines"
                  onChange={(e) => run(setHyphenation(e.target.value === "auto"))()}>
                  <option value="none">Hyphenation: None</option>
                  <option value="auto">Hyphenation: Automatic</option>
                </select>
              </div>
              <div className="group" aria-label="Line Numbers">
                <select value={lineNumbersOf(state)} aria-label="Line Numbers" title="Numbers in the margin beside each line" onChange={(e) => {
                  if (e.target.value === "suppress") run(toggleSuppressLineNumbers)();
                  else run(setLineNumbers(e.target.value as LineNumbering))();
                }}>
                  <option value="none">Line Numbers: None</option>
                  <option value="continuous">Continuous</option>
                  <option value="newPage">Restart Each Page</option>
                  <option value="newSection">Restart Each Section</option>
                  <option value="suppress">{lineNumbersSuppressed(state) ? "✓ " : ""}Suppress for Current Paragraph</option>
                </select>
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

        {activeTab === "references" && (
          <>
            <div className="group" aria-label="Table of Contents">
              <select value="" onChange={(e) => { if (e.target.value) onToc(Number(e.target.value)); }} title="Insert a table of contents from the headings" aria-label="Table of Contents" className="toc-select">
                <option value="">Table of Contents</option>
                <option value="3">Contents: Headings 1–3</option>
                <option value="2">Contents: Headings 1–2</option>
                <option value="1">Contents: Heading 1 only</option>
              </select>
              <button onClick={onUpdateFields} title="Update the contents, captions, cross-references and page numbers (F9)">Update Table</button>
            </div>
            <div className="group" aria-label="Footnotes">
              <button onClick={() => onNote("footnote")} title="Insert Footnote (Alt+Ctrl+F): a numbered note at the foot of the page">Insert Footnote</button>
              <button onClick={() => onNote("endnote")} title="Insert Endnote (Alt+Ctrl+D): a note at the end of the document">Insert Endnote</button>
              <button onClick={onShowNotes} title="Show and edit all footnotes and endnotes">Show Notes</button>
            </div>
            <div className="group" aria-label="Citations and Bibliography">
              <select value="" onChange={(e) => {
                const v = e.target.value;
                if (v === "new") onSources(true);
                else if (v) run(insertCitation(docRef, v))();
              }} title="Insert a citation of one of your sources" aria-label="Insert Citation" className="cite-select">
                <option value="">Insert Citation</option>
                {readSources(doc).sources.map((s) => <option key={s.tag} value={s.tag}>{s.authors[0]?.last ? `${s.authors[0].last} (${s.year ?? "n.d."}): ` : ""}{s.title.slice(0, 50)}</option>)}
                <option value="new">Add New Source…</option>
              </select>
              <button onClick={() => onSources(false)} title="Add, edit and cite the sources of this document">Manage Sources…</button>
              <label className="choice" title="Citation style">Style <select value={readSources(doc).style} onChange={(e) => onCitationStyle(e.target.value as CitationStyle)} aria-label="Citation style">
                <option value="APA">APA</option><option value="Harvard">Harvard</option><option value="IEEE">IEEE</option>
              </select></label>
              <button onClick={onBibliography} title="Insert the reference list of the sources cited">Bibliography</button>
            </div>
            <div className="group" aria-label="Captions">
              <button onClick={onCaption} title="Number a picture or table: Figure 1, Table 1…">Insert Caption…</button>
              <button onClick={onTableOfFigures} title="A list of the figures (or tables) with their pages">Insert Table of Figures</button>
              <button onClick={onCrossRef} title="Refer to a heading or caption: its text, number or page, kept up to date">Cross-reference…</button>
            </div>
          </>
        )}

        {activeTab === "view" && (
          <>
            <div className="group" aria-label="Views">
              <button onClick={onReadMode} title="Read the pages, one or two at a time (Esc to return)">Read Mode</button>
              <button className={mode === "print" ? "on" : ""} onClick={() => onMode("print")}>Print Layout</button>
              <button className={mode === "outline" ? "on" : ""} onClick={() => onMode("outline")} title="The document as an outline of its headings">Outline</button>
              <button className={mode === "web" ? "on" : ""} onClick={() => onMode("web")}>Web Layout</button>
            </div>
            <div className="group" aria-label="Show">
              <label className="choice"><input type="checkbox" checked={navPane} onChange={onNavPane} /> Navigation Pane</label>
              <label className="choice"><input type="checkbox" checked={ruler} onChange={onRuler} /> Ruler</label>
              <label className="choice"><input type="checkbox" checked={showMarks} onChange={onShowMarks} /> Formatting marks</label>
            </div>
            <div className="group" aria-label="Zoom">
              <button onClick={() => onZoomFit("page")} title="Zoom so a whole page fits in the window">One Page</button>
              <button onClick={() => onZoomFit("width")} title="Zoom so the page fills the window's width">Page Width</button>
            </div>
          </>
        )}

        {activeTab === "outlining" && state && (
          <>
            <div className="group" aria-label="Outline tools">
              <button onClick={run(changeOutlineLevel(() => doc, 0, 1))} title="Promote to Heading 1">«</button>
              <button onClick={run(changeOutlineLevel(() => doc, -1))} title="Promote (Alt+Shift+←)" aria-label="Promote">←</button>
              <select value={String(outlineLevelAt(state, doc))} aria-label="Outline level" title="Outline level"
                onChange={(e) => run(changeOutlineLevel(() => doc, 0, Number(e.target.value)))()}>
                {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((l) => <option key={l} value={l}>Level {l}</option>)}
                <option value={BODY}>Body Text</option>
              </select>
              <button onClick={run(changeOutlineLevel(() => doc, 1))} title="Demote (Alt+Shift+→)" aria-label="Demote">→</button>
              <button onClick={run(changeOutlineLevel(() => doc, 0, BODY))} title="Demote to body text">»</button>
            </div>
            <div className="group" aria-label="Move">
              <button onClick={run(moveBlock(-1))} title="Move up (Alt+Shift+↑)">▲ Move Up</button>
              <button onClick={run(moveBlock(1))} title="Move down (Alt+Shift+↓)">▼ Move Down</button>
            </div>
            <div className="group" aria-label="Show">
              <label className="field">Show Level{" "}
                <select value={String(outlineShown(state))} aria-label="Show level" onChange={(e) => run(setOutlineLevel(Number(e.target.value)))()}>
                  {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((l) => <option key={l} value={l}>Level {l}</option>)}
                  <option value={BODY}>All Levels</option>
                </select>
              </label>
            </div>
            <div className="group">
              <button onClick={() => onMode("print")} title="Back to Print Layout">Close Outline View</button>
            </div>
          </>
        )}

        {activeTab === "review" && (
          <>
            <div className="group stacked" aria-label="Comments">
              <button onClick={onNewComment} title="Comment on the selected text">New Comment</button>
              <button onClick={onComments} title="Show or hide the comments">Comments</button>
            </div>
            <div className="group stacked" aria-label="Proofing">
              <button onClick={onAutoCorrect} title="Choose what is corrected as you type, and add your own replacements">AutoCorrect…</button>
              <button onClick={onThesaurus} title="Words of like meaning for the word at the cursor (Shift+F7)">Thesaurus</button>
              <button onClick={onReadAloud} title="Read the document aloud from the cursor (Ctrl+Alt+Space)">Read Aloud</button>
            </div>
            <div className="group stacked" aria-label="Tracking">
              <button className={state && trackingOn(state) ? "on" : ""} aria-pressed={!!(state && trackingOn(state))} onClick={() => state && run(setTracking(!trackingOn(state)))()}
                title="Track Changes: record insertions and deletions as changes others can accept or reject (Ctrl+Shift+E)">Track Changes</button>
              <select className="markup-select" value={markup} onChange={(e) => onMarkup(e.target.value as Markup)} title="How tracked changes are shown" aria-label="Markup">
                <option value="all">All Markup</option>
                <option value="none">No Markup (final)</option>
                <option value="original">Original</option>
              </select>
            </div>
            <div className="group stacked" aria-label="Tracked changes">
              <button onClick={run(gotoChange(-1))} title="Previous change">◀ Previous</button>
              <button onClick={run(gotoChange(1))} title="Next change">Next ▶</button>
              <button onClick={run(acceptOrReject(true))} title="Accept the change at the cursor">Accept</button>
              <button onClick={run(acceptOrReject(false))} title="Reject the change at the cursor">Reject</button>
              <button onClick={run(acceptOrReject(true, true))} title="Accept every change in the document">Accept All</button>
              <button onClick={run(acceptOrReject(false, true))} title="Reject every change in the document">Reject All</button>
            </div>
            <div className="group" aria-label="Compare">
              <button onClick={onCompare} title="Compare this document with an earlier version: the differences as tracked changes">Compare…</button>
            </div>
          </>
        )}

        {activeTab === "picture" && picture && (
          <>
            {shape && (
              <div className="group" aria-label="Shape styles">
                <label className="field">Fill{" "}
                  <input type="color" aria-label="Shape fill" value={`#${shape.float.shape.fill ?? "FFFFFF"}`}
                    onChange={(e) => run(setShapeFill(e.target.value.slice(1).toUpperCase()))()} disabled={shape.float.shape.geometry === "line"} />
                </label>
                <button onClick={run(setShapeFill(undefined))} disabled={!shape.float.shape.fill} title="No fill (see-through)">No Fill</button>
                <label className="field">Outline{" "}
                  <input type="color" aria-label="Shape outline" value={`#${shape.float.shape.line?.color ?? "000000"}`}
                    onChange={(e) => run(setShapeLine({ color: e.target.value.slice(1).toUpperCase(), width: shape.float.shape.line?.width ?? 0.75 }))()} />
                </label>
                <select aria-label="Outline weight" title="Outline weight" value={String(shape.float.shape.line?.width ?? "")}
                  onChange={(e) => run(setShapeLine({ color: shape.float.shape.line?.color ?? "000000", width: Number(e.target.value) }))()}>
                  {!shape.float.shape.line && <option value="">—</option>}
                  {[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => <option key={w} value={w}>{w} pt</option>)}
                </select>
                <button onClick={run(setShapeLine(undefined))} disabled={!shape.float.shape.line} title="No outline">No Outline</button>
              </div>
            )}
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
              <span className="hint">{shape ? (shape.float.kind === "textbox" ? "Drag to move, a corner to resize; double-click to type in it." : "Drag to move; drag a corner to resize.") : pictureWrap(picture.node) !== "inline" ? "Drag the picture to move it; drag a corner to resize." : "Drag a corner to resize (keeps its shape), or a side to stretch it."}</span>
            </div>
            {!shape && <div className="group">
              <button onClick={onEditPicture} disabled={picture.kept} title={picture.kept ? "Pictures with effects from Word cannot be cropped yet" : "Crop and rotate"}>Crop &amp; Rotate…</button>
              <button disabled={picture.kept} onClick={() => {
                const text = window.prompt("Alternative text (read aloud by screen readers):", picture.node.attrs.description ?? "");
                if (text !== null) run(setAltText(text))();
              }}>Alt Text…</button>
              <button onClick={() => v && (v.dispatch(v.state.tr.deleteSelection()), v.focus())}>Delete</button>
            </div>}
            {shape && (
              <div className="group">
                <button onClick={() => v && (v.dispatch(v.state.tr.deleteSelection()), v.focus())}>Delete</button>
              </div>
            )}
          </>
        )}

        {activeTab === "tableDesign" && look && (
          <>
            <div className="group" aria-label="Table Style Options">
              {LOOK_OPTIONS.map(([k, label]) => (
                <label key={k} className="choice"><input type="checkbox" checked={look[k]} onChange={(e) => run(setTableStyle(docRef, undefined, { ...look, [k]: e.target.checked }))()} /> {label}</label>
              ))}
            </div>
            <div className="group" aria-label="Table Styles">
              <select value={currentTableStyle ?? ""} onChange={(e) => run(setTableStyle(docRef, e.target.value, look))()} title="Table style" aria-label="Table style" className="table-style-select">
                {!TABLE_STYLES.some((s) => s.id === currentTableStyle) && <option value={currentTableStyle ?? ""}>{currentTableStyle ? displayName(doc.styles.styles.find((s) => s.id === currentTableStyle)?.name ?? currentTableStyle) : "(no style)"}</option>}
                {TABLE_STYLES.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </div>
            <div className="group" aria-label="Borders">
              <select value={pen.width} onChange={(e) => setPen((p) => ({ ...p, width: Number(e.target.value) }))} title="Pen weight" aria-label="Pen weight">
                {[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3].map((w) => <option key={w} value={w}>{w} pt</option>)}
              </select>
              <label className="color" title="Pen colour" aria-label="Pen colour"><span style={{ borderBottom: `3px solid #${pen.color}` }}>✎</span>
                <input type="color" value={`#${pen.color.toLowerCase()}`} onChange={(e) => setPen((p) => ({ ...p, color: e.target.value.slice(1).toUpperCase() }))} /></label>
              <select value="" onChange={(e) => { if (e.target.value) run(setCellBorders(e.target.value as BorderChoice, { width: pen.width, color: pen.color }))(); }} title="Borders of the selected cells" aria-label="Borders" className="border-select">
                <option value="">Borders</option>
                {BORDER_CHOICES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
              </select>
            </div>
            <div className="group" aria-label="Shading">
              <label className="color" title="Shading colour" aria-label="Shading colour"><span style={{ background: `#${shade}`, padding: "0 4px" }}>▦</span>
                <input type="color" value={`#${shade.toLowerCase()}`} onChange={(e) => { const c = e.target.value.slice(1).toUpperCase(); setShade(c); run(setCellShading(c))(); }} /></label>
              <button onClick={run(setCellShading(shade))} title="Shade the selected cells">Shading</button>
              <button onClick={run(setCellShading(undefined))} title="Remove the shading of the selected cells">No Colour</button>
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
            <div className="group" aria-label="Merge">
              <button onClick={run(mergeCells)} title="Merge the selected cells into one">Merge Cells</button>
              <button onClick={run(splitCell)} title="Split a merged cell back, or a cell into two columns">Split Cell</button>
            </div>
            <div className="group" aria-label="Cell Size">
              <button onClick={run(distributeColumns)} title="Make all columns the same width">Distribute Columns</button>
              <button onClick={run(autofitWindow(columnWidth))} title="Make the table as wide as the page's text">AutoFit Window</button>
            </div>
            <div className="group" aria-label="Alignment">
              <select value="" onChange={(e) => { const [v, h] = e.target.value.split(" "); if (v) run(setCellAlign(v as "top", h as "left"))(); }} title="Cell alignment" aria-label="Cell alignment" className="cell-align-select">
                <option value="">Align</option>
                {(["top", "center", "bottom"] as const).flatMap((v) => (["left", "center", "right"] as const).map((h) => (
                  <option key={`${v}${h}`} value={`${v} ${h}`}>{`${v === "center" ? "Middle" : v[0].toUpperCase() + v.slice(1)} ${h === "center" ? "centre" : h}`}</option>
                )))}
              </select>
            </div>
            <div className="group" aria-label="Data">
              <button className={state && isHeaderRow(state) ? "on" : ""} aria-pressed={!!(state && isHeaderRow(state))} onClick={run(toggleHeaderRows)} title="Repeat the selected top rows on every page the table runs onto">Repeat Header Rows</button>
              <button onClick={onSortTable} title="Sort the rows by a column">Sort…</button>
              <button onClick={run(tableToText())} title="Turn the table into text, cells separated by tabs">Convert to Text</button>
            </div>
          </>
        )}
      </div>
    </header>
  );
}


