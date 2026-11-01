import { readDocx, writeDocx } from "@prime/docx";
import {
  blankDocument, changeLevel, createRender, createSchema, exitEmptyListItem, findPlugin, fromProseMirror,
  insertPageBreak, moveCell, PAGE_GAP, pageSetup, pagination, paragraphStyles, pastePlugin, toggleRun, toProseMirror,
  wordCount,
} from "@prime/editor";
import type { LayoutResult } from "@prime/layout";
import { CanvasMeasurer } from "@prime/layout/dist/browser/canvas.js";
import type { PrimeDocument } from "@prime/pdm";
import type { Template } from "@prime/editor";
import { invoke } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { baseKeymap, chainCommands } from "prosemirror-commands";
import { history, redo, undo } from "prosemirror-history";
import { keymap } from "prosemirror-keymap";
import { type Command, EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { useCallback, useEffect, useReducer, useRef, useState } from "react";
import { FindPanel } from "./FindPanel";
import { HeaderFooterEditor } from "./HeaderFooterEditor";
import { ParagraphDialog } from "./ParagraphDialog";
import { TemplatePicker } from "./TemplatePicker";
import { SpellMenu } from "./SpellMenu";
import { documentLanguage, loadSpeller, type Misspelling, misspellingAt, recheck, type Speller, spellKey, spellPlugin } from "./spelling";
import { loadCloudFonts } from "./fonts";
import { PageSheets } from "./PageSheets";
import { exportPdf, makePdf, printDocument } from "./exportPdf";
import { ChangedOnDiskError, confirmAction, type DocumentLocation, isDesktop, openDocument, openPath, saveDocument } from "./platform";
import { deleteSnapshot, orphanedSnapshots, SESSION, saveSnapshot, type Snapshot } from "./recovery";
import { LINK_EVENT, Ribbon } from "./Ribbon";

/** Raised by Ctrl+P in the editor. */
const PRINT_EVENT = "prime-word:print";

interface Loaded {
  base: PrimeDocument;
  name: string;
  location?: DocumentLocation;
  warnings: string[];
  /** Bumped when fonts finish loading, so text is re-measured. */
  fonts: number;
}

/** Text in which the document's font family names appear (styles, theme, runs). */
function fontUsage(doc: PrimeDocument): string {
  const theme = doc.package.parts.find((p) => /theme\/theme\d*\.xml$/.test(p.name));
  return JSON.stringify(doc.styles) + JSON.stringify(doc.sections) + (theme ? atob(theme.data) : "");
}

export function App() {
  const mount = useRef<HTMLDivElement>(null);
  const view = useRef<EditorView | null>(null);
  const [file, setFile] = useState<Loaded>({ base: blankDocument(), name: "Document1.docx", warnings: [], fonts: 0 });
  const [dirty, setDirty] = useState(false);
  const [showWarnings, setShowWarnings] = useState(false);
  /** Print Layout shows pages as they will print; Web Layout is one continuous column. */
  const [mode, setMode] = useState<"print" | "web">("print");
  const [layout, setLayout] = useState<LayoutResult>();
  /** The document the current layout was made from (automation waits for the startup document's). */
  const [layoutOf, setLayoutOf] = useState<PrimeDocument>();
  /** Desktop: the startup document (if any) has been loaded. */
  const [startupDone, setStartupDone] = useState(!isDesktop);
  const [overflow, setOverflow] = useState<Set<number>>(new Set());
  const [find, setFind] = useState<"find" | "replace" | null>(null);
  const [editingHeaderFooter, setEditingHeaderFooter] = useState(false);
  const [editingParagraph, setEditingParagraph] = useState(false);
  const [pickingTemplate, setPickingTemplate] = useState(false);
  // Spelling: on by default; the dictionary loads in the background.
  const [spelling, setSpelling] = useState(true);
  const speller = useRef<Speller>();
  const ignoredWords = useRef(new Set<string>());
  const [spellMenu, setSpellMenu] = useState<{ issue: Misspelling; x: number; y: number }>();
  // Zoom (percent). ?zoom= sets the initial value (handy for overview screenshots).
  const [zoom, setZoom] = useState(() => Number(new URLSearchParams(location.search).get("zoom")) || 100);
  /** Unsaved work found from an earlier session that did not close cleanly. */
  const [recovery, setRecovery] = useState<Snapshot[]>([]);
  const [notice, setNotice] = useState("");
  const [, refresh] = useReducer((x: number) => x + 1, 0);
  const fileRef = useRef(file);
  fileRef.current = file;
  const dirtyRef = useRef(dirty);
  dirtyRef.current = dirty;
  const measurerRef = useRef<CanvasMeasurer>();
  const snapshotTimer = useRef<ReturnType<typeof setTimeout>>();
  const automationDone = useRef(false);

  /** Writes a recovery snapshot of the current editor content (autosave). */
  const snapshot = useCallback(() => {
    const v = view.current;
    if (!v || !dirtyRef.current) return;
    const { base, name, location, warnings } = fileRef.current;
    void saveSnapshot({
      id: SESSION, name, savedAt: Date.now(), warnings, doc: fromProseMirror(v.state.doc, base),
      ...(location?.kind === "path" ? { path: location.path } : {}),
    }).catch(() => undefined);
  }, []);

  const flash = (message: string) => {
    setNotice(message);
    setTimeout(() => setNotice((n) => (n === message ? "" : n)), 4000);
  };

  const save = useCallback(async (saveAs = false) => {
    const v = view.current;
    if (!v) return;
    const { base, location, name } = fileRef.current;
    try {
      const bytes = writeDocx(fromProseMirror(v.state.doc, base));
      let saved;
      try {
        saved = await saveDocument(bytes, name, saveAs ? undefined : location);
      } catch (e) {
        if (!(e instanceof ChangedOnDiskError)) throw e;
        // Another program saved this file after we opened it: never overwrite that silently.
        const overwrite = await confirmAction(
          `${name} was changed by another program after you opened it.\n\n` +
          "Yes: replace it with your version (the other changes are lost).\nNo: save your version as a new file.",
          "File changed on disk");
        saved = await saveDocument(bytes, name, overwrite ? location : undefined, true);
      }
      if (!saved) return;
      setFile((f) => ({ ...f, name: saved.name, location: saved.location ?? f.location }));
      setDirty(false);
      clearTimeout(snapshotTimer.current);
      void deleteSnapshot(SESSION);
    } catch (e) {
      alert(`Prime Word could not save ${name}:\n${(e as Error).message}`);
    }
  }, []);

  // (Re)create the editor whenever a different document is loaded, its fonts arrive, or the view changes.
  useEffect(() => {
    if (!mount.current) return;
    const measurer = new CanvasMeasurer();
    measurerRef.current = measurer;
    const schema = createSchema(createRender(file.base, measurer));
    const command = (fn: () => void): Command => () => { fn(); return true; };
    const state = EditorState.create({
      doc: toProseMirror(file.base, schema),
      plugins: [
        history(),
        keymap({
          "Mod-z": undo, "Mod-y": redo, "Mod-Shift-z": redo,
          "Mod-b": toggleRun("bold"), "Mod-i": toggleRun("italic"), "Mod-u": toggleRun("underline"),
          "Mod-s": command(() => void save()),
          "Mod-f": command(() => setFind("find")),
          "Mod-h": command(() => setFind("replace")),
          "Mod-k": command(() => window.dispatchEvent(new Event(LINK_EVENT))),
          "Mod-p": command(() => window.dispatchEvent(new Event(PRINT_EVENT))),
          "Mod-Enter": insertPageBreak,
          // Tab moves between table cells, or changes a list item's level.
          Tab: chainCommands(moveCell(1), changeLevel(1)),
          "Shift-Tab": chainCommands(moveCell(-1), changeLevel(-1)),
          // Enter on an empty list item ends the list (as in Word).
          Enter: exitEmptyListItem,
        }),
        keymap(baseKeymap),
        paragraphStyles(file.base, measurer),
        findPlugin(),
        pastePlugin(() => fileRef.current.base),
        ...(spelling ? [spellPlugin(() => speller.current, ignoredWords.current)] : []),
        ...(mode === "print" ? [pagination({ base: () => fileRef.current.base, measurer, onLayout: (l) => { setLayout(l); setLayoutOf(fileRef.current.base); }, onFit: setOverflow })] : []),
      ],
    });
    const v = new EditorView(mount.current, {
      state,
      // Our own checker underlines misspellings; turn off the browser's so words are not underlined twice.
      attributes: { spellcheck: "false" },
      dispatchTransaction(tr) {
        v.updateState(v.state.apply(tr));
        if (tr.docChanged) {
          setDirty(true);
          dirtyRef.current = true;
          // Autosave: a recovery snapshot 3 seconds after typing stops.
          clearTimeout(snapshotTimer.current);
          snapshotTimer.current = setTimeout(snapshot, 3000);
        }
        refresh();
      },
    });
    view.current = v;
    v.focus();
    refresh();
    return () => v.destroy();
  }, [file.base, file.fonts, save, mode, snapshot, spelling]);

  // Load the document's dictionary (offline, shipped with the app), then check the document.
  useEffect(() => {
    if (!spelling) return;
    let cancelled = false;
    void loadSpeller(documentLanguage(file.base)).then((s) => {
      if (cancelled) return;
      speller.current = s;
      if (view.current) recheck(view.current);
    }).catch(() => undefined);
    return () => { cancelled = true; };
  }, [file.base, spelling]);

  // Load any Office cloud fonts the document uses, then re-measure.
  useEffect(() => {
    let cancelled = false;
    void loadCloudFonts(fontUsage(file.base)).then((added) => {
      if (added && !cancelled) setFile((f) => ({ ...f, fonts: f.fonts + 1 }));
    });
    return () => { cancelled = true; };
  }, [file.base]);

  const load = (bytes: Uint8Array, name: string, location?: DocumentLocation) => {
    try {
      const { doc, warnings } = readDocx(bytes);
      setFile({ base: doc, name, location, warnings, fonts: 0 });
      setDirty(false);
    } catch (e) {
      alert(`Prime Word could not open ${name}:\n${(e as Error).message}`);
    }
  };

  /** Replacing the current document: ask first if it has unsaved changes, then drop its snapshot. */
  const discardCurrent = async () => {
    if (dirtyRef.current && !(await confirmAction(`Discard unsaved changes to ${fileRef.current.name}?`))) return false;
    clearTimeout(snapshotTimer.current);
    await deleteSnapshot(SESSION);
    return true;
  };

  const open = async () => {
    const opened = await openDocument();
    if (opened && (await discardCurrent())) load(opened.bytes, opened.name, opened.location);
  };

  const openRecent = async (path: string) => {
    if (!(await discardCurrent())) return;
    try {
      const opened = await openPath(path);
      load(opened.bytes, opened.name, opened.location);
    } catch (e) {
      alert(`Prime Word could not open ${path}:\n${String(e)}`);
    }
  };

  /** A new, unsaved document from a template (blank by default). */
  const newDocument = async (template?: Template) => {
    if (!(await discardCurrent())) return;
    const base = template ? template.create() : blankDocument();
    const name = template && template.id !== "blank" ? `${template.name.replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, " ").trim()}1.docx` : "Document1.docx";
    setFile({ base, name, warnings: [], fonts: 0 });
    setDirty(false);
  };

  /** Restores unsaved work from an earlier session (kept as unsaved: nothing is overwritten). */
  const restore = async (s: Snapshot) => {
    if (!(await discardCurrent())) return;
    setFile({ base: s.doc, name: s.name, warnings: s.warnings, fonts: 0, ...(s.path ? { location: { kind: "path" as const, path: s.path, modified: s.savedAt } } : {}) });
    setDirty(true);
    dirtyRef.current = true;
    await saveSnapshot({ ...s, id: SESSION, savedAt: Date.now() }); // now this session's
    await deleteSnapshot(s.id);
    setRecovery((r) => r.filter((x) => x.id !== s.id));
  };

  const discardRecovery = async (s: Snapshot) => {
    if (!(await confirmAction(`Permanently discard the unsaved changes to ${s.name}?`))) return;
    await deleteSnapshot(s.id);
    setRecovery((r) => r.filter((x) => x.id !== s.id));
  };

  // Offer unsaved work from a session that did not close cleanly.
  useEffect(() => { void orphanedSnapshots().then(setRecovery); }, []);

  // Desktop: closing the window asks about unsaved changes; a clean close removes the snapshot.
  useEffect(() => {
    if (!isDesktop) return;
    const pending = getCurrentWindow().onCloseRequested(async (event) => {
      if (dirtyRef.current && !(await confirmAction(`Close without saving ${fileRef.current.name}?`))) {
        event.preventDefault();
        return;
      }
      await deleteSnapshot(SESSION);
    });
    return () => { void pending.then((unlisten) => unlisten()); };
  }, []);

  const exportToPdf = async () => {
    const v = view.current;
    if (!v || !measurerRef.current) return;
    try {
      if (await exportPdf(fromProseMirror(v.state.doc, fileRef.current.base), measurerRef.current, fileRef.current.name)) flash("PDF exported");
    } catch (e) {
      alert(`Prime Word could not export the PDF:\n${(e as Error).message}`);
    }
  };

  // Ctrl+P from the editor.
  useEffect(() => {
    const onPrint = () => void printRef.current();
    window.addEventListener(PRINT_EVENT, onPrint);
    return () => window.removeEventListener(PRINT_EVENT, onPrint);
  }, []);

  const print = async () => {
    const v = view.current;
    if (!v || !measurerRef.current) return;
    try {
      flash("Preparing to print…");
      await printDocument(fromProseMirror(v.state.doc, fileRef.current.base), measurerRef.current, fileRef.current.name);
    } catch (e) {
      alert(`Prime Word could not print:\n${(e as Error).message}`);
    }
  };

  // Desktop: open the document the app was launched with ("Open with", double-click).
  useEffect(() => {
    if (!isDesktop) return;
    void invoke<[string, number[]] | null>("startup_document").then((doc) => {
      if (!doc) return setStartupDone(true);
      void invoke<number | null>("document_modified", { path: doc[0] }).then((modified) => {
        load(Uint8Array.from(doc[1]), doc[0].split(/[\\/]/).pop()!, { kind: "path", path: doc[0], ...(modified ? { modified } : {}) });
        setStartupDone(true);
      });
    });
  }, []);

  // Dev self-tests (browser only): ?selftest=paste | copy | spell, after the document has loaded.
  useEffect(() => {
    const test = new URLSearchParams(location.search).get("selftest");
    if (!test || isDesktop || !layout || layoutOf !== file.base || automationDone.current) return;
    automationDone.current = true;
    void import("./selftest").then(({ pasteSelfTest, copySelfTest, spellSelfTest }) => {
      const v = view.current;
      if (!v) return;
      if (test === "paste") pasteSelfTest(v, () => fileRef.current.base);
      if (test === "copy") copySelfTest(v, () => fileRef.current.base);
      if (test === "spell") spellSelfTest(v);
    });
  }, [layout, layoutOf, file.base]);

  // Dev convenience (browser only): ?file=<absolute path> opens a document from disk.
  useEffect(() => {
    const path = new URLSearchParams(location.search).get("file");
    if (path && !isDesktop) {
      void fetch(`/@fs/${path.replace(/\\/g, "/")}`).then(async (r) => {
        if (r.ok) load(new Uint8Array(await r.arrayBuffer()), path.split(/[\\/]/).pop()!);
      });
    }
  }, []);

  useEffect(() => {
    const warn = (e: BeforeUnloadEvent) => { if (dirtyRef.current && !isDesktop) e.preventDefault(); };
    window.addEventListener("beforeunload", warn);
    return () => window.removeEventListener("beforeunload", warn);
  }, [dirty]);

  useEffect(() => {
    const title = `${dirty ? "• " : ""}${file.name} — Prime Word`;
    document.title = title;
    if (isDesktop) void getCurrentWindow().setTitle(title);
  }, [file.name, dirty]);

  // Automation (desktop command line): --pw-type=<text> edits the opened document;
  // --pw-export-pdf=<path> exports it as PDF and closes the app. Used by end-to-end tests.
  useEffect(() => {
    // Wait for a layout of the document the app was started with (not the initial blank one).
    if (!isDesktop || !startupDone || !layout || layoutOf !== file.base) return;
    const done = automationDone.current;
    if (done) return;
    automationDone.current = true;
    void invoke<string[]>("automation_args").then(async (args) => {
      const v = view.current;
      if (!v || !measurerRef.current) return;
      for (const a of args) {
        if (a.startsWith("--pw-type=")) v.dispatch(v.state.tr.insertText(a.slice("--pw-type=".length), 1));
        if (a.startsWith("--pw-export-pdf=")) {
          const { bytes } = await makePdf(fromProseMirror(v.state.doc, fileRef.current.base), measurerRef.current);
          await invoke("save_document", bytes, { headers: { "x-path": encodeURIComponent(a.slice("--pw-export-pdf=".length)) } });
          await getCurrentWindow().destroy();
        }
      }
    });
  }, [layout, layoutOf, startupDone, file.base]);

  const printRef = useRef(print);
  printRef.current = print;

  const v = view.current;
  const section = v ? pageSetup(v.state) : file.base.sections[0]?.props ?? {};
  const margins = section.margins ?? {};

  return (
    <div className="app">
      <Ribbon
        view={v} doc={file.base} mode={mode}
        onNew={() => setPickingTemplate(true)} onOpen={() => void open()} onOpenRecent={(p) => void openRecent(p)} onSave={(saveAs) => void save(saveAs)}
        onExportPdf={() => void exportToPdf()} onPrint={() => void print()}
        onFind={(replace) => setFind(replace ? "replace" : "find")}
        onHeaderFooter={() => setEditingHeaderFooter(true)}
        onParagraph={() => setEditingParagraph(true)}
        onMode={(m) => { setMode(m); if (m === "web") setLayout(undefined); }}
      />

      <main className="canvas" onContextMenu={(e) => {
        const v2 = view.current;
        if (!v2 || !speller.current || !(e.target as HTMLElement).closest(".pw-misspelled")) return;
        const pos = v2.posAtCoords({ left: e.clientX, top: e.clientY })?.pos;
        const issue = pos === undefined ? undefined : misspellingAt(v2, pos);
        if (!issue) return;
        e.preventDefault();
        setSpellMenu({ issue, x: e.clientX, y: e.clientY });
      }} onKeyDown={(e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "o") { e.preventDefault(); void open(); }
      }}>
        {recovery.length > 0 && (
          <div className="recovery">
            <strong>Prime Word recovered unsaved work.</strong>
            {recovery.map((s) => (
              <div key={s.id} className="recovery-item">
                <span>{s.name} — last autosaved {new Date(s.savedAt).toLocaleString()}</span>
                <button className="primary" onClick={() => void restore(s)}>Restore</button>
                <button onClick={() => void discardRecovery(s)}>Discard</button>
              </div>
            ))}
            <button className="link" onClick={() => setRecovery([])}>Decide later</button>
          </div>
        )}
        {find && <FindPanel view={v} replace={find === "replace"} onClose={() => setFind(null)} />}
        <div className="zoomed" style={{ zoom: zoom / 100 }}>
          {mode === "print" ? (
            <PageSheets layout={layout} doc={file.base} gap={PAGE_GAP} overflow={overflow}>
              <div ref={mount} className="pw-editor paged" />
            </PageSheets>
          ) : (
            <div className="page" style={{
              width: `${section.pageWidth ?? 612}pt`,
              minHeight: `${section.pageHeight ?? 792}pt`,
              padding: `${margins.top ?? 72}pt ${margins.right ?? 72}pt ${margins.bottom ?? 72}pt ${margins.left ?? 72}pt`,
            }}>
              <div ref={mount} className="pw-editor" />
            </div>
          )}
        </div>
      </main>

      <footer className="status">
        <span>{file.name}</span>
        {mode === "print" && layout && <span>{layout.pages.length} page{layout.pages.length > 1 ? "s" : ""}</span>}
        <span>{v ? wordCount(v.state.doc) : 0} words</span>
        <button className="link" onClick={() => setSpelling((on) => !on)} title="Turn spell checking on or off">
          {spelling ? `Spelling: ${v ? (spellKey.getState(v.state)?.issues.length ?? 0) : 0} to review` : "Spelling: off"}
        </button>
        <span className={dirty ? "unsaved" : "saved"}>{dirty ? "Unsaved changes (autosaved for recovery)" : "Saved"}</span>
        {notice && <span className="notice">{notice}</span>}
        {file.warnings.length > 0 && (
          <button className="link" onClick={() => setShowWarnings((s) => !s)}>
            {file.warnings.length} item{file.warnings.length > 1 ? "s" : ""} kept unchanged
          </button>
        )}
        <span className="zoom">
          <button className="link" onClick={() => setZoom((z) => Math.max(25, z - 10))} title="Zoom out">−</button>
          <input type="range" min={25} max={200} step={5} value={zoom} onChange={(e) => setZoom(Number(e.target.value))} title="Zoom" />
          <button className="link" onClick={() => setZoom((z) => Math.min(200, z + 10))} title="Zoom in">+</button>
          <span className="zoom-value">{zoom}%</span>
        </span>
      </footer>
      {spellMenu && v && speller.current && (
        <SpellMenu view={v} speller={speller.current} issue={spellMenu.issue} x={spellMenu.x} y={spellMenu.y}
          ignored={ignoredWords.current} onClose={() => setSpellMenu(undefined)} />
      )}
      {pickingTemplate && (
        <TemplatePicker onClose={() => setPickingTemplate(false)} onOpenRecent={(p) => { setPickingTemplate(false); void openRecent(p); }} onPick={(t) => { setPickingTemplate(false); void newDocument(t); }} />
      )}
      {editingParagraph && v && <ParagraphDialog view={v} doc={file.base} onClose={() => setEditingParagraph(false)} />}
      {editingHeaderFooter && v && measurerRef.current && (
        <HeaderFooterEditor view={v} doc={file.base} measurer={measurerRef.current} onClose={() => setEditingHeaderFooter(false)} />
      )}
      {showWarnings && (
        <aside className="warnings">
          <h3>Kept exactly as they were</h3>
          <p>Prime Word preserves these and saves them back unchanged; it cannot edit them yet.</p>
          <ul>{file.warnings.map((w) => <li key={w}>{w}</li>)}</ul>
        </aside>
      )}
    </div>
  );
}
