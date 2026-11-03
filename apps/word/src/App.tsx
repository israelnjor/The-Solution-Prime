import { readDocx, writeDocx } from "@prime/docx";
import {
  blankDocument, changeLevel, createRender, createSchema, exitEmptyListItem, findPlugin, fromProseMirror,
  insertPageBreak, moveCell, PAGE_GAP, pageSetup, pagination, paragraphStyles, pastePlugin, toggleRun, toProseMirror,
  wordCount, floatEditing, reviewPlugin, blockHandle, moveBlock, insertImage, setPictureWrap, autoFormat,
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
import { type Command, EditorState, NodeSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { Profiler, useCallback, useEffect, useReducer, useRef, useState } from "react";
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
import { LINK_EVENT, type Markup, Ribbon } from "./Ribbon";
import { CommentsPane } from "./CommentsPane";
import { Home } from "./Home";
import type { ConversionResult } from "@prime/pdf-import";
import { PdfConverter } from "./PdfConverter";
import { PictureEditor } from "./PictureEditor";
import { LetterheadDesigner } from "./LetterheadDesigner";
import { type SavedSignature, SignatureDialog } from "./SignatureDialog";
import { BlocksDialog } from "./BlocksDialog";
import { AutoCorrectDialog, loadAutoCorrect } from "./AutoCorrectDialog";
import { saveTemplate } from "./MyTemplates";

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

/** Dev self-test run (?selftest=…): turns on the "pw:" performance timings the tests read. */
const SELFTEST = new URLSearchParams(location.search).has("selftest");
if (SELFTEST) (globalThis as { primeTiming?: boolean }).primeTiming = true;

/** Word count, remembered per document version (the status bar asks on every render). */
const wordCounts = new WeakMap<object, number>();
function countWords(doc: Parameters<typeof wordCount>[0]): number {
  let n = wordCounts.get(doc);
  if (n === undefined) wordCounts.set(doc, (n = wordCount(doc)));
  return n;
}

export function App() {
  const mount = useRef<HTMLDivElement>(null);
  const view = useRef<EditorView | null>(null);
  const [file, setFile] = useState<Loaded>({ base: blankDocument(), name: "Document1.docx", warnings: [], fonts: 0 });
  const [dirty, setDirty] = useState(false);
  const [showWarnings, setShowWarnings] = useState(false);
  /** Print Layout shows pages as they will print; Web Layout is one continuous column. */
  // (Dev: ?mode=web opens in Web Layout.)
  const [mode, setMode] = useState<"print" | "web">(() => (new URLSearchParams(location.search).get("mode") === "web" ? "web" : "print"));
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
  const [editingPicture, setEditingPicture] = useState(false);
  const [designingLetterhead, setDesigningLetterhead] = useState(false);
  const [signing, setSigning] = useState(false);
  const [choosingBlock, setChoosingBlock] = useState(false);
  const [autoCorrecting, setAutoCorrecting] = useState(false);
  const autoCorrect = useRef(loadAutoCorrect());
  // Bumped when a template is saved, so the template lists reload.
  const [templatesVersion, setTemplatesVersion] = useState(0);
  // PDF → Word: the PDF being converted, then what the converted document came from.
  const [converter, setConverter] = useState<{ initial?: { bytes: Uint8Array; name: string } }>();
  const [conversion, setConversion] = useState<{ from: string; warnings: string[] }>();
  // The home screen: at start (unless a document was given, or a dev test runs), and from the Prime Word button.
  const [home, setHome] = useState(() => !/[?&](selftest|file)=/.test(location.search));
  const [hasDocument, setHasDocument] = useState(false);
  const leaveHome = () => { setHome(false); setHasDocument(true); };
  // Review: how tracked changes are shown, and the comments pane (writing = a new comment is being typed).
  const [markup, setMarkup] = useState<Markup>("all");
  const [comments, setComments] = useState<"closed" | "open" | "writing">("closed");
  const baseOf = useCallback(() => fileRef.current.base, []);
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
        // Before the keymaps: it handles Enter (to correct the last word) and Backspace after a change.
        autoFormat(() => autoCorrect.current, () => fileRef.current.base),
        keymap({
          "Mod-z": undo, "Mod-y": redo, "Mod-Shift-z": redo,
          "Mod-b": toggleRun("bold"), "Mod-i": toggleRun("italic"), "Mod-u": toggleRun("underline"),
          "Mod-s": command(() => void save()),
          "Mod-f": command(() => setFind("find")),
          "Mod-h": command(() => setFind("replace")),
          "Mod-k": command(() => window.dispatchEvent(new Event(LINK_EVENT))),
          "Mod-p": command(() => window.dispatchEvent(new Event(PRINT_EVENT))),
          "Mod-Enter": insertPageBreak,
          // Word's "move paragraph up / down".
          "Alt-Shift-ArrowUp": moveBlock(-1), "Alt-Shift-ArrowDown": moveBlock(1),
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
        floatEditing(),
        blockHandle(),
        reviewPlugin(),
        ...(spelling ? [spellPlugin(() => speller.current, ignoredWords.current)] : []),
        ...(mode === "print" ? [pagination({ base: () => fileRef.current.base, measurer, onLayout: (l) => { setLayout(l); setLayoutOf(fileRef.current.base); }, onFit: (next) => setOverflow((prev) => (prev.size === next.size && [...next].every((p) => prev.has(p)) ? prev : next)) })] : []),
      ],
    });
    const v = new EditorView(mount.current, {
      state,
      // Our own checker underlines misspellings; turn off the browser's so words are not underlined twice.
      // For screen readers: the document is one multi-line text field.
      attributes: { spellcheck: "false", role: "textbox", "aria-multiline": "true", "aria-label": "Document" },
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
      leaveHome();
      setConversion(undefined);
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

  /** A converted PDF opens as a new, unsaved Word document (the PDF itself is not changed). */
  const openConverted = async (result: ConversionResult, pdfName: string) => {
    if (!(await discardCurrent())) return;
    const { doc, warnings } = result;
    setFile({ base: doc, name: pdfName.replace(/\.pdf$/i, ".docx"), warnings: [], fonts: 0 });
    setDirty(true);
    dirtyRef.current = true;
    setConversion({ from: pdfName, warnings });
    setConverter(undefined);
    leaveHome();
    if (SELFTEST) {
      // Dev self-test: what the editor shows once the converted document is laid out.
      setTimeout(() => console.info(`[selftest] align ${JSON.stringify({
        converted: pdfName, blocks: doc.sections[0].blocks.length, warnings,
        sheets: document.querySelectorAll(".pw-sheet").length,
        tables: document.querySelectorAll(".ProseMirror table").length,
        text: document.querySelector(".ProseMirror")?.textContent?.slice(0, 160),
      })}`), 3000);
    }
  };

  /**
   * A saved signature at the cursor: a floating picture in front of the text (about 150pt wide),
   * selected, so it can be dragged onto the signature line at once.
   */
  const insertSignature = (s: SavedSignature) => {
    const v2 = view.current;
    if (!v2) return;
    const width = 150 * (96 / 72); // insertImage takes pixels at 96 dpi
    insertImage(() => fileRef.current.base, {
      data: s.data.png, contentType: "image/png", width, height: (width / s.data.width) * s.data.height, name: s.name,
    })(v2.state, v2.dispatch);
    const at = v2.state.selection.from - 1;
    if (v2.state.doc.nodeAt(at)?.type.name !== "image") return;
    v2.dispatch(v2.state.tr.setSelection(NodeSelection.create(v2.state.doc, at)));
    setPictureWrap("inFront")(v2.state, v2.dispatch);
    v2.focus();
  };

  /** Saves a converted PDF straight to a .docx, without opening it. */
  const saveConverted = async (result: ConversionResult, pdfName: string) => {
    try {
      const saved = await saveDocument(writeDocx(result.doc), pdfName.replace(/\.pdf$/i, ".docx"));
      if (saved) flash(`Saved ${saved.name}`);
    } catch (e) {
      alert(`Prime Word could not save the converted document:\n${(e as Error).message}`);
    }
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
    leaveHome();
    setConversion(undefined);
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
    leaveHome();
  };

  const discardRecovery = async (s: Snapshot) => {
    if (!(await confirmAction(`Permanently discard the unsaved changes to ${s.name}?`))) return;
    await deleteSnapshot(s.id);
    setRecovery((r) => r.filter((x) => x.id !== s.id));
  };

  // F6 / Shift+F6: move the keyboard between the ribbon, the document and the status bar (as in Word).
  useEffect(() => {
    const regions = [
      () => document.querySelector<HTMLElement>(".toolbar button:not(:disabled), .toolbar select"),
      () => view.current?.dom as HTMLElement | undefined,
      () => document.querySelector<HTMLElement>(".status button, .status input"),
    ];
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "F6") return;
      e.preventDefault();
      const at = document.activeElement;
      const current = at?.closest(".toolbar, .tabs") ? 0 : at?.closest(".status") ? 2 : 1;
      const next = (current + (e.shiftKey ? 2 : 1)) % 3;
      regions[next]()?.focus();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

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

  // Dev self-tests (browser only): ?selftest=paste | copy | spell | perf | align | scroll | lines | floatdrag | review | keyboard, after the document has loaded.
  useEffect(() => {
    const test = new URLSearchParams(location.search).get("selftest");
    if (!test || isDesktop || (mode === "print" && (!layout || layoutOf !== file.base)) || automationDone.current) return;
    // With ?file=, wait for that document (not the blank one shown first).
    const wanted = new URLSearchParams(location.search).get("file");
    if (wanted && !wanted.replace(/\\/g, "/").endsWith(`/${file.name}`)) return;
    automationDone.current = true;
    void import("./selftest").then(({ pasteSelfTest, copySelfTest, spellSelfTest, perfSelfTest, alignSelfTest, scrollSelfTest, linesSelfTest, floatDragSelfTest, reviewSelfTest, keyboardSelfTest, pictureSelfTest, rotateSelfTest, dragBlockSelfTest, letterheadSelfTest, signatureSelfTest, blocksSelfTest, autoCorrectSelfTest }) => {
      const v = view.current;
      if (!v) return;
      if (test === "paste") pasteSelfTest(v, () => fileRef.current.base);
      if (test === "copy") copySelfTest(v, () => fileRef.current.base);
      if (test === "spell") spellSelfTest(v);
      if (test === "perf") perfSelfTest(v);
      if (test === "align") alignSelfTest();
      if (test === "scroll") scrollSelfTest();
      if (test === "floatdrag") floatDragSelfTest();
      if (test === "review") reviewSelfTest(v);
      if (test === "keyboard") keyboardSelfTest();
      if (test === "picture") pictureSelfTest(v);
      if (test === "rotate") rotateSelfTest(v);
      if (test === "dragblock") dragBlockSelfTest(v);
      if (test === "letterhead") letterheadSelfTest();
      if (test === "signature") signatureSelfTest(v);
      if (test === "blocks") blocksSelfTest(v);
      if (test === "autocorrect") autoCorrectSelfTest(v);
      if (test === "lines" && measurerRef.current) linesSelfTest(v, () => fileRef.current.base, measurerRef.current);
    });
  }, [layout, layoutOf, file.base]);

  // Dev convenience (browser only): ?file=<absolute path> opens a document from disk.
  useEffect(() => {
    const path = new URLSearchParams(location.search).get("file");
    if (path && !isDesktop) {
      void fetch(`/@fs/${path.replace(/\\/g, "/")}`).then(async (r) => {
        if (!r.ok) return;
        const bytes = new Uint8Array(await r.arrayBuffer());
        const name = path.split(/[\\/]/).pop()!;
        if (/\.pdf$/i.test(name)) setConverter({ initial: { bytes, name } });
        else load(bytes, name);
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
    <Profiler id="app" onRender={(_id, _phase, ms) => { if (SELFTEST) performance.measure("pw:render", { start: performance.now() - ms, end: performance.now() }); }}>
    <div className="app">
      <Ribbon
        view={v} doc={file.base} mode={mode}
        onNew={() => setPickingTemplate(true)} onOpen={() => void open()} onOpenRecent={(p) => void openRecent(p)} onSave={(saveAs) => void save(saveAs)}
        onExportPdf={() => void exportToPdf()} onPrint={() => void print()}
        onFind={(replace) => setFind(replace ? "replace" : "find")}
        onHeaderFooter={() => setEditingHeaderFooter(true)}
        onParagraph={() => setEditingParagraph(true)}
        onMode={(m) => { setMode(m); if (m === "web") setLayout(undefined); }}
        markup={markup} onMarkup={setMarkup} onHome={() => setHome(true)} onConvertPdf={() => setConverter({})} onEditPicture={() => setEditingPicture(true)}
        onLetterhead={() => setDesigningLetterhead(true)}
        onSignature={() => setSigning(true)}
        onBlocks={() => setChoosingBlock(true)}
        onAutoCorrect={() => setAutoCorrecting(true)}
        onSaveTemplate={() => {
          if (!v) return;
          const name = window.prompt("Save as a template named:", fileRef.current.name.replace(/\.docx$/i, ""));
          if (!name?.trim()) return;
          void saveTemplate(fromProseMirror(v.state.doc, fileRef.current.base), name.trim())
            .then(() => { setTemplatesVersion((n) => n + 1); flash(`"${name.trim()}" saved to My templates`); })
            .catch((e) => alert(`Prime Word could not save the template:\n${(e as Error).message}`));
        }}
        onComments={() => setComments((c) => (c === "closed" ? "open" : "closed"))}
        onNewComment={() => {
          if (!v || v.state.selection.empty) { flash("Select the text you want to comment on."); return; }
          setComments("writing");
        }}
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
        {conversion && (
          <div className="conversion" role="status">
            <strong>Converted from {conversion.from}.</strong>{" "}
            Check the layout against the PDF; the PDF itself is unchanged. Save to keep this as a Word document.
            {conversion.warnings.length > 0 && <ul>{conversion.warnings.map((w) => <li key={w}>{w}</li>)}</ul>}
            <button className="primary" onClick={() => void save(true)}>Save as Word document…</button>
            <button className="link" onClick={() => setConversion(undefined)}>Dismiss</button>
          </div>
        )}
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
              <div ref={mount} className={`pw-editor paged pw-markup-${markup}`} />
            </PageSheets>
          ) : (
            <div className="page" style={{
              width: `${section.pageWidth ?? 612}pt`,
              minHeight: `${section.pageHeight ?? 792}pt`,
              padding: `${margins.top ?? 72}pt ${margins.right ?? 72}pt ${margins.bottom ?? 72}pt ${margins.left ?? 72}pt`,
            }}>
              <div ref={mount} className={`pw-editor pw-markup-${markup}`} />
            </div>
          )}
        </div>
      </main>
      {comments !== "closed" && v && (
        <CommentsPane view={v} base={baseOf} writing={comments === "writing"}
          onWritten={() => setComments("open")} onClose={() => setComments("closed")} />
      )}

      <footer className="status" aria-label="Status bar">
        <span>{file.name}</span>
        {mode === "print" && layout && <span>{layout.pages.length} page{layout.pages.length > 1 ? "s" : ""}</span>}
        <span>{v ? countWords(v.state.doc) : 0} words</span>
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
          <button className="link" onClick={() => setZoom((z) => Math.max(25, z - 10))} title="Zoom out" aria-label="Zoom out">−</button>
          <input type="range" min={25} max={200} step={5} value={zoom} onChange={(e) => setZoom(Number(e.target.value))} title="Zoom" aria-label="Zoom" />
          <button className="link" onClick={() => setZoom((z) => Math.min(200, z + 10))} title="Zoom in" aria-label="Zoom in">+</button>
          <span className="zoom-value">{zoom}%</span>
        </span>
      </footer>
      {spellMenu && v && speller.current && (
        <SpellMenu view={v} speller={speller.current} issue={spellMenu.issue} x={spellMenu.x} y={spellMenu.y}
          ignored={ignoredWords.current} onClose={() => setSpellMenu(undefined)} />
      )}
      {editingPicture && v && <PictureEditor view={v} base={baseOf} onClose={() => setEditingPicture(false)} />}
      {autoCorrecting && <AutoCorrectDialog onChange={(o) => { autoCorrect.current = o; }} onClose={() => { setAutoCorrecting(false); view.current?.focus(); }} />}
      {choosingBlock && view.current && <BlocksDialog view={view.current} base={() => fileRef.current.base} onClose={() => setChoosingBlock(false)} />}
      {signing && <SignatureDialog onClose={() => setSigning(false)} onInsert={(s) => { setSigning(false); insertSignature(s); }} />}
      {designingLetterhead && (
        <LetterheadDesigner view={hasDocument && !home ? v ?? undefined : undefined} base={baseOf}
          onNewDocument={(doc, name) => { setDesigningLetterhead(false); void newDocument({ id: "letterhead", name: name.replace(/\.docx$/, ""), description: "", create: () => doc }); }}
          onSaved={() => { setDesigningLetterhead(false); setTemplatesVersion((n) => n + 1); flash("Letterhead saved to My templates"); }}
          onClose={() => setDesigningLetterhead(false)} />
      )}
      {converter && (
        <PdfConverter measurer={() => measurerRef.current ?? new CanvasMeasurer()} initial={converter.initial}
          onOpenInEditor={(r, n) => void openConverted(r, n)} onSave={(r, n) => void saveConverted(r, n)}
          onClose={() => setConverter(undefined)} />
      )}
      {home && (
        <Home recovery={recovery} onPick={(t) => void newDocument(t)} onOpen={() => void open()} onOpenRecent={(p) => void openRecent(p)}
          onRestore={(s) => void restore(s)} onDiscard={(s) => void discardRecovery(s)}
          {...(hasDocument ? { onClose: () => setHome(false) } : {})}  onConvertPdf={() => setConverter({})}
          onLetterhead={() => setDesigningLetterhead(true)} templatesVersion={templatesVersion} />
      )}
      {pickingTemplate && (
        <TemplatePicker templatesVersion={templatesVersion} onClose={() => setPickingTemplate(false)} onOpenRecent={(p) => { setPickingTemplate(false); void openRecent(p); }} onPick={(t) => { setPickingTemplate(false); void newDocument(t); }} />
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
    </Profiler>
  );
}
