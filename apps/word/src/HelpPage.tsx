import { type ReactNode, useMemo, useState } from "react";

/**
 * Home → Help: a short guide to Prime Word. The everyday tasks, and above all what Prime Word does that
 * people would not think to look for (PDF to Word, letterheads, signatures, blocks). Searchable; kept in
 * the app so every update can improve it.
 */
interface Topic { group: string; title: string; summary: string; body: ReactNode; words?: string }

const K = ({ children }: { children: ReactNode }) => <kbd>{children}</kbd>;
const B = ({ children }: { children: ReactNode }) => <strong>{children}</strong>;

const TOPICS: Topic[] = [
  // --- Getting started ------------------------------------------------------------------------------
  {
    group: "Getting started", title: "Start or open a document", summary: "A blank page, a template, or a file you already have.",
    body: <ol>
      <li>On the home screen, click <B>New document</B>, or pick a template under <B>Start new</B>.</li>
      <li>To open a file, click <B>Open…</B> (<K>Ctrl+O</K>). Prime Word opens Word documents (.docx and older .doc), Rich Text (.rtf), OpenDocument (.odt) and plain text.</li>
      <li>You can also drop a file onto the Prime Word window, or right-click a .docx in File Explorer and choose <B>Open with → Prime Word</B>.</li>
      <li>Documents you used lately are under <B>Recent</B> on the home screen. Several documents stay open as tabs at the top.</li>
    </ol>,
  },
  {
    group: "Getting started", title: "Save, and never lose your work", summary: "Word files by default; your unsaved changes are kept as you type.",
    body: <>
      <p>Press <K>Ctrl+S</K> to save. Documents are saved as Word documents (.docx): anyone with Microsoft Word can open them as they are.
        <B> File → Save As</B> also saves as .rtf, .odt or .txt.</p>
      <p>Prime Word keeps your unsaved changes as you type. If it closes unexpectedly (a power cut, a crash), the home screen shows them under
        <B> Recovered</B> next time: restore the document and carry on, or discard it.</p>
    </>,
    words: "autosave recovery crash backup",
  },
  {
    group: "Getting started", title: "Find your way around", summary: "The ribbon, the home screen, and clicking where you want to type.",
    body: <ul>
      <li>The tabs across the top (<B>File, Home, Insert, Design, Layout, References, Review, View</B>) work as in Word. Tabs for tables, pictures, shapes and charts appear when you select one.</li>
      <li>Click <B>Prime Word</B> at the top left to go back to the home screen.</li>
      <li>Click anywhere on a page (beside the text, below it, in a margin) and the cursor goes to the nearest place you can type.</li>
      <li><K>F6</K> moves between the ribbon, the document and the status bar, for working without a mouse.</li>
      <li>The status bar at the bottom shows the page, the word count, spelling and grammar to review, and the zoom.</li>
    </ul>,
    words: "ribbon menu cursor zoom status bar keyboard",
  },

  // --- What Prime Word does for you -------------------------------------------------------------------
  {
    group: "Made for your documents", title: "Turn a PDF into a Word document", summary: "Edit a PDF as a normal document, even a scanned one.",
    body: <ol>
      <li>On the home screen, click <B>Convert PDF to Word…</B> (or <B>File → Convert PDF</B>, or drop a PDF onto the window).</li>
      <li>Check the first page shown, the number of pages and the size, then convert.</li>
      <li>The result opens as an editable document. Compare it with the PDF, then save it as a Word document. The PDF itself is never changed.</li>
    </ol>,
    words: "pdf convert scan scanned ocr text recognition",
  },
  {
    group: "Made for your documents", title: "Design a letterhead", summary: "Your organisation's letterhead, made once and used for every letter.",
    body: <ol>
      <li>Go to <B>Insert → Letterhead…</B> (or <B>Design a letterhead…</B> on the home screen).</li>
      <li>Fill in the name, tagline, address and contacts, add a logo, and choose the layout, colour and font. The preview shows the page as you go.</li>
      <li>Click <B>Save as template</B>. It appears under <B>My templates</B>: start each new letter from it.</li>
    </ol>,
    words: "letterhead logo header company",
  },
  {
    group: "Made for your documents", title: "Put a letterhead or template on a document you are writing", summary: "First page only, or every page, without retyping anything.",
    body: <ol>
      <li>Open the document and choose <B>Insert → Apply Template…</B> (also on the File tab).</li>
      <li>Pick a saved letterhead or template, or <B>From a file…</B> (any Word document or template).</li>
      <li>Choose where it goes: <B>the first page only</B> or <B>every page</B>, and what the following pages show (the template's own header, yours, or nothing). You can also take its page size and styles.</li>
      <li>The preview shows the first two pages. Click <B>Apply</B>: your text stays as it is.</li>
    </ol>,
    words: "template letterhead apply first page header footer",
  },
  {
    group: "Made for your documents", title: "Sign a document", summary: "Your handwritten signature, from a photo, ready to place on any document.",
    body: <ol>
      <li>Sign on plain white paper and take a photo of it (or scan it).</li>
      <li>Go to <B>Insert → Signature…</B> and choose the photo. Prime Word removes the paper and the shadows and keeps the ink, then saves the signature on this computer.</li>
      <li>From then on, <B>Insert → Signature…</B> and click your signature to put it on any document. Drag it into place, or drag a corner to resize it.</li>
    </ol>,
    words: "signature sign signed handwriting",
  },
  {
    group: "Made for your documents", title: "Signature blocks, addresses and stamps", summary: "Make them once; they go where they belong in a letter.",
    body: <>
      <p>Go to <B>Insert → Blocks…</B> and make, once: a <B>signature block</B> (closing, line, name, title), an <B>address</B> (your own or the recipient's) or a <B>stamp</B>. You can also keep any selected text as a block.</p>
      <p>Click a block to insert it. It goes where it belongs in a letter: a signature block after the last paragraph, your address at the top on the right,
        the recipient's address above “Dear…”. Or choose the place yourself: at the bottom of the last page, at the cursor, or in the header or footer.</p>
    </>,
    words: "blocks signature block address stamp closing yours faithfully",
  },
  {
    group: "Made for your documents", title: "Projects", summary: "Everything for one piece of work in one place.",
    body: <p>On the home screen, open <B>Projects</B> and click <B>New project…</B>. A project is a folder on your computer (a thesis, a client, a sale):
      its documents, PDFs and pictures stay together, where you and your backups can find them.</p>,
    words: "project folder thesis client",
  },

  // --- Writing and formatting --------------------------------------------------------------------------
  {
    group: "Writing and formatting", title: "Text, paragraphs and styles", summary: "Fonts, lists, alignment, spacing, and styles for a consistent look.",
    body: <ul>
      <li>The <B>Home</B> tab has fonts, sizes, colours, highlighting, bullets and numbering, alignment, spacing and borders.</li>
      <li>Use the styles (<B>Normal, Heading 1, Heading 2…</B>) for a consistent look and a table of contents. <B>Design</B> changes the look of the whole document at once.</li>
      <li>Copy formatting with the paintbrush (or <K>Ctrl+Shift+C</K>, then <K>Ctrl+Shift+V</K> on other text).</li>
      <li>Text pasted from Word or a web page keeps its formatting. The small <B>(Ctrl)</B> button after a paste offers <B>Keep Text Only</B>.</li>
    </ul>,
    words: "font bold italic bullets numbering heading style paste format painter",
  },
  {
    group: "Writing and formatting", title: "Shapes, text boxes and pictures", summary: "Draw shapes, type in them, and place pictures anywhere.",
    body: <ul>
      <li><B>Insert → Shapes</B>: pick a shape, then drag on the page to draw it (hold <K>Shift</K> for a perfect square, circle or straight line). A click gives the standard size.</li>
      <li>With a shape selected, just start typing to put text in it, or double-click it.</li>
      <li>The <B>Shape Format</B> tab has styles, fill, transparency, outline, dashes and arrows, rotate and flip, and the size. Drag the round handle to turn a shape, and the yellow diamond to reshape it (rounder corners, a longer arrow head, a speech bubble's pointer).</li>
      <li><B>Insert → Text Box</B> works the same way. <B>Insert → Picture</B> adds a picture: choose how text wraps around it, drag it, resize it, or crop and rotate it.</li>
      <li>Arrow keys nudge a selected shape or picture; <K>Delete</K> removes it.</li>
    </ul>,
    words: "shape draw arrow box picture image photo wrap rotate transparency",
  },
  {
    group: "Writing and formatting", title: "Tables and charts", summary: "Tables with Word's styles, and charts from your numbers.",
    body: <ul>
      <li><B>Insert → Table</B>. Click in a table for the <B>Table Design</B> (styles, shading, borders) and <B>Table Layout</B> (rows, columns, merge, sort) tabs.</li>
      <li><B>Insert → Chart…</B> makes a column, bar, line or pie chart from a small table of numbers. Double-click a chart to change its data.</li>
    </ul>,
    words: "table rows columns merge chart graph",
  },
  {
    group: "Writing and formatting", title: "Equations and symbols", summary: "Fractions, roots, powers, sums and Greek letters, saved as Word's own maths.",
    body: <ul>
      <li>Press <B>Alt+=</B> (or <B>Insert → Equation…</B>). Click the pieces you need (fraction, power, root, sum, integral, Greek letters…), or type them, and watch the preview as you go.</li>
      <li>Typing works like this: <code>x^{"{2}"}</code> for x², <code>\frac{"{a}{b}"}</code> for a fraction, <code>\sqrt{"{x}"}</code> for a root, <code>\alpha</code> for α.</li>
      <li>It sits in the line of text; tick <B>Display: on its own line, centred</B> for a standalone equation. Double-click an equation to change it.</li>
      <li>Equations open in Microsoft Word as real equations, and Word's equations open here.</li>
      <li>For a single character such as ©, ° or €, use <B>Insert → Symbol…</B>.</li>
    </ul>,
    words: "equation math maths mathematics formula fraction square root power exponent integral sum greek latex symbol alt=",
  },
  {
    group: "Writing and formatting", title: "Spelling, grammar, thesaurus and reading aloud", summary: "Checks as you type, and help finding the right word.",
    body: <ul>
      <li>Misspelled words are underlined in red, grammar in blue. Right-click one for suggestions. The status bar shows how many are left to review.</li>
      <li><K>Shift+F7</K> opens the thesaurus for the word at the cursor.</li>
      <li><B>Review → Read Aloud</B> reads the document to you, highlighting each word.</li>
    </ul>,
    words: "spelling grammar thesaurus synonyms read aloud proofread",
  },

  // --- Pages and documents ---------------------------------------------------------------------------------
  {
    group: "Pages and documents", title: "Page layout, headers, footers and page numbers", summary: "Margins, orientation, columns, and what repeats on every page.",
    body: <ul>
      <li>The <B>Layout</B> tab: margins, portrait or landscape, paper size, columns and breaks. <K>Ctrl+Enter</K> starts a new page.</li>
      <li><B>Insert → Header &amp; Footer</B>: type the header and footer, and add <B>Page number</B> or <B>Page X of Y</B>. A different first page is an option.</li>
      <li>The <B>Design</B> tab adds a page colour, page borders or a watermark (DRAFT, CONFIDENTIAL…).</li>
      <li><B>Insert → Delete Page</B> removes the page the cursor is on (<K>Ctrl+Z</K> brings it back).</li>
    </ul>,
    words: "margins orientation landscape columns header footer page numbers watermark",
  },
  {
    group: "Pages and documents", title: "Contents, footnotes and references", summary: "For reports, papers and theses.",
    body: <ul>
      <li>Give your headings the Heading styles (<K>Ctrl+Alt+1</K>, <K>2</K>, <K>3</K>), then <B>References → Table of Contents</B>. Update it after editing.</li>
      <li><B>Insert Footnote</B> (<K>Ctrl+Alt+F</K>) and <B>Insert Endnote</B> (<K>Ctrl+Alt+D</K>) number themselves.</li>
      <li>Captions and a table of figures, cross-references, and citations with a bibliography (APA, Harvard or IEEE) are on the <B>References</B> tab.</li>
    </ul>,
    words: "table of contents toc footnote endnote citation bibliography caption reference",
  },
  {
    group: "Pages and documents", title: "Comments, track changes and comparing", summary: "Review documents with others.",
    body: <ul>
      <li><B>Review → New Comment</B> on selected text. Comments show beside the page.</li>
      <li><B>Track Changes</B> (<K>Ctrl+Shift+E</K>) marks every insertion and deletion; <B>Accept</B> or <B>Reject</B> them one by one or all at once.</li>
      <li><B>Review → Compare…</B> shows what changed between two versions of a document.</li>
    </ul>,
    words: "comment track changes review compare accept reject",
  },

  // --- Sharing ----------------------------------------------------------------------------------------------
  {
    group: "Sharing and printing", title: "PDF, printing and sharing", summary: "Send it, print it, or lock it with a password.",
    body: <ul>
      <li><B>File → Export PDF</B> makes a PDF of the document.</li>
      <li><B>File → Print</B> (<K>Ctrl+P</K>): choose the pages and see a preview; the printer, copies and two-sided printing are in Windows' print window.</li>
      <li><B>Share</B> (top right) sends a copy, as a Word document or a PDF, through Windows' share panel: email, WhatsApp, Nearby Share and the like.</li>
      <li><B>File → Protect</B> sets a password. Microsoft Word opens the document with the same password.</li>
    </ul>,
    words: "pdf export print share email whatsapp password protect",
  },
  {
    group: "Sharing and printing", title: "Updates", summary: "Prime Word keeps itself up to date.",
    body: <p>Once a day Prime Word checks for a newer version. When there is one, the home screen shows <B>Update available</B>: open it, read what is new, and click
      <B> Install and restart</B>. <B>About Prime Word</B> (bottom left of the home screen) shows your version and has <B>Check for updates</B>. Checking sends nothing about you or your documents.</p>,
    words: "update version about new",
  },
];

const SHORTCUTS: [string, string][] = [
  ["Ctrl+S", "Save"], ["Ctrl+O", "Open"], ["Ctrl+P", "Print"], ["Ctrl+Z / Ctrl+Y", "Undo / Redo"],
  ["Ctrl+F", "Find"], ["Ctrl+H", "Replace"], ["Ctrl+K", "Insert a link"],
  ["Ctrl+B / I / U", "Bold / Italic / Underline"], ["Ctrl+D", "Font…"], ["Ctrl+Space", "Clear character formatting"],
  ["Ctrl+E / L / R / J", "Centre / Left / Right / Justify"], ["Ctrl+1 / 2 / 5", "Single / Double / 1.5 line spacing"],
  ["Ctrl+Shift+L", "Bullets"], ["Ctrl+Alt+1 / 2 / 3", "Heading 1 / 2 / 3"], ["Ctrl+Shift+N", "Normal style"],
  ["Shift+F3", "Change case"], ["Ctrl+Shift+C / V", "Copy / paste formatting"],
  ["Alt+=", "Equation"], ["Ctrl+Enter", "Page break"], ["Ctrl+Alt+F / D", "Footnote / Endnote"], ["Ctrl+Shift+E", "Track changes"],
  ["Shift+F7", "Thesaurus"], ["Ctrl+Shift+8", "Show formatting marks"], ["F6", "Move between ribbon, document and status bar"], ["F1", "This help"],
];

export function HelpPage({ onFeedback }: { onFeedback(): void }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState<string>();
  const q = query.trim().toLowerCase();
  const found = useMemo(() => TOPICS.filter((t) => !q || `${t.title} ${t.summary} ${t.words ?? ""}`.toLowerCase().includes(q)), [q]);
  const groups = [...new Set(found.map((t) => t.group))];
  return (
    <main className="home-main">
      <div className="home-content help-page">
        <div className="home-title">
          <h1>Help</h1>
          <input className="help-search" type="search" placeholder="Search help (e.g. signature, PDF, page numbers)" aria-label="Search help"
            value={query} onChange={(e) => setQuery(e.target.value)} autoFocus />
        </div>
        <p className="home-hint">Prime Word works much like Microsoft Word. These are the everyday tasks, and the things Prime Word does that you might not think to look for.
          Something not covered, not working, or hard to use? <button className="link" onClick={onFeedback}>Send feedback</button></p>
        {groups.map((g) => (
          <section key={g} className="help-group" aria-label={g}>
            <h2>{g}</h2>
            {found.filter((t) => t.group === g).map((t) => (
              <details key={t.title} className="help-topic" open={!!q || open === t.title}
                onToggle={(e) => { if ((e.target as HTMLDetailsElement).open) setOpen(t.title); else if (open === t.title) setOpen(undefined); }}>
                <summary><strong>{t.title}</strong><span>{t.summary}</span></summary>
                <div className="help-body">{t.body}</div>
              </details>
            ))}
          </section>
        ))}
        {!found.length && <p className="recent-empty">Nothing in the guide matches “{query}”. Try another word.</p>}
        {!q && (
          <section className="help-group" aria-label="Keyboard shortcuts">
            <h2>Keyboard shortcuts</h2>
            <table className="help-keys"><tbody>
              {SHORTCUTS.map(([keys, what]) => <tr key={keys}><td><kbd>{keys}</kbd></td><td>{what}</td></tr>)}
            </tbody></table>
          </section>
        )}
      </div>
    </main>
  );
}
