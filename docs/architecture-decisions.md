# Architecture Decision Records — Prime Word

Revision 2 · 2026-10-31 · replaces revision 1 (the "adopt an editor SDK" approach).

**Strategy:** Prime Productivity owns the infrastructure that makes Prime Word valuable — the
document model, Word compatibility, page layout, PDF reconstruction and storage. Third-party code
is used only for mature, low-level building blocks under permissive licences.

**Constraint:** one developer. Every decision below favours fewer languages, fewer moving parts,
and work that can be verified automatically.

Status key: **Accepted** — strategic, in force. **Proposed** — direction chosen, pending the named
experiment. **Deferred** — decided later.

| ADR | Decision | Status |
|---|---|---|
| [000](#adr-000--own-vs-borrow-rule) | Own vs borrow rule | Accepted |
| [001](#adr-001--dependency-licence-policy) | Permissive licences only | Accepted |
| [002](#adr-002--languages-typescript-core-minimal-rust) | TypeScript core, minimal Rust | Accepted |
| [003](#adr-003--desktop-shell-tauri) | Desktop shell: Tauri | Accepted |
| [004](#adr-004--prime-document-model-pdm) | Prime Document Model (PDM) | Accepted |
| [005](#adr-005--editing-engine-prosemirror-used-directly) | Editor: ProseMirror, used directly | Proposed |
| [006](#adr-006--docx-compatibility-engine-built-in-house) | DOCX engine built in-house | Accepted |
| [007](#adr-007--layout-engine-built-in-house-pdf-first) | Layout engine built in-house, PDF first | Proposed |
| [008](#adr-008--pdf-export-and-printing) | PDF export and printing | Proposed |
| [009](#adr-009--pdf-to-docx-and-ocr) | PDF→DOCX and OCR | Proposed |
| [010](#adr-010--local-storage-and-recovery) | Local storage and recovery | Proposed |
| [011](#adr-011--windows-first-platform-neutral-core) | Windows first, platform-neutral core | Accepted |
| [012](#adr-012--ai-integration) | AI integration | Deferred |
| [013](#adr-013--evidence-driven-testing) | Evidence-driven testing | Accepted |

---

## ADR-000 — Own vs borrow rule

**Decision.** Build a component when it encodes Prime Word's product knowledge or behaviour users
will notice. Borrow it when all of the following hold: it is low-level, mature, permissively
licensed, and rebuilding it would take years without making the product different.

| Own (build) | Borrow (permissive) |
|---|---|
| Document model, DOCX read/write, layout and pagination, PDF page output, PDF→DOCX reconstruction, storage and recovery, all product UI and editing behaviour | Text input/selection/IME/undo core (ProseMirror), desktop shell (Tauri), database (SQLite), PDF file writing and font embedding (pdf-lib), PDF text extraction (pdf.js), OCR recognition (Tesseract), zip/XML parsing |

**Not building:** browser input handling, IME, accessibility plumbing, OCR models, font
rasterisation, a PDF parser.

**Consequence.** Every borrowed dependency sits behind our own interface, so it can be replaced
without touching product code.

## ADR-001 — Dependency licence policy

**Decision.** Shipped dependencies must be MIT, BSD, Apache-2.0, ISC, zlib or public domain.
GPL, LGPL-static, AGPL and SSPL are excluded from the shipped product. Development-only tools
(e.g. the Python test harness) may use other licences but never ship.

**Why.** Prime Word is a closed-source commercial product; this removes licence fees and
copyleft obligations.

**Rules out:** Tiptap Pro, PyMuPDF, pdf2docx, OnlyOffice and other AGPL editors/converters.

**Enforcement.** A licence check runs in CI over the dependency tree once the app exists.

## ADR-002 — Languages: TypeScript core, minimal Rust

**Decision.** The document model, DOCX engine, layout engine, PDF output and PDF reconstruction
are pure TypeScript packages with no DOM or OS dependencies. Rust is used only for what the
webview cannot do: file system access, atomic writes, SQLite, printing, font enumeration.

**Why.** A solo developer should work mostly in one language. Pure TypeScript runs in the
editor, in Node tests, and later on the web unchanged.

**Alternatives.** Rust core compiled to WASM (faster, but two languages for every feature);
Python sidecar (ruled out for the shipped product: packaging weight and licence exposure).

## ADR-003 — Desktop shell: Tauri

**Decision.** Tauri 2 with React + TypeScript UI.

**Why.** Small installer, scoped native permissions, Rust backend for the few native needs.
On Windows it renders with WebView2 (Chromium), giving consistent layout and measurement.

**Risk.** On macOS/Linux Tauri uses WebKit. Because our layout engine (ADR-007) computes pages
itself from font metrics instead of relying on browser layout, engine differences matter less —
but this must be re-tested before the cross-platform phase. Electron remains the fallback.

## ADR-004 — Prime Document Model (PDM)

**Decision.** A versioned, JSON-serialisable model owned by us and independent of the editor:
document → sections (page size, margins, orientation, headers/footers) → blocks (paragraph,
heading, list item, table, image, page break) → inline runs with marks. Styles are first-class
and resolved by inheritance, as in Word.

- The ProseMirror schema is **derived from** the PDM, not the other way round.
- Every PDM document carries a `schemaVersion`; migrations are explicit functions.
- Content the PDM cannot represent is carried as an opaque **preserved fragment** with enough
  information to write it back (see ADR-006).

**Why.** The model is the heart of the product: DOCX, layout, PDF, storage and a future web app
all speak it. Owning it keeps the editor replaceable.

**Native file format.** `.docx` is the user-facing format. PDM JSON is used internally for
autosave snapshots and recovery (ADR-010).

## ADR-005 — Editing engine: ProseMirror, used directly

**Decision.** Use ProseMirror (MIT) directly rather than Tiptap, with our own schema, commands,
input rules, keymaps and React toolbar.

**Why.** We need tight control of the boundary between the editor and our layout engine; an
extra abstraction layer gets in the way, and Tiptap's main paid value-adds (pagination, DOCX) are
exactly what we are building ourselves. Tiptap's open-source extensions remain useful as
reference implementations.

**Validation (Experiment 3).** A PDM-derived schema can edit fixtures A–D after import, and
transactions map cleanly back to PDM without loss.

## ADR-006 — DOCX compatibility engine, built in-house

**Decision.** Our own OOXML reader (DOCX → PDM) and writer (PDM → DOCX), following ECMA-376,
using only zip and XML parsing libraries.

- **Reader:** resolves styles, numbering, theme fonts, sections and relationships into PDM.
- **Writer:** emits clean, valid OOXML that Word opens without repair.
- **Preservation:** unsupported elements (e.g. footnotes, equations, tracked changes in v1) are
  kept as preserved fragments and written back unchanged where their position still exists.
  Anything dropped produces a user-visible warning — never silent loss.

**Supported feature matrix** grows release by release; it is defined by the fixtures, not by
promises.

**Validation (Experiment 1).** Candidate `prime-docx` in `experiments/e1-docx-roundtrip` must
score all ✅ on fixtures A–D (zero CRITICAL/MAJOR findings) before ADR-007 work starts.

## ADR-007 — Layout engine, built in-house, PDF first

**Decision.** A pure layout module: `layout(PDM, fontMetrics) → pages → lines → positioned boxes`.
It handles line breaking (UAX #14), paragraph splitting, keep-with-next / widow-orphan control,
table row splitting with repeating header rows, headers/footers, sections and page-number fields.

**Sequencing — one engine, two consumers:**
1. **PDF export and print** consume it first — no live editing, fully testable.
2. **The paginated editor view** consumes the same output later, so screen and PDF cannot
   disagree.

**Text measurement.** Behind an interface: browser canvas measurement first (zero dependency),
HarfBuzz/WASM (MIT) if shaping accuracy requires it.

**Risk.** The hardest component in the product. Editing a paragraph that is split across a page
boundary is the specific unsolved problem; it is tackled only after (1) is proven.

**Validation (Experiment 2).** Rendering fixtures A–D matches Word's page counts (2, 2, 5, 6) and
page-break positions within an agreed tolerance, measured against the Word baselines the harness
already produces.

## ADR-008 — PDF export and printing

**Decision.** Our layout output is drawn to PDF with pdf-lib (MIT) + fontkit for font embedding
and subsetting. Printing sends the same PDF to the OS print dialog via the Rust backend.

**Why.** One rendering path for export, print and preview.

**Validation.** Exported PDFs are text-searchable, embed fonts, and match the layout engine's
page breaks exactly.

## ADR-009 — PDF→DOCX and OCR

**Decision.** A staged pipeline we own: classify (digital vs scanned) → extract positioned text
→ OCR when needed → reconstruct layout (lines → paragraphs → columns → tables → styles) → PDM →
DOCX writer (ADR-006) → confidence report.

- **Extraction:** pdf.js (Apache-2.0) for text, positions and fonts.
- **OCR:** Tesseract via tesseract.js (Apache-2.0, WASM) — runs offline in the app, no native
  sidecar. PaddleOCR via ONNX is the upgrade path if accuracy requires it.
- **Reconstruction** is ours and is the differentiating know-how.
- Low-confidence regions are flagged in the conversion preview; the source PDF is never modified.

**Validation (Experiment 4).** A representative PDF corpus (digital single-column, multi-column,
tables, scanned, poor-quality scan, failure cases), scored with the same rubric approach as
Experiment 1.

## ADR-010 — Local storage and recovery

**Decision.**
- User documents are ordinary `.docx` files on the file system, written atomically
  (write to temp → flush → rename) by the Rust backend.
- An app database (SQLite) holds the document library, recent files and settings.
- Autosave writes PDM snapshots to an app-private recovery store every few seconds of
  inactivity; after a crash the user is offered the newest snapshot, never silently overwriting
  a newer file on disk.

**Validation.** Kill-the-process tests during save and autosave; disk-full and permission-denied
cases produce visible errors.

## ADR-011 — Windows first, platform-neutral core

**Decision.** Ship on Windows first. Everything except the Rust backend and packaging stays
platform-neutral (ADR-002), so macOS/web later reuse the core unchanged.

## ADR-012 — AI integration

**Deferred** (MVP excludes AI writing features). Direction retained: one internal AI interface
with local and cloud implementations; document content leaves the device only on explicit user
action; AI failure never blocks editing or saving.

## ADR-013 — Evidence-driven testing

**Decision.** Every owned engine has an automated, scored benchmark before it is called done:
DOCX round-trip (E1), layout vs Word (E2), editor mapping (E3), PDF conversion (E4). Microsoft
Word is the reference renderer. Fixtures are generated deterministically and versioned.

The Python harness is development tooling only (ADR-001).

---

## Build order and solo-developer estimate

Rough order-of-magnitude for one full-time developer; to be replaced with measured velocity
after the first phase.

| Phase | Work | Exit gate | Rough solo effort |
|---|---|---|---|
| 1 | PDM + DOCX reader/writer (TypeScript) | E1 all ✅ on A–D | 2–4 months |
| 2 | Layout engine → PDF export and print | E2 page counts/breaks match Word | 3–6 months |
| 3 | ProseMirror editor on PDM, flow view | E3 lossless editing | 2–3 months |
| 4 | Paginated editing view | Editing across page boundaries | 3–6 months |
| 5 | Storage, autosave, recovery, Windows packaging | Crash and save tests pass | 1–2 months |
| 6 | PDF→DOCX and OCR | E4 thresholds met | 3–6 months |

**Total: roughly 14–27 months to the full locked MVP.** That is the honest cost of owning the
stack alone.

**Recommendation (does not change the locked scope; it stages its release):**
- **Prime Word 1.0** — phases 1–5: editing, DOCX, pagination, PDF export/print, recovery.
- **Prime Word 1.1** — phase 6: PDF→DOCX and OCR, shipped as a headline update.

This gets a dependable product to users sooner and lets real documents shape the conversion
engine.
