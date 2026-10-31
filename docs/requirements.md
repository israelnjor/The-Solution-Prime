# Prime Word — MVP Requirements v0.2

Status: **scope locked** as the product baseline. Technical feasibility of individual features is
still being validated (milestone M1).

## Product principles

- **Windows desktop first**, cross-platform later; document core stays platform-neutral.
- **Offline-first:** create, edit, save and reopen documents with no internet connection.
- **Familiar:** conventions users know from Microsoft Word.
- **Honest compatibility:** never claim perfect fidelity; never silently lose content.
- **Ambient, optional AI:** helpful, never in the way, never required for core work.

## First-release scope

| Area | Included |
|---|---|
| Core editing | Create/open/edit/save · text formatting, fonts, colors, alignment · headings, styles, bullets, numbered lists · tables, hyperlinks, images · undo/redo, find/replace · keyboard shortcuts, word count |
| Files & reliability | Open, rename, duplicate, organize · recent documents · automatic local saving and crash recovery · clear save status and save-error warnings · full offline core editing |
| Portability | DOCX import/export · export and print to PDF |
| PDF conversion | Digital PDF → editable DOCX · scanned PDFs via OCR · conversion preview and warnings |
| Presentation | Page layout, margins, page breaks, pagination · headers, footers, page numbers · professional templates |

**Deferred:** AI writing refinement and tone adjustment · cloud sync · real-time collaboration.

## Quality requirements

- **Reliability:** saving never silently loses work; recovery after unexpected shutdown.
- **DOCX fidelity:** common documents keep text, formatting, tables and images; unsupported elements are reported, not dropped.
- **PDF conversion honesty:** limitations are surfaced; OCR and complex layouts may need manual correction.
- **Performance:** ordinary editing stays responsive in long documents.
- **Privacy:** content leaves the device only on explicit user action.
- **Accessibility:** keyboard navigation, readable controls.
- **Maintainability:** editing, conversion, storage and AI are separable components.

## Acceptance criteria (thresholds set after M1 baselines)

| Requirement | Criterion |
|---|---|
| Editing | Documents with supported text, formatting, lists, tables, links and images can be created and edited. |
| Local saving | Save to a user-chosen location; saving / saved / failed states are always visible. |
| Autosave & recovery | Latest recoverable work restored after a crash, without silently overwriting a newer file. |
| DOCX compatibility | Supported content survives import → edit → export → reopen; altered elements are reported. **Zero unexplained text loss.** |
| Pagination | Page boundaries, breaks, margins, headers, footers and page numbers consistent between editor and exported PDF. |
| Templates | Professional starting points that save as ordinary documents. |
| PDF export & print | PDF export, print preview, standard printer settings. |
| PDF → DOCX | Digital and scanned PDFs convert to editable documents with preview and limitation warnings. |
| Offline | Core editing, local files, autosave and bundled templates work without a connection. |

## Milestones

| | Deliverable | Gate |
|---|---|---|
| M0 | PRD, journeys, wireframes | Scope agreed ✅ |
| **M1** | **Editing-engine and conversion feasibility report** | **Core technical risks understood** |
| M2 | Reliable local editor | Save/recovery and editing tests pass |
| M3 | DOCX and PDF conversion workflows | Compatibility benchmark meets thresholds |
| M4 | Desktop release candidate | Usability, security, accessibility, regression |
| M5 | Ambient AI features | Quality, privacy, offline behavior validated |
| M6 | Ecosystem expansion plan | User feedback supports next product |
