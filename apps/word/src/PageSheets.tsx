import type { LayoutResult, Page, PlacedImage } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { memo, type ReactNode } from "react";

/** Clicking a footnote or endnote on the page opens the Notes pane. */
export const NOTES_EVENT = "prime-word:notes";

const IMAGE_TYPES: Record<string, string> = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp" };

function imageUrl(doc: PrimeDocument, img: PlacedImage): string | undefined {
  const res = img.imageId ? doc.images[img.imageId] : undefined;
  if (res) return `data:${res.contentType};base64,${res.data}`;
  const part = img.part ? doc.package.parts.find((p) => p.name === img.part) : undefined;
  const ext = part?.name.split(".").pop()?.toLowerCase() ?? "";
  return part && IMAGE_TYPES[ext] ? `data:${IMAGE_TYPES[ext]};base64,${part.data}` : undefined;
}

/**
 * Header and footer of one page, positioned exactly as the layout engine placed them. (Floating
 * pictures and text boxes are drawn by the editor, with their anchor paragraph.)
 */
function Chrome({ page, total, doc, top }: { page: Page; total: number; doc: PrimeDocument; top: number }) {
  return (
    <>
      {/* The short rule above footnotes and endnotes. */}
      {page.rects.filter((r) => r.story).map((r, i) => (
        <div key={`r${i}`} className="pw-chrome-rule" style={{ left: `${r.x}pt`, top: `${top + r.y}pt`, width: `${r.width}pt` }} />
      ))}
      {page.lines.filter((l) => l.story === "header" || l.story === "footer" || l.story === "footnote" || l.story === "endnote").map((line, i) => (
        <div key={i} className={line.story === "footnote" || line.story === "endnote" ? "pw-chrome-note" : undefined}
          onClick={line.story === "footnote" || line.story === "endnote" ? () => window.dispatchEvent(new Event(NOTES_EVENT)) : undefined}>
          {line.segments.map((s, j) => (
            <span key={j} className="pw-chrome-text" style={{
              left: `${s.x}pt`, top: `${top + line.y}pt`, lineHeight: `${line.height}pt`,
              fontFamily: `"${s.font.family}"`, fontSize: `${s.size}pt`,
              fontWeight: s.font.bold ? 700 : 400, fontStyle: s.font.italic ? "italic" : "normal",
              color: s.color ? `#${s.color}` : undefined, textDecoration: s.underline ? "underline" : undefined,
              ...(s.vertAlign ? { fontSize: `${s.size * 0.65}pt`, verticalAlign: s.vertAlign === "superscript" ? "super" : "sub" } : {}),
            }}>
              {s.field === "PAGE" ? page.number : s.field === "NUMPAGES" ? total : s.text}
            </span>
          ))}
          {line.images.map((img, j) => {
            const src = imageUrl(doc, img);
            return src ? (
              <img key={`i${j}`} src={src} alt="" className="pw-chrome-image"
                style={{ left: `${img.x}pt`, top: `${top + img.y}pt`, width: `${img.width}pt`, height: `${img.height}pt` }} />
            ) : null;
          })}
        </div>
      ))}
    </>
  );
}

/**
 * Print Layout, in four layers (back to front): page sheets; the editor, starting at the first
 * page's body area, with page-gap spacers moving content onto the following sheets; margin masks,
 * which hide what runs through the space between pages (e.g. table borders); headers and footers.
 * A mask is dropped where the screen's content overflowed a page, so text is never hidden.
 */
export const PageSheets = memo(function PageSheets({ layout, doc, gap, overflow, children }: {
  layout?: LayoutResult; doc: PrimeDocument; gap: number; overflow: Set<number>; children: ReactNode;
}) {
  const section = doc.sections[0]?.props ?? {};
  const margins = section.margins ?? {};
  const pages = layout?.pages ?? [];
  const first = pages[0];
  const width = first?.width ?? section.pageWidth ?? 612;
  const height = first?.height ?? section.pageHeight ?? 792;
  const pitch = height + gap;
  const count = Math.max(1, pages.length);
  const mixed = pages.some((p) => p.width !== width || p.height !== height);
  return (
    <div className="pw-sheets" style={{ width: `${width}pt`, height: `${count * pitch - gap}pt` }}>
      {mixed && <div className="pw-notice">This document has sections with different page sizes; all pages are shown at the first page's size.</div>}
      {Array.from({ length: count }, (_, i) => (
        <div key={i} className="pw-sheet" style={{ top: `${i * pitch}pt`, width: `${width}pt`, height: `${height}pt` }} />
      ))}
      <div className="pw-flow" style={{
        top: `${first?.bodyTop ?? margins.top ?? 72}pt`,
        left: `${first?.bodyLeft ?? margins.left ?? 72}pt`,
        width: `${first ? first.bodyRight - first.bodyLeft : width - (margins.left ?? 72) - (margins.right ?? 72)}pt`,
      }}>
        {children}
      </div>
      {pages.slice(1).map((next, i) => {
        if (overflow.has(next.number)) return null;
        const prev = pages[i];
        const from = i * pitch + prev.bodyBottom;
        return (
          <div key={`m${next.number}`} className="pw-mask" style={{ top: `${from}pt`, height: `${(i + 1) * pitch + next.bodyTop - from}pt` }}>
            <div className="pw-mask-page" style={{ height: `${height - prev.bodyBottom}pt` }} />
            <div className="pw-mask-gap" style={{ height: `${gap}pt` }} />
            <div className="pw-mask-page" style={{ height: `${next.bodyTop}pt` }} />
          </div>
        );
      })}
      <div className="pw-chrome">
        {pages.map((p, i) => <Chrome key={i} page={p} total={pages.length} doc={doc} top={i * pitch} />)}
      </div>
    </div>
  );
},
// Redraw only for a new layout or overflow: the editor inside (children) manages its own content, and
// re-rendering a long document's sheets on every keystroke made typing slow.
(a, b) => a.layout === b.layout && a.doc === b.doc && a.gap === b.gap && a.overflow === b.overflow);
