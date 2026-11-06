import type { LayoutResult, Page, PlacedImage } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { memo, useEffect, useRef, useState, type ReactNode } from "react";

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

/** Pages drawn beyond the visible ones, above and below (scrolling a little shows no blank page). */
const AROUND = 2;

/**
 * Where page `page`'s sheet starts on screen (client y), drawn or not: sheets are spaced evenly.
 * (Only the sheets near the visible pages are drawn: see PageSheets.)
 */
export function sheetTop(page: number): number | undefined {
  const box = document.querySelector<HTMLElement>(".pw-sheets");
  if (!box) return undefined;
  const pitch = Number(box.dataset.pitch);
  const total = Number(box.dataset.count) * pitch - Number(box.dataset.gap);
  const r = box.getBoundingClientRect();
  return r.top + ((page - 1) * pitch * r.height) / total;
}

/**
 * Print Layout, in four layers (back to front): page sheets; the editor, starting at the first
 * page's body area, with page-gap spacers moving content onto the following sheets; margin masks,
 * which hide what runs through the space between pages (e.g. table borders); headers and footers.
 * A mask is dropped where the screen's content overflowed a page, so text is never hidden.
 *
 * Only the sheets, masks and headers and footers of the pages on screen (and AROUND either side) are
 * drawn, as Word draws only what is on screen: each is a positioned box the browser composites on
 * every frame, and all of a 855-page book's (some 3,400) cost ~60 ms a frame while typing. The editor
 * itself stays whole; the drawn range follows scrolling.
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
  const box = useRef<HTMLDivElement>(null);
  const [shown, setShown] = useState<[number, number]>([0, AROUND * 2]);
  useEffect(() => {
    const el = box.current;
    const scroller = el?.closest(".canvas");
    if (!el) return;
    const update = () => {
      const r = el.getBoundingClientRect();
      const view = scroller ? scroller.getBoundingClientRect() : { top: 0, bottom: innerHeight };
      const perPage = (pitch * r.height) / (count * pitch - gap); // px
      if (!perPage) return;
      const first = Math.max(0, Math.floor((view.top - r.top) / perPage) - AROUND);
      const last = Math.min(count - 1, Math.floor((view.bottom - r.top) / perPage) + AROUND);
      setShown((old) => (old[0] === first && old[1] === last ? old : [first, last]));
    };
    update();
    const target = scroller ?? window;
    target.addEventListener("scroll", update, { passive: true });
    window.addEventListener("resize", update);
    // Zooming changes the sheets' size without scrolling.
    const resize = new ResizeObserver(update);
    resize.observe(el);
    return () => { target.removeEventListener("scroll", update); window.removeEventListener("resize", update); resize.disconnect(); };
  }, [count, pitch, gap]);
  const [from, to] = shown;
  const near = (i: number) => i >= from && i <= to;
  return (
    <div ref={box} className="pw-sheets" data-pitch={pitch} data-count={count} data-gap={gap} data-masked={Math.max(0, pages.length - 1 - overflow.size)}
      style={{ width: `${width}pt`, height: `${count * pitch - gap}pt` }}>
      {mixed && <div className="pw-notice">This document has sections with different page sizes; all pages are shown at the first page's size.</div>}
      {Array.from({ length: count }, (_, i) => near(i) && (
        <div key={i} className="pw-sheet" data-page={i + 1} style={{ top: `${i * pitch}pt`, width: `${width}pt`, height: `${height}pt` }} />
      ))}
      <div className="pw-flow" style={{
        top: `${first?.bodyTop ?? margins.top ?? 72}pt`,
        left: `${first?.bodyLeft ?? margins.left ?? 72}pt`,
        width: `${first ? first.bodyRight - first.bodyLeft : width - (margins.left ?? 72) - (margins.right ?? 72)}pt`,
      }}>
        {children}
      </div>
      {pages.slice(1).map((next, i) => {
        if (overflow.has(next.number) || !(near(i) || near(i + 1))) return null;
        const prev = pages[i];
        const from = i * pitch + prev.bodyBottom;
        return (
          <div key={`m${next.number}`} className="pw-mask" data-page={next.number} style={{ top: `${from}pt`, height: `${(i + 1) * pitch + next.bodyTop - from}pt` }}>
            <div className="pw-mask-page" style={{ height: `${height - prev.bodyBottom}pt` }} />
            <div className="pw-mask-gap" style={{ height: `${gap}pt` }} />
            <div className="pw-mask-page" style={{ height: `${next.bodyTop}pt` }} />
          </div>
        );
      })}
      <div className="pw-chrome">
        {pages.map((p, i) => near(i) && <Chrome key={i} page={p} total={pages.length} doc={doc} top={i * pitch} />)}
      </div>
    </div>
  );
},
// Redraw only for a new layout or overflow: the editor inside (children) manages its own content, and
// re-rendering a long document's sheets on every keystroke made typing slow.
(a, b) => a.layout === b.layout && a.doc === b.doc && a.gap === b.gap && a.overflow === b.overflow);
