import type { LayoutResult, Page, PageBorderLine, PlacedImage, Watermark } from "@prime/layout";
import type { PrimeDocument } from "@prime/pdm";
import { pageTops } from "@prime/editor";
import { memo, useEffect, useMemo, useRef, useState, type ReactNode } from "react";

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

/** Text across the page behind the body text, as the layout placed it (rotated about its centre). */
function WatermarkText({ w }: { w: Watermark }) {
  return (
    <div className="pw-watermark" style={{
      left: `${w.x}pt`, top: `${w.y}pt`, transform: `translate(-50%, -50%) rotate(${w.rotation}deg)`,
      fontFamily: `"${w.font.family}"`, fontSize: `${w.size}pt`, color: `#${w.color}`, opacity: w.opacity,
    }}>{w.text}</div>
  );
}

/** Page borders over the page (in front of the text, as Word draws them by default). */
function Borders({ lines, width, height, top }: { lines: PageBorderLine[]; width: number; height: number; top: number }) {
  const dash = (l: PageBorderLine) => (l.style === "dotted" ? `${l.width} ${l.width * 2}` : l.style.startsWith("dash") ? `${l.width * 4} ${l.width * 3}` : undefined);
  return (
    <svg className="pw-page-borders" style={{ top: `${top}pt`, width: `${width}pt`, height: `${height}pt` }} viewBox={`0 0 ${width} ${height}`}>
      {lines.flatMap((l, i) => {
        const vertical = l.x1 === l.x2;
        // Double: two lines, each a third of the width, a third apart (Word's w:sz is the whole width).
        const offsets = l.style === "double" ? [-l.width / 3, l.width / 3] : [0];
        const stroke = l.style === "double" ? l.width / 3 : l.width;
        return offsets.map((d, k) => (
          <line key={`${i}-${k}`} x1={l.x1 + (vertical ? d : 0)} y1={l.y1 + (vertical ? 0 : d)} x2={l.x2 + (vertical ? d : 0)} y2={l.y2 + (vertical ? 0 : d)}
            stroke={`#${l.color ?? "000000"}`} strokeWidth={stroke} strokeDasharray={dash(l)} />
        ));
      })}
    </svg>
  );
}

/** Pages drawn beyond the visible ones, above and below (scrolling a little shows no blank page). */
const AROUND = 2;

/**
 * Where page `page`'s sheet starts on screen (client y), drawn or not. (Only the sheets near the
 * visible pages are drawn: see PageSheets.)
 */
export function sheetTop(page: number): number | undefined {
  const box = document.querySelector<HTMLElement>(".pw-sheets");
  if (!box) return undefined;
  const tops = JSON.parse(box.dataset.tops ?? "[]") as number[];
  const r = box.getBoundingClientRect();
  return r.top + ((tops[page - 1] ?? 0) * r.height) / Number(box.dataset.height);
}

/**
 * Print Layout, in four layers (back to front): page sheets; the editor, starting at the first
 * page's body area, with page-gap spacers moving content onto the following sheets; margin masks,
 * which hide what runs through the space between pages (e.g. table borders); headers and footers.
 * A mask is dropped where the screen's content overflowed a page, so text is never hidden.
 *
 * Pages are stacked at their own sizes (a section can be landscape), centred on the widest; the
 * text of a section with another page width or margins is moved into its own column by the editor
 * (sectionOffsets).
 *
 * Only the sheets, masks and headers and footers of the pages on screen (and AROUND either side) are
 * drawn, as Word draws only what is on screen. The editor itself stays whole; the drawn range
 * follows scrolling.
 */
export const PageSheets = memo(function PageSheets({ layout, doc, gap, overflow, children }: {
  layout?: LayoutResult; doc: PrimeDocument; gap: number; overflow: Set<number>; children: ReactNode;
}) {
  const section = doc.sections[0]?.props ?? {};
  const margins = section.margins ?? {};
  const laid = layout?.pages ?? [];
  // Before the first layout: one page of the first section's size.
  const pages: { width: number; height: number }[] = laid.length ? laid : [{ width: section.pageWidth ?? 612, height: section.pageHeight ?? 792 }];
  const first = laid[0];
  const tops = useMemo(() => pageTops(pages, gap), [pages, gap]);
  const widest = Math.max(...pages.map((p) => p.width));
  const total = tops[tops.length - 1] + pages[pages.length - 1].height;
  const offset = (i: number) => (widest - pages[i].width) / 2; // each sheet, centred
  const count = pages.length;
  const box = useRef<HTMLDivElement>(null);
  const [shown, setShown] = useState<[number, number]>([0, AROUND * 2]);
  useEffect(() => {
    const el = box.current;
    const scroller = el?.closest(".canvas");
    if (!el) return;
    // The page at a height (pt) from the top: the last one starting above it.
    const at = (y: number) => { let lo = 0; let hi = tops.length - 1; while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (tops[mid] <= y) lo = mid; else hi = mid - 1; } return lo; };
    const update = () => {
      const r = el.getBoundingClientRect();
      const view = scroller ? scroller.getBoundingClientRect() : { top: 0, bottom: innerHeight };
      const scale = r.height / total; // px per pt
      if (!scale) return;
      const from = Math.max(0, at((view.top - r.top) / scale) - AROUND);
      const to = Math.min(count - 1, at((view.bottom - r.top) / scale) + AROUND);
      setShown((old) => (old[0] === from && old[1] === to ? old : [from, to]));
    };
    update();
    const target = scroller ?? window;
    target.addEventListener("scroll", update, { passive: true });
    window.addEventListener("resize", update);
    // Zooming changes the sheets' size without scrolling.
    const resize = new ResizeObserver(update);
    resize.observe(el);
    return () => { target.removeEventListener("scroll", update); window.removeEventListener("resize", update); resize.disconnect(); };
  }, [tops, total, count]);
  const [from, to] = shown;
  const near = (i: number) => i >= from && i <= to;
  return (
    <div ref={box} className="pw-sheets" data-tops={JSON.stringify(tops)} data-height={total} data-count={count} data-masked={Math.max(0, laid.length - 1 - overflow.size)}
      style={{ width: `${widest}pt`, height: `${total}pt` }}>
      {pages.map((p, i) => near(i) && (
        <div key={i} className="pw-sheet" data-page={i + 1} style={{
          top: `${tops[i]}pt`, left: `${offset(i)}pt`, width: `${p.width}pt`, height: `${p.height}pt`,
          ...(laid[i]?.decor?.background ? { background: `#${laid[i].decor!.background}` } : {}),
        }}>
          {laid[i]?.decor?.watermark && <WatermarkText w={laid[i].decor!.watermark!} />}
        </div>
      ))}
      <div className="pw-flow" style={{
        top: `${first?.bodyTop ?? margins.top ?? 72}pt`,
        left: `${offset(0) + (first?.bodyLeft ?? margins.left ?? 72)}pt`,
        width: `${first ? first.bodyRight - first.bodyLeft : pages[0].width - (margins.left ?? 72) - (margins.right ?? 72)}pt`,
      }}>
        {children}
      </div>
      {laid.slice(1).map((next, i) => {
        if (overflow.has(next.number) || !(near(i) || near(i + 1))) return null;
        const prev = laid[i];
        const start = tops[i] + prev.bodyBottom;
        // White over the bottom margin of one page and the top margin of the next (each its own width).
        const sheet = (k: number) => ({
          marginLeft: `calc(40px + ${offset(k)}pt)`, width: `${laid[k].width}pt`,
          ...(laid[k].decor?.background ? { background: `#${laid[k].decor!.background}` } : {}),
        });
        return (
          <div key={`m${next.number}`} className="pw-mask" data-page={next.number} style={{ top: `${start}pt`, height: `${tops[i + 1] + next.bodyTop - start}pt` }}>
            <div className="pw-mask-page" style={{ height: `${prev.height - prev.bodyBottom}pt`, ...sheet(i) }} />
            <div className="pw-mask-gap" style={{ height: `${gap}pt` }} />
            <div className="pw-mask-page" style={{ height: `${next.bodyTop}pt`, ...sheet(i + 1) }} />
          </div>
        );
      })}
      <div className="pw-chrome">
        {laid.map((p, i) => near(i) && (
          <div key={i} className="pw-chrome-page" style={{ left: `${offset(i)}pt` }}>
            {p.decor?.borders && <Borders lines={p.decor.borders} width={p.width} height={p.height} top={tops[i]} />}
            <Chrome page={p} total={laid.length} doc={doc} top={tops[i]} />
          </div>
        ))}
      </div>
    </div>
  );
},
// Redraw only for a new layout or overflow: the editor inside (children) manages its own content, and
// re-rendering a long document's sheets on every keystroke made typing slow.
(a, b) => a.layout === b.layout && a.doc === b.doc && a.gap === b.gap && a.overflow === b.overflow);
