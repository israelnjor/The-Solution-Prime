/**
 * Charts made in Prime Word, as Word makes them: a chart part (c:chartSpace: the data cached in it, so
 * every program shows it), the spreadsheet holding the data embedded beside it (so Word's "Edit Data"
 * works), and the inline drawing that places it. Added to (or replaced in) a document's package.
 */
import type { PrimeDocument } from "@prime/pdm";
import { strToU8, zipSync } from "fflate";
import { EMU_PER_PT } from "./ooxml.js";
import { toBase64 } from "./xml.js";

export interface ChartData {
  kind: "column" | "bar" | "line" | "pie" | "area" | "doughnut" | "scatter";
  title?: string;
  /** Column, bar and area: series stacked on each other. */
  stacked?: boolean;
  categories: string[];
  series: { name: string; values: number[]; color?: string }[];
}

/** The colours new charts are drawn in (Word's current default theme, accent 1–6). */
export const NEW_CHART_COLORS = ["156082", "E97132", "196B24", "0F9ED5", "A02B93", "4EA72E"];

const CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const PACKAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
/** Spreadsheet column letters: 0 → A, 25 → Z, 26 → AA. */
const col = (i: number): string => (i < 26 ? String.fromCharCode(65 + i) : col(Math.floor(i / 26) - 1) + String.fromCharCode(65 + (i % 26)));
const num = (v: number) => (Number.isFinite(v) ? String(v) : "");

/** The chart part's XML (data and colours given explicitly, as Word writes them). */
export function chartXml(c: ChartData, workbookRelId: string): string {
  const n = c.categories.length;
  const strCache = (vals: string[]) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("")}</c:strCache>`;
  const numCache = (vals: number[]) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (Number.isFinite(v) ? `<c:pt idx="${i}"><c:v>${num(v)}</c:v></c:pt>` : "")).join("")}</c:numCache>`;
  const fill = (hex: string) => `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  const round = c.kind === "pie" || c.kind === "doughnut";
  const stacked = !!c.stacked && (c.kind === "column" || c.kind === "bar" || c.kind === "area");
  const ser = (s: ChartData["series"][number], j: number) => {
    const color = s.color ?? NEW_CHART_COLORS[j % NEW_CHART_COLORS.length];
    const letter = col(j + 1);
    const look = c.kind === "line"
      ? `<c:spPr><a:ln w="28575" cap="rnd">${fill(color)}<a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`
      : round
        ? `${c.categories.map((_, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fill(NEW_CHART_COLORS[i % NEW_CHART_COLORS.length])}<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join("")}`
        : c.kind === "scatter"
          // Markers only (Word's default scatter): no line between the points.
          ? `<c:spPr><a:ln w="25400" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${fill(color)}<a:ln w="9525">${fill(color)}</a:ln></c:spPr></c:marker>`
          : c.kind === "area"
            ? `<c:spPr>${fill(color)}<a:ln><a:noFill/></a:ln></c:spPr>`
            : `<c:spPr>${fill(color)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>`;
    const name = `<c:tx><c:strRef><c:f>Sheet1!$${letter}$1</c:f>${strCache([s.name])}</c:strRef></c:tx>`;
    const values = `<c:numRef><c:f>Sheet1!$${letter}$2:$${letter}$${n + 1}</c:f>${numCache(s.values)}</c:numRef>`;
    if (c.kind === "scatter") {
      // The categories are the x values (numbers).
      return `<c:ser><c:idx val="${j}"/><c:order val="${j}"/>${name}${look}`
        + `<c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f>${numCache(c.categories.map(Number))}</c:numRef></c:xVal>`
        + `<c:yVal>${values}</c:yVal><c:smooth val="0"/></c:ser>`;
    }
    return `<c:ser><c:idx val="${j}"/><c:order val="${j}"/>${name}${look}`
      + `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${n + 1}</c:f>${strCache(c.categories)}</c:strRef></c:cat>`
      + `<c:val>${values}</c:val>`
      + `${c.kind === "line" ? '<c:smooth val="0"/>' : ""}</c:ser>`;
  };
  const series = (round ? c.series.slice(0, 1) : c.series).map(ser).join("");
  const axes = `<c:axId val="508012001"/><c:axId val="508012002"/>`;
  const axisText = `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
  // A scatter chart's x axis holds values too.
  const xValAx = `<c:valAx><c:axId val="508012001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>${axisText}<c:crossAx val="508012002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
  const catAx = `<c:catAx><c:axId val="508012001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${c.kind === "bar" ? "l" : "b"}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>${axisText}<c:crossAx val="508012002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAx = `<c:valAx><c:axId val="508012002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${c.kind === "bar" ? "b" : "l"}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${axisText}<c:crossAx val="508012001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const group = c.kind === "pie"
    ? `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`
    : c.kind === "doughnut"
      ? `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`
      : c.kind === "line"
        ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axes}</c:lineChart>${catAx}${valAx}`
        : c.kind === "area"
          ? `<c:areaChart><c:grouping val="${stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>${series}${axes}</c:areaChart>${catAx}${valAx}`
          : c.kind === "scatter"
            ? `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${axes}</c:scatterChart>${xValAx}${valAx.replace('<c:crossBetween val="between"/>', '<c:crossBetween val="midCat"/>')}`
            : `<c:barChart><c:barDir val="${c.kind === "bar" ? "bar" : "col"}"/><c:grouping val="${stacked ? "stacked" : "clustered"}"/><c:varyColors val="0"/>${series}<c:gapWidth val="${stacked ? 150 : 219}"/><c:overlap val="${stacked ? 100 : -27}"/>${axes}</c:barChart>${catAx}${valAx}`;
  const title = c.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:rPr><a:t>${esc(c.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : `<c:autoTitleDeleted val="1"/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>`
    + `<c:chart>${title}<c:plotArea><c:layout/>${group}</c:plotArea>`
    + `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${axisText}</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
    + `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>`
    + `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`
    + `<c:externalData r:id="${workbookRelId}"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`;
}

/** The chart's data as a spreadsheet (Sheet1: categories down column A, a series per column). */
export function chartWorkbook(c: ChartData): Uint8Array {
  const cell = (ref: string, v: string | number) => (typeof v === "number"
    ? (Number.isFinite(v) ? `<c r="${ref}"><v>${v}</v></c>` : "")
    : `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`);
  const header = `<row r="1">${cell("A1", "")}${c.series.map((s, j) => cell(`${col(j + 1)}1`, s.name)).join("")}</row>`;
  const rows = c.categories.map((cat, i) => `<row r="${i + 2}">${cell(`A${i + 2}`, c.kind === "scatter" && cat.trim() !== "" && Number.isFinite(Number(cat)) ? Number(cat) : cat)}${c.series.map((s, j) => cell(`${col(j + 1)}${i + 2}`, s.values[i])).join("")}</row>`).join("");
  const main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
    "xl/workbook.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
    "xl/_rels/workbook.xml.rels": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
    "xl/worksheets/sheet1.xml": strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${main}"><sheetData>${header}${rows}</sheetData></worksheet>`),
  });
}

/**
 * Adds a chart to the document's package (its part, its spreadsheet, their relationships and content
 * types) and returns the drawing that places it inline (w × h pt). `drawingId`: a unique wp:docPr id.
 */
export function addChart(doc: PrimeDocument, c: ChartData, w: number, h: number, drawingId: number): string {
  const pkg = doc.package;
  let k = 1;
  while (pkg.parts.some((p) => p.name === `word/charts/chart${k}.xml`)) k++;
  const chart = `word/charts/chart${k}.xml`;
  const workbook = `word/embeddings/Microsoft_Excel_Worksheet${k}.xlsx`;
  let relId = `rIdChart${k}`;
  for (let i = 2; pkg.preservedRels.some((r) => r.id === relId); i++) relId = `rIdChart${k}_${i}`;
  pkg.parts.push(
    { name: chart, data: toBase64(new TextEncoder().encode(chartXml(c, "rId1"))) },
    { name: `word/charts/_rels/chart${k}.xml.rels`, data: toBase64(new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${PACKAGE_REL}" Target="../embeddings/Microsoft_Excel_Worksheet${k}.xlsx"/></Relationships>`)) },
    { name: workbook, data: toBase64(chartWorkbook(c)) },
  );
  pkg.preservedRels.push({ id: relId, type: CHART_REL, target: chart });
  pkg.contentTypes = {
    defaults: { ...pkg.contentTypes.defaults, xlsx: pkg.contentTypes.defaults.xlsx ?? XLSX_TYPE },
    overrides: { ...pkg.contentTypes.overrides, [chart]: CHART_TYPE },
  };
  const emu = (pt: number) => Math.round(pt * EMU_PER_PT);
  return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${emu(w)}" cy="${emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:docPr id="${drawingId}" name="Chart ${drawingId}"${c.title ? ` descr="${esc(c.title)}"` : ""}/><wp:cNvGraphicFramePr/>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${relId}"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

/** A chart's data changed (Edit Data): its part and spreadsheet rewritten in place. */
export function updateChart(doc: PrimeDocument, relId: string, c: ChartData): boolean {
  const rel = doc.package.preservedRels.find((r) => r.id === relId);
  const part = rel && doc.package.parts.find((p) => p.name === rel.target);
  if (!rel || !part) return false;
  const name = rel.target.split("/").pop()!;
  const relsName = rel.target.replace(name, `_rels/${name}.rels`);
  const relsPart = doc.package.parts.find((p) => p.name === relsName);
  const relsXml = relsPart ? new TextDecoder().decode(Uint8Array.from(atob(relsPart.data), (ch) => ch.charCodeAt(0))) : "";
  // Its spreadsheet (a package relationship of the chart part), if it has one of its own.
  const wbRel = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Type="[^"]*\/package"[^>]*Target="([^"]+)"/.exec(relsXml)
    ?? /<Relationship\b[^>]*Type="[^"]*\/package"[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/.exec(relsXml);
  part.data = toBase64(new TextEncoder().encode(chartXml(c, wbRel?.[1] ?? "rId1")));
  if (wbRel) {
    const dir = rel.target.split("/").slice(0, -1);
    for (const seg of wbRel[2].split("/")) { if (seg === "..") dir.pop(); else if (seg !== ".") dir.push(seg); }
    const wb = doc.package.parts.find((p) => p.name === dir.join("/"));
    if (wb) wb.data = toBase64(chartWorkbook(c));
  }
  return true;
}
