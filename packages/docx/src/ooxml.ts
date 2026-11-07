/** OOXML constants: relationship types, content types and schema child order (ECMA-376 Part 1). */

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

export const RelType = {
  officeDocument: `${REL}/officeDocument`,
  styles: `${REL}/styles`,
  numbering: `${REL}/numbering`,
  header: `${REL}/header`,
  footer: `${REL}/footer`,
  image: `${REL}/image`,
  hyperlink: `${REL}/hyperlink`,
  footnotes: `${REL}/footnotes`,
  endnotes: `${REL}/endnotes`,
  settings: `${REL}/settings`,
} as const;

/** Relationship types whose targets the reader turns into model content (the writer regenerates them). */
export const MODELED_REL_TYPES = new Set<string>([RelType.styles, RelType.numbering, RelType.header, RelType.footer, RelType.footnotes, RelType.endnotes]);

export const ContentType = {
  document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
  endnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
  settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  rels: "application/vnd.openxmlformats-package.relationships+xml",
  xml: "application/xml",
} as const;

export const IMAGE_TYPES: Record<string, string> = {
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  emf: "image/x-emf",
  wmf: "image/x-wmf",
  svg: "image/svg+xml",
};

export const RPR_ORDER = [
  "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", "w:dstrike",
  "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden",
  "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
  "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
  "w:specVanish", "w:oMath", "w:ins", "w:del", "w:moveFrom", "w:moveTo", "w:rPrChange",
] as const;

export const PPR_ORDER = [
  "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl", "w:numPr",
  "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
  "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
  "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
  "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
  "w:sectPr", "w:pPrChange",
] as const;

export const SECTPR_ORDER = [
  "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz", "w:pgMar",
  "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
  "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange",
] as const;

export const TBLPR_ORDER = [
  "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize", "w:tblStyleColBandSize",
  "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar",
  "w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange",
] as const;

export const TRPR_ORDER = [
  "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore", "w:wAfter", "w:cantSplit", "w:trHeight",
  "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange",
] as const;

export const TCPR_ORDER = [
  "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap", "w:tcMar",
  "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
  "w:cellMerge", "w:tcPrChange",
] as const;

export const STYLE_ORDER = [
  "w:name", "w:aliases", "w:basedOn", "w:next", "w:link", "w:autoRedefine", "w:hidden", "w:uiPriority",
  "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked", "w:personal", "w:personalCompose",
  "w:personalReply", "w:rsid", "w:pPr", "w:rPr", "w:tblPr", "w:trPr", "w:tcPr", "w:tblStylePr",
] as const;

export const ABSTRACT_NUM_ORDER = [
  "w:nsid", "w:multiLevelType", "w:tmpl", "w:name", "w:styleLink", "w:numStyleLink", "w:lvl",
] as const;

export const LVL_ORDER = [
  "w:start", "w:numFmt", "w:lvlRestart", "w:pStyle", "w:isLgl", "w:suff", "w:lvlText", "w:lvlPicBulletId",
  "w:legacy", "w:lvlJc", "w:pPr", "w:rPr",
] as const;

export const NUMBERING_ORDER = ["w:numPicBullet", "w:abstractNum", "w:num", "w:numIdMacAtCleanup"] as const;

export const TWIPS_PER_PT = 20;
export const EMU_PER_PT = 12700;

export const pt = (twips: string | undefined) => (twips === undefined ? undefined : Number(twips) / TWIPS_PER_PT);
export const twips = (points: number) => String(Math.round(points * TWIPS_PER_PT));
