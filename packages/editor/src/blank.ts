/** A new, empty document built from the model alone (no template file). */
import { type PrimeDocument, SCHEMA_VERSION, type Style } from "@prime/pdm";

const heading = (id: string, name: string, size: number, before: number, outline: number): Style => ({
  id, type: "paragraph", name, basedOn: "Normal", next: "Normal",
  paragraph: { keepNext: true, keepLines: true, spacing: { before, after: 4 }, outlineLevel: outline },
  run: { font: { ascii: "Calibri Light", hAnsi: "Calibri Light" }, size, color: "2F5496" },
});

export function blankDocument(): PrimeDocument {
  return {
    schemaVersion: SCHEMA_VERSION,
    styles: {
      defaults: {
        run: { font: { ascii: "Calibri", hAnsi: "Calibri", eastAsia: "Calibri", cs: "Times New Roman" }, size: 11 },
        paragraph: { spacing: { after: 8, line: 259, lineRule: "auto" } },
      },
      styles: [
        { id: "Normal", type: "paragraph", name: "Normal", isDefault: true },
        { id: "Title", type: "paragraph", name: "Title", basedOn: "Normal", next: "Normal",
          paragraph: { spacing: { after: 0, line: 240, lineRule: "auto" } },
          run: { font: { ascii: "Calibri Light", hAnsi: "Calibri Light" }, size: 28 } },
        heading("Heading1", "heading 1", 16, 12, 0),
        heading("Heading2", "heading 2", 13, 2, 1),
        heading("Heading3", "heading 3", 12, 2, 2),
        { id: "DefaultParagraphFont", type: "character", name: "Default Paragraph Font", isDefault: true },
        { id: "TableNormal", type: "table", name: "Normal Table", isDefault: true },
      ],
    },
    lists: { definitions: [], instances: [] },
    sections: [{
      props: {
        pageWidth: 612, pageHeight: 792,
        margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36, gutter: 0 },
      },
      headers: {},
      footers: {},
      blocks: [{ type: "paragraph", props: {}, content: [] }],
    }],
    images: {},
    package: {
      mainPart: "word/document.xml",
      parts: [],
      preservedRels: [],
      contentTypes: { defaults: {}, overrides: {} },
      rootAttrs: {},
    },
  };
}
