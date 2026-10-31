/**
 * Built-in templates: professional starting points built from the document model alone (no files).
 * Each is an ordinary document — users edit and save it like any other.
 */
import type { Block, Inline, ParagraphProps, PrimeDocument, RunProps, TableCell } from "@prime/pdm";
import { blankDocument } from "./blank.js";
import { ensureList } from "./lists.js";

export interface Template {
  id: string;
  name: string;
  description: string;
  create(): PrimeDocument;
}

const ACCENT = "1F3864";


const text = (t: string, props?: RunProps): Inline => ({ type: "text", text: t, ...(props ? { props } : {}) });
const para = (content: Inline[] | string, props: ParagraphProps = {}): Block => ({
  type: "paragraph", props, content: typeof content === "string" ? (content ? [text(content)] : []) : content,
});
const heading = (t: string, level = 1): Block => para(t, { style: `Heading${level}` });

const BORDERS = '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders>';
const NO_BORDERS = '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
const twips = (pt: number) => String(Math.round(pt * 20));

/** A table; `widths` in pt. Header cells get a shaded fill. */
function table(widths: number[], rows: (string | Inline[])[][], opts: { header?: boolean; borders?: boolean; align?: (string | undefined)[] } = {}): Block {
  return {
    type: "table",
    props: { width: { w: "0", type: "auto" }, extra: [{ name: "w:tblBorders", xml: opts.borders === false ? NO_BORDERS : BORDERS }] },
    grid: widths,
    rows: rows.map((cells, r) => ({
      props: r === 0 && opts.header ? { header: true } : {},
      cells: cells.map((c, i): TableCell => ({
        props: {
          width: { w: twips(widths[i]), type: "dxa" },
          ...(r === 0 && opts.header ? { shading: { val: "clear", color: "auto", fill: ACCENT } } : {}),
        },
        blocks: [para(
          typeof c === "string" ? [text(c, r === 0 && opts.header ? { bold: true, color: "FFFFFF" } : undefined)] : c,
          { spacing: { before: 2, after: 2 }, ...(opts.align?.[i] ? { align: opts.align[i] } : {}) },
        )],
      })),
    })),
  };
}

function withBody(build: (doc: PrimeDocument) => Block[]): () => PrimeDocument {
  return () => {
    const doc = blankDocument();
    doc.sections[0].blocks = build(doc);
    return doc;
  };
}

const bullets = (doc: PrimeDocument, items: string[]): Block[] => {
  const numId = ensureList(doc, "bullet");
  return items.map((t) => para(t, { list: { numId, level: 0 } }));
};

export const TEMPLATES: Template[] = [
  {
    id: "blank", name: "Blank document", description: "An empty page.",
    create: blankDocument,
  },
  {
    id: "letter", name: "Letter", description: "A formal business letter.",
    create: withBody(() => [
      para([text("Your Name", { bold: true, size: 14, color: ACCENT })], { spacing: { after: 0 } }),
      para("Street address · City", { spacing: { after: 0 } }),
      para("Phone · Email", { spacing: { after: 18 } }),
      para(new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }), { spacing: { after: 18 } }),
      para("Recipient Name", { spacing: { after: 0 } }),
      para("Title, Company", { spacing: { after: 0 } }),
      para("Address", { spacing: { after: 18 } }),
      para([text("Subject: ", { bold: true }), text("Write the subject of your letter")]),
      para("Dear Recipient,"),
      para("Start with the purpose of the letter. Keep paragraphs short and focused on one point each."),
      para("Add supporting details here. Replace this text with your own; the formatting stays."),
      para("Thank you for your time and consideration."),
      para("Yours sincerely,", { spacing: { after: 36 } }),
      para("Your Name", { spacing: { after: 0 } }),
      para("Your Title"),
    ]),
  },
  {
    id: "report", name: "Report", description: "Title page heading, sections and a findings list.",
    create: withBody((doc) => [
      para("Report Title", { style: "Title" }),
      para([text("Subtitle or short description · Author · Date", { color: "595959", italic: true })], { spacing: { after: 24 } }),
      heading("1. Executive summary"),
      para("Summarise the purpose, main findings and recommendations in a few sentences."),
      heading("2. Background"),
      para("Describe the context and why this report was written."),
      heading("3. Findings"),
      ...bullets(doc, ["First finding", "Second finding", "Third finding"]),
      heading("4. Recommendations"),
      para("State what should be done next, by whom and by when."),
    ]),
  },
  {
    id: "cv", name: "CV / Résumé", description: "Profile, experience, education and skills.",
    create: withBody((doc) => [
      para([text("YOUR NAME", { bold: true, size: 26, color: ACCENT })], { spacing: { after: 0 } }),
      para([text("Professional title · City · Phone · Email · LinkedIn", { color: "595959" })], {
        spacing: { after: 12 },
        extra: [{ name: "w:pBdr", xml: `<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="${ACCENT}"/></w:pBdr>` }],
      }),
      heading("Profile", 2),
      para("Two or three sentences about who you are, what you do best and what you are looking for."),
      heading("Experience", 2),
      para([text("Job Title", { bold: true }), text(" — Company, City")], { spacing: { after: 0 } }),
      para([text("Month Year – Present", { italic: true, color: "595959" })], { spacing: { after: 4 } }),
      ...bullets(doc, ["Achievement with a measurable result", "Responsibility or project you led"]),
      heading("Education", 2),
      para([text("Degree, Field of Study", { bold: true }), text(" — University, Year")]),
      heading("Skills", 2),
      ...bullets(doc, ["Skill one", "Skill two", "Skill three"]),
    ]),
  },
  {
    id: "invoice", name: "Invoice", description: "Company details, line items, totals and payment terms.",
    create: withBody(() => {
      const w = 612 - 144;
      return [
        table([w * 0.6, w * 0.4], [
          [[text("YOUR COMPANY", { bold: true, size: 18, color: ACCENT })], [text("INVOICE", { bold: true, size: 24, color: ACCENT })]],
          ["Address · City · Phone · Email", [text("No. INV-0001", { bold: true })]],
          ["", `Date: ${new Date().toLocaleDateString("en-GB")}`],
        ], { borders: false, align: [undefined, "right"] }),
        para("", { spacing: { after: 6 } }),
        para([text("BILL TO", { bold: true, color: ACCENT, size: 9 })], { spacing: { after: 0 } }),
        para([text("Client Name", { bold: true })], { spacing: { after: 0 } }),
        para("Client address · Phone · Email", { spacing: { after: 12 } }),
        table([w * 0.46, w * 0.12, w * 0.2, w * 0.22], [
          ["Description", "Qty", "Unit price", "Amount"],
          ["Item or service", "1", "0.00", "0.00"],
          ["Item or service", "1", "0.00", "0.00"],
          ["Item or service", "1", "0.00", "0.00"],
        ], { header: true, align: [undefined, "center", "right", "right"] }),
        para("", { spacing: { after: 6 } }),
        table([w * 0.58, w * 0.2, w * 0.22], [
          ["", "Subtotal", "0.00"],
          ["", "Tax", "0.00"],
          ["", [text("Total", { bold: true })], [text("0.00", { bold: true })]],
        ], { borders: false, align: [undefined, "right", "right"] }),
        heading("Payment details", 3),
        para("Bank · Account name · Account number · Reference: INV-0001"),
        heading("Terms", 3),
        para("Payment is due within 14 days of the invoice date. Thank you for your business."),
      ];
    }),
  },
  {
    id: "memo", name: "Memo", description: "An internal memorandum.",
    create: withBody(() => [
      para([text("MEMO", { bold: true, size: 28, color: ACCENT })], { spacing: { after: 12 } }),
      table([1.6 * 72, 612 - 144 - 1.6 * 72], [
        [[text("To:", { bold: true })], "Recipient"],
        [[text("From:", { bold: true })], "Your name"],
        [[text("Date:", { bold: true })], new Date().toLocaleDateString("en-GB")],
        [[text("Subject:", { bold: true })], "Subject of the memo"],
      ], { borders: false }),
      para("", {
        spacing: { after: 12 },
        extra: [{ name: "w:pBdr", xml: `<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="${ACCENT}"/></w:pBdr>` }],
      }),
      para("Write the purpose of the memo in the first sentence, then the details and any action needed."),
    ]),
  },
];

