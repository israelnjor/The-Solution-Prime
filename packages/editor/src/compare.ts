/**
 * Review → Compare: an original and a revised document, as one document holding the differences as
 * tracked changes (what the revised version took out, marked deleted; what it put in, inserted), as
 * Word's Compare does. Paragraphs are matched first (by their text); a changed paragraph is then
 * compared word by word. Formatting-only differences are not shown.
 */
import { Fragment, type Mark, type Node, type Schema } from "prosemirror-model";

type Op = "equal" | "delete" | "insert";
interface Edit { op: Op; a?: number; b?: number }

/** Longest common subsequence by table (small inputs). */
function lcs(a: string[], b: string[], a0: number, b0: number): Edit[] {
  const n = a.length;
  const m = b.length;
  const t: Uint32Array[] = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) t[i][j] = a[i] === b[j] ? t[i + 1][j + 1] + 1 : Math.max(t[i + 1][j], t[i][j + 1]);
  const out: Edit[] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) { out.push({ op: "equal", a: a0 + i++, b: b0 + j++ }); }
    else if (t[i + 1][j] >= t[i][j + 1]) out.push({ op: "delete", a: a0 + i++ });
    else out.push({ op: "insert", b: b0 + j++ });
  }
  while (i < n) out.push({ op: "delete", a: a0 + i++ });
  while (j < m) out.push({ op: "insert", b: b0 + j++ });
  return out;
}

/**
 * The edits turning `a` into `b`: common start and end first; a small middle by table; a large one
 * anchored on items that occur once in each (patience diff), the stretches between compared again.
 */
export function diff(a: string[], b: string[], a0 = 0, b0 = 0): Edit[] {
  let s = 0;
  while (s < a.length && s < b.length && a[s] === b[s]) s++;
  let e = 0;
  while (e < a.length - s && e < b.length - s && a[a.length - 1 - e] === b[b.length - 1 - e]) e++;
  const head: Edit[] = Array.from({ length: s }, (_, i) => ({ op: "equal" as const, a: a0 + i, b: b0 + i }));
  const tail: Edit[] = Array.from({ length: e }, (_, i) => ({ op: "equal" as const, a: a0 + a.length - e + i, b: b0 + b.length - e + i }));
  const ma = a.slice(s, a.length - e);
  const mb = b.slice(s, b.length - e);
  let middle: Edit[];
  if (!ma.length || !mb.length || ma.length * mb.length <= 4_000_000) middle = lcs(ma, mb, a0 + s, b0 + s);
  else {
    // Items once in each: pairs in order (the longest increasing run of them) anchor the comparison.
    const count = (xs: string[]) => { const c = new Map<string, number>(); xs.forEach((x, i) => c.set(x, c.has(x) ? -1 : i)); return c; };
    const ca = count(ma);
    const cb = count(mb);
    const pairs: [number, number][] = [];
    for (const [k, i] of ca) { const j = cb.get(k); if (i >= 0 && j !== undefined && j >= 0) pairs.push([i, j]); }
    pairs.sort((x, y) => x[0] - y[0]);
    // Longest increasing subsequence of the second indexes.
    const tails: number[] = [];
    const prev: number[] = new Array(pairs.length).fill(-1);
    pairs.forEach(([, j], k) => {
      let lo = 0;
      let hi = tails.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (pairs[tails[mid]][1] < j) lo = mid + 1; else hi = mid; }
      if (lo > 0) prev[k] = tails[lo - 1];
      tails[lo] = k;
    });
    const anchors: [number, number][] = [];
    for (let k = tails.length ? tails[tails.length - 1] : -1; k >= 0; k = prev[k]) anchors.unshift(pairs[k]);
    if (!anchors.length) {
      middle = [...ma.map((_, i) => ({ op: "delete" as const, a: a0 + s + i })), ...mb.map((_, j) => ({ op: "insert" as const, b: b0 + s + j }))];
    } else {
      middle = [];
      let pi = 0;
      let pj = 0;
      for (const [i, j] of [...anchors, [ma.length, mb.length] as [number, number]]) {
        middle.push(...diff(ma.slice(pi, i), mb.slice(pj, j), a0 + s + pi, b0 + s + pj));
        if (i < ma.length) middle.push({ op: "equal", a: a0 + s + i, b: b0 + s + j });
        pi = i + 1;
        pj = j + 1;
      }
    }
  }
  return [...head, ...middle, ...tail];
}

/** A block's key for matching: its kind and text. */
const keyOf = (n: Node) => `${n.type.name}:${n.textContent}`;

/** A paragraph's content as tokens (words, spaces, punctuation, objects), each with its marks. */
interface Token { key: string; node: Node }
function tokens(p: Node): Token[] {
  const out: Token[] = [];
  p.forEach((child) => {
    if (child.isText) {
      for (const m of child.text!.matchAll(/\s+|[\p{L}\p{N}'’]+|[^\s\p{L}\p{N}]/gu)) out.push({ key: m[0], node: child.type.schema.text(m[0], child.marks) });
    } else out.push({ key: `\u0000${child.type.name}:${JSON.stringify(child.attrs)}`, node: child });
  });
  return out;
}

/** Inline content marked as a revision (keeping its other marks; revisions it had are replaced). */
function marked(nodes: Node[], revision: Mark): Node[] {
  return nodes.map((n) => n.mark(revision.addToSet(n.marks.filter((m) => m.type !== revision.type))));
}

function withParagraphRevision(p: Node, kind: "ins" | "del", author: string, date: string, id: () => number): Node {
  const props = { ...(p.attrs.props ?? {}) };
  const markRun = { ...(props.markRun ?? {}) };
  const extra = ((markRun.extra ?? []) as { name: string; xml: string }[]).filter((x) => x.name !== "w:ins" && x.name !== "w:del");
  const esc = author.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  extra.unshift({ name: `w:${kind}`, xml: `<w:${kind} w:id="${id()}" w:author="${esc}" w:date="${date}"/>` });
  return p.type.create({ ...p.attrs, props: { ...props, markRun: { ...markRun, extra } } }, p.content, p.marks);
}

/** A whole block as deleted or inserted: its text marked, its paragraph marks too. */
function wholeBlock(block: Node, kind: "ins" | "del", schema: Schema, author: string, date: string, id: () => number): Node {
  const revision = schema.marks.revision.create({ kind, attrs: { "w:author": author, "w:date": date } });
  const walk = (n: Node): Node => {
    if (n.isTextblock) {
      const inner: Node[] = [];
      n.forEach((c) => inner.push(c));
      const p = n.type.create(n.attrs, Fragment.fromArray(marked(inner, revision)), n.marks);
      return n.type.name === "paragraph" ? withParagraphRevision(p, kind, author, date, id) : p;
    }
    if (n.isLeaf) return n;
    const kids: Node[] = [];
    n.forEach((c) => kids.push(walk(c)));
    return n.type.create(n.attrs, Fragment.fromArray(kids), n.marks);
  };
  return walk(block);
}

/** A paragraph changed: compared word by word (the revised paragraph's settings kept). */
function mergedParagraph(a: Node, b: Node, schema: Schema, author: string, date: string): Node {
  const ta = tokens(a);
  const tb = tokens(b);
  const del = schema.marks.revision.create({ kind: "del", attrs: { "w:author": author, "w:date": date } });
  const ins = schema.marks.revision.create({ kind: "ins", attrs: { "w:author": author, "w:date": date } });
  const out: Node[] = [];
  // Deletions in a run go before the insertions that replace them (as Word shows "old new").
  let dels: Node[] = [];
  let inss: Node[] = [];
  const flush = () => { out.push(...marked(dels, del), ...marked(inss, ins)); dels = []; inss = []; };
  for (const e of diff(ta.map((t) => t.key), tb.map((t) => t.key))) {
    if (e.op === "equal") { flush(); out.push(tb[e.b!].node); }
    else if (e.op === "delete") dels.push(ta[e.a!].node);
    else inss.push(tb[e.b!].node);
  }
  flush();
  return b.type.create(b.attrs, Fragment.fromArray(out), b.marks);
}

/** How alike two paragraphs' words are (0–1): worth comparing word by word, or a replacement. */
function likeness(a: Node, b: Node): number {
  const wa = a.textContent.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
  const wb = b.textContent.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
  if (!wa.length && !wb.length) return 1;
  const sb = new Map<string, number>();
  for (const w of wb) sb.set(w, (sb.get(w) ?? 0) + 1);
  let common = 0;
  for (const w of wa) { const c = sb.get(w); if (c) { common++; sb.set(w, c - 1); } }
  return (2 * common) / (wa.length + wb.length);
}

/**
 * The comparison of two documents (same schema): the revised document, with what changed from the
 * original as tracked changes by `author` at `date`.
 */
export function compareDocuments(original: Node, revised: Node, author: string, date: string): Node {
  const schema = revised.type.schema;
  let n = 700000 + Math.floor(Math.random() * 50000);
  const id = () => n++;
  const A: Node[] = [];
  const B: Node[] = [];
  original.forEach((c) => A.push(c));
  revised.forEach((c) => B.push(c));
  const edits = diff(A.map(keyOf), B.map(keyOf));
  const out: Node[] = [];
  // A run of deleted and inserted blocks: like paragraphs paired and compared word by word.
  let gone: Node[] = [];
  let added: Node[] = [];
  const flush = () => {
    let i = 0;
    let j = 0;
    while (i < gone.length || j < added.length) {
      const a = gone[i];
      const b = added[j];
      if (a && b && a.type.name === "paragraph" && b.type.name === "paragraph" && likeness(a, b) >= 0.4) {
        out.push(mergedParagraph(a, b, schema, author, date)); i++; j++;
      } else if (a && (!b || gone.length - i >= added.length - j)) {
        out.push(wholeBlock(a, "del", schema, author, date, id)); i++;
      } else {
        out.push(wholeBlock(b!, "ins", schema, author, date, id)); j++;
      }
    }
    gone = [];
    added = [];
  };
  for (const e of edits) {
    if (e.op === "equal") { flush(); out.push(B[e.b!]); }
    else if (e.op === "delete") gone.push(A[e.a!]);
    else added.push(B[e.b!]);
  }
  flush();
  return revised.type.create(revised.attrs, Fragment.fromArray(out));
}
