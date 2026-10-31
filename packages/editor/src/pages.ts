/**
 * Maps the layout engine's page starts onto editor positions (pure, testable in Node).
 */
import type { PageStart } from "@prime/layout";
import type { Node } from "prosemirror-model";

export type ResolvedStart =
  | { page: number; kind: "block"; pos: number }      // a top-level paragraph starts the page: break before it
  | { page: number; kind: "inline"; pos: number }     // a paragraph continues onto the page: break inside it
  | { page: number; kind: "row"; pos: number };       // a table row starts the page

/** Position of the node at `path` (top-level index, then child indexes), or undefined. */
function nodeAt(doc: Node, path: number[]): { node: Node; pos: number } | undefined {
  let node = doc;
  let pos = 0;
  for (const [depth, index] of path.entries()) {
    if (index >= node.childCount) return undefined;
    let offset = 0;
    for (let i = 0; i < index; i++) offset += node.child(i).nodeSize;
    pos += offset + (depth === 0 ? 0 : 1);
    node = node.child(index);
  }
  return { node, pos };
}

export function resolvePageStarts(doc: Node, starts: PageStart[]): ResolvedStart[] {
  const out: ResolvedStart[] = [];
  for (const s of starts) {
    const target = nodeAt(doc, s.path);
    if (!target) continue;
    if (s.row !== undefined && target.node.type.name === "table") {
      if (s.row >= target.node.childCount) continue;
      let pos = target.pos + 1;
      for (let i = 0; i < s.row; i++) pos += target.node.child(i).nodeSize;
      out.push({ page: s.page, kind: "row", pos });
    } else if (target.node.type.name === "paragraph") {
      const offset = Math.min(s.offset ?? 0, target.node.content.size);
      out.push(offset === 0
        ? { page: s.page, kind: "block", pos: target.pos }
        : { page: s.page, kind: "inline", pos: target.pos + 1 + offset });
    }
  }
  return out;
}
