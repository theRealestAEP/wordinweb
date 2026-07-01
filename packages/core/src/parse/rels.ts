import { XmlElement, attr, children } from "../xml.js";
import { resolvePartPath } from "../zip.js";

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

export type Relationships = Map<string, Relationship>;

export function parseRelationships(root: XmlElement | undefined, sourcePart: string): Relationships {
  const rels: Relationships = new Map();
  if (!root) return rels;
  for (const r of children(root, "Relationship")) {
    const id = attr(r, "Id");
    const type = attr(r, "Type") ?? "";
    const rawTarget = attr(r, "Target") ?? "";
    if (!id) continue;
    const external = attr(r, "TargetMode") === "External";
    rels.set(id, {
      id,
      type,
      target: external ? rawTarget : resolvePartPath(sourcePart, rawTarget),
      external,
    });
  }
  return rels;
}

/** Path of the .rels part for a given part ("word/document.xml" → "word/_rels/document.xml.rels"). */
export function relsPathFor(part: string): string {
  const idx = part.lastIndexOf("/");
  const dir = idx === -1 ? "" : part.slice(0, idx + 1);
  const file = idx === -1 ? part : part.slice(idx + 1);
  return `${dir}_rels/${file}.rels`;
}
