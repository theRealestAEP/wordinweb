import { DocxDocument } from "../docx.js";
import { XmlElement, attr, localName } from "../xml.js";
import { SelectionSegment, applyRunFormat } from "./commands.js";

/**
 * Hyperlink edit commands. Like all commands, these mutate the source XML;
 * callers checkpoint history and refresh/relayout afterwards.
 */

function prefixOf(e: XmlElement): string {
  return e.name.includes(":") ? e.name.slice(0, e.name.indexOf(":") + 1) : "";
}

/** The w:hyperlink ancestor of a node, or null. */
function hyperlinkOf(doc: DocxDocument, t: XmlElement): XmlElement | null {
  let cur: XmlElement | undefined = doc.findParentOf(t);
  while (cur && localName(cur.name) !== "p") {
    if (localName(cur.name) === "hyperlink") return cur;
    cur = doc.findParentOf(cur);
  }
  return null;
}

/** The external URL of the hyperlink containing `t`, or null. */
export function linkAt(doc: DocxDocument, t: XmlElement): string | null {
  const h = hyperlinkOf(doc, t);
  if (!h) return null;
  const rid = Object.entries(h.attrs).find(([k]) => localName(k) === "id")?.[1];
  const rel = rid ? doc.documentRels.get(rid) : undefined;
  return rel?.external ? rel.target : null;
}

/**
 * Wrap the selection in a hyperlink to `url` (Word: w:hyperlink r:id +
 * Hyperlink character style). If the selection is already inside a link,
 * only the URL is retargeted. Multi-paragraph selections wrap per paragraph.
 */
export function setLink(doc: DocxDocument, segments: SelectionSegment[], url: string): boolean {
  if (!url || segments.length === 0) return false;

  // Editing an existing link: retarget its relationship.
  const existingT = segments.find((s) => s.t)?.t;
  const existing = existingT ? hyperlinkOf(doc, existingT) : null;
  if (existing) {
    const rid = Object.entries(existing.attrs).find(([k]) => localName(k) === "id")?.[1];
    if (rid && doc.setRelTarget(rid, url)) {
      doc.refresh();
      return true;
    }
    return false;
  }

  // Split partially covered runs so the link lands on run boundaries.
  const ranges = applyRunFormat(doc, segments, {});
  if (ranges.length === 0) return false;

  // Collect the runs covering the selection, grouped by paragraph.
  const runsByPara = new Map<XmlElement, XmlElement[]>();
  for (const range of ranges) {
    const rEl = doc.findParentOf(range.t);
    if (!rEl || localName(rEl.name) !== "r") continue;
    const pEl = doc.findParentOf(rEl);
    if (!pEl || localName(pEl.name) !== "p") continue; // already inside a hyperlink etc.
    const list = runsByPara.get(pEl) ?? [];
    if (!list.includes(rEl)) list.push(rEl);
    runsByPara.set(pEl, list);
  }
  if (runsByPara.size === 0) return false;

  const relId = doc.addHyperlinkRel(url);
  for (const [pEl, runs] of runsByPara) {
    const w = prefixOf(pEl);
    const first = pEl.children.indexOf(runs[0]);
    // Word styles linked runs with the Hyperlink character style.
    for (const rEl of runs) {
      let rPr = rEl.children.find((c) => localName(c.name) === "rPr");
      if (!rPr) {
        rPr = { name: `${w}rPr`, attrs: {}, children: [], text: "" };
        rEl.children.unshift(rPr);
      }
      if (!rPr.children.some((c) => localName(c.name) === "rStyle")) {
        rPr.children.unshift({ name: `${w}rStyle`, attrs: { [`${w}val`]: "Hyperlink" }, children: [], text: "" });
      }
    }
    const link: XmlElement = {
      name: `${w}hyperlink`,
      attrs: { "r:id": relId, [`${w}history`]: "1" },
      children: runs,
      text: "",
    };
    // The runs are contiguous post-split; replace them with the wrapper.
    pEl.children = pEl.children.filter((c) => !runs.includes(c));
    pEl.children.splice(first, 0, link);
    if (!pEl.attrs["xmlns:r"] && !doc.findParentOf(pEl)) {
      // namespace declared on the document root in practice; nothing to do
    }
  }

  doc.refresh();
  return true;
}

/** Unwrap the hyperlink containing `t` (text and formatting stay). */
export function removeLink(doc: DocxDocument, t: XmlElement): boolean {
  const h = hyperlinkOf(doc, t);
  if (!h) return false;
  const pEl = doc.findParentOf(h);
  if (!pEl) return false;
  const idx = pEl.children.indexOf(h);
  // Drop the Hyperlink character style from the unwrapped runs.
  for (const rEl of h.children) {
    const rPr = rEl.children.find((c) => localName(c.name) === "rPr");
    if (rPr) {
      rPr.children = rPr.children.filter(
        (c) => !(localName(c.name) === "rStyle" && attr(c, "val") === "Hyperlink"),
      );
    }
  }
  pEl.children.splice(idx, 1, ...h.children);
  doc.refresh();
  return true;
}
