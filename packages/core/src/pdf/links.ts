import { LaidOutPage } from "../layout/types.js";
import { isSafeUrl } from "../url-safety.js";
import { pdfString } from "./writer.js";

/**
 * The clickable boxes of a page's links, for its /Link annotations: one per
 * line of each link, over the text the link draws, as Word's PDFs have
 * them. A link's text items on one line (same lineTop, same href, no other
 * source text between them) make one box from the leftmost to the
 * rightmost, the spaces and tab leaders between them covered too; a second
 * link to the same place after other text on the line is a box of its own.
 */
export interface LinkBox {
  href: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

export function linkBoxes(page: LaidOutPage): LinkBox[] {
  const boxes: LinkBox[] = [];
  let open: LinkBox | null = null;
  for (const item of page.items) {
    if (item.kind !== "text" || item.hidden || !item.text) continue;
    if (!item.href) {
      // A space between a link's words is laid without the link, and a tab
      // leader (no source text) between a TOC entry and its page number is
      // part of the entry, as Word's one box over the whole line has it.
      if (item.src && item.text.trim()) open = null;
      continue;
    }
    if (open && open.href === item.href && Math.abs(open.y - item.lineTop) < 0.5) {
      const right = Math.max(open.x + open.width, item.x + item.width);
      open.x = Math.min(open.x, item.x);
      open.width = right - open.x;
      open.height = Math.max(open.height, item.lineHeight);
      continue;
    }
    open = { href: item.href, x: item.x, y: item.lineTop, width: item.width, height: item.lineHeight };
    boxes.push(open);
  }
  return boxes;
}

/** The bookmarks that begin on a page's drawn text, each with the top of its line. */
export function bookmarkTops(page: LaidOutPage): Map<string, number> {
  const tops = new Map<string, number>();
  for (const item of page.items) {
    if (item.kind !== "text" || !item.bookmarks) continue;
    for (const name of item.bookmarks) if (!tops.has(name)) tops.set(name, item.lineTop);
  }
  return tops;
}

/**
 * What a link annotation does for an href: opens the URL (one isSafeUrl
 * allows: http, https, mailto, or no scheme), or goes to the named
 * destination of the bookmark a "#name" href points at, when the document
 * has the bookmark. Undefined for any other href, which gets no annotation.
 */
export function linkAction(href: string, bookmarks: Map<string, number>): string | undefined {
  if (href.startsWith("#")) return bookmarks.has(href.slice(1)) ? `/Dest ${pdfName(href.slice(1))}` : undefined;
  if (!isSafeUrl(href)) return undefined;
  // A URI string is ASCII: other characters go percent-encoded as UTF-8.
  return `/A << /S /URI /URI ${pdfString(href.replace(/[^\x21-\x7e]/g, (c) => encodeURIComponent(c)))} >>`;
}

/** A PDF name object for a bookmark name: `/` and the name's UTF-8 bytes,
 * each outside the regular printable characters as #XX. */
export function pdfName(name: string): string {
  let out = "/";
  for (const byte of Buffer.from(name, "utf8")) {
    const regular = byte > 0x20 && byte < 0x7f && !"()<>[]{}/%#".includes(String.fromCharCode(byte));
    out += regular ? String.fromCharCode(byte) : `#${byte.toString(16).padStart(2, "0")}`;
  }
  return out;
}
