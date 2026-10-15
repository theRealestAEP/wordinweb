import { DocxDocument } from "../docx.js";
import { DocComment } from "../model.js";
import { FontSpec, LaidOutPage, TextItem } from "../layout/types.js";
import { XmlElement } from "../xml.js";
import { PageDrawer, parseColor } from "./draw.js";
import { NodeMeasurer } from "./measure.js";
import { Resources } from "./resources.js";
import { num } from "./writer.js";

/**
 * The review-comment column of a sheet, as renderToDom's renderComments
 * lays it out on screen: the commented text highlighted, each top-level
 * comment in a balloon in the rail right of the page beside its first
 * line, replies nested under their parent, balloons stacked down so none
 * overlaps the one above.
 *
 * Balloons that do not fit: a balloon is a stack of lines, placed in order
 * down the rail. The lines that would cross the sheet's bottom continue at
 * the top of the next sheet's rail, ahead of that page's own balloons, and
 * after the last page further sheets with an empty page area carry what is
 * left. No comment is ever cut.
 *
 * A page needs only its own comments: the column keeps the ids already
 * given a balloon and the lines carried to the next sheet, nothing else.
 */
export const RAIL_WIDTH = 256;
const CARD_WIDTH = 220;
const CARD_GAP = 12;
const PAD_X = 12;
const PAD_Y = 10;
const RAIL_PAD = 8;
const TEXT_SIZE = 12;
const LINE_HEIGHT = 15;
const AVATAR = 24;
const UI_FONT = "Arial";

/** One line of a balloon: its height and how to paint it at (x, y). */
interface Line {
  height: number;
  paint: (drawer: PageDrawer, x: number, y: number) => void;
}

interface Balloon {
  id: string;
  lines: Line[];
}

export class CommentColumn {
  private readonly placed = new Set<string>();
  /** Lines of balloons that did not fit the sheet before. */
  private carried: Line[] = [];
  private readonly replies = new Map<string, DocComment[]>();
  private readonly byId = new Map<string, DocComment>();

  constructor(
    private readonly doc: DocxDocument,
    private readonly measurer: NodeMeasurer,
    private readonly resources: Resources,
    private readonly commentIds: (t: XmlElement) => string[] | undefined,
  ) {
    for (const comment of doc.comments) {
      this.byId.set(comment.id, comment);
      if (!comment.parentId) continue;
      const list = this.replies.get(comment.parentId);
      if (list) list.push(comment);
      else this.replies.set(comment.parentId, [comment]);
    }
  }

  /** Draw the page's highlights and its rail into `drawer`. */
  draw(page: LaidOutPage, drawer: PageDrawer, sheetHeight: number): void {
    // Each top-level comment's text items on this page, in paint order.
    const items = new Map<string, TextItem[]>();
    for (const item of page.items) {
      if (item.kind !== "text" || item.hidden || !item.src?.t) continue;
      const ids = this.commentIds(item.src.t);
      if (!ids) continue;
      for (const id of ids) {
        const comment = this.byId.get(id);
        if (!comment || comment.parentId) continue;
        const list = items.get(id);
        if (list) list.push(item);
        else items.set(id, [item]);
      }
    }
    for (const [id, list] of items) {
      const resolved = this.byId.get(id)?.resolved;
      // One continuous rect per line, as word-granular spans would gap at spaces.
      let run: { top: number; height: number; x0: number; x1: number } | null = null;
      const flush = (): void => {
        if (run && run.x1 > run.x0) {
          drawer.ops.push(
            `q /${this.resources.alpha(resolved ? 0.22 : 0.38)} gs ${resolved ? "0.627 0.627 0.627" : "1 0.784 0.353"} rg ${drawer.rect(run.x0, run.top, run.x1 - run.x0, run.height)} f Q`,
          );
        }
        run = null;
      };
      for (const item of list) {
        if (run && run.top !== item.lineTop) flush();
        if (!run) run = { top: item.lineTop, height: item.lineHeight, x0: item.x, x1: item.x + item.width };
        else {
          run.x0 = Math.min(run.x0, item.x);
          run.x1 = Math.max(run.x1, item.x + item.width);
          run.height = Math.max(run.height, item.lineHeight);
        }
      }
      flush();
    }

    // Balloons for the comments whose first line is on this page, by that line.
    const balloons: { top: number; balloon: Balloon }[] = [];
    for (const [id, list] of items) {
      if (this.placed.has(id)) continue;
      this.placed.add(id);
      balloons.push({ top: list[0].lineTop, balloon: this.balloon(this.byId.get(id) as DocComment) });
    }
    balloons.sort((a, b) => a.top - b.top);

    const x = page.width + CARD_GAP;
    const bottom = sheetHeight - RAIL_PAD;
    let y = RAIL_PAD;
    let rest: Line[] = [];
    const place = (lines: Line[], at: number): number => {
      if (rest.length > 0) {
        rest.push(...lines);
        return at;
      }
      let fit = 0;
      let height = 0;
      while (fit < lines.length && at + height + lines[fit].height <= bottom) height += lines[fit++].height;
      if (fit > 0) {
        drawer.ops.push(`q 1 g 0.878 0.878 0.878 RG 0.75 w ${drawer.rect(x, at, CARD_WIDTH, height)} B Q`);
        let lineY = at;
        for (const line of lines.slice(0, fit)) {
          line.paint(drawer, x, lineY);
          lineY += line.height;
        }
      }
      rest = lines.slice(fit);
      return at + height;
    };
    let stacked = false;
    if (this.carried.length > 0) {
      y = place(this.carried, y);
      this.carried = [];
      stacked = true;
    }
    for (const { top, balloon } of balloons) {
      y = place(balloon.lines, Math.max(top, stacked ? y + RAIL_PAD : y));
      stacked = true;
    }
    this.carried = rest;
  }

  /** Whether lines wait for another sheet after the last page. */
  hasCarried(): boolean {
    return this.carried.length > 0;
  }

  /** The lines of a comment's balloon: its head, text, and replies. */
  private balloon(comment: DocComment): Balloon {
    const lines: Line[] = [{ height: PAD_Y, paint: () => undefined }];
    const inner = CARD_WIDTH - 2 * PAD_X;
    const resolved = !!comment.resolved;
    const fg = resolved ? "#5f6368" : "#3c4043";
    const push = (indent: number, source: DocComment): void => {
      const left = PAD_X + indent;
      const width = inner - indent;
      lines.push({
        height: AVATAR + 6,
        paint: (drawer, x, y) => {
          drawer.ops.push(`q ${rgb(avatarColor(source.author))} rg ${circle(drawer, x + left + AVATAR / 2, y + AVATAR / 2, AVATAR / 2)} f Q`);
          const initials = initialsOf(source.author, source.initials);
          const initialsFont: FontSpec = { family: UI_FONT, size: 10, bold: true, italic: false };
          const iw = this.measurer.width(initials, initialsFont);
          this.text(drawer, initials, initialsFont, "#ffffff", x + left + AVATAR / 2 - iw / 2, y + AVATAR / 2 + 3.5);
          const author = source.author || (indent ? "Reply" : "Comment");
          const authorFont: FontSpec = { family: UI_FONT, size: TEXT_SIZE, bold: true, italic: false };
          const date = source.date ? new Date(source.date) : null;
          const dateText = date && !Number.isNaN(date.getTime()) ? date.toLocaleDateString() : "";
          const textX = x + left + AVATAR + 8;
          this.text(drawer, this.clip(author, authorFont, width - AVATAR - 8), authorFont, fg, textX, y + (dateText ? 10 : 16));
          if (dateText) this.text(drawer, dateText, { family: UI_FONT, size: 11, bold: false, italic: false }, "#5f6368", textX, y + 22);
        },
      });
      const font: FontSpec = { family: UI_FONT, size: TEXT_SIZE, bold: false, italic: false };
      for (const text of this.wrap(source.text, font, width)) {
        lines.push({ height: LINE_HEIGHT, paint: (drawer, x, y) => this.text(drawer, text, font, fg, x + left, y + 12) });
      }
    };
    push(0, comment);
    for (const reply of this.replies.get(comment.id) ?? []) {
      // The reply block: a rule above, a bar down its left side.
      lines.push({
        height: 8 + 8,
        paint: (drawer, x, y) => {
          drawer.ops.push(`q 0.945 0.953 0.957 rg ${drawer.rect(x + PAD_X, y + 8, inner, 1)} f Q`);
        },
      });
      const start = lines.length;
      push(10, reply);
      const bar = lines.slice(start);
      for (const line of bar) {
        const paint = line.paint;
        line.paint = (drawer, x, y) => {
          drawer.ops.push(`q 0.91 0.918 0.929 rg ${drawer.rect(x + PAD_X, y, 2, line.height)} f Q`);
          paint(drawer, x, y);
        };
      }
    }
    lines.push({ height: PAD_Y, paint: () => undefined });
    return { id: comment.id, lines };
  }

  private text(drawer: PageDrawer, text: string, font: FontSpec, color: string, x: number, baseline: number): void {
    if (!text) return;
    const item: TextItem = {
      kind: "text",
      x,
      baseline,
      width: this.measurer.width(text, font),
      text,
      props: { color },
      font,
      lineTop: baseline - font.size,
      lineHeight: font.size * 1.25,
    };
    drawer.drawItem(item, () => null);
  }

  /** Greedy word wrap of a comment's paragraphs to `width` px; a word
   * longer than the line breaks where it must. */
  private wrap(text: string, font: FontSpec, width: number): string[] {
    const out: string[] = [];
    for (const paragraph of text.split("\n")) {
      let line = "";
      for (const word of paragraph.split(/(\s+)/)) {
        if (!word) continue;
        if (this.measurer.width(line + word, font) <= width || (!line && !/\S/.test(word))) {
          line += word;
          continue;
        }
        if (/^\s+$/.test(word)) {
          out.push(line);
          line = "";
          continue;
        }
        if (line) out.push(line.trimEnd());
        line = "";
        for (const ch of word) {
          if (line && this.measurer.width(line + ch, font) > width) {
            out.push(line);
            line = "";
          }
          line += ch;
        }
      }
      out.push(line.trimEnd());
    }
    return out;
  }

  private clip(text: string, font: FontSpec, width: number): string {
    if (this.measurer.width(text, font) <= width) return text;
    let out = "";
    for (const ch of text) {
      if (this.measurer.width(out + ch + "…", font) > width) return out + "…";
      out += ch;
    }
    return out;
  }
}

function circle(drawer: PageDrawer, cx: number, cy: number, r: number): string {
  const k = 0.5523 * r;
  const p = (x: number, y: number): string => `${num(drawer.x(x))} ${num(drawer.y(y))}`;
  return `${p(cx + r, cy)} m ${p(cx + r, cy + k)} ${p(cx + k, cy + r)} ${p(cx, cy + r)} c ${p(cx - k, cy + r)} ${p(cx - r, cy + k)} ${p(cx - r, cy)} c ${p(cx - r, cy - k)} ${p(cx - k, cy - r)} ${p(cx, cy - r)} c ${p(cx + k, cy - r)} ${p(cx + r, cy - k)} ${p(cx + r, cy)} c`;
}

function rgb(css: string): string {
  const c = parseColor(css) ?? { r: 0, g: 0, b: 0, a: 1 };
  return `${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)}`;
}

function initialsOf(author: string, initials?: string): string {
  if (initials) return initials.slice(0, 2).toUpperCase();
  const parts = author.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  return ((parts[0][0] ?? "") + (parts.length > 1 ? parts[parts.length - 1][0] : "")).toUpperCase();
}

/** Stable per-author avatar color, the DOM renderer's palette. */
function avatarColor(author: string): string {
  const palette = ["#1a73e8", "#188038", "#a50e0e", "#8430ce", "#007b83", "#b06000"];
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) | 0;
  return palette[Math.abs(h) % palette.length];
}
