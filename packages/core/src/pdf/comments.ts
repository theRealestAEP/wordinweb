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
 * A dashed leader in the author's color runs from the end of the comment's
 * first highlighted line across to the rail and on to the card's left edge.
 *
 * Balloons that do not fit: cards are placed in order down the rail. A
 * card that would cross the sheet's bottom moves whole to the top of the
 * next sheet's rail, ahead of that page's own cards (its leader stays on
 * the sheet with its text, reaching the rail); after the last page further
 * sheets with an empty page area carry what is left. Only a card taller
 * than a whole rail splits, at a line, and continues the same way. No
 * comment is ever cut.
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
  /** The author's color: the avatar's, and the leader line's. */
  color: string;
  lines: Line[];
}

export class CommentColumn {
  private readonly placed = new Set<string>();
  /** Lines of balloons that did not fit the sheet before. */
  private carried: Balloon[] = [];
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
    // The end of each comment's first highlighted line, where its leader starts.
    const firstLine = new Map<string, { x: number; y: number; top: number; bottom: number }>();
    for (const [id, list] of items) {
      const resolved = this.byId.get(id)?.resolved;
      // One continuous rect per line, as word-granular spans would gap at spaces.
      let run: { top: number; height: number; x0: number; x1: number } | null = null;
      const flush = (): void => {
        if (run && run.x1 > run.x0) {
          drawer.ops.push(
            `q /${this.resources.alpha(resolved ? 0.22 : 0.38)} gs ${resolved ? "0.627 0.627 0.627" : "1 0.784 0.353"} rg ${drawer.rect(run.x0, run.top, run.x1 - run.x0, run.height)} f Q`,
          );
          if (!firstLine.has(id)) firstLine.set(id, { x: run.x1, y: run.top + run.height / 2, top: run.top, bottom: run.top + run.height });
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

    const cardX = page.width + CARD_GAP;
    const bottom = sheetHeight - RAIL_PAD;
    const rail = bottom - RAIL_PAD;
    let y = RAIL_PAD;
    let stacked = false;
    const carried: Balloon[] = [];
    // The leader stays in the white margin: the right edge of the page's
    // text area is past every text item and column on the page.
    let textRight = 0;
    for (const band of page.columnBands) band.colXs.forEach((x, i) => (textRight = Math.max(textRight, x + band.colWidths[i])));
    for (const item of page.items) if (item.kind === "text" && !item.hidden) textRight = Math.max(textRight, item.x + item.width);
    textRight = Math.min(textRight + 6, page.width - 2);
    /** Mark the highlight's end with a short tick in the author's color,
     * and draw the dashed leader from the text area's right edge across the
     * margin to the rail, and on to a card at `cardTop` when it is on this
     * sheet. No segment crosses or underlines text. */
    const leader = (balloon: Balloon, cardTop: number | null): void => {
      const from = firstLine.get(balloon.id);
      if (!from) return;
      drawer.ops.push(`q 1.5 w ${rgb(balloon.color)} RG ${num(drawer.x(from.x - 1))} ${num(drawer.y(from.top + 2))} m ${num(drawer.x(from.x - 1))} ${num(drawer.y(from.bottom - 2))} l S Q`);
      const ops = [
        `q [${num(1.5)} ${num(1.5)}] 0 d 0.75 w ${rgb(balloon.color)} RG ${num(drawer.x(textRight))} ${num(drawer.y(from.y))} m ${num(drawer.x(page.width))} ${num(drawer.y(from.y))} l`,
      ];
      if (cardTop !== null) ops.push(`${num(drawer.x(cardX))} ${num(drawer.y(cardTop + PAD_Y + AVATAR / 2))} l`);
      ops.push("S Q");
      drawer.ops.push(ops.join(" "));
    };
    /** Draw a card at `at` and give its top, or carry it to the next sheet
     * (null). Only a card taller than the rail splits, at a line. */
    const place = (balloon: Balloon, at: number): number | null => {
      const height = balloon.lines.reduce((sum, line) => sum + line.height, 0);
      let lines = balloon.lines;
      if (carried.length > 0 || (at + height > bottom && height <= rail)) {
        carried.push(balloon);
        return null;
      }
      if (at + height > bottom) {
        let fit = 0;
        let fitHeight = 0;
        while (fit < lines.length && at + fitHeight + lines[fit].height <= bottom) fitHeight += lines[fit++].height;
        if (fit === 0) {
          carried.push(balloon);
          return null;
        }
        carried.push({ ...balloon, lines: lines.slice(fit) });
        lines = lines.slice(0, fit);
      }
      const drawn = lines.reduce((sum, line) => sum + line.height, 0);
      drawer.ops.push(`q 1 g 0.878 0.878 0.878 RG 0.75 w ${drawer.rect(cardX, at, CARD_WIDTH, drawn)} B Q`);
      let lineY = at;
      for (const line of lines) {
        line.paint(drawer, cardX, lineY);
        lineY += line.height;
      }
      y = at + drawn;
      stacked = true;
      return at;
    };
    for (const balloon of this.carried) place(balloon, stacked ? y + RAIL_PAD : y);
    this.carried = [];
    for (const { top, balloon } of balloons) leader(balloon, place(balloon, Math.max(top, stacked ? y + RAIL_PAD : y)));
    this.carried = carried;
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
          const dateText = date && !Number.isNaN(date.getTime()) ? `${date.toLocaleDateString("en-US", { timeZone: "UTC" })} UTC` : "";
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
    return { id: comment.id, color: avatarColor(comment.author), lines };
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
