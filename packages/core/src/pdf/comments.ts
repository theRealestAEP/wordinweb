import { DocxDocument } from "../docx.js";
import { DocComment, Revision } from "../model.js";
import { FontSpec, LaidOutPage, TextItem } from "../layout/types.js";
import { StreamingDocument } from "../stream/body.js";
import { XmlElement, attr, localName } from "../xml.js";
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
 * Balloons that do not fit: each sheet holds all of its page's cards,
 * whole, in order down the rail. When they reach past the sheet's bottom,
 * that sheet grows taller (layout gives the height): the page keeps its
 * size at the top, and the rail goes on below it. So no card leaves its
 * page's sheet, every leader reaches its card, and the PDF has one sheet
 * per page.
 *
 * A page needs only its own comments: the column keeps the ids already
 * given a balloon, nothing else.
 *
 * Every comment the document marks gets its card. One whose text is never
 * drawn (clipped by an exact-height row, hidden past a text box's bottom)
 * goes beside the nearest drawn text of its block, where the counting pass
 * found it (UndrawnComments); one with no drawn text near it at all
 * (marked at no body text, or in a header or footer) goes after the last
 * page's own cards, on the last sheet.
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
  /** How far below the card's top the leader meets it: the avatar's
   * middle, or with no avatar the first line's. */
  leaderAt: number;
}

/** The line a card goes beside: its page, and where its leader starts. */
export interface LineAnchor {
  page: number;
  x: number;
  top: number;
  bottom: number;
}

/** What the column asks of the streamed document about its comments. */
export type CommentSource = Pick<StreamingDocument, "commentIds" | "textPosition" | "commentAnchor">;

/** The body w:t, w:delText or equation a text item draws, when it draws one. */
function textOf(item: TextItem): XmlElement | undefined {
  return item.src?.t ?? item.mathSrc;
}

/**
 * The counting pass's look for the comments none of whose text is drawn:
 * it sees every page, so it knows which comments no page shows and, for
 * each, the drawn text nearest its anchor in the same block (the row's
 * visible lines, the paragraph's), by the texts' order; the one before the
 * anchor wins a tie. The rendering pass puts the card beside that line.
 */
export class UndrawnComments {
  private drawn = new Set<string>();
  private nearest = new Map<string, { distance: number; anchor: LineAnchor }>();
  /** The top-level comments anchored in each block. */
  private readonly byBlock = new Map<number, string[]>();

  constructor(
    doc: DocxDocument,
    private readonly source: CommentSource,
  ) {
    for (const comment of doc.comments) {
      if (comment.parentId) continue;
      const at = source.commentAnchor(comment.id);
      if (!at) continue;
      const list = this.byBlock.get(at.block);
      if (list) list.push(comment.id);
      else this.byBlock.set(at.block, [comment.id]);
    }
  }

  /** Forget the pages seen: a pass starts. */
  reset(): void {
    this.drawn = new Set();
    this.nearest = new Map();
  }

  visit(page: LaidOutPage): void {
    for (const item of page.items) {
      if (item.kind !== "text" || item.hidden) continue;
      const t = textOf(item);
      if (!t) continue;
      for (const id of this.source.commentIds(t) ?? []) this.drawn.add(id);
      const position = this.source.textPosition(t);
      const ids = position && this.byBlock.get(position.block);
      if (!ids) continue;
      for (const id of ids) {
        if (this.drawn.has(id)) continue;
        const anchor = this.source.commentAnchor(id) as { ordinal: number };
        const distance = Math.abs(anchor.ordinal - position.ordinal) * 2 + (position.ordinal > anchor.ordinal ? 1 : 0);
        const best = this.nearest.get(id);
        if (best && best.distance <= distance) continue;
        this.nearest.set(id, { distance, anchor: { page: page.index, x: item.x + item.width, top: item.lineTop, bottom: item.lineTop + item.lineHeight } });
      }
    }
  }

  /** After the pass: where each comment no page drew goes. */
  anchors(): Map<string, LineAnchor> {
    const out = new Map<string, LineAnchor>();
    for (const [id, { anchor }] of this.nearest) if (!this.drawn.has(id)) out.set(id, anchor);
    return out;
  }
}

export class CommentColumn {
  private readonly placed = new Set<string>();
  private readonly replies = new Map<string, DocComment[]>();
  private readonly byId = new Map<string, DocComment>();

  constructor(
    private readonly doc: DocxDocument,
    private readonly measurer: NodeMeasurer,
    private readonly resources: Resources,
    private readonly source: CommentSource,
    /** The cards of the comments no page draws the text of (UndrawnComments.anchors). */
    private readonly undrawn: Map<string, LineAnchor>,
    /** Name the authors (RenderPdfOptions.authors). Without, a card has no
     * name line and no avatar, and every comment's leader is one color. */
    private readonly authors = true,
  ) {
    // The column prints the review comments; with markup, the tracked
    // changes too: a balloon per change (Run.changes) with its author,
    // date and what it is, in the author's revision color, beside the
    // change's first line. The inline marks say where the change is, so
    // no highlight is drawn for it.
    for (const comment of doc.comments) {
      this.byId.set(comment.id, comment);
      if (!comment.parentId) continue;
      const list = this.replies.get(comment.parentId);
      if (list) list.push(comment);
      else this.replies.set(comment.parentId, [comment]);
    }
    // A card names its author as people.xml does: the person's name, found
    // by the name the comment carries or by the identity its presence
    // provider knows (the userId, or the address in an "S::address::id" one).
    for (const person of doc.people) {
      if (!person.author) continue;
      this.names.set(person.author, person.author);
      if (!person.userId) continue;
      this.names.set(person.userId, person.author);
      const address = /^S::([^:]+)::/.exec(person.userId)?.[1];
      if (address) this.names.set(address, person.author);
    }
    // A comment the body never marks may be marked in a header, footer or
    // note; one marked nowhere is not in the document (Word shows none).
    const mark = (el: XmlElement): void => {
      const name = localName(el.name);
      if (name === "commentRangeStart" || name === "commentReference") {
        const id = attr(el, "id");
        if (id !== undefined) this.markedElsewhere.add(id);
        return;
      }
      for (const child of el.children) mark(child);
    };
    for (const root of doc.revisionRoots()) mark(root);
  }

  private readonly markedElsewhere = new Set<string>();
  private readonly names = new Map<string, string>();

  /** Lay out the page's highlights and its rail: the height of the sheet
   * that holds them (`sheetHeight`, or more when the cards reach past its
   * bottom), and how to draw them on a drawer for a sheet that tall. On the
   * last page, the cards of the comments no page placed go after its own. */
  layout(page: LaidOutPage, sheetHeight: number, last = false): { height: number; draw: (drawer: PageDrawer) => void } {
    // Each top-level comment's text items on this page, in paint order.
    const items = new Map<string, TextItem[]>();
    for (const item of page.items) {
      if (item.kind !== "text" || item.hidden) continue;
      const t = textOf(item);
      const ids = t && this.source.commentIds(t);
      if (!ids) continue;
      for (const id of ids) {
        const comment = this.byId.get(id);
        if (!comment || comment.parentId) continue;
        const list = items.get(id);
        if (list) list.push(item);
        else items.set(id, [item]);
      }
    }
    // Each tracked change's first line on this page, in paint order (a
    // changed picture's box counts as a line), and where the change ends on
    // it: its last item, its last item that is more than spaces (ink), and
    // the middle of the clear space the layout keeps after a change that
    // ends inside a word (gap, TextItem.changeGap).
    type ChangeLine = { top: number; bottom: number; x1: number; ink?: number; gap?: number };
    const changes = new Map<string, ChangeLine & { change: Revision }>();
    for (const item of page.items) {
      let list: Revision[] | undefined;
      let line: ChangeLine;
      if (item.kind === "text" && !item.hidden) {
        list = item.src?.run.changes;
        line = { top: item.lineTop, bottom: item.lineTop + item.lineHeight, x1: item.x + item.width };
        if (item.changeGap) line.gap = item.x + item.width / 2;
        else if (item.text.trim()) line.ink = item.x + item.width;
      } else if (item.kind === "image") {
        list = item.changes;
        line = { top: item.y, bottom: item.y + item.height, x1: item.x + item.width, ink: item.x + item.width };
      } else continue;
      for (const change of list ?? []) {
        const first = changes.get(change.id);
        if (!first) changes.set(change.id, { change, ...line });
        else if (line.top === first.top) {
          first.x1 = Math.max(first.x1, line.x1);
          if (line.ink !== undefined) first.ink = Math.max(first.ink ?? line.ink, line.ink);
          if (line.gap !== undefined) first.gap = line.gap;
        }
      }
    }
    // The end of each comment's first highlighted line, where its leader
    // starts. A change's bar goes in the clear space after it when it ends
    // inside a word, else just past its last letter, in the space or the
    // margin after it.
    const firstLine = new Map<string, { x: number; y: number; top: number; bottom: number }>();
    for (const [id, { top, bottom, x1, ink, gap }] of changes) {
      let x = x1;
      if (gap !== undefined && gap > (ink ?? -Infinity)) x = gap;
      else if (ink !== undefined) x = ink + 1;
      firstLine.set(`rev:${id}`, { x, y: (top + bottom) / 2, top, bottom });
    }
    const highlights: { top: number; height: number; x0: number; x1: number; resolved: boolean }[] = [];
    for (const [id, list] of items) {
      const resolved = !!this.byId.get(id)?.resolved;
      // One continuous rect per line, as word-granular spans would gap at spaces.
      let run: { top: number; height: number; x0: number; x1: number } | null = null;
      const flush = (): void => {
        if (run && run.x1 > run.x0) {
          highlights.push({ ...run, resolved });
          if (!firstLine.has(id)) firstLine.set(id, { x: run.x1 - 1, y: run.top + run.height / 2, top: run.top, bottom: run.top + run.height });
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
    for (const [id, { change, top }] of changes) {
      if (this.placed.has(`rev:${id}`)) continue;
      this.placed.add(`rev:${id}`);
      const card: DocComment = { id: `rev:${id}`, author: change.author, date: change.date, text: change.label };
      balloons.push({ top, balloon: this.balloon(card, this.doc.revisionColor(change.author)) });
    }
    // The comments whose text is not drawn, beside the nearest text that is.
    for (const [id, at] of this.undrawn) {
      const comment = this.byId.get(id);
      if (at.page !== page.index || !comment || this.placed.has(id)) continue;
      this.placed.add(id);
      firstLine.set(id, { x: at.x - 1, y: (at.top + at.bottom) / 2, top: at.top, bottom: at.bottom });
      balloons.push({ top: at.top, balloon: this.balloon(comment) });
    }
    balloons.sort((a, b) => a.top - b.top);
    if (last) balloons.push(...this.remaining().map((balloon) => ({ top: 0, balloon })));

    // Each card at its line, or just below the card above, whichever is lower.
    const cards: { top: number; balloon: Balloon }[] = [];
    let y = RAIL_PAD;
    for (const { top, balloon } of balloons) {
      const at = Math.max(top, y);
      cards.push({ top: at, balloon });
      y = at + balloon.lines.reduce((sum, line) => sum + line.height, 0) + RAIL_PAD;
    }

    const cardX = page.width + CARD_GAP;
    // The leader stays in the white margin: the right edge of the page's
    // text area is past every text item and column on the page.
    let textRight = 0;
    for (const band of page.columnBands) band.colXs.forEach((x, i) => (textRight = Math.max(textRight, x + band.colWidths[i])));
    for (const item of page.items) if (item.kind === "text" && !item.hidden) textRight = Math.max(textRight, item.x + item.width);
    textRight = Math.min(textRight + 6, page.width - 2);
    /** Mark the highlight's or the change's end with a hairline bar in the
     * author's color (a heavier bar hid the letters beside it), and draw
     * the dashed leader from the text area's right edge across the margin
     * to the rail, and on to the card at `cardTop`. No segment crosses or
     * underlines text. */
    const leader = (drawer: PageDrawer, balloon: Balloon, cardTop: number): void => {
      const from = firstLine.get(balloon.id);
      if (!from) return;
      drawer.ops.push(`q 0.75 w ${rgb(balloon.color)} RG ${num(drawer.x(from.x))} ${num(drawer.y(from.top + 2))} m ${num(drawer.x(from.x))} ${num(drawer.y(from.bottom - 2))} l S Q`);
      drawer.ops.push(
        `q [${num(1.5)} ${num(1.5)}] 0 d 0.75 w ${rgb(balloon.color)} RG ${num(drawer.x(textRight))} ${num(drawer.y(from.y))} m ${num(drawer.x(page.width))} ${num(drawer.y(from.y))} l ` +
          `${num(drawer.x(cardX))} ${num(drawer.y(cardTop + balloon.leaderAt))} l S Q`,
      );
    };
    const card = (drawer: PageDrawer, balloon: Balloon, top: number): void => {
      const height = balloon.lines.reduce((sum, line) => sum + line.height, 0);
      drawer.ops.push(`q 1 g 0.878 0.878 0.878 RG 0.75 w ${drawer.rect(cardX, top, CARD_WIDTH, height)} B Q`);
      let lineY = top;
      for (const line of balloon.lines) {
        line.paint(drawer, cardX, lineY);
        lineY += line.height;
      }
    };
    return {
      height: Math.max(sheetHeight, Math.ceil(y)),
      draw: (drawer) => {
        for (const { top, height, x0, x1, resolved } of highlights) {
          drawer.ops.push(`q /${this.resources.alpha(resolved ? 0.22 : 0.38)} gs ${resolved ? "0.627 0.627 0.627" : "1 0.784 0.353"} rg ${drawer.rect(x0, top, x1 - x0, height)} f Q`);
        }
        for (const { top, balloon } of cards) {
          card(drawer, balloon, top);
          leader(drawer, balloon, top);
        }
      },
    };
  }

  /** The balloons of the comments the document marks that are not yet placed, in comments.xml order. */
  private remaining(): Balloon[] {
    const out: Balloon[] = [];
    for (const comment of this.doc.comments) {
      if (comment.parentId || this.placed.has(comment.id)) continue;
      if (!this.source.commentAnchor(comment.id) && !this.markedElsewhere.has(comment.id)) continue;
      this.placed.add(comment.id);
      out.push(this.balloon(comment));
    }
    return out;
  }

  /** The lines of a comment's balloon: its head, text, and its thread, each
   * reply under the comment it answers, stepped in by its depth. A tracked
   * change's balloon takes its author's revision color. The head is the
   * avatar, the author's name and the date; without authors, or for an
   * author with no name, the date alone, on a line of its own (none when
   * there is no date). As in Word, a comment with no author shows no date. */
  private balloon(comment: DocComment, color = this.authors ? avatarColor(comment.author) : AVATAR_COLORS[0]): Balloon {
    const lines: Line[] = [{ height: PAD_Y, paint: () => undefined }];
    const inner = CARD_WIDTH - 2 * PAD_X;
    const resolved = !!comment.resolved;
    const fg = resolved ? "#5f6368" : "#3c4043";
    const change = comment.id.startsWith("rev:");
    const nameOf = (source: DocComment): string => this.names.get(source.author) ?? source.author;
    const push = (indent: number, source: DocComment): void => {
      const left = PAD_X + indent;
      const width = inner - indent;
      const name = nameOf(source);
      const date = source.date && (name || change) ? new Date(source.date) : null;
      const dateText = date && !Number.isNaN(date.getTime()) ? `${date.toLocaleDateString("en-US", { timeZone: "UTC" })} UTC` : "";
      const dateFont: FontSpec = { family: UI_FONT, size: 11, bold: false, italic: false };
      if (!this.authors || !name) {
        if (dateText) lines.push({ height: LINE_HEIGHT + 4, paint: (drawer, x, y) => this.text(drawer, dateText, dateFont, "#5f6368", x + left, y + 12) });
      } else lines.push({
        height: AVATAR + 6,
        paint: (drawer, x, y) => {
          drawer.ops.push(`q ${rgb(source === comment ? color : avatarColor(source.author))} rg ${circle(drawer, x + left + AVATAR / 2, y + AVATAR / 2, AVATAR / 2)} f Q`);
          const initials = initialsOf(name, source.initials);
          const initialsFont: FontSpec = { family: UI_FONT, size: 10, bold: true, italic: false };
          const iw = this.measurer.width(initials, initialsFont);
          this.text(drawer, initials, initialsFont, "#ffffff", x + left + AVATAR / 2 - iw / 2, y + AVATAR / 2 + 3.5);
          const authorFont: FontSpec = { family: UI_FONT, size: TEXT_SIZE, bold: true, italic: false };
          const textX = x + left + AVATAR + 8;
          this.text(drawer, this.clip(name, authorFont, width - AVATAR - 8), authorFont, fg, textX, y + (dateText ? 10 : 16));
          if (dateText) this.text(drawer, dateText, dateFont, "#5f6368", textX, y + 22);
        },
      });
      const font: FontSpec = { family: UI_FONT, size: TEXT_SIZE, bold: false, italic: false };
      for (const text of this.wrap(source.text, font, width)) {
        lines.push({ height: LINE_HEIGHT, paint: (drawer, x, y) => this.text(drawer, text, font, fg, x + left, y + 12) });
      }
    };
    push(0, comment);
    const seen = new Set<string>();
    const thread = (parent: DocComment, indent: number): void => {
      for (const reply of this.replies.get(parent.id) ?? []) {
        // A malformed thread may answer itself.
        if (seen.has(reply.id)) continue;
        seen.add(reply.id);
        // The reply block: a rule above, a bar down its left side, both at
        // the depth of the comment it answers.
        const bar = PAD_X + indent - 10;
        lines.push({
          height: 8 + 8,
          paint: (drawer, x, y) => {
            drawer.ops.push(`q 0.945 0.953 0.957 rg ${drawer.rect(x + bar, y + 8, inner - (bar - PAD_X), 1)} f Q`);
          },
        });
        const start = lines.length;
        push(indent, reply);
        for (const line of lines.slice(start)) {
          const paint = line.paint;
          line.paint = (drawer, x, y) => {
            drawer.ops.push(`q 0.91 0.918 0.929 rg ${drawer.rect(x + bar, y, 2, line.height)} f Q`);
            paint(drawer, x, y);
          };
        }
        thread(reply, Math.min(indent + 10, 30));
      }
    };
    thread(comment, 10);
    lines.push({ height: PAD_Y, paint: () => undefined });
    return { id: comment.id, color, lines, leaderAt: PAD_Y + (this.authors && nameOf(comment) ? AVATAR : LINE_HEIGHT) / 2 };
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

/** The DOM renderer's avatar palette. */
const AVATAR_COLORS = ["#1a73e8", "#188038", "#a50e0e", "#8430ce", "#007b83", "#b06000"];

/** Stable per-author avatar color, the DOM renderer's. */
function avatarColor(author: string): string {
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) | 0;
  return AVATAR_COLORS[Math.abs(h) % AVATAR_COLORS.length];
}
