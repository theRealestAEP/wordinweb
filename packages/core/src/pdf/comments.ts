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
 * The balloon column of a sheet, as Word prints "Show Revisions in
 * Balloons": a thin grey rule along the page's right edge, and beside the
 * page each comment and each change that Word balloons, in plain text with
 * a thin bar in its author's color down its left side. A balloon reads the
 * author's name in bold, then, stepped in, what changed: "Deleted: the"
 * (the deleted text in its own formatting, wrapped, never cut short),
 * "Formatted: Font: Bold", "Moved (insertion) [1]", or the comment's text
 * with its replies under it. No date, no avatar: Word prints neither.
 *
 * Which changes get a balloon: a deletion whose text left the line (see
 * Run.inBalloon), a deleted paragraph mark ("Deleted: ¶"), a deleted
 * table row, a move's destination, and every format change. An insertion
 * stays on its line, underlined in its author's color, with no balloon;
 * so does a deleted picture or shape, struck through.
 *
 * The commented text is highlighted. A dashed leader in the author's color
 * runs from the change or the comment along the top of its first line to
 * the text area's right edge, then across to the balloon's bar; it slopes
 * down when the balloon above pushed this one down.
 *
 * Balloons that do not fit: each sheet holds all of its page's balloons,
 * whole, in order down the column. When they reach past the sheet's
 * bottom, that sheet grows taller (layout gives the height): the page
 * keeps its size at the top, and the column goes on below it. So no
 * balloon leaves its page's sheet, every leader reaches its balloon, and
 * the PDF has one sheet per page.
 *
 * A page needs only its own comments and changes: the column keeps the ids
 * already given a balloon, nothing else.
 *
 * Every comment the document marks gets its balloon. One whose text is
 * never drawn (deleted, clipped by an exact-height row, hidden past a text
 * box's bottom) goes beside the nearest drawn text of its block, where the
 * counting pass found it (UndrawnComments); one with no drawn text near it
 * at all (marked at no body text, or in a header or footer) goes after the
 * last page's own balloons, on the last sheet.
 */
export const RAIL_WIDTH = 256;
/** The bar down a balloon's left side: its left edge right of the page, and its width (1.5 pt). */
const BAR_X = 3;
const BAR_WIDTH = 2;
/** The author's name, right of the bar; the change or comment text, stepped in about 1.5 em. */
const NAME_X = BAR_X + BAR_WIDTH + 8;
const TEXT_X = NAME_X + 16;
/** The balloons' right edge, right of the page. */
const RIGHT = RAIL_WIDTH - 8;
/** Space above and below a balloon's text, inside its bar. */
const PAD_Y = 2;
/** The least space between two balloons, and above the first. */
const GAP = 12;
/** Word prints balloons at about 8 pt. */
const TEXT_SIZE = 11;
const LINE_HEIGHT = 13;
const UI_FONT = "Arial";

/** One line of a balloon: its height and how to paint it at (x, y), x the page's right edge. */
interface Line {
  height: number;
  paint: (drawer: PageDrawer, x: number, y: number) => void;
}

interface Balloon {
  id: string;
  /** The author's color: the bar's, and the leader's. */
  color: string;
  lines: Line[];
  /** How far below the balloon's top the leader meets its bar: under the
   * name line, or with no name the middle of the first line. */
  leaderAt: number;
}

/** A piece of a balloon's text in its font. */
interface Piece {
  text: string;
  font: FontSpec;
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
    /** The balloons of the comments no page draws the text of (UndrawnComments.anchors). */
    private readonly undrawn: Map<string, LineAnchor>,
    /** Name the authors (RenderPdfOptions.authors). Without, a balloon has
     * no name line, and every author's bar and leader are one color. */
    private readonly authors = true,
  ) {
    // The column prints the review comments; with markup, the tracked
    // changes Word balloons too (Run.changes), beside the change's first
    // line. The marks on the line say where the change is, so no
    // highlight is drawn for it.
    for (const comment of doc.comments) {
      this.byId.set(comment.id, comment);
      if (!comment.parentId) continue;
      const list = this.replies.get(comment.parentId);
      if (list) list.push(comment);
      else this.replies.set(comment.parentId, [comment]);
    }
    // A balloon names its author as people.xml does: the person's name, found
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

  /** Lay out the page's highlights and its column: the height of the sheet
   * that holds them (`sheetHeight`, or more when the balloons reach past its
   * bottom), and how to draw them on a drawer for a sheet that tall. On the
   * last page, the balloons of the comments no page placed go after its own. */
  layout(page: LaidOutPage, sheetHeight: number, last = false): { height: number; draw: (drawer: PageDrawer) => void } {
    // Each top-level comment's text items on this page, in paint order; and
    // where a comment on deleted text goes, which is in a balloon and not
    // on the line: at the item that stands where the text was.
    const items = new Map<string, TextItem[]>();
    const onDeleted = new Map<string, { x: number; top: number }>();
    for (const item of page.items) {
      if (item.kind !== "text" || item.hidden) continue;
      if (item.src?.run.inBalloon) {
        for (const content of item.src.run.content) {
          const ids = content.kind === "text" && content.srcT ? this.source.commentIds(content.srcT) : undefined;
          for (const id of ids ?? []) if (!onDeleted.has(id)) onDeleted.set(id, { x: item.x, top: item.lineTop });
        }
      }
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
    // Where each ballooned change's leader starts: where the change begins
    // on its first line on this page (a changed picture's box counts as a
    // line). Deleted text left the line, and an empty item marks where it
    // was (Run.inBalloon).
    const firstLine = new Map<string, { x: number; top: number }>();
    const changes: { top: number; change: Revision }[] = [];
    for (const item of page.items) {
      let list: Revision[] | undefined;
      let at: { x: number; top: number };
      if (item.kind === "text" && !item.hidden) {
        list = item.src?.run.changes;
        at = { x: item.x, top: item.lineTop };
      } else if (item.kind === "image") {
        list = item.changes;
        at = { x: item.x, top: item.y };
      } else continue;
      for (const change of list ?? []) {
        if (!ballooned(change)) continue;
        const id = `rev:${change.id}`;
        const first = firstLine.get(id);
        if (!first) {
          firstLine.set(id, at);
          changes.push({ top: at.top, change });
        } else if (at.top === first.top) first.x = Math.min(first.x, at.x);
      }
    }
    // The end of each comment's first highlighted line, where its leader starts.
    const highlights: { top: number; height: number; x0: number; x1: number; resolved: boolean }[] = [];
    for (const [id, list] of items) {
      const resolved = !!this.byId.get(id)?.resolved;
      // One continuous rect per line, as word-granular spans would gap at spaces.
      let run: { top: number; height: number; x0: number; x1: number } | null = null;
      const flush = (): void => {
        if (run && run.x1 > run.x0) {
          highlights.push({ ...run, resolved });
          if (!firstLine.has(id)) firstLine.set(id, { x: run.x1, top: run.top });
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

    // Balloons for the comments and changes whose first line is on this page, by that line.
    const balloons: { top: number; balloon: Balloon }[] = [];
    for (const [id, list] of items) {
      if (this.placed.has(id)) continue;
      this.placed.add(id);
      balloons.push({ top: list[0].lineTop, balloon: this.commentBalloon(this.byId.get(id) as DocComment) });
    }
    for (const [id, at] of onDeleted) {
      const comment = this.byId.get(id);
      if (!comment || comment.parentId || this.placed.has(id)) continue;
      this.placed.add(id);
      firstLine.set(id, at);
      balloons.push({ top: at.top, balloon: this.commentBalloon(comment) });
    }
    for (const { top, change } of changes) {
      if (this.placed.has(`rev:${change.id}`)) continue;
      this.placed.add(`rev:${change.id}`);
      balloons.push({ top, balloon: this.changeBalloon(change) });
    }
    // The comments whose text is not drawn, beside the nearest text that is.
    for (const [id, at] of this.undrawn) {
      const comment = this.byId.get(id);
      if (at.page !== page.index || !comment || this.placed.has(id)) continue;
      this.placed.add(id);
      firstLine.set(id, { x: at.x, top: at.top });
      balloons.push({ top: at.top, balloon: this.commentBalloon(comment) });
    }
    balloons.sort((a, b) => a.top - b.top);
    if (last) balloons.push(...this.remaining().map((balloon) => ({ top: 0, balloon })));

    // Each balloon with its leader level with the top of its line, or just
    // below the balloon above, whichever is lower.
    const placed: { top: number; height: number; balloon: Balloon }[] = [];
    let y = GAP;
    for (const { top, balloon } of balloons) {
      const height = balloon.lines.reduce((sum, line) => sum + line.height, 0);
      const at = Math.max(top - balloon.leaderAt, y);
      placed.push({ top: at, height, balloon });
      y = at + height + GAP;
    }

    // The leader turns toward the column at the right edge of the page's
    // text area (past every text item and column on the page).
    let textRight = 0;
    for (const band of page.columnBands) band.colXs.forEach((x, i) => (textRight = Math.max(textRight, x + band.colWidths[i])));
    for (const item of page.items) if (item.kind === "text" && !item.hidden) textRight = Math.max(textRight, item.x + item.width);
    textRight = Math.min(textRight + 6, page.width - 2);
    const barX = page.width + BAR_X;
    const sheet = Math.max(sheetHeight, Math.ceil(y));
    return {
      height: sheet,
      draw: (drawer) => {
        for (const { top, height, x0, x1, resolved } of highlights) {
          drawer.ops.push(`q /${this.resources.alpha(resolved ? 0.22 : 0.38)} gs ${resolved ? "0.627 0.627 0.627" : "1 0.784 0.353"} rg ${drawer.rect(x0, top, x1 - x0, height)} f Q`);
        }
        // The rule between the page and the column, the sheet's full height.
        drawer.ops.push(`q 0.75 w 0.8 G ${num(drawer.x(page.width))} ${num(drawer.y(0))} m ${num(drawer.x(page.width))} ${num(drawer.y(sheet))} l S Q`);
        for (const { top, height, balloon } of placed) {
          drawer.ops.push(`q ${rgb(balloon.color)} rg ${drawer.rect(barX, top, BAR_WIDTH, height)} f Q`);
          let lineY = top;
          for (const line of balloon.lines) {
            line.paint(drawer, page.width, lineY);
            lineY += line.height;
          }
          const from = firstLine.get(balloon.id);
          if (!from) continue;
          const turn = Math.max(textRight, from.x);
          drawer.ops.push(
            `q [${num(1.5)} ${num(1.5)}] 0 d 0.75 w ${rgb(balloon.color)} RG ${num(drawer.x(from.x))} ${num(drawer.y(from.top))} m ` +
              `${num(drawer.x(turn))} ${num(drawer.y(from.top))} l ${num(drawer.x(barX))} ${num(drawer.y(top + balloon.leaderAt))} l S Q`,
          );
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
      out.push(this.commentBalloon(comment));
    }
    return out;
  }

  /** A change's balloon: its author, then what changed ("Deleted: the",
   * "Formatted: Font: Bold"), the label in bold. As in Word, a change with
   * no w:author is by "Unknown" (the parse names it), and one with an
   * empty w:author has no name line. */
  private changeBalloon(change: Revision): Balloon {
    const bold: FontSpec = { family: UI_FONT, size: TEXT_SIZE, bold: true, italic: false };
    const plain: FontSpec = { ...bold, bold: false };
    const pieces: Piece[] = [];
    if (change.deleted) {
      pieces.push({ text: `${change.label}: `, font: bold });
      for (const piece of change.deleted) {
        pieces.push({ text: piece.text, font: { family: piece.props.font ?? UI_FONT, size: TEXT_SIZE, bold: !!piece.props.bold, italic: !!piece.props.italic } });
      }
    } else {
      const colon = change.label.indexOf(": ");
      if (colon < 0) pieces.push({ text: change.label, font: bold });
      else pieces.push({ text: change.label.slice(0, colon + 1), font: bold }, { text: change.label.slice(colon + 1), font: plain });
    }
    const lines: Line[] = [{ height: PAD_Y, paint: () => undefined }];
    const named = this.authors && !!change.author;
    if (named) lines.push(this.nameLine(change.author, NAME_X));
    lines.push(...this.textLines(pieces, TEXT_X, "#000000"));
    lines.push({ height: PAD_Y, paint: () => undefined });
    return { id: `rev:${change.id}`, color: this.doc.revisionColor(change.author), lines, leaderAt: PAD_Y + (named ? LINE_HEIGHT : LINE_HEIGHT / 2) };
  }

  /** A comment's balloon: its author and text, then its thread, each reply
   * under the comment it answers, stepped in by its depth. Without
   * authors, or for an author with no name, no name line. */
  private commentBalloon(comment: DocComment): Balloon {
    const lines: Line[] = [{ height: PAD_Y, paint: () => undefined }];
    const fg = comment.resolved ? "#5f6368" : "#000000";
    const nameOf = (source: DocComment): string => (this.authors ? this.names.get(source.author) ?? source.author : "");
    const push = (indent: number, source: DocComment): void => {
      const name = nameOf(source);
      if (name) lines.push(this.nameLine(name, NAME_X + indent, fg));
      lines.push(...this.textLines([{ text: source.text, font: { family: UI_FONT, size: TEXT_SIZE, bold: false, italic: false } }], TEXT_X + indent, fg));
    };
    push(0, comment);
    const seen = new Set<string>();
    const thread = (parent: DocComment, indent: number): void => {
      for (const reply of this.replies.get(parent.id) ?? []) {
        // A malformed thread may answer itself.
        if (seen.has(reply.id)) continue;
        seen.add(reply.id);
        lines.push({ height: 4, paint: () => undefined });
        push(indent, reply);
        thread(reply, Math.min(indent + 10, 30));
      }
    };
    thread(comment, 10);
    lines.push({ height: PAD_Y, paint: () => undefined });
    // Without authors every comment is one color; with them, its author's, as Word colors a reviewer's comments and changes alike.
    const color = this.doc.revisionColor(this.authors ? comment.author : "");
    return { id: comment.id, color, lines, leaderAt: PAD_Y + (nameOf(comment) ? LINE_HEIGHT : LINE_HEIGHT / 2) };
  }

  /** The author's name in bold, cut short with an ellipsis when it does not fit. */
  private nameLine(name: string, left: number, color = "#000000"): Line {
    const font: FontSpec = { family: UI_FONT, size: TEXT_SIZE, bold: true, italic: false };
    const text = this.clip(name, font, RIGHT - left);
    return { height: LINE_HEIGHT, paint: (drawer, x, y) => this.text(drawer, text, font, color, x + left, y + LINE_HEIGHT - 3) };
  }

  /** `pieces` wrapped to the column from `left`, as lines that paint them. */
  private textLines(pieces: Piece[], left: number, color: string): Line[] {
    return this.wrap(pieces, RIGHT - left).map((line) => ({
      height: LINE_HEIGHT,
      paint: (drawer, x, y) => {
        let at = x + left;
        for (const piece of line) {
          this.text(drawer, piece.text, piece.font, color, at, y + LINE_HEIGHT - 3);
          at += this.measurer.width(piece.text, piece.font);
        }
      },
    }));
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

  /** Greedy word wrap of the pieces to `width` px: a line breaks at a
   * space, at a line break ("\n"), and inside a word longer than the line.
   * A tab prints as a space. Each line is its pieces, a piece per font. */
  private wrap(pieces: Piece[], width: number): Piece[][] {
    const out: Piece[][] = [];
    let line: Piece[] = [];
    let used = 0;
    const add = (text: string, font: FontSpec): void => {
      const last = line.at(-1);
      if (last && last.font === font) last.text += text;
      else line.push({ text, font });
      used += this.measurer.width(text, font);
    };
    const newLine = (): void => {
      const last = line.at(-1);
      if (last) last.text = last.text.trimEnd();
      out.push(line);
      line = [];
      used = 0;
    };
    for (const { text, font } of pieces) {
      for (const word of text.replace(/\t/g, " ").split(/(\n|[^\S\n]+)/)) {
        if (!word) continue;
        if (word === "\n") {
          newLine();
          continue;
        }
        const w = this.measurer.width(word, font);
        if (used + w <= width || (used === 0 && !/\S/.test(word))) {
          add(word, font);
          continue;
        }
        if (!/\S/.test(word)) {
          newLine();
          continue;
        }
        if (used > 0) newLine();
        for (const ch of word) {
          if (used > 0 && used + this.measurer.width(ch, font) > width) newLine();
          add(ch, font);
        }
      }
    }
    newLine();
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

/** Whether Word balloons a change: a deletion whose text left the line
 * (Revision.deleted), a deleted row, a move's destination, a format
 * change. An insertion, and a deleted picture or shape, stay on the line
 * with no balloon. */
function ballooned(change: Revision): boolean {
  if (change.kind === "format" || change.kind === "moveTo") return true;
  if (change.kind === "ins") return false;
  return !!change.deleted || change.id.startsWith("row-");
}

function rgb(css: string): string {
  const c = parseColor(css) ?? { r: 0, g: 0, b: 0, a: 1 };
  return `${num(c.r / 255)} ${num(c.g / 255)} ${num(c.b / 255)}`;
}
