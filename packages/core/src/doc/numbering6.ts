// Word 6.0/95 lists as numbering.xml. Word 6 has no list tables: each
// numbered paragraph carries its own autonumbering, an ANLD (sprmPAnld)
// and a level (sprmPNLvlAnm), in its PAPX or its style's: 10 numbered,
// 11 bulleted, 12 in the list but skipped (no number), 1-9 a heading
// level numbered as its section's OLST (sprmSOlstAnm) says. In a run of
// autonumbered paragraphs, those with the same level and ANLD are one list
// (a numbered list runs on past the bullets and skipped paragraphs in it),
// and a paragraph without autonumbering ends the run; heading numbers run
// through the document. Each list becomes a num with its own abstractNum.
import { u16 } from "./bytes.js";
import { ParaProps, XmlContext, escapeAttr, icoColor, numFormat } from "./sprm.js";
import type { Font } from "./styles.js";
import { highHalf } from "./word6.js";

/** An ANLD's or an ANLV's number format, text and character properties (the first 16 bytes of both are the ANLV). */
interface Level {
  nfc: number;
  before: string;
  after: string;
  jc: number;
  /** Include the higher levels' numbers (legal numbering). */
  prev: boolean;
  startAt: number;
  rPr: string;
}

const UNDERLINE = ["none", "single", "words", "double", "dotted"];

export class Word6Lists {
  /** The abstractNum (and num) XML of each list, numId 1 first. */
  private readonly lists: string[] = [];
  /** The lists of the run of autonumbered paragraphs being read, by level and ANLD. */
  private readonly open = new Map<string, number>();
  private readonly headingLists = new Map<string, number>();

  constructor(private readonly fonts: Font[], private readonly mac: boolean, private readonly ctx: XmlContext) {}

  /** A new story: a list does not run on into it. */
  startStory(): void {
    this.open.clear();
  }

  /**
   * Number paragraph `p` (set its ilfo and ilvl) from its autonumbering:
   * `nLvlAnm` and `anld` its own or its style's, `olst` its section's.
   */
  number(p: ParaProps, nLvlAnm: number, anld: Uint8Array | undefined, olst: Uint8Array | undefined): void {
    if (nLvlAnm >= 1 && nLvlAnm <= 9 && olst && olst.length >= 9 * 16) {
      const key = hex(olst);
      let numId = this.headingLists.get(key);
      if (numId === undefined) {
        const levels: Level[] = [];
        let textAt = 9 * 16 + 4;
        for (let i = 0; i < 9; i++) {
          const level = this.level(olst.subarray(i * 16, i * 16 + 16), olst.subarray(textAt));
          textAt += olst[i * 16 + 1] + olst[i * 16 + 2];
          levels.push(level);
        }
        numId = this.add(levels, "multilevel");
        this.headingLists.set(key, numId);
      }
      p.ilfo = numId;
      p.ilvl = nLvlAnm - 1;
      this.open.clear();
      return;
    }
    // 12 skips numbering: the paragraph has no number but stays in the lists.
    if (nLvlAnm === 12) return;
    if ((nLvlAnm !== 10 && nLvlAnm !== 11) || !anld || anld.length < 20) {
      this.open.clear();
      return;
    }
    const key = `${nLvlAnm}:${hex(anld)}`;
    let numId = this.open.get(key);
    if (numId === undefined) {
      const level = this.level(anld, anld.subarray(20));
      if (nLvlAnm === 11) level.nfc = 23;
      numId = this.add([level], "singleLevel");
      this.open.set(key, numId);
    }
    p.ilfo = numId;
    p.ilvl = 0;
  }

  /** numbering.xml, or undefined when no paragraph is numbered. */
  xml(): string | undefined {
    if (this.lists.length === 0) return undefined;
    const nums = this.lists.map((_, i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${i}"/></w:num>`);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${this.lists.join("")}${nums.join("")}</w:numbering>`;
  }

  private add(levels: Level[], type: string): number {
    const id = this.lists.length;
    let s = `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="${type}"/>`;
    levels.forEach((level, ilvl) => {
      let number = `%${ilvl + 1}`;
      if (level.prev) for (let k = ilvl; k > 0; k--) number = `%${k}.${number}`;
      const text = level.nfc === 23 ? level.before + level.after : level.before + number + level.after;
      s += `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${level.startAt}"/><w:numFmt w:val="${level.nfc === 23 ? "bullet" : numFormat(level.nfc)}"/>`;
      s += `<w:lvlText w:val="${escapeAttr(text)}"/><w:lvlJc w:val="${["left", "center", "right", "left"][level.jc]}"/>`;
      if (level.rPr) s += `<w:rPr>${level.rPr}</w:rPr>`;
      s += "</w:lvl>";
    });
    this.lists.push(s + "</w:abstractNum>");
    return id + 1;
  }

  /** A level from an ANLV (`v`) and the text it takes from `rgch`: its text before the number, then after it. */
  private level(v: Uint8Array, rgch: Uint8Array): Level {
    const ftc = u16(v, 6);
    const font = ftc !== 0xffff ? this.fonts[ftc] : undefined;
    const chars = (bytes: Uint8Array): string => {
      let s = "";
      for (const b of bytes) {
        if (b === 0) continue;
        if (font?.charset === 2) s += String.fromCharCode(0xf000 + b);
        else s += b < 0x80 ? String.fromCharCode(b) : highHalf(font?.charset ?? 0, this.mac)[b - 0x80];
      }
      return s;
    };
    const cbBefore = v[1];
    const cbAfter = v[2];
    const set = v[3];
    const values = v[4];
    const rPr: string[] = [];
    const fontName = font && this.ctx.fontName(ftc);
    if (fontName) rPr.push(`<w:rFonts w:ascii="${escapeAttr(fontName)}" w:hAnsi="${escapeAttr(fontName)}"/>`);
    if (set & 0x10) rPr.push(values & 0x08 ? "<w:b/>" : '<w:b w:val="0"/>');
    if (set & 0x20) rPr.push(values & 0x10 ? "<w:i/>" : '<w:i w:val="0"/>');
    if (set & 0x80) rPr.push(values & 0x40 ? "<w:caps/>" : '<w:caps w:val="0"/>');
    if (set & 0x40) rPr.push(values & 0x20 ? "<w:smallCaps/>" : '<w:smallCaps w:val="0"/>');
    if (values & 0x01) rPr.push(values & 0x80 ? "<w:strike/>" : '<w:strike w:val="0"/>');
    const ico = v[5] >>> 3;
    if (ico) rPr.push(`<w:color w:val="${icoColor(ico)}"/>`);
    const hps = u16(v, 8);
    if (hps) rPr.push(`<w:sz w:val="${hps}"/>`);
    if (values & 0x02) rPr.push(`<w:u w:val="${UNDERLINE[v[5] & 7] ?? "single"}"/>`);
    return {
      nfc: v[0],
      before: chars(rgch.subarray(0, cbBefore)),
      after: chars(rgch.subarray(cbBefore, cbBefore + cbAfter)),
      jc: set & 3,
      prev: (set & 0x04) !== 0,
      startAt: u16(v, 10),
      rPr: rPr.join(""),
    };
  }
}

function hex(bytes: Uint8Array): string {
  let s = "";
  for (const b of bytes) s += b.toString(16).padStart(2, "0");
  return s;
}
