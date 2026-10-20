// The lists (PlfLst, PlfLfo; [MS-DOC] 2.9.200-201) as numbering.xml: an
// abstractNum per LSTF with its nine (or one) levels, a num per LFO with
// its level overrides. A paragraph's sprmPIlfo is its numId.
import { i32, u16, u32 } from "./bytes.js";
import { FC, Fib, tableBlob } from "./fib.js";
import { XmlContext, applyParaSprms, applyRunSprms, escapeAttr, levelText, newParaProps, newRunProps, numFormat, pPrXml, rPrXml } from "./sprm.js";

interface Level {
  startAt: number;
  nfc: number;
  jc: number;
  legal: boolean;
  noRestart: boolean;
  follow: number;
  restartLim: number;
  papx: Uint8Array;
  chpx: Uint8Array;
  text: string;
}

function readLevel(b: Uint8Array, at: number): { level: Level; next: number } | null {
  if (at + 28 > b.length) return null;
  const flags = b[at + 5];
  const cbChpx = b[at + 24];
  const cbPapx = b[at + 25];
  const papx = b.subarray(at + 28, at + 28 + cbPapx);
  const chpx = b.subarray(at + 28 + cbPapx, at + 28 + cbPapx + cbChpx);
  const xstAt = at + 28 + cbPapx + cbChpx;
  if (xstAt + 2 > b.length) return null;
  const { text, count } = levelText(b, xstAt, b.subarray(at + 6, at + 15));
  const level: Level = {
    startAt: i32(b, at),
    nfc: b[at + 4],
    jc: flags & 3,
    legal: (flags & 4) !== 0,
    noRestart: (flags & 8) !== 0,
    follow: b[at + 15],
    restartLim: b[at + 26],
    papx,
    chpx,
    text,
  };
  return { level, next: xstAt + 2 + count * 2 };
}

function levelXml(level: Level, ilvl: number, ctx: XmlContext, pStyle: string | undefined, tentative: boolean): string {
  let s = `<w:lvl w:ilvl="${ilvl}"${tentative ? ' w:tentative="1"' : ""}>`;
  s += `<w:start w:val="${level.startAt}"/>`;
  s += `<w:numFmt w:val="${numFormat(level.nfc)}"/>`;
  if (level.noRestart) s += `<w:lvlRestart w:val="${level.restartLim}"/>`;
  if (pStyle) s += `<w:pStyle w:val="${escapeAttr(pStyle)}"/>`;
  if (level.legal) s += "<w:isLgl/>";
  if (level.follow !== 0) s += `<w:suff w:val="${level.follow === 1 ? "space" : "nothing"}"/>`;
  s += `<w:lvlText w:val="${escapeAttr(level.text)}"/>`;
  s += `<w:lvlJc w:val="${["left", "center", "right"][level.jc] ?? "left"}"/>`;
  const para = newParaProps(0);
  applyParaSprms(level.papx, para);
  const ppr = pPrXml(para, ctx).before;
  if (ppr) s += `<w:pPr>${ppr}</w:pPr>`;
  const run = newRunProps();
  applyRunSprms(level.chpx, run);
  const rpr = rPrXml(run, ctx, 0);
  if (rpr) s += `<w:rPr>${rpr}</w:rPr>`;
  return s + "</w:lvl>";
}

/** numbering.xml, or undefined for a document with no lists. */
export function numberingXml(fib: Fib, table: Uint8Array, ctx: XmlContext): string | undefined {
  // The LVLs follow the PlfLst in the Table stream, past its lcb.
  const lstAt = fib.fcLcb(FC.plfLst);
  const plfLst = lstAt.lcb && lstAt.fc < table.length ? table.subarray(lstAt.fc) : new Uint8Array(0);
  const plfLfo = tableBlob(fib, table, FC.plfLfo);
  if (plfLst.length < 2 || plfLfo.length < 4) return undefined;
  const cLst = u16(plfLst, 0);
  const lists: { lsid: number; simple: boolean; hybrid: boolean; styles: number[]; levels: Level[] }[] = [];
  let at = 2;
  for (let i = 0; i < cLst && at + 28 <= plfLst.length; i++) {
    const styles: number[] = [];
    for (let k = 0; k < 9; k++) styles.push(u16(plfLst, at + 8 + k * 2));
    const flags = plfLst[at + 26];
    lists.push({ lsid: i32(plfLst, at), simple: (flags & 1) !== 0, hybrid: (flags & 0x10) !== 0, styles, levels: [] });
    at += 28;
  }
  for (const list of lists) {
    for (let k = 0; k < (list.simple ? 1 : 9); k++) {
      const read = readLevel(plfLst, at);
      if (!read) break;
      list.levels.push(read.level);
      at = read.next;
    }
  }
  const out: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'];
  lists.forEach((list, i) => {
    let s = `<w:abstractNum w:abstractNumId="${i}"><w:nsid w:val="${(list.lsid >>> 0).toString(16).toUpperCase().padStart(8, "0")}"/>`;
    s += `<w:multiLevelType w:val="${list.simple ? "singleLevel" : list.hybrid ? "hybridMultilevel" : "multilevel"}"/>`;
    list.levels.forEach((level, ilvl) => {
      const istd = list.styles[ilvl];
      s += levelXml(level, ilvl, ctx, istd !== 0x0fff ? ctx.styleId(istd) : undefined, false);
    });
    out.push(s + "</w:abstractNum>");
  });
  const lfoMac = u32(plfLfo, 0);
  const lfos: { lsid: number; clfolvl: number }[] = [];
  at = 4;
  for (let i = 0; i < lfoMac && at + 16 <= plfLfo.length; i++) {
    lfos.push({ lsid: i32(plfLfo, at), clfolvl: plfLfo[at + 12] });
    at += 16;
  }
  lfos.forEach((lfo, i) => {
    const abstractId = lists.findIndex((list) => list.lsid === lfo.lsid);
    let s = `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${abstractId >= 0 ? abstractId : 0}"/>`;
    // LFOData: cp, then the overrides.
    at += 4;
    for (let k = 0; k < lfo.clfolvl && at + 8 <= plfLfo.length; k++) {
      const startAt = i32(plfLfo, at);
      const flags = plfLfo[at + 4];
      const ilvl = flags & 0x0f;
      const fStartAt = (flags & 0x10) !== 0;
      const fFormatting = (flags & 0x20) !== 0;
      at += 8;
      let o = `<w:lvlOverride w:ilvl="${ilvl}">`;
      if (fStartAt && !fFormatting) o += `<w:startOverride w:val="${startAt}"/>`;
      if (fFormatting) {
        const read = readLevel(plfLfo, at);
        if (read) {
          o += levelXml(read.level, ilvl, ctx, undefined, false);
          at = read.next;
        }
      }
      s += o + "</w:lvlOverride>";
    }
    out.push(s + "</w:num>");
  });
  out.push("</w:numbering>");
  return out.join("");
}
