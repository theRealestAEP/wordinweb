// The stylesheet (STSH, [MS-DOC] 2.9.271) as styles.xml, and the fonts
// table (SttbfFfn) the styles and runs name their fonts in.
import { parseSttb, u16, utf16 } from "./bytes.js";
import { FC, Fib, tableBlob } from "./fib.js";
import { RunProps, ToggleProp, XmlContext, applyParaSprms, applyRunSprms, escapeAttr, newParaProps, newRunProps, pPrXml, rPrXml, tblPrXml } from "./sprm.js";

export interface Font {
  name: string;
  /** The Windows character set: 2 is a symbol font. */
  charset: number;
}

/** The fonts of SttbfFfn, by ftc. */
export function parseFonts(fib: Fib, table: Uint8Array): Font[] {
  const blob = tableBlob(fib, table, FC.sttbfFfn);
  const fonts: Font[] = [];
  if (blob.length < 4) return fonts;
  const count = u16(blob, 0);
  let at = 4;
  for (let i = 0; i < count && at < blob.length; i++) {
    const cch = blob[at];
    const ffn = blob.subarray(at + 1, at + 1 + cch);
    at += 1 + cch;
    // ffid, wWeight, chs, ixchSzAlt, panose (10), fs (24), then the name.
    let name = "";
    for (let k = 39; k + 1 < ffn.length; k += 2) {
      const code = u16(ffn, k);
      if (code === 0) break;
      name += String.fromCharCode(code);
    }
    fonts.push({ name, charset: ffn[3] ?? 0 });
  }
  return fonts;
}

interface Style {
  istd: number;
  sti: number;
  /** 1 paragraph, 2 character, 3 table, 4 numbering. */
  kind: number;
  base: number;
  next: number;
  link: number;
  name: string;
  id: string;
  hidden: boolean;
  semiHidden: boolean;
  unhideWhenUsed: boolean;
  qFormat: boolean;
  priority: number | undefined;
  papx: Uint8Array | null;
  chpx: Uint8Array | null;
  tapx: Uint8Array | null;
}

/** Word's own ids for the built-in styles whose names it localizes, by name. */
const BUILT_IN_IDS: Record<string, string> = {
  "annotation text": "CommentText",
  "annotation reference": "CommentReference",
  "annotation subject": "CommentSubject",
  "Default Paragraph Font": "DefaultParagraphFont",
  "Normal Table": "TableNormal",
  "No List": "NoList",
  "Normal (Web)": "NormalWeb",
};

export class StyleSheet {
  readonly styles: (Style | null)[] = [];
  /** The default fonts' ftcs from the Stshif. */
  readonly defaultFtc: { ascii: number; eastAsia: number; other: number; bi: number | undefined };
  private readonly toggleCache = new Map<string, boolean>();

  constructor(fib: Fib, table: Uint8Array) {
    const stsh = tableBlob(fib, table, FC.stshf);
    const cbStshi = stsh.length >= 2 ? u16(stsh, 0) : 0;
    const stshi = stsh.subarray(2, 2 + cbStshi);
    const cstd = stshi.length >= 2 ? u16(stshi, 0) : 0;
    const cbStdBase = stshi.length >= 4 ? u16(stshi, 2) : 10;
    this.defaultFtc = {
      ascii: stshi.length >= 14 ? u16(stshi, 12) : 0,
      eastAsia: stshi.length >= 16 ? u16(stshi, 14) : 0,
      other: stshi.length >= 18 ? u16(stshi, 16) : 0,
      bi: stshi.length >= 20 ? u16(stshi, 18) : undefined,
    };
    let at = 2 + cbStshi;
    const ids = new Set<string>();
    for (let i = 0; i < cstd && at + 2 <= stsh.length; i++) {
      const cbStd = u16(stsh, at);
      const std = stsh.subarray(at + 2, at + 2 + cbStd);
      at += 2 + cbStd;
      if (cbStd === 0) {
        this.styles.push(null);
        continue;
      }
      this.styles.push(parseStd(std, i, cbStdBase, ids));
    }
  }

  style(istd: number): Style | null {
    return this.styles[istd] ?? null;
  }

  id(istd: number): string | undefined {
    return this.styles[istd]?.id;
  }

  /** Whether the paragraph style, then the character style, leaves toggle `prop` on. */
  toggle(prop: ToggleProp, istdPara: number, istdChar: number | undefined): boolean {
    let on = this.chainToggle(prop, istdPara);
    if (istdChar !== undefined && istdChar !== 10 && istdChar !== istdPara) {
      // A character style's toggle flips the paragraph style's value.
      if (this.chainToggle(prop, istdChar)) on = !on;
    }
    return on;
  }

  private chainToggle(prop: ToggleProp, istd: number): boolean {
    const key = `${prop}:${istd}`;
    const cached = this.toggleCache.get(key);
    if (cached !== undefined) return cached;
    let on = false;
    const chain: Style[] = [];
    let seen = 0;
    for (let i = istd; i !== 0x0fff && seen < 32; seen++) {
      const style = this.styles[i];
      if (!style) break;
      chain.push(style);
      i = style.base;
    }
    for (const style of chain.reverse()) {
      if (!style.chpx) continue;
      const props = newRunProps();
      applyRunSprms(style.chpx, props);
      const v = props.toggles[prop];
      if (v === undefined || v === 0x80) continue;
      on = v === 0x81 ? !on : v !== 0;
    }
    this.toggleCache.set(key, on);
    return on;
  }

  /** styles.xml. */
  xml(ctx: XmlContext, fonts: Font[]): string {
    const out: string[] = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'];
    const font = (ftc: number): string | undefined => fonts[ftc]?.name;
    const defaults: string[] = [];
    const rfonts: string[] = [];
    if (font(this.defaultFtc.ascii)) rfonts.push(`w:ascii="${escapeAttr(font(this.defaultFtc.ascii)!)}"`);
    if (font(this.defaultFtc.eastAsia)) rfonts.push(`w:eastAsia="${escapeAttr(font(this.defaultFtc.eastAsia)!)}"`);
    if (font(this.defaultFtc.other)) rfonts.push(`w:hAnsi="${escapeAttr(font(this.defaultFtc.other)!)}"`);
    if (this.defaultFtc.bi !== undefined && font(this.defaultFtc.bi)) rfonts.push(`w:cs="${escapeAttr(font(this.defaultFtc.bi)!)}"`);
    if (rfonts.length) defaults.push(`<w:rFonts ${rfonts.join(" ")}/>`);
    // Word's defaults before any style: 10 points.
    defaults.push('<w:sz w:val="20"/><w:szCs w:val="20"/>');
    out.push(`<w:docDefaults><w:rPrDefault><w:rPr>${defaults.join("")}</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>`);
    for (const style of this.styles) {
      if (!style) continue;
      const type = ["paragraph", "paragraph", "character", "table", "numbering"][style.kind] ?? "paragraph";
      const isDefault = (style.kind === 1 && style.istd === 0) || (style.kind === 2 && style.istd === 10) || (style.kind === 3 && style.istd === 11) || (style.kind === 4 && style.istd === 12);
      let s = `<w:style w:type="${type}" w:styleId="${escapeAttr(style.id)}"${isDefault ? ' w:default="1"' : ""}>`;
      s += `<w:name w:val="${escapeAttr(style.name)}"/>`;
      if (style.base !== 0x0fff && this.styles[style.base]) s += `<w:basedOn w:val="${escapeAttr(this.styles[style.base]!.id)}"/>`;
      if (style.kind === 1 && style.next !== style.istd && this.styles[style.next]) s += `<w:next w:val="${escapeAttr(this.styles[style.next]!.id)}"/>`;
      if (style.link && this.styles[style.link]) s += `<w:link w:val="${escapeAttr(this.styles[style.link]!.id)}"/>`;
      if (style.priority !== undefined) s += `<w:uiPriority w:val="${style.priority}"/>`;
      if (style.semiHidden) s += "<w:semiHidden/>";
      if (style.unhideWhenUsed) s += "<w:unhideWhenUsed/>";
      if (style.qFormat) s += "<w:qFormat/>";
      if (style.hidden) s += "<w:hidden/>";
      if (style.tapx) {
        const para = newParaProps(style.istd);
        applyParaSprms(style.tapx, para);
        if (para.table) {
          s += `<w:tblPr>${tblPrXml(para.table, ctx, undefined, true)}</w:tblPr>`;
        }
      }
      if (style.papx) {
        const para = newParaProps(style.istd);
        applyParaSprms(style.papx, para);
        para.istd = 0;
        const ppr = pPrXml(para, ctx).before;
        if (ppr) s += `<w:pPr>${ppr}</w:pPr>`;
      }
      if (style.chpx) {
        const run: RunProps = newRunProps();
        applyRunSprms(style.chpx, run);
        const rpr = styleRPrXml(run, ctx);
        if (rpr) s += `<w:rPr>${rpr}</w:rPr>`;
      }
      out.push(s + "</w:style>");
    }
    out.push("</w:styles>");
    return out.join("");
  }
}

/** A style's rPr: its toggles are written as set (on/off), never resolved against another style. */
function styleRPrXml(run: RunProps, ctx: XmlContext): string {
  for (const key of Object.keys(run.toggles) as ToggleProp[]) {
    const v = run.toggles[key];
    if (v === 0x80) delete run.toggles[key];
    else if (v === 0x81) run.toggles[key] = 1;
  }
  const istd = run.istd;
  delete run.istd;
  const xml = rPrXml(run, ctx, 0);
  run.istd = istd;
  return xml;
}

function parseStd(std: Uint8Array, istd: number, cbStdBase: number, ids: Set<string>): Style | null {
  if (std.length < 10) return null;
  const w0 = u16(std, 0);
  const w1 = u16(std, 2);
  const w2 = u16(std, 4);
  const w3 = u16(std, 6);
  const grfstd = u16(std, 8);
  const sti = w0 & 0x0fff;
  const kind = w1 & 0x000f;
  const base = (w1 >>> 4) & 0x0fff;
  const cupx = w2 & 0x000f;
  const next = (w2 >>> 4) & 0x0fff;
  let link = 0;
  let priority: number | undefined;
  if (cbStdBase >= 18 && std.length >= 18) {
    link = u16(std, 10) & 0x0fff;
    priority = u16(std, 16) >>> 4;
  }
  let at = cbStdBase;
  // xstzName: cch, the characters, a null.
  const cch = u16(std, at);
  const fullName = utf16(std, at + 2, cch);
  at += 2 + cch * 2 + 2;
  const name = fullName.split(",")[0] || `Style${istd}`;
  if (at & 1) at++;
  const upx = (): Uint8Array | null => {
    if (at + 2 > std.length) return null;
    const cb = u16(std, at);
    const data = std.subarray(at + 2, at + 2 + cb);
    at += 2 + cb;
    if (at & 1) at++;
    return data;
  };
  let papx: Uint8Array | null = null;
  let chpx: Uint8Array | null = null;
  let tapx: Uint8Array | null = null;
  if (kind === 1) {
    if (cupx >= 1) papx = upx();
    if (cupx >= 2) chpx = upx();
  } else if (kind === 2) {
    if (cupx >= 1) chpx = upx();
  } else if (kind === 3) {
    if (cupx >= 1) tapx = upx();
    if (cupx >= 2) papx = upx();
    if (cupx >= 3) chpx = upx();
  } else if (kind === 4) {
    if (cupx >= 1) papx = upx();
  }
  // A UpxPapx starts with the style's own istd.
  if (papx && papx.length >= 2) papx = papx.subarray(2);
  if (tapx && tapx.length === 0) tapx = null;
  let id = BUILT_IN_IDS[name] ?? styleIdOf(name);
  while (ids.has(id)) id += "0";
  ids.add(id);
  return {
    istd,
    sti,
    kind,
    base,
    next,
    link,
    name,
    id,
    hidden: (grfstd & 0x0002) !== 0,
    semiHidden: (grfstd & 0x0100) !== 0,
    unhideWhenUsed: (grfstd & 0x0800) !== 0,
    qFormat: (grfstd & 0x1000) !== 0,
    priority,
    papx: papx && papx.length ? papx : null,
    chpx: chpx && chpx.length ? chpx : null,
    tapx,
  };
}

/** A style id as Word derives it from the name: each word capitalized, punctuation and spaces dropped. */
function styleIdOf(name: string): string {
  const id = name.replace(/\b\w/g, (c) => c.toUpperCase()).replace(/[^A-Za-z0-9À-￿]/g, "");
  return id || "Style";
}

/** The sttb strings at `index`: a document's author names (SttbfRMark) and the like. */
export function sttbStrings(fib: Fib, table: Uint8Array, index: number): string[] {
  return parseSttb(tableBlob(fib, table, index)).strings;
}
