import { Face } from "./sfnt.js";
import { PdfWriter, num } from "./writer.js";

/**
 * A face embedded in the PDF as a Type0 font with Identity-H encoding. A
 * TrueType face is subset to the glyphs the pages used, renumbered in the
 * order they were first used, so a glyph's CID is known the moment it is
 * drawn and the subset is built once at the end. A CFF face goes in whole
 * (no CFF subsetter here) with CID = glyph id.
 */
export class EmbeddedFont {
  /** The glyph each CID stands for; CID 0 is .notdef. */
  private readonly gids: number[] = [0];
  private readonly cidOf = new Map<number, number>([[0, 0]]);
  /** The code point a CID shows, for the ToUnicode map. */
  private readonly unicode = new Map<number, number>();

  constructor(readonly face: Face, readonly name: string, readonly objNum: number, private readonly tag: string) {}

  cid(gid: number, codePoint: number): number {
    if (this.face.cff) {
      if (!this.unicode.has(gid)) this.unicode.set(gid, codePoint);
      return gid;
    }
    let cid = this.cidOf.get(gid);
    if (cid === undefined) {
      cid = this.gids.length;
      this.gids.push(gid);
      this.cidOf.set(gid, cid);
      this.unicode.set(cid, codePoint);
      // Composite glyphs draw other glyphs, which the subset must hold too.
      for (const component of this.face.components(gid)) this.cid(component, 0);
    }
    return cid;
  }

  async write(writer: PdfWriter): Promise<void> {
    const face = this.face;
    const descendant = writer.allocate();
    const descriptor = writer.allocate();
    const file = writer.allocate();
    const toUnicode = writer.allocate();
    const baseFont = `/${this.tag}+${face.postscriptName.replace(/[^!-~]|[/#()<>[\]{}%]/g, "")}`;
    await writer.object(
      this.objNum,
      `<< /Type /Font /Subtype /Type0 /BaseFont ${baseFont} /Encoding /Identity-H /DescendantFonts [${descendant} 0 R] /ToUnicode ${toUnicode} 0 R >>`,
    );
    const scale = 1000 / face.unitsPerEm;
    const widths: string[] = [];
    if (face.cff) {
      for (const gid of [...this.unicode.keys()].sort((a, b) => a - b)) widths.push(`${gid} [${num(face.advance(gid) * scale)}]`);
    } else {
      widths.push(`0 [${this.gids.map((gid) => num(face.advance(gid) * scale)).join(" ")}]`);
    }
    await writer.object(
      descendant,
      `<< /Type /Font /Subtype /${face.cff ? "CIDFontType0" : "CIDFontType2"} /BaseFont ${baseFont} ` +
        `/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${descriptor} 0 R ` +
        `/DW 1000 /W [${widths.join(" ")}]${face.cff ? "" : " /CIDToGIDMap /Identity"} >>`,
    );
    const flags = 4 | (face.italic ? 64 : 0);
    const box = face.bbox.map((v) => num(v * scale)).join(" ");
    await writer.object(
      descriptor,
      `<< /Type /FontDescriptor /FontName ${baseFont} /Flags ${flags} /FontBBox [${box}] /ItalicAngle ${num(face.italicAngle)} ` +
        `/Ascent ${num(face.ascent * scale)} /Descent ${num(face.descent * scale)} /CapHeight ${num(face.capHeight * scale)} /StemV 80 ` +
        `/${face.cff ? "FontFile3" : "FontFile2"} ${file} 0 R >>`,
    );
    if (face.cff) await writer.deflatedStream(file, "/Subtype /OpenType", face.sfnt);
    else {
      const subset = face.subset(this.gids);
      await writer.deflatedStream(file, `/Length1 ${subset.length}`, subset);
    }
    await writer.deflatedStream(toUnicode, "", Buffer.from(toUnicodeCMap(this.unicode), "latin1"));
  }
}

function toUnicodeCMap(unicode: Map<number, number>): string {
  const entries = [...unicode].filter(([, cp]) => cp > 0);
  const lines = [
    "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n",
    "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n",
    "/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n",
    "1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n",
  ];
  for (let at = 0; at < entries.length; at += 100) {
    const block = entries.slice(at, at + 100);
    lines.push(`${block.length} beginbfchar\n`);
    for (const [cid, cp] of block) {
      const utf16 = cp > 0xffff ? String.fromCodePoint(cp) : String.fromCharCode(cp);
      let hex = "";
      for (let i = 0; i < utf16.length; i++) hex += utf16.charCodeAt(i).toString(16).padStart(4, "0");
      lines.push(`<${cid.toString(16).padStart(4, "0")}> <${hex}>\n`);
    }
    lines.push("endbfchar\n");
  }
  lines.push("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
  return lines.join("");
}

/**
 * The fonts, graphics states and images the pages share: named here as
 * they are first used, written as objects at the end (fonts: their glyphs
 * are known only then) or as they come (images). Every page's Resources
 * entry points at the one font and state dictionaries.
 */
export class Resources {
  readonly fontDict: number;
  readonly stateDict: number;
  private readonly fonts = new Map<Face, EmbeddedFont>();
  /** Alpha (ca and CA both) -> graphics state name. */
  private readonly states = new Map<number, string>();

  constructor(private readonly writer: PdfWriter) {
    this.fontDict = writer.allocate();
    this.stateDict = writer.allocate();
  }

  font(face: Face): EmbeddedFont {
    let font = this.fonts.get(face);
    if (!font) {
      const index = this.fonts.size;
      font = new EmbeddedFont(face, `F${index + 1}`, this.writer.allocate(), subsetTag(index));
      this.fonts.set(face, font);
    }
    return font;
  }

  /** The graphics state name for a constant alpha. */
  alpha(value: number): string {
    const key = Math.round(value * 1000);
    let name = this.states.get(key);
    if (!name) {
      name = `GS${this.states.size + 1}`;
      this.states.set(key, name);
    }
    return name;
  }

  /** The Resources dictionary entries every page shares. */
  shared(): string {
    return `/Font ${this.fontDict} 0 R /ExtGState ${this.stateDict} 0 R`;
  }

  async write(): Promise<void> {
    const fonts: string[] = [];
    for (const font of this.fonts.values()) {
      await font.write(this.writer);
      fonts.push(`/${font.name} ${font.objNum} 0 R`);
    }
    await this.writer.object(this.fontDict, `<< ${fonts.join(" ")} >>`);
    const states: string[] = [];
    for (const [key, name] of this.states) states.push(`/${name} << /Type /ExtGState /ca ${num(key / 1000)} /CA ${num(key / 1000)} >>`);
    await this.writer.object(this.stateDict, `<< ${states.join(" ")} >>`);
  }

  /** Families that embedded, for a report. */
  embedded(): string[] {
    return [...this.fonts.keys()].map((face) => `${face.family} ${face.subfamily}`);
  }
}

/** The six-letter tag a subset font's name carries. */
function subsetTag(index: number): string {
  let tag = "";
  let n = index;
  for (let i = 0; i < 6; i++) {
    tag = String.fromCharCode(65 + (n % 26)) + tag;
    n = Math.floor(n / 26);
  }
  return tag;
}
