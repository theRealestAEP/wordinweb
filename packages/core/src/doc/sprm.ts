// The single property modifiers ([MS-DOC] 2.6) decoded into property sets,
// and the property sets written as WordprocessingML: a run's rPr, a
// paragraph's pPr, a table row's tblPr/trPr/tcPr, a section's sectPr.
// Values keep the file's units (twips, half-points, eighths of a point),
// which are the .docx units too.
import { i16, i32, u16, u32, utf16 } from "./bytes.js";
import { sprms } from "./props.js";
import { SPRM_ANLD, SPRM_NLVLANM } from "./word6.js";

export interface Border {
  /** A BrcType ([MS-DOC] 2.9.22); 0 for none, 0xff for "nil" (leave as inherited). */
  type: number;
  /** Eighths of a point. */
  width: number;
  /** "auto" or RRGGBB. */
  color: string;
  /** Points. */
  space: number;
  shadow: boolean;
  frame: boolean;
}

export interface Shading {
  /** An Ipat ([MS-DOC] 2.9.121). */
  pattern: number;
  fore: string;
  fill: string;
}

export interface Tab {
  pos: number;
  /** A TabJC, or "clear" for a cleared tab. */
  jc: number | "clear";
  /** A TabLC. */
  leader: number;
}

/** Revision marking on a run, paragraph, row or section property set (PropRMark). */
export interface PropRevision {
  ibst: number;
  dttm: number;
}

/** Properties that are toggles ([MS-DOC] ToggleOperand): 0 off, 1 on, 0x80 as the style, 0x81 the opposite of the style. */
export type ToggleProp = "b" | "i" | "strike" | "dstrike" | "outline" | "shadow" | "smallCaps" | "caps" | "vanish" | "emboss" | "imprint" | "bCs" | "iCs" | "webHidden" | "noProof" | "specVanish" | "snapToGrid" | "rtl" | "cs" | "ins" | "del" | "fldVanish" | "fSpec";

export interface RunProps {
  istd?: number;
  toggles: Partial<Record<ToggleProp, number>>;
  ftc0?: number;
  ftc1?: number;
  ftc2?: number;
  ftcBi?: number;
  /** Half-points. */
  hps?: number;
  hpsBi?: number;
  /** RRGGBB or "auto". */
  color?: string;
  colorBi?: string;
  underlineColor?: string;
  /** An Ico. */
  highlight?: number;
  /** A Kul. */
  kul?: number;
  /** Twips. */
  dxaSpace?: number;
  /** sprmCFitText: the width in twips the run is fitted to, and the id that joins the runs fitted together. */
  fitText?: { width: number; id: number };
  charScale?: number;
  hpsKern?: number;
  hpsPos?: number;
  /** 0 normal, 1 superscript, 2 subscript. */
  iss?: number;
  lid0?: number;
  lid1?: number;
  lidBi?: number;
  /** Emphasis mark: 0 none, 1 dot, 2 comma, 3 circle, 4 underDot. */
  kcd?: number;
  /** Text effect: 1 lights, 2 blinking background, 3 sparkle, 4 marching black ants, 5 marching red ants, 6 shimmer. */
  sfx?: number;
  border?: Border;
  shading?: Shading;
  /** sprmCSymbol: the font and character of a symbol (the text is U+0028). */
  symbol?: { ftc: number; char: number };
  /** sprmCChs (Word 6 sprmCChse): the character set of pasted 8-bit text, 256 the Macintosh one, when it is not the document's. */
  chse?: number;
  /** sprmCPicLocation: the offset in the Data stream of a picture (U+0001) or an OLE object's name (U+0014). */
  picLocation?: number;
  /** sprmCFData: the U+0001 is binary data, not a picture. */
  fData?: boolean;
  fOle2?: boolean;
  fObj?: boolean;
  /** Who inserted the text and when (sprmCIbstRMark, sprmCDttmRMark), by SttbfRMark index and DTTM. */
  insAuthor?: number;
  insDate?: number;
  delAuthor?: number;
  delDate?: number;
  propRevision?: PropRevision;
}

export interface FrameProps {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  minHeight?: boolean;
  /** PositionCodeOperand: pcVert and pcHorz. */
  pcVert?: number;
  pcHorz?: number;
  wrap?: number;
  dxaFromText?: number;
  dyaFromText?: number;
  locked?: boolean;
  /** DCS: drop cap type and lines. */
  dropCap?: number;
  dropLines?: number;
  textFlow?: number;
}

export interface ParaProps {
  istd: number;
  jc?: number;
  keepLines?: boolean;
  keepNext?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  noLineNumbers?: boolean;
  noAutoHyph?: boolean;
  kinsoku?: boolean;
  wordWrap?: boolean;
  overflowPunct?: boolean;
  topLinePunct?: boolean;
  autoSpaceDE?: boolean;
  autoSpaceDN?: boolean;
  adjustRight?: boolean;
  snapToGrid?: boolean;
  bidi?: boolean;
  contextualSpacing?: boolean;
  mirrorIndents?: boolean;
  noAllowOverlap?: boolean;
  outlineLevel?: number;
  /** 0 top, 1 center, 2 baseline, 3 bottom, 4 auto. */
  textAlignment?: number;
  ilfo?: number;
  ilvl?: number;
  /** Word 6 autonumbering (word6.ts): the ANLD, and the level: 1-9 a heading level, 10 numbered, 11 bulleted, 0 none. */
  anld?: Uint8Array;
  nLvlAnm?: number;
  dxaLeft?: number;
  dxaRight?: number;
  dxaLeft1?: number;
  /** LSPD: line spacing and whether it is a multiple. */
  lineSpacing?: number;
  lineMultiple?: boolean;
  dyaBefore?: number;
  dyaAfter?: number;
  beforeAuto?: boolean;
  afterAuto?: boolean;
  dylBefore?: number;
  dylAfter?: number;
  tabs: Tab[];
  borders: Partial<Record<"top" | "left" | "bottom" | "right" | "between" | "bar", Border>>;
  shading?: Shading;
  frame: FrameProps;
  /** Table nesting: fInTable, itap (depth), the cell and row marks. */
  inTable?: boolean;
  itap?: number;
  dtap?: number;
  ttp?: boolean;
  innerTableCell?: boolean;
  innerTtp?: boolean;
  propRevision?: PropRevision;
  /** The table row's properties, when this paragraph is a row mark. */
  table?: TableProps;
}

export interface CellProps {
  /** The cell's logical right edge, in twips from the row's origin (rgdxaCenter[i+1]). */
  right: number;
  /** TC80 fields. */
  /** 1: the first cell of a horizontally merged range; 2 or 3: merged into the cell before (as Word 97 files hold it). */
  horzMerge: number;
  vertMerge: number;
  vertAlign: number;
  textFlow: number;
  ftsWidth: number;
  width: number;
  fitText: boolean;
  noWrap: boolean;
  hideMark: boolean;
  borders: Partial<Record<"top" | "left" | "bottom" | "right" | "tl2br" | "tr2bl", Border>>;
  shading?: Shading;
  /** Per-side cell margins (sprmTCellPadding), twips. */
  margins: Partial<Record<"top" | "left" | "bottom" | "right", number>>;
  /** sprmTCellWidth: fts and width. */
  prefWidth?: { fts: number; w: number };
}

export interface TableProps {
  istd?: number;
  jc?: number;
  dxaGapHalf?: number;
  dxaLeft?: number;
  /** Signed: negative is an exact height. */
  dyaRowHeight?: number;
  header?: boolean;
  cantSplit?: boolean;
  bidi?: boolean;
  autofit?: boolean;
  /** rgdxaCenter[0]: the logical left edge of the row's first cell. */
  left?: number;
  cells: CellProps[];
  borders: Partial<Record<"top" | "left" | "bottom" | "right" | "insideH" | "insideV", Border>>;
  shading?: Shading;
  /** Twips. */
  cellMargins: Partial<Record<"top" | "left" | "bottom" | "right", number>>;
  cellSpacing?: number;
  tableWidth?: { fts: number; w: number };
  widthBefore?: { fts: number; w: number };
  widthAfter?: { fts: number; w: number };
  indent?: { fts: number; w: number };
  /** Floating table position. */
  pcVert?: number;
  pcHorz?: number;
  dxaAbs?: number;
  dyaAbs?: number;
  dxaFromText?: number;
  dyaFromText?: number;
  dxaFromTextRight?: number;
  dyaFromTextBottom?: number;
  noAllowOverlap?: boolean;
  /** The TLP's grfatl: which table style options are on. */
  tlpFlags?: number;
  propRevision?: PropRevision;
}

export interface SectionProps {
  /** SBkcOperand: 0 continuous, 1 new column, 2 new page, 3 even page, 4 odd page. */
  bkc?: number;
  titlePage?: boolean;
  columns?: number;
  dxaColumns?: number;
  evenlySpaced?: boolean;
  colWidths: number[];
  colSpacings: number[];
  lineBetween?: boolean;
  nfcPgn?: number;
  pgnRestart?: boolean;
  pgnStart?: number;
  /** SLncOperand: 0 per page, 1 restart per section, 2 continuous. */
  lnc?: number;
  nLnnMod?: number;
  dxaLnn?: number;
  lnnMin?: number;
  dyaHdrTop?: number;
  dyaHdrBottom?: number;
  vjc?: number;
  landscape?: boolean;
  xaPage?: number;
  yaPage?: number;
  dxaLeft?: number;
  dxaRight?: number;
  dyaTop?: number;
  dyaBottom?: number;
  dzaGutter?: number;
  bidi?: boolean;
  rtlGutter?: boolean;
  borders: Partial<Record<"top" | "left" | "bottom" | "right", Border>>;
  /** SPgbPropOperand bits. */
  pgbProp?: number;
  dxtCharSpace?: number;
  dyaLinePitch?: number;
  clm?: number;
  /** MSOTXFL ([MS-ODRAW] 2.4.5): the direction text flows on the page. */
  textFlow?: number;
  /** The printer's paper code (sprmSDmPaperReq). */
  paperCode?: number;
  /** SFpcOperand: 0 end of section, 1 bottom of page, 2 beneath text. */
  fpc?: number;
  rncFtn?: number;
  rncEdn?: number;
  nFtn?: number;
  nfcFtnRef?: number;
  nEdn?: number;
  nfcEdnRef?: number;
  endnotesHere?: boolean;
  propRevision?: PropRevision;
  /** Word 6 heading numbering (sprmSOlstAnm): an OLST. */
  olst?: Uint8Array;
  /** The headers and footers the section has (sprmSGprfIhdt), which a Word 6 file stores only. */
  grpfIhdt?: number;
}

export function newRunProps(): RunProps {
  return { toggles: {} };
}

export function newParaProps(istd: number): ParaProps {
  return { istd, tabs: [], borders: {}, frame: {} };
}

export function newTableProps(): TableProps {
  return { cells: [], borders: {}, cellMargins: {} };
}

export function newSectionProps(): SectionProps {
  return { colWidths: [], colSpacings: [], borders: {} };
}

function newCell(right: number): CellProps {
  return { right, horzMerge: 0, vertMerge: 0, vertAlign: 0, textFlow: 0, ftsWidth: 0, width: 0, fitText: false, noWrap: false, hideMark: false, borders: {}, margins: {} };
}

// ---- Colors, borders, shading -------------------------------------------

/** The Ico palette ([MS-DOC] 2.9.119) as RRGGBB; index 0 is auto. */
const ICO: string[] = ["auto", "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF", "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0"];

/** The Ico names of w:highlight. */
const HIGHLIGHT: string[] = ["none", "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white", "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray"];

export function icoColor(ico: number): string {
  return ICO[ico] ?? "auto";
}

/** A COLORREF (red, green, blue, fAuto) at `at`. */
export function colorref(b: Uint8Array, at: number): string {
  if (b[at + 3] === 0xff) return "auto";
  return hex2(b[at]) + hex2(b[at + 1]) + hex2(b[at + 2]);
}

function hex2(n: number): string {
  return n.toString(16).toUpperCase().padStart(2, "0");
}

/** A cell border that is explicitly none (Brc80MayBeNil all ones). */
const NIL_BORDER: Border = { width: 0, type: 0xff, color: "auto", space: 0, shadow: false, frame: false };

/** A Brc80 (4 bytes) at `at`; undefined for the nil value. */
export function brc80(b: Uint8Array, at: number): Border | undefined {
  if (at + 4 > b.length) return undefined;
  if (u32(b, at) === 0xffffffff) return undefined;
  return { width: b[at], type: b[at + 1], color: icoColor(b[at + 2]), space: b[at + 3] & 0x1f, shadow: (b[at + 3] & 0x20) !== 0, frame: (b[at + 3] & 0x40) !== 0 };
}

/** A Brc (8 bytes) at `at`; undefined for the nil value. */
export function brc(b: Uint8Array, at: number): Border | undefined {
  if (at + 8 > b.length) return undefined;
  if (u32(b, at) === 0xffffffff && u32(b, at + 4) === 0xffffffff) return undefined;
  const type = b[at + 5];
  const widthRaw = b[at + 4];
  return { width: type >= 0x40 ? widthRaw * 8 : widthRaw, type, color: colorref(b, at), space: b[at + 6] & 0x1f, shadow: (b[at + 6] & 0x20) !== 0, frame: (b[at + 6] & 0x40) !== 0 };
}

export function shd80(b: Uint8Array, at: number): Shading | undefined {
  const v = u16(b, at);
  if (v === 0xffff) return undefined;
  return { fore: icoColor(v & 0x1f), fill: icoColor((v >>> 5) & 0x1f), pattern: v >>> 10 };
}

export function shd(b: Uint8Array, at: number): Shading | undefined {
  if (at + 10 > b.length) return undefined;
  const pattern = u16(b, at + 8);
  if (pattern === 0xffff) return undefined;
  return { fore: colorref(b, at), fill: colorref(b, at + 4), pattern };
}

const BORDER_NAMES: Record<number, string> = {
  0: "none", 1: "single", 2: "thick", 3: "double", 5: "hairline", 6: "dotted", 7: "dashed", 8: "dotDash", 9: "dotDotDash", 10: "triple",
  11: "thinThickSmallGap", 12: "thickThinSmallGap", 13: "thinThickThinSmallGap", 14: "thinThickMediumGap", 15: "thickThinMediumGap",
  16: "thinThickThinMediumGap", 17: "thinThickLargeGap", 18: "thickThinLargeGap", 19: "thinThickThinLargeGap", 20: "wave", 21: "doubleWave",
  22: "dashSmallGap", 23: "dashDotStroked", 24: "threeDEmboss", 25: "threeDEngrave", 26: "outset", 27: "inset",
};

const ART_BORDERS = "apples archedScallops babyPacifier babyRattle balloons3Colors balloonsHotAir basicBlackDashes basicBlackDots basicBlackSquares basicThinLines basicWhiteDashes basicWhiteDots basicWhiteSquares basicWideInline basicWideMidline basicWideOutline bats birds birdsFlight cabins cakeSlice candyCorn celticKnotwork certificateBanner chainLink champagneBottle checkedBarBlack checkedBarColor checkered christmasTree circlesLines circlesRectangles classicalWave clocks compass confetti confettiGrays confettiOutline confettiStreamers confettiWhite cornerTriangles couponCutoutDashes couponCutoutDots crazyMaze creaturesButterfly creaturesFish creaturesInsects creaturesLadyBug crossStitch cup decoArch decoArchColor decoBlocks diamondsGray doubleD doubleDiamonds earth1 earth2 eclipsingSquares1 eclipsingSquares2 eggsBlack fans film firecrackers flowersBlockPrint flowersDaisies flowersModern1 flowersModern2 flowersPansy flowersRedRose flowersRoses flowersTeacup flowersTiny gems gingerbreadMan gradient handmade1 handmade2 heartBalloon heartGray hearts heebieJeebies holly houseFunky hypnotic iceCreamCones lightBulb lightning1 lightning2 mapPins mapleLeaf mapleMuffins marquee marqueeToothed moons mosaic musicNotes northwest ovals packages palmsBlack palmsColor paperClips papyrus partyFavor partyGlass pencils people peopleWaving peopleHats poinsettias postageStamp pumpkin1 pushPinNote2 pushPinNote1 pyramids pyramidsAbove quadrants rings safari sawtooth sawtoothGray scaredCat seattle shadowedSquares sharksTeeth shorebirdTracks skyrocket snowflakeFancy snowflakes sombrero southwest stars starsTop stars3d starsBlack starsShadowed sun swirligig tornPaper tornPaperBlack trees triangleParty triangles tribal1 tribal2 tribal3 tribal4 tribal5 tribal6 twistedLines1 twistedLines2 vine waveline weavingAngles weavingBraid weavingRibbon weavingStrips whiteFlowers woodwork xIllusions zanyTriangles zigZag zigZagStitch".split(" ");

export function borderName(type: number): string {
  if (type >= 0x40 && type - 0x40 < ART_BORDERS.length) return ART_BORDERS[type - 0x40];
  return BORDER_NAMES[type] ?? "single";
}

const SHADING_NAMES: Record<number, string> = {
  0: "clear", 1: "solid", 2: "pct5", 3: "pct10", 4: "pct20", 5: "pct25", 6: "pct30", 7: "pct40", 8: "pct50", 9: "pct60", 10: "pct70", 11: "pct75",
  12: "pct80", 13: "pct90", 14: "horzStripe", 15: "vertStripe", 16: "reverseDiagStripe", 17: "diagStripe", 18: "horzCross", 19: "diagCross",
  20: "thinHorzStripe", 21: "thinVertStripe", 22: "thinReverseDiagStripe", 23: "thinDiagStripe", 24: "thinHorzCross", 25: "thinDiagCross",
  35: "pct5", 36: "pct10", 37: "pct12", 38: "pct15", 39: "pct20", 40: "pct20", 41: "pct25", 42: "pct30", 43: "pct35", 44: "pct37", 45: "pct40",
  46: "pct45", 47: "pct50", 48: "pct50", 49: "pct55", 50: "pct60", 51: "pct62", 52: "pct65", 53: "pct70", 54: "pct70", 55: "pct75", 56: "pct80",
  57: "pct85", 58: "pct87", 59: "pct90", 60: "pct95", 61: "pct95",
};

export function borderXml(name: string, border: Border | undefined): string {
  // No border, or nil: nothing of its own, so the table's or the style's shows through.
  if (!border || border.type === 0xff || border.type === 0) return "";
  let out = `<w:${name} w:val="${borderName(border.type)}" w:sz="${border.width}" w:space="${border.space}" w:color="${border.color}"`;
  if (border.shadow) out += ' w:shadow="1"';
  if (border.frame) out += ' w:frame="1"';
  return out + "/>";
}

export function shadingXml(shading: Shading | undefined): string {
  if (!shading) return "";
  return `<w:shd w:val="${SHADING_NAMES[shading.pattern] ?? "clear"}" w:color="${shading.fore}" w:fill="${shading.fill}"/>`;
}

// ---- Character properties -----------------------------------------------

const TOGGLE_SPRMS: Record<number, ToggleProp> = {
  0x0800: "del", 0x0801: "ins", 0x0802: "fldVanish", 0x0811: "webHidden", 0x0818: "specVanish", 0x0835: "b", 0x0836: "i", 0x0837: "strike",
  0x0838: "outline", 0x0839: "shadow", 0x083a: "smallCaps", 0x083b: "caps", 0x083c: "vanish", 0x2a53: "dstrike", 0x0854: "imprint",
  0x0855: "fSpec", 0x0858: "emboss", 0x085a: "rtl", 0x085c: "bCs", 0x085d: "iCs", 0x0868: "snapToGrid", 0x0875: "noProof", 0x0882: "cs",
};

/** Apply a grpprl of character sprms to `p`. */
export function applyRunSprms(grpprl: Uint8Array, p: RunProps): void {
  for (const { sprm, operand: o } of sprms(grpprl)) {
    const toggle = TOGGLE_SPRMS[sprm];
    if (toggle) {
      p.toggles[toggle] = o[0];
      continue;
    }
    switch (sprm) {
      case 0x6a03: p.picLocation = i32(o, 0); break;
      case 0x4804: p.insAuthor = u16(o, 0); break;
      case 0x6805: p.insDate = u32(o, 0); break;
      case 0x0806: p.fData = o[0] !== 0; break;
      case 0x6a09: p.symbol = { ftc: u16(o, 0), char: u16(o, 2) }; break;
      case 0xea08: if (o[0]) p.chse = u16(o, 1); else delete p.chse; break;
      case 0x080a: p.fOle2 = o[0] !== 0; break;
      case 0x0856: p.fObj = o[0] !== 0; break;
      case 0x2a0c: p.highlight = o[0]; break;
      case 0x4a30: p.istd = u16(o, 0); break;
      case 0x2a33: {
        // sprmCPlain: back to the paragraph style's character properties.
        const istd = p.istd;
        for (const key of Object.keys(p) as (keyof RunProps)[]) delete p[key];
        p.toggles = {};
        if (istd !== undefined) p.istd = istd;
        break;
      }
      case 0x2a34: p.kcd = o[0]; break;
      case 0x2a3e: p.kul = o[0]; break;
      case 0x8840: p.dxaSpace = i16(o, 0); break;
      case 0xca76: p.fitText = { width: i32(o, 1), id: i32(o, 5) }; break;
      case 0x2a42: p.color = icoColor(o[0]); break;
      case 0x4a43: p.hps = u16(o, 0); break;
      case 0x4845: p.hpsPos = i16(o, 0); break;
      case 0x2a48: p.iss = o[0]; break;
      case 0x484b: p.hpsKern = u16(o, 0); break;
      case 0x4a4f: p.ftc0 = u16(o, 0); break;
      case 0x4a50: p.ftc1 = u16(o, 0); break;
      case 0x4a51: p.ftc2 = u16(o, 0); break;
      case 0x4852: p.charScale = u16(o, 0); break;
      case 0xca57:
      case 0xca89:
        if (o[1]) p.propRevision = { ibst: u16(o, 2), dttm: u32(o, 4) };
        else delete p.propRevision;
        break;
      case 0x2859: p.sfx = o[0]; break;
      case 0x4a5e: p.ftcBi = u16(o, 0); break;
      case 0x485f: p.lidBi = u16(o, 0); break;
      case 0x4a60: p.colorBi = icoColor(o[0]); break;
      case 0x4a61: p.hpsBi = u16(o, 0); break;
      case 0x4863: p.delAuthor = u16(o, 0); break;
      case 0x6864: p.delDate = u32(o, 0); break;
      case 0x6865: p.border = brc80(o, 0); break;
      case 0x4866: p.shading = shd80(o, 0); break;
      case 0x486d: case 0x4873: p.lid0 = u16(o, 0); break;
      case 0x486e: case 0x4874: p.lid1 = u16(o, 0); break;
      case 0x6870: p.color = colorref(o, 0); break;
      case 0xca71: p.shading = shd(o, 1); break;
      case 0xca72: p.border = brc(o, 1); break;
      case 0x6877: p.underlineColor = colorref(o, 0); break;
      default: break;
    }
  }
}

/** The names a document gives the things the properties refer to. */
export interface XmlContext {
  styleId(istd: number): string | undefined;
  fontName(ftc: number): string | undefined;
  /** Whether the style stack (the paragraph style, then the character style when given) turns `prop` on. */
  styleToggle(prop: ToggleProp, istdPara: number, istdChar: number | undefined): boolean;
}

const UNDERLINES: Record<number, string> = {
  1: "single", 2: "words", 3: "double", 4: "dotted", 6: "thick", 7: "dash", 9: "dotDash", 10: "dotDotDash", 11: "wave", 20: "dottedHeavy",
  23: "dashedHeavy", 25: "dashDotHeavy", 26: "dashDotDotHeavy", 27: "wavyHeavy", 39: "dashLong", 43: "wavyDouble", 55: "dashLongHeavy",
};

const EFFECTS: Record<number, string> = { 1: "lights", 2: "blinkBackground", 3: "sparkle", 4: "antsBlack", 5: "antsRed", 6: "shimmer" };
const EMPHASIS: Record<number, string> = { 1: "dot", 2: "comma", 3: "circle", 4: "underDot" };

/** A toggle's effective direct value: undefined when it is the style's. */
export function toggleValue(p: RunProps, prop: ToggleProp, ctx: XmlContext, istdPara: number): boolean | undefined {
  const v = p.toggles[prop];
  if (v === undefined || v === 0x80) return undefined;
  if (v === 0x81) return !ctx.styleToggle(prop, istdPara, p.istd);
  return v !== 0;
}

function onOff(name: string, value: boolean | undefined): string {
  if (value === undefined) return "";
  return value ? `<w:${name}/>` : `<w:${name} w:val="0"/>`;
}

function lang(lid: number | undefined): string | undefined {
  if (lid === undefined || lid === 0x0400 || lid === 0) return undefined;
  return LANGS[lid] ?? `${lid}`;
}

/** The common LIDs as RFC 4646 tags, which Word writes; an unknown LID is written as its number. */
const LANGS: Record<number, string> = {
  0x0409: "en-US", 0x0809: "en-GB", 0x0c09: "en-AU", 0x1009: "en-CA", 0x1409: "en-NZ", 0x1809: "en-IE", 0x1c09: "en-ZA", 0x2009: "en-JM",
  0x2409: "en-029", 0x2809: "en-BZ", 0x2c09: "en-TT", 0x3009: "en-ZW", 0x3409: "en-PH", 0x4009: "en-IN", 0x040c: "fr-FR", 0x080c: "fr-BE",
  0x0c0c: "fr-CA", 0x100c: "fr-CH", 0x140c: "fr-LU", 0x0407: "de-DE", 0x0807: "de-CH", 0x0c07: "de-AT", 0x0410: "it-IT", 0x0810: "it-CH",
  0x040a: "es-ES_tradnl", 0x0c0a: "es-ES", 0x080a: "es-MX", 0x0416: "pt-BR", 0x0816: "pt-PT", 0x0413: "nl-NL", 0x0813: "nl-BE", 0x0406: "da-DK",
  0x041d: "sv-SE", 0x0414: "nb-NO", 0x0814: "nn-NO", 0x040b: "fi-FI", 0x0415: "pl-PL", 0x0419: "ru-RU", 0x0405: "cs-CZ", 0x040e: "hu-HU",
  0x0408: "el-GR", 0x041f: "tr-TR", 0x040d: "he-IL", 0x0401: "ar-SA", 0x0411: "ja-JP", 0x0412: "ko-KR", 0x0804: "zh-CN", 0x0404: "zh-TW",
  0x0c04: "zh-HK", 0x1004: "zh-SG", 0x041e: "th-TH", 0x042a: "vi-VN", 0x0439: "hi-IN", 0x0418: "ro-RO", 0x0402: "bg-BG", 0x041a: "hr-HR",
  0x0424: "sl-SI", 0x041b: "sk-SK", 0x0425: "et-EE", 0x0426: "lv-LV", 0x0427: "lt-LT", 0x0422: "uk-UA", 0x0421: "id-ID", 0x043e: "ms-MY",
  0x0403: "ca-ES", 0x0456: "gl-ES", 0x042d: "eu-ES", 0x043a: "mt-MT", 0x083c: "ga-IE", 0x0452: "cy-GB", 0x041c: "sq-AL", 0x042f: "mk-MK",
  0x0c1a: "sr-Cyrl-CS", 0x081a: "sr-Latn-CS", 0x0437: "ka-GE", 0x042b: "hy-AM", 0x042c: "az-Latn-AZ", 0x0423: "be-BY", 0x043f: "kk-KZ",
  0x0443: "uz-Latn-UZ", 0x0429: "fa-IR", 0x0420: "ur-PK", 0x0445: "bn-IN", 0x0449: "ta-IN", 0x044a: "te-IN", 0x044b: "kn-IN", 0x044e: "mr-IN",
  0x044c: "ml-IN", 0x0447: "gu-IN", 0x0446: "pa-IN", 0x0436: "af-ZA", 0x0441: "sw-KE", 0x043b: "se-NO", 0x040f: "is-IS", 0x0438: "fo-FO",
};

/**
 * A run's rPr from its direct properties. `istdPara` resolves the toggles
 * set to the style's opposite. `paragraphMark` leaves out what a mark
 * cannot carry (nothing yet, kept for the caller's intent).
 */
export function rPrXml(p: RunProps, ctx: XmlContext, istdPara: number): string {
  const out: string[] = [];
  if (p.istd !== undefined && p.istd !== 10) {
    const id = ctx.styleId(p.istd);
    if (id) out.push(`<w:rStyle w:val="${escapeAttr(id)}"/>`);
  }
  if (p.ftc0 !== undefined || p.ftc1 !== undefined || p.ftc2 !== undefined || p.ftcBi !== undefined) {
    let fonts = "<w:rFonts";
    const name = (ftc: number | undefined, attr: string): void => {
      if (ftc === undefined) return;
      const font = ctx.fontName(ftc);
      if (font) fonts += ` w:${attr}="${escapeAttr(font)}"`;
    };
    name(p.ftc0, "ascii");
    name(p.ftc1, "eastAsia");
    name(p.ftc2, "hAnsi");
    name(p.ftcBi, "cs");
    if (fonts.length > 8) out.push(fonts + "/>");
  }
  const t = (prop: ToggleProp): boolean | undefined => toggleValue(p, prop, ctx, istdPara);
  out.push(onOff("b", t("b")), onOff("bCs", t("bCs")), onOff("i", t("i")), onOff("iCs", t("iCs")), onOff("caps", t("caps")), onOff("smallCaps", t("smallCaps")));
  out.push(onOff("strike", t("strike")), onOff("dstrike", t("dstrike")), onOff("outline", t("outline")), onOff("shadow", t("shadow")), onOff("emboss", t("emboss")), onOff("imprint", t("imprint")));
  out.push(onOff("noProof", t("noProof")), onOff("snapToGrid", t("snapToGrid")), onOff("vanish", t("vanish")), onOff("webHidden", t("webHidden")));
  if (p.color !== undefined) out.push(`<w:color w:val="${p.color}"/>`);
  if (p.dxaSpace !== undefined) out.push(`<w:spacing w:val="${p.dxaSpace}"/>`);
  if (p.charScale !== undefined && p.charScale !== 100) out.push(`<w:w w:val="${p.charScale}"/>`);
  if (p.hpsKern !== undefined) out.push(`<w:kern w:val="${p.hpsKern}"/>`);
  if (p.hpsPos !== undefined) out.push(`<w:position w:val="${p.hpsPos}"/>`);
  if (p.hps !== undefined) out.push(`<w:sz w:val="${p.hps}"/>`);
  if (p.hpsBi !== undefined) out.push(`<w:szCs w:val="${p.hpsBi}"/>`);
  if (p.highlight !== undefined && HIGHLIGHT[p.highlight]) out.push(`<w:highlight w:val="${HIGHLIGHT[p.highlight]}"/>`);
  if (p.kul !== undefined) {
    const val = p.kul === 0 ? "none" : UNDERLINES[p.kul] ?? "single";
    out.push(`<w:u w:val="${val}"${p.underlineColor && p.underlineColor !== "auto" ? ` w:color="${p.underlineColor}"` : ""}/>`);
  }
  if (p.sfx !== undefined && EFFECTS[p.sfx]) out.push(`<w:effect w:val="${EFFECTS[p.sfx]}"/>`);
  if (p.border) out.push(borderXml("bdr", p.border));
  if (p.shading) out.push(shadingXml(p.shading));
  if (p.fitText && p.fitText.width > 0) out.push(`<w:fitText w:val="${p.fitText.width}" w:id="${p.fitText.id}"/>`);
  if (p.iss !== undefined) out.push(`<w:vertAlign w:val="${p.iss === 1 ? "superscript" : p.iss === 2 ? "subscript" : "baseline"}"/>`);
  out.push(onOff("rtl", t("rtl")), onOff("cs", t("cs")));
  if (p.kcd !== undefined && EMPHASIS[p.kcd]) out.push(`<w:em w:val="${EMPHASIS[p.kcd]}"/>`);
  const l0 = lang(p.lid0);
  const l1 = lang(p.lid1);
  const lb = lang(p.lidBi);
  if (l0 || l1 || lb) out.push(`<w:lang${l0 ? ` w:val="${l0}"` : ""}${l1 ? ` w:eastAsia="${l1}"` : ""}${lb ? ` w:bidi="${lb}"` : ""}/>`);
  out.push(onOff("specVanish", t("specVanish")));
  return out.join("");
}

export function escapeAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
}

export function escapeText(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// ---- Paragraph and table properties --------------------------------------

/** Apply a grpprl of paragraph sprms to `p`; its table sprms go to `p.table` when the paragraph is a row mark. */
export function applyParaSprms(grpprl: Uint8Array, p: ParaProps): void {
  for (const { sprm, operand: o } of sprms(grpprl)) {
    if ((sprm & 0x1c00) === 0x1400) {
      // sgc 5: a table property, kept on the row mark only.
      if (!p.table) p.table = newTableProps();
      applyTableSprm(sprm, o, p.table);
      continue;
    }
    switch (sprm) {
      case 0x4600: p.istd = u16(o, 0); break;
      case 0x2403: case 0x2461: p.jc = o[0]; break;
      case 0x2405: p.keepLines = o[0] !== 0; break;
      case 0x2406: p.keepNext = o[0] !== 0; break;
      case 0x2407: p.pageBreakBefore = o[0] !== 0; break;
      case 0x260a: p.ilvl = o[0]; break;
      case SPRM_ANLD: p.anld = o.subarray(1); break;
      case SPRM_NLVLANM: p.nLvlAnm = o[0]; break;
      case 0x460b: {
        const v = u16(o, 0);
        p.ilfo = v === 0xf801 ? 0 : v >= 0xf802 ? 0x10000 - v : v;
        break;
      }
      case 0x240c: p.noLineNumbers = o[0] !== 0; break;
      case 0xc60d: applyChgTabsPapx(o, p); break;
      case 0x840e: case 0x845d: p.dxaRight = i16(o, 0); break;
      case 0x840f: case 0x845e: p.dxaLeft = i16(o, 0); break;
      case 0x4610: case 0x465f: p.dxaLeft = (p.dxaLeft ?? 0) + i16(o, 0); break;
      case 0x8411: case 0x8460: p.dxaLeft1 = i16(o, 0); break;
      case 0x6412: p.lineSpacing = i16(o, 0); p.lineMultiple = u16(o, 2) !== 0; break;
      case 0xa413: p.dyaBefore = u16(o, 0); break;
      case 0xa414: p.dyaAfter = u16(o, 0); break;
      case 0xc615: applyChgTabs(o, p); break;
      case 0x2416: p.inTable = o[0] !== 0; break;
      case 0x2417: p.ttp = o[0] !== 0; break;
      case 0x8418: p.frame.x = u16(o, 0); break;
      case 0x8419: p.frame.y = u16(o, 0); break;
      case 0x841a: p.frame.width = u16(o, 0); break;
      case 0x261b: p.frame.pcVert = (o[0] >>> 4) & 3; p.frame.pcHorz = (o[0] >>> 6) & 3; break;
      case 0x2423: p.frame.wrap = o[0]; break;
      case 0x6424: p.borders.top = brc80(o, 0); break;
      case 0x6425: p.borders.left = brc80(o, 0); break;
      case 0x6426: p.borders.bottom = brc80(o, 0); break;
      case 0x6427: p.borders.right = brc80(o, 0); break;
      case 0x6428: p.borders.between = brc80(o, 0); break;
      case 0x6629: p.borders.bar = brc80(o, 0); break;
      case 0x242a: p.noAutoHyph = o[0] !== 0; break;
      case 0x442b: { const v = u16(o, 0); p.frame.height = v & 0x7fff; p.frame.minHeight = (v & 0x8000) !== 0; break; }
      case 0x442c: { const v = u16(o, 0); p.frame.dropCap = v & 7; p.frame.dropLines = (v >>> 3) & 0x1f; break; }
      case 0x442d: p.shading = shd80(o, 0); break;
      case 0x842e: p.frame.dyaFromText = u16(o, 0); break;
      case 0x842f: p.frame.dxaFromText = u16(o, 0); break;
      case 0x2430: p.frame.locked = o[0] !== 0; break;
      case 0x2431: p.widowControl = o[0] !== 0; break;
      case 0x2433: p.kinsoku = o[0] !== 0; break;
      case 0x2434: p.wordWrap = o[0] !== 0; break;
      case 0x2435: p.overflowPunct = o[0] !== 0; break;
      case 0x2436: p.topLinePunct = o[0] !== 0; break;
      case 0x2437: p.autoSpaceDE = o[0] !== 0; break;
      case 0x2438: p.autoSpaceDN = o[0] !== 0; break;
      case 0x4439: p.textAlignment = u16(o, 0); break;
      case 0x443a: p.frame.textFlow = u16(o, 0) & 7; break;
      case 0x2640: p.outlineLevel = o[0]; break;
      case 0x2441: p.bidi = o[0] !== 0; break;
      case 0x2447: p.snapToGrid = o[0] !== 0; break;
      case 0x2448: p.adjustRight = o[0] !== 0; break;
      case 0x6649: p.itap = i32(o, 0); break;
      case 0x664a: p.dtap = (p.dtap ?? 0) + i32(o, 0); break;
      case 0x244b: p.innerTableCell = o[0] !== 0; break;
      case 0x244c: p.innerTtp = o[0] !== 0; break;
      case 0xc64d: p.shading = shd(o, 1); break;
      case 0xc64e: p.borders.top = brc(o, 1); break;
      case 0xc64f: p.borders.left = brc(o, 1); break;
      case 0xc650: p.borders.bottom = brc(o, 1); break;
      case 0xc651: p.borders.right = brc(o, 1); break;
      case 0xc652: p.borders.between = brc(o, 1); break;
      case 0xc653: p.borders.bar = brc(o, 1); break;
      case 0x4458: p.dylBefore = i16(o, 0); break;
      case 0x4459: p.dylAfter = i16(o, 0); break;
      case 0x245b: p.beforeAuto = o[0] !== 0; break;
      case 0x245c: p.afterAuto = o[0] !== 0; break;
      case 0x2462: p.noAllowOverlap = o[0] !== 0; break;
      case 0x246d: p.contextualSpacing = o[0] !== 0; break;
      case 0xc66f:
        if (o[1]) p.propRevision = { ibst: u16(o, 2), dttm: u32(o, 4) };
        else delete p.propRevision;
        break;
      case 0x2470: p.mirrorIndents = o[0] !== 0; break;
      default: break;
    }
  }
}

function applyChgTabsPapx(o: Uint8Array, p: ParaProps): void {
  let at = 1;
  const cDel = o[at++];
  for (let i = 0; i < cDel; i++) setTab(p, { pos: i16(o, at + i * 2), jc: "clear", leader: 0 });
  at += cDel * 2;
  const cAdd = o[at++];
  for (let i = 0; i < cAdd; i++) {
    const tbd = o[at + cAdd * 2 + i];
    setTab(p, { pos: i16(o, at + i * 2), jc: tbd & 7, leader: (tbd >>> 3) & 7 });
  }
}

function applyChgTabs(o: Uint8Array, p: ParaProps): void {
  let at = 1;
  const cDel = o[at++];
  for (let i = 0; i < cDel; i++) setTab(p, { pos: i16(o, at + i * 2), jc: "clear", leader: 0 });
  at += cDel * 4;
  const cAdd = o[at++];
  for (let i = 0; i < cAdd; i++) {
    const tbd = o[at + cAdd * 2 + i];
    setTab(p, { pos: i16(o, at + i * 2), jc: tbd & 7, leader: (tbd >>> 3) & 7 });
  }
}

function setTab(p: ParaProps, tab: Tab): void {
  const i = p.tabs.findIndex((t) => t.pos === tab.pos);
  if (i >= 0) p.tabs[i] = tab;
  else p.tabs.push(tab);
}

function cellRange(o: Uint8Array, at: number, cells: CellProps[]): CellProps[] {
  const first = o[at];
  const lim = o[at + 1];
  while (cells.length < lim) cells.push(newCell(cells.length ? cells[cells.length - 1].right : 0));
  return cells.slice(first, lim);
}

function ftsWidth(o: Uint8Array, at: number): { fts: number; w: number } {
  return { fts: o[at], w: i16(o, at + 1) };
}

function applyTableSprm(sprm: number, o: Uint8Array, t: TableProps): void {
  switch (sprm) {
    case 0x5400: case 0x548a: t.jc = u16(o, 0); break;
    case 0x9601: t.dxaLeft = i16(o, 0); break;
    case 0x9602: t.dxaGapHalf = i16(o, 0); break;
    case 0x3403: case 0x3466: t.cantSplit = o[0] !== 0; break;
    case 0x3404: t.header = o[0] !== 0; break;
    case 0xd605: {
      const sides = ["top", "left", "bottom", "right", "insideH", "insideV"] as const;
      sides.forEach((side, i) => { t.borders[side] = brc80(o, 1 + i * 4); });
      break;
    }
    case 0x9407: t.dyaRowHeight = i16(o, 0); break;
    case 0xd608: applyDefTable(o, t); break;
    case 0xd609: {
      const n = Math.floor(o[0] / 2);
      for (let i = 0; i < n && i < t.cells.length; i++) t.cells[i].shading = shd80(o, 1 + i * 2);
      break;
    }
    case 0x740a: t.tlpFlags = u16(o, 2); break;
    case 0x560b: case 0x5664: t.bidi = u16(o, 0) !== 0; break;
    case 0xd612: case 0xd616: case 0xd60c: case 0xd670: case 0xd671: case 0xd672: {
      const base = sprm === 0xd612 || sprm === 0xd670 ? 0 : sprm === 0xd616 || sprm === 0xd671 ? 22 : 44;
      const n = Math.floor(o[0] / 10);
      for (let i = 0; i < n && base + i < t.cells.length; i++) t.cells[base + i].shading = shd(o, 1 + i * 10);
      break;
    }
    case 0x360d: t.pcVert = (o[0] >>> 4) & 3; t.pcHorz = (o[0] >>> 6) & 3; break;
    case 0x940e: t.dxaAbs = u16(o, 0); break;
    case 0x940f: t.dyaAbs = u16(o, 0); break;
    case 0x9410: t.dxaFromText = u16(o, 0); break;
    case 0x9411: t.dyaFromText = u16(o, 0); break;
    case 0x941e: t.dxaFromTextRight = u16(o, 0); break;
    case 0x941f: t.dyaFromTextBottom = u16(o, 0); break;
    case 0xd613: {
      const sides = ["top", "left", "bottom", "right", "insideH", "insideV"] as const;
      sides.forEach((side, i) => { t.borders[side] = brc(o, 1 + i * 8); });
      break;
    }
    case 0xf614: t.tableWidth = ftsWidth(o, 0); break;
    case 0x3615: t.autofit = o[0] !== 0; break;
    case 0xf617: t.widthBefore = ftsWidth(o, 0); break;
    case 0xf618: t.widthAfter = ftsWidth(o, 0); break;
    case 0xd61a: case 0xd61b: case 0xd61c: case 0xd61d: {
      // sprmTBrc*Cv: the full color of each cell's border, which its TC80 holds only as an Ico.
      const side = (["top", "left", "bottom", "right"] as const)[sprm - 0xd61a];
      for (let i = 0; i < Math.floor(o[0] / 4) && i < t.cells.length; i++) {
        const border = t.cells[i].borders[side];
        const color = colorref(o, 1 + i * 4);
        if (border && color !== "auto") t.cells[i].borders[side] = { ...border, color };
      }
      break;
    }
    case 0xd620: case 0xd62f: {
      const wide = sprm === 0xd62f;
      const border = wide ? brc(o, 4) : brc80(o, 4);
      const which = o[3];
      for (const cell of cellRange(o, 1, t.cells)) {
        if (which & 0x01) cell.borders.top = border;
        if (which & 0x02) cell.borders.left = border;
        if (which & 0x04) cell.borders.bottom = border;
        if (which & 0x08) cell.borders.right = border;
        if (wide && which & 0x10) cell.borders.tl2br = border;
        if (wide && which & 0x20) cell.borders.tr2bl = border;
      }
      break;
    }
    case 0x7621: {
      // sprmTInsert: new cells, only when they extend the row (a TDefTable usually came first).
      const first = o[0];
      const count = o[1];
      const width = i16(o, 2);
      if (first >= t.cells.length) {
        for (let i = 0; i < count; i++) {
          const right = (t.cells.length ? t.cells[t.cells.length - 1].right : t.left ?? 0) + width;
          const cell = newCell(right);
          cell.ftsWidth = 3;
          cell.width = width;
          t.cells.push(cell);
        }
      }
      break;
    }
    case 0x7623: {
      // sprmTDxaCol: each cell in [itcFirst, itcLim) becomes dxaCol wide, and the cell edges after it move with it.
      const width = i16(o, 2);
      cellRange(o, 0, t.cells);
      for (let i = o[0]; i < o[1]; i++) {
        const left = i === 0 ? t.left ?? 0 : t.cells[i - 1].right;
        const delta = width - (t.cells[i].right - left);
        for (let j = i; j < t.cells.length; j++) t.cells[j].right += delta;
        t.cells[i].width = width;
      }
      break;
    }
    case 0x5624: {
      const cells = cellRange(o, 0, t.cells);
      cells.forEach((cell, i) => { cell.horzMerge = i === 0 ? 1 : 2; });
      break;
    }
    case 0x5625: for (const cell of cellRange(o, 0, t.cells)) cell.horzMerge = 0; break;
    case 0x7629: { const tf = u16(o, 2); for (const cell of cellRange(o, 0, t.cells)) cell.textFlow = tf; break; }
    case 0xd62b: { const cell = t.cells[o[1]]; if (cell) cell.vertMerge = o[2]; break; }
    case 0xd62c: { const va = o[3]; for (const cell of cellRange(o, 1, t.cells)) cell.vertAlign = va; break; }
    case 0xd62d: { const s = shd(o, 3); for (const cell of cellRange(o, 1, t.cells)) cell.shading = s; break; }
    case 0xd62e: { const s = shd(o, 3); cellRange(o, 1, t.cells).forEach((cell, i) => { if (i % 2 === 0) cell.shading = s; }); break; }
    case 0xd632: case 0xd634: case 0xd63e: {
      const sides = o[3];
      const fts = o[4];
      const w = fts === 3 ? u16(o, 5) : 0;
      const targets = sprm === 0xd632 ? cellRange(o, 1, t.cells).map((c) => c.margins) : [t.cellMargins];
      for (const m of targets) {
        if (sides & 0x01) m.top = w;
        if (sides & 0x02) m.left = w;
        if (sides & 0x04) m.bottom = w;
        if (sides & 0x08) m.right = w;
      }
      break;
    }
    case 0xd633: if (o[4] === 3 || o[4] === 0x13) t.cellSpacing = u16(o, 5); break;
    case 0xd635: { const w = ftsWidth(o, 3); for (const cell of cellRange(o, 1, t.cells)) cell.prefWidth = w; break; }
    case 0xf636: { const on = o[2] !== 0; for (const cell of cellRange(o, 0, t.cells)) cell.fitText = on; break; }
    case 0xd639: { const on = o[3] !== 0; for (const cell of cellRange(o, 1, t.cells)) cell.noWrap = on; break; }
    case 0x563a: t.istd = u16(o, 0); break;
    case 0xd642: { const on = o[3] !== 0; for (const cell of cellRange(o, 1, t.cells)) cell.hideMark = on; break; }
    case 0xd660: t.shading = shd(o, 1); break;
    case 0xf661: t.indent = ftsWidth(o, 0); break;
    case 0xd662: {
      const n = Math.floor(o[0] / 4);
      for (let i = 0; i < n && i < t.cells.length; i++) {
        const b = t.cells[i].borders;
        const at = 1 + i * 4;
        const set = (side: "top" | "left" | "bottom" | "right", type: number): void => {
          const current = b[side];
          if (type === 0xff) return;
          b[side] = current ? { ...current, type } : { type, width: 4, color: "auto", space: 0, shadow: false, frame: false };
        };
        set("top", o[at]);
        set("left", o[at + 1]);
        set("bottom", o[at + 2]);
        set("right", o[at + 3]);
      }
      break;
    }
    case 0x3465: t.noAllowOverlap = o[0] !== 0; break;
    case 0xd667:
      if (o[1]) t.propRevision = { ibst: u16(o, 2), dttm: u32(o, 4) };
      else delete t.propRevision;
      break;
    default: break;
  }
}

function applyDefTable(o: Uint8Array, t: TableProps): void {
  const count = o[2];
  const centersAt = 3;
  const tcsAt = centersAt + (count + 1) * 2;
  t.left = i16(o, centersAt);
  t.cells = [];
  for (let i = 0; i < count; i++) {
    const cell = newCell(i16(o, centersAt + (i + 1) * 2));
    const at = tcsAt + i * 20;
    if (at + 20 <= o.length) {
      const grf = u16(o, at);
      cell.horzMerge = grf & 3;
      cell.textFlow = (grf >>> 2) & 7;
      cell.vertMerge = (grf >>> 5) & 3;
      cell.vertAlign = (grf >>> 7) & 3;
      cell.ftsWidth = (grf >>> 9) & 7;
      cell.fitText = (grf & 0x1000) !== 0;
      cell.noWrap = (grf & 0x2000) !== 0;
      cell.hideMark = (grf & 0x4000) !== 0;
      cell.width = u16(o, at + 2);
      // A TC80's borders are Brc80MayBeNil: all ones is nil, which hides the table's border there.
      const cellBorder = (at: number): Border | undefined => (u32(o, at) === 0xffffffff ? NIL_BORDER : brc80(o, at));
      cell.borders.top = cellBorder(at + 4);
      cell.borders.left = cellBorder(at + 8);
      cell.borders.bottom = cellBorder(at + 12);
      cell.borders.right = cellBorder(at + 16);
    }
    t.cells.push(cell);
  }
}

const JC: Record<number, string> = { 0: "left", 1: "center", 2: "right", 3: "both", 4: "distribute", 5: "mediumKashida", 7: "highKashida", 8: "lowKashida", 9: "thaiDistribute" };
const TAB_JC: Record<number, string> = { 0: "left", 1: "center", 2: "right", 3: "decimal", 4: "bar", 6: "num" };
const TAB_LEADER: Record<number, string> = { 0: "none", 1: "dot", 2: "hyphen", 3: "underscore", 4: "heavy", 5: "middleDot" };
const TEXT_ALIGN: Record<number, string> = { 0: "top", 1: "center", 2: "baseline", 3: "bottom", 4: "auto" };
const WRAP: Record<number, string> = { 0: "auto", 1: "notBeside", 2: "around", 3: "none", 4: "tight", 5: "through" };
const X_ALIGN: Record<number, string> = { 0xfffc: "center", 0xfff8: "right", 0xfff4: "inside", 0xfff0: "outside" };
// [MS-DOC] 2.6.2 sprmPDyaAbs: 0 is "inline" (the frame keeps its place in the text), not an offset of 0.
const Y_ALIGN: Record<number, string> = { 0: "inline", 0xfffc: "top", 0xfff8: "center", 0xfff4: "bottom", 0xfff0: "inside", 0xffec: "outside" };

/** A paragraph's pPr body (without the rPr of its mark or a sectPr, which the caller adds in their places). */
export function pPrXml(p: ParaProps, ctx: XmlContext): { before: string; after: string } {
  const out: string[] = [];
  if (p.istd !== 0) {
    const id = ctx.styleId(p.istd);
    if (id) out.push(`<w:pStyle w:val="${escapeAttr(id)}"/>`);
  }
  out.push(onOff("keepNext", p.keepNext), onOff("keepLines", p.keepLines), onOff("pageBreakBefore", p.pageBreakBefore));
  out.push(framePrXml(p.frame));
  out.push(onOff("widowControl", p.widowControl));
  if (p.ilfo !== undefined || p.ilvl !== undefined) {
    out.push(`<w:numPr>${p.ilvl !== undefined ? `<w:ilvl w:val="${p.ilvl}"/>` : ""}${p.ilfo !== undefined ? `<w:numId w:val="${p.ilfo}"/>` : ""}</w:numPr>`);
  }
  out.push(onOff("suppressLineNumbers", p.noLineNumbers));
  const borders = (["top", "left", "bottom", "right", "between", "bar"] as const).map((side) => borderXml(side, p.borders[side])).join("");
  if (borders) out.push(`<w:pBdr>${borders}</w:pBdr>`);
  out.push(shadingXml(p.shading));
  if (p.tabs.length) {
    const tabs = [...p.tabs].sort((a, b) => a.pos - b.pos).map((t) => `<w:tab w:val="${t.jc === "clear" ? "clear" : TAB_JC[t.jc] ?? "left"}"${t.jc !== "clear" && t.leader ? ` w:leader="${TAB_LEADER[t.leader] ?? "none"}"` : ""} w:pos="${t.pos}"/>`);
    out.push(`<w:tabs>${tabs.join("")}</w:tabs>`);
  }
  out.push(onOff("suppressAutoHyphens", p.noAutoHyph), onOff("kinsoku", p.kinsoku), onOff("wordWrap", p.wordWrap), onOff("overflowPunct", p.overflowPunct));
  out.push(onOff("topLinePunct", p.topLinePunct), onOff("autoSpaceDE", p.autoSpaceDE), onOff("autoSpaceDN", p.autoSpaceDN), onOff("bidi", p.bidi));
  out.push(onOff("adjustRightInd", p.adjustRight), onOff("snapToGrid", p.snapToGrid));
  const spacing: string[] = [];
  if (p.dyaBefore !== undefined) spacing.push(`w:before="${p.dyaBefore}"`);
  if (p.dylBefore !== undefined) spacing.push(`w:beforeLines="${p.dylBefore}"`);
  if (p.beforeAuto !== undefined) spacing.push(`w:beforeAutospacing="${p.beforeAuto ? 1 : 0}"`);
  if (p.dyaAfter !== undefined) spacing.push(`w:after="${p.dyaAfter}"`);
  if (p.dylAfter !== undefined) spacing.push(`w:afterLines="${p.dylAfter}"`);
  if (p.afterAuto !== undefined) spacing.push(`w:afterAutospacing="${p.afterAuto ? 1 : 0}"`);
  if (p.lineSpacing !== undefined) {
    if (p.lineMultiple) spacing.push(`w:line="${p.lineSpacing}" w:lineRule="auto"`);
    else if (p.lineSpacing < 0) spacing.push(`w:line="${-p.lineSpacing}" w:lineRule="exact"`);
    else spacing.push(`w:line="${p.lineSpacing}" w:lineRule="atLeast"`);
  }
  if (spacing.length) out.push(`<w:spacing ${spacing.join(" ")}/>`);
  const ind: string[] = [];
  if (p.dxaLeft !== undefined) ind.push(`w:left="${p.dxaLeft}"`);
  if (p.dxaRight !== undefined) ind.push(`w:right="${p.dxaRight}"`);
  if (p.dxaLeft1 !== undefined) ind.push(p.dxaLeft1 < 0 ? `w:hanging="${-p.dxaLeft1}"` : `w:firstLine="${p.dxaLeft1}"`);
  if (ind.length) out.push(`<w:ind ${ind.join(" ")}/>`);
  out.push(onOff("contextualSpacing", p.contextualSpacing), onOff("mirrorIndents", p.mirrorIndents), onOff("suppressOverlap", p.noAllowOverlap));
  if (p.jc !== undefined && JC[p.jc]) out.push(`<w:jc w:val="${JC[p.jc]}"/>`);
  if (p.textAlignment !== undefined && TEXT_ALIGN[p.textAlignment]) out.push(`<w:textAlignment w:val="${TEXT_ALIGN[p.textAlignment]}"/>`);
  if (p.outlineLevel !== undefined && p.outlineLevel < 9) out.push(`<w:outlineLvl w:val="${p.outlineLevel}"/>`);
  return { before: out.join(""), after: "" };
}

function framePrXml(f: FrameProps): string {
  const attrs: string[] = [];
  if (f.dropCap) attrs.push(`w:dropCap="${f.dropCap === 2 ? "margin" : "drop"}" w:lines="${f.dropLines ?? 3}"`);
  if (f.width) attrs.push(`w:w="${f.width}"`);
  if (f.height) attrs.push(`w:h="${f.height}" w:hRule="${f.minHeight ? "atLeast" : "exact"}"`);
  if (f.dxaFromText !== undefined) attrs.push(`w:hSpace="${f.dxaFromText}"`);
  if (f.dyaFromText !== undefined) attrs.push(`w:vSpace="${f.dyaFromText}"`);
  if (f.wrap !== undefined && WRAP[f.wrap]) attrs.push(`w:wrap="${WRAP[f.wrap]}"`);
  if (f.pcVert !== undefined && f.pcVert !== 3) attrs.push(`w:vAnchor="${["margin", "page", "text"][f.pcVert]}"`);
  if (f.pcHorz !== undefined && f.pcHorz !== 3) attrs.push(`w:hAnchor="${["text", "margin", "page"][f.pcHorz]}"`);
  if (f.x !== undefined) attrs.push(X_ALIGN[f.x] ? `w:xAlign="${X_ALIGN[f.x]}"` : `w:x="${(f.x << 16) >> 16}"`);
  if (f.y !== undefined) attrs.push(Y_ALIGN[f.y] ? `w:yAlign="${Y_ALIGN[f.y]}"` : `w:y="${(f.y << 16) >> 16}"`);
  if (f.locked) attrs.push('w:anchorLock="1"');
  // A frame needs a position or a drop cap to be one: a lone wrap or width is noise.
  const positioned = f.dropCap || f.x !== undefined || f.y !== undefined || (f.pcHorz !== undefined && f.pcHorz !== 3) || (f.pcVert !== undefined && f.pcVert !== 3) || f.width || f.height;
  return positioned && attrs.length ? `<w:framePr ${attrs.join(" ")}/>` : "";
}

const FTS: Record<number, string> = { 0: "nil", 1: "auto", 2: "pct", 3: "dxa" };

function widthXml(name: string, w: { fts: number; w: number } | undefined): string {
  if (!w || w.fts === 0) return "";
  return `<w:${name} w:w="${w.w}" w:type="${FTS[w.fts] ?? "dxa"}"/>`;
}

/**
 * A row mark's tblPr: the table-wide properties this row carries. A table
 * style's tblPr (`inStyle`) holds only what its sprms set.
 */
export function tblPrXml(t: TableProps, ctx: XmlContext, tblInd: number | undefined, inStyle = false): string {
  const out: string[] = [];
  if (t.istd !== undefined) {
    const id = ctx.styleId(t.istd);
    if (id) out.push(`<w:tblStyle w:val="${escapeAttr(id)}"/>`);
  }
  if (t.pcVert !== undefined || t.pcHorz !== undefined || t.dxaAbs !== undefined || t.dyaAbs !== undefined) {
    const attrs: string[] = [];
    if (t.dxaFromText !== undefined) attrs.push(`w:leftFromText="${t.dxaFromText}"`);
    if (t.dxaFromTextRight !== undefined) attrs.push(`w:rightFromText="${t.dxaFromTextRight}"`);
    if (t.dyaFromText !== undefined) attrs.push(`w:topFromText="${t.dyaFromText}"`);
    if (t.dyaFromTextBottom !== undefined) attrs.push(`w:bottomFromText="${t.dyaFromTextBottom}"`);
    if (t.pcHorz !== undefined && t.pcHorz !== 3) attrs.push(`w:horzAnchor="${["text", "margin", "page"][t.pcHorz]}"`);
    if (t.pcVert !== undefined && t.pcVert !== 3) attrs.push(`w:vertAnchor="${["margin", "page", "text"][t.pcVert]}"`);
    if (t.dxaAbs !== undefined) attrs.push(X_ALIGN[t.dxaAbs] ? `w:tblpXSpec="${X_ALIGN[t.dxaAbs]}"` : `w:tblpX="${(t.dxaAbs << 16) >> 16}"`);
    if (t.dyaAbs !== undefined) attrs.push(Y_ALIGN[t.dyaAbs] ? `w:tblpYSpec="${Y_ALIGN[t.dyaAbs]}"` : `w:tblpY="${(t.dyaAbs << 16) >> 16}"`);
    if (attrs.length) out.push(`<w:tblpPr ${attrs.join(" ")}/>`);
  }
  if (t.noAllowOverlap) out.push('<w:tblOverlap w:val="never"/>');
  out.push(onOff("bidiVisual", t.bidi));
  out.push(widthXml("tblW", t.tableWidth) || (inStyle ? "" : '<w:tblW w:w="0" w:type="auto"/>'));
  if (t.jc !== undefined && t.jc !== 0) out.push(`<w:jc w:val="${t.jc === 1 ? "center" : "right"}"/>`);
  if (t.cellSpacing !== undefined) out.push(`<w:tblCellSpacing w:w="${Math.floor(t.cellSpacing / 2)}" w:type="dxa"/>`);
  if (t.indent && t.indent.fts === 3) {
    if (t.indent.w !== 0) out.push(`<w:tblInd w:w="${t.indent.w}" w:type="dxa"/>`);
  } else if (tblInd !== undefined && tblInd !== 0) out.push(`<w:tblInd w:w="${tblInd}" w:type="dxa"/>`);
  const borders = (["top", "left", "bottom", "right", "insideH", "insideV"] as const).map((side) => borderXml(side, t.borders[side])).join("");
  if (borders) out.push(`<w:tblBorders>${borders}</w:tblBorders>`);
  out.push(shadingXml(t.shading));
  if (inStyle ? t.autofit === false : !t.autofit) out.push('<w:tblLayout w:type="fixed"/>');
  // A table with no style of its own leaves out the margins Normal Table has (0, 108, 0, 108), as Word writes it.
  const normal = !inStyle && t.istd === undefined;
  const margins = { top: t.cellMargins.top, left: t.cellMargins.left ?? t.dxaGapHalf, bottom: t.cellMargins.bottom, right: t.cellMargins.right ?? t.dxaGapHalf };
  const normalMargins = { top: 0, left: 108, bottom: 0, right: 108 };
  const mar = (["top", "left", "bottom", "right"] as const)
    .filter((side) => margins[side] !== undefined && !(normal && margins[side] === normalMargins[side]))
    .map((side) => `<w:${side} w:w="${margins[side]}" w:type="dxa"/>`);
  if (mar.length) out.push(`<w:tblCellMar>${mar.join("")}</w:tblCellMar>`);
  // Word writes a table with no look (no sprmTTlp) as tblLook 0000.
  if (inStyle && t.tlpFlags === undefined) return out.join("");
  const look = t.tlpFlags ?? 0;
  const bit = (mask: number): string => (look & mask ? "1" : "0");
  out.push(`<w:tblLook w:val="${look.toString(16).toUpperCase().padStart(4, "0")}" w:firstRow="${bit(0x20)}" w:lastRow="${bit(0x40)}" w:firstColumn="${bit(0x80)}" w:lastColumn="${bit(0x100)}" w:noHBand="${bit(0x200)}" w:noVBand="${bit(0x400)}"/>`);
  return out.join("");
}

/** A row's trPr. */
export function trPrXml(t: TableProps, gridBefore: number, gridAfter: number): string {
  const out: string[] = [];
  if (t.cellSpacing !== undefined) out.push(`<w:tblCellSpacing w:w="${Math.floor(t.cellSpacing / 2)}" w:type="dxa"/>`);
  if (gridBefore) out.push(`<w:gridBefore w:val="${gridBefore}"/>`);
  if (gridAfter) out.push(`<w:gridAfter w:val="${gridAfter}"/>`);
  if (gridBefore) out.push(widthXml("wBefore", t.widthBefore));
  if (gridAfter) out.push(widthXml("wAfter", t.widthAfter));
  out.push(onOff("cantSplit", t.cantSplit));
  if (t.dyaRowHeight) out.push(`<w:trHeight${t.dyaRowHeight < 0 ? ' w:hRule="exact"' : ""} w:val="${Math.abs(t.dyaRowHeight)}"/>`);
  out.push(onOff("tblHeader", t.header));
  if (t.jc !== undefined && t.jc !== 0) out.push(`<w:jc w:val="${t.jc === 1 ? "center" : "right"}"/>`);
  return out.join("");
}

const TEXT_DIRECTION: Record<number, string> = { 1: "tbRl", 3: "btLr", 4: "lrTbV", 5: "tbRlV" };

/** A cell's tcPr: its width, span, merges, borders, shading and margins. */
export function tcPrXml(cell: CellProps, width: number, gridSpan: number, autofit: boolean): string {
  const out: string[] = [];
  // The preferred width: sprmTCellWidth's, else the TC80's own (wWidth, ftsWidth).
  const prefWidth = cell.prefWidth ?? { fts: cell.ftsWidth, w: cell.width };
  if (prefWidth.fts !== 0) out.push(widthXml("tcW", prefWidth));
  else if (autofit) out.push('<w:tcW w:w="0" w:type="auto"/>');
  else out.push(`<w:tcW w:w="${width}" w:type="dxa"/>`);
  if (gridSpan > 1) out.push(`<w:gridSpan w:val="${gridSpan}"/>`);
  if (cell.vertMerge === 3) out.push('<w:vMerge w:val="restart"/>');
  else if (cell.vertMerge === 1) out.push("<w:vMerge/>");
  const borders = (["top", "left", "bottom", "right", "tl2br", "tr2bl"] as const).map((side) => (cell.borders[side] === NIL_BORDER ? `<w:${side} w:val="nil"/>` : borderXml(side, cell.borders[side]))).join("");
  if (borders) out.push(`<w:tcBorders>${borders}</w:tcBorders>`);
  out.push(shadingXml(cell.shading));
  if (cell.noWrap) out.push("<w:noWrap/>");
  const mar = (["top", "left", "bottom", "right"] as const).filter((side) => cell.margins[side] !== undefined).map((side) => `<w:${side} w:w="${cell.margins[side]}" w:type="dxa"/>`).join("");
  if (mar) out.push(`<w:tcMar>${mar}</w:tcMar>`);
  if (cell.textFlow && TEXT_DIRECTION[cell.textFlow]) out.push(`<w:textDirection w:val="${TEXT_DIRECTION[cell.textFlow]}"/>`);
  if (cell.fitText) out.push("<w:tcFitText/>");
  if (cell.vertAlign === 1) out.push('<w:vAlign w:val="center"/>');
  else if (cell.vertAlign === 2) out.push('<w:vAlign w:val="bottom"/>');
  if (cell.hideMark) out.push("<w:hideMark/>");
  return out.join("");
}

// ---- Section properties --------------------------------------------------

export function applySectSprms(grpprl: Uint8Array, s: SectionProps): void {
  for (const { sprm, operand: o } of sprms(grpprl)) {
    switch (sprm) {
      case 0xf203: s.colWidths[o[0]] = u16(o, 1); break;
      case 0xf204: s.colSpacings[o[0]] = u16(o, 1); break;
      case 0x3005: s.evenlySpaced = o[0] !== 0; break;
      case 0x3009: s.bkc = o[0]; break;
      case 0x300a: s.titlePage = o[0] !== 0; break;
      case 0x500b: s.columns = u16(o, 0) + 1; break;
      case 0x900c: s.dxaColumns = u16(o, 0); break;
      case 0x300e: s.nfcPgn = o[0]; break;
      case 0x3011: s.pgnRestart = o[0] !== 0; break;
      case 0x3012: s.endnotesHere = o[0] !== 0; break;
      case 0x3013: s.lnc = o[0]; break;
      case 0x5015: s.nLnnMod = u16(o, 0); break;
      case 0x9016: s.dxaLnn = u16(o, 0); break;
      case 0xb017: s.dyaHdrTop = u16(o, 0); break;
      case 0xb018: s.dyaHdrBottom = u16(o, 0); break;
      case 0x3019: s.lineBetween = o[0] !== 0; break;
      case 0x301a: s.vjc = o[0]; break;
      case 0x501b: s.lnnMin = u16(o, 0); break;
      case 0x501c: s.pgnStart = u16(o, 0); break;
      case 0x301d: s.landscape = o[0] === 2; break;
      case 0xb01f: s.xaPage = u16(o, 0); break;
      case 0xb020: s.yaPage = u16(o, 0); break;
      case 0xb021: s.dxaLeft = u16(o, 0); break;
      case 0xb022: s.dxaRight = u16(o, 0); break;
      case 0x9023: s.dyaTop = i16(o, 0); break;
      case 0x9024: s.dyaBottom = i16(o, 0); break;
      case 0xb025: s.dzaGutter = u16(o, 0); break;
      case 0x3228: s.bidi = o[0] !== 0; break;
      case 0x322a: s.rtlGutter = o[0] !== 0; break;
      case 0x702b: s.borders.top = brc80(o, 0); break;
      case 0x702c: s.borders.left = brc80(o, 0); break;
      case 0x702d: s.borders.bottom = brc80(o, 0); break;
      case 0x702e: s.borders.right = brc80(o, 0); break;
      case 0x522f: s.pgbProp = u16(o, 0); break;
      case 0x7030: s.dxtCharSpace = i32(o, 0); break;
      case 0x9031: s.dyaLinePitch = i16(o, 0); break;
      case 0x5032: s.clm = u16(o, 0); break;
      case 0x5033: s.textFlow = u16(o, 0); break;
      case 0x5026: s.paperCode = u16(o, 0); break;
      case 0xd234: s.borders.top = brc(o, 1); break;
      case 0xd235: s.borders.left = brc(o, 1); break;
      case 0xd236: s.borders.bottom = brc(o, 1); break;
      case 0xd237: s.borders.right = brc(o, 1); break;
      case 0x303b: s.fpc = o[0]; break;
      case 0x303c: s.rncFtn = o[0]; break;
      case 0x303e: s.rncEdn = o[0]; break;
      case 0x503f: s.nFtn = u16(o, 0); break;
      case 0x5040: s.nfcFtnRef = u16(o, 0); break;
      case 0x5041: s.nEdn = u16(o, 0); break;
      case 0x5042: s.nfcEdnRef = u16(o, 0); break;
      case 0xd243:
        if (o[1]) s.propRevision = { ibst: u16(o, 2), dttm: u32(o, 4) };
        else delete s.propRevision;
        break;
      case 0x7044: s.pgnStart = u32(o, 0); break;
      case 0xd202: s.olst = o.subarray(1); break;
      case 0x3014: s.grpfIhdt = o[0]; break;
      default: break;
    }
  }
}

/** The MSONFC numbering formats ([MS-OSHARED] 2.2.1.3) as ST_NumberFormat. */
export const NUM_FORMATS: string[] = [
  "decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter", "ordinal", "cardinalText", "ordinalText", "hex", "chicago",
  "ideographDigital", "japaneseCounting", "aiueo", "iroha", "decimalFullWidth", "decimalHalfWidth", "japaneseLegal", "japaneseDigitalTenThousand",
  "decimalEnclosedCircle", "decimalFullWidth2", "aiueoFullWidth", "irohaFullWidth", "decimalZero", "bullet", "ganada", "chosung",
  "decimalEnclosedFullstop", "decimalEnclosedParen", "decimalEnclosedCircleChinese", "ideographEnclosedCircle", "ideographTraditional",
  "ideographZodiac", "ideographZodiacTraditional", "taiwaneseCounting", "ideographLegalTraditional", "taiwaneseCountingThousand",
  "taiwaneseDigital", "chineseCounting", "chineseLegalSimplified", "chineseCountingThousand", "koreanDigital", "koreanCounting", "koreanLegal",
  "koreanDigital2", "vietnameseCounting", "russianLower", "russianUpper", "none", "numberInDash", "hebrew1", "hebrew2", "arabicAlpha",
  "arabicAbjad", "hindiVowels", "hindiConsonants", "hindiNumbers", "hindiCounting", "thaiLetters", "thaiNumbers", "thaiCounting",
];

export function numFormat(nfc: number): string {
  if (nfc === 0xff) return "none";
  return NUM_FORMATS[nfc] ?? "decimal";
}

/** A section's sectPr body; `headerRefs` is the headerReference/footerReference XML. */
export function sectPrXml(s: SectionProps, headerRefs: string, footnotePr: string, endnotePr: string): string {
  const out: string[] = [headerRefs];
  out.push(footnotePr, endnotePr);
  const types = ["continuous", "nextColumn", "nextPage", "evenPage", "oddPage"];
  if (s.bkc !== undefined && s.bkc !== 2) out.push(`<w:type w:val="${types[s.bkc] ?? "nextPage"}"/>`);
  out.push(`<w:pgSz w:w="${s.xaPage ?? 12240}" w:h="${s.yaPage ?? 15840}"${s.landscape ? ' w:orient="landscape"' : ""}${s.paperCode ? ` w:code="${s.paperCode}"` : ""}/>`);
  out.push(`<w:pgMar w:top="${s.dyaTop ?? 1440}" w:right="${s.dxaRight ?? 1800}" w:bottom="${s.dyaBottom ?? 1440}" w:left="${s.dxaLeft ?? 1800}" w:header="${s.dyaHdrTop ?? 720}" w:footer="${s.dyaHdrBottom ?? 720}" w:gutter="${s.dzaGutter ?? 0}"/>`);
  const pb = (["top", "left", "bottom", "right"] as const).map((side) => borderXml(side, s.borders[side])).join("");
  if (pb) {
    const prop = s.pgbProp ?? 0;
    // Only the values that differ from the defaults (front, allPages, text), as Word writes them.
    const display = ["allPages", "firstPage", "notFirstPage"][prop & 7] ?? "allPages";
    const offsetFrom = (prop >>> 5) & 7 ? "page" : "text";
    const zOrder = (prop >>> 3) & 3 ? "back" : "front";
    const attrs = (zOrder !== "front" ? ` w:zOrder="${zOrder}"` : "") + (display !== "allPages" ? ` w:display="${display}"` : "") + (offsetFrom !== "text" ? ` w:offsetFrom="${offsetFrom}"` : "");
    out.push(`<w:pgBorders${attrs}>${pb}</w:pgBorders>`);
  }
  if (s.nLnnMod) {
    const restart = ["newPage", "newSection", "continuous"][s.lnc ?? 0] ?? "newPage";
    out.push(`<w:lnNumType w:countBy="${s.nLnnMod}"${s.dxaLnn ? ` w:distance="${s.dxaLnn}"` : ""}${s.lnnMin ? ` w:start="${s.lnnMin + 1}"` : ""} w:restart="${restart}"/>`);
  }
  if (s.pgnRestart || (s.nfcPgn !== undefined && s.nfcPgn !== 0)) {
    out.push(`<w:pgNumType${s.nfcPgn !== undefined && s.nfcPgn !== 0 ? ` w:fmt="${numFormat(s.nfcPgn)}"` : ""}${s.pgnRestart ? ` w:start="${s.pgnStart ?? 1}"` : ""}/>`);
  }
  // The column spacing is 720 twips unless the section says otherwise.
  const num = s.columns ?? 1;
  let cols = `<w:cols${num > 1 ? ` w:num="${num}"` : ""} w:space="${s.dxaColumns ?? 720}"`;
  if (s.lineBetween) cols += ' w:sep="1"';
  if (num > 1 && s.evenlySpaced === false && s.colWidths.length) {
    cols += ' w:equalWidth="0">';
    for (let i = 0; i < num; i++) cols += `<w:col w:w="${s.colWidths[i] ?? 0}"${i < num - 1 ? ` w:space="${s.colSpacings[i] ?? s.dxaColumns ?? 0}"` : ""}/>`;
    cols += "</w:cols>";
  } else {
    cols += "/>";
  }
  out.push(cols);
  if (s.vjc) out.push(`<w:vAlign w:val="${["top", "center", "both", "bottom"][s.vjc] ?? "top"}"/>`);
  if (s.titlePage) out.push("<w:titlePg/>");
  if (s.textFlow && TEXT_DIRECTION[s.textFlow]) out.push(`<w:textDirection w:val="${TEXT_DIRECTION[s.textFlow]}"/>`);
  if (s.bidi) out.push("<w:bidi/>");
  if (s.rtlGutter) out.push("<w:rtlGutter/>");
  if (s.clm !== undefined && s.clm !== 0 && s.dyaLinePitch) {
    // SClmOperand: 1 chars and lines, 2 lines only, 3 every character in its own square.
    const type = ["default", "linesAndChars", "lines", "snapToChars"][s.clm & 3];
    out.push(`<w:docGrid w:type="${type}" w:linePitch="${s.dyaLinePitch}"${s.dxtCharSpace ? ` w:charSpace="${s.dxtCharSpace}"` : ""}/>`);
  } else if (s.dyaLinePitch) {
    // No grid: the pitch and spacing are kept for when one is turned on.
    out.push(`<w:docGrid w:linePitch="${s.dyaLinePitch}"${s.dxtCharSpace ? ` w:charSpace="${s.dxtCharSpace}"` : ""}/>`);
  }
  return out.join("");
}

/** The footnotePr or endnotePr of a section from its numbering sprms. */
export function notePrXml(name: "footnotePr" | "endnotePr", s: SectionProps): string {
  const footnote = name === "footnotePr";
  const parts: string[] = [];
  if (footnote && s.fpc !== undefined && s.fpc !== 1) parts.push(`<w:pos w:val="${s.fpc === 0 ? "sectEnd" : "beneathText"}"/>`);
  if (!footnote && s.endnotesHere) parts.push('<w:pos w:val="sectEnd"/>');
  const nfc = footnote ? s.nfcFtnRef : s.nfcEdnRef;
  if (nfc !== undefined && nfc !== (footnote ? 0 : 2)) parts.push(`<w:numFmt w:val="${numFormat(nfc)}"/>`);
  const start = footnote ? s.nFtn : s.nEdn;
  if (start) parts.push(`<w:numStart w:val="${start + 1}"/>`);
  const rnc = footnote ? s.rncFtn : s.rncEdn;
  if (rnc) parts.push(`<w:numRestart w:val="${rnc === 1 ? "eachSect" : "eachPage"}"/>`);
  return parts.length ? `<w:${name}>${parts.join("")}</w:${name}>` : "";
}

/** The level text of a list level: its placeholders become %1..%9 ([MS-DOC] 2.9.149). */
export function levelText(xst: Uint8Array, at: number, placeholders: Uint8Array): { text: string; count: number } {
  const cch = u16(xst, at);
  const chars = utf16(xst, at + 2, cch);
  const positions = new Set<number>();
  for (let i = 0; i < 9 && placeholders[i] !== 0; i++) positions.add(placeholders[i] - 1);
  let text = "";
  for (let i = 0; i < chars.length; i++) {
    if (positions.has(i)) text += `%${chars.charCodeAt(i) + 1}`;
    else text += chars[i];
  }
  return { text, count: cch };
}
