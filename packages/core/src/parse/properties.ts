import {
  XmlElement,
  attr,
  child,
  children,
  childVal,
  intAttr,
  localName,
  onOff,
} from "../xml.js";
import {
  Alignment,
  Border,
  BorderStyle,
  LineSpacing,
  ParaProps,
  ParagraphBorders,
  RunProps,
  TabStop,
  Theme,
} from "../model.js";
import { eighthPtToPx, halfPtToPx, ptToPx, twipsToPx } from "../units.js";

export interface ParseContext {
  theme?: Theme;
  /** Tracked-changes display: "final" hides deletions and shows insertions
   * plain (default); "markup" colors both. */
  revisionView?: "final" | "markup";
}

// ---------- colors ----------

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: "#ffff00",
  green: "#00ff00",
  cyan: "#00ffff",
  magenta: "#ff00ff",
  blue: "#0000ff",
  red: "#ff0000",
  darkBlue: "#00008b",
  darkCyan: "#008b8b",
  darkGreen: "#006400",
  darkMagenta: "#800080",
  darkRed: "#8b0000",
  darkYellow: "#808000",
  darkGray: "#a9a9a9",
  lightGray: "#d3d3d3",
  black: "#000000",
  white: "#ffffff",
};

export function parseColor(
  el: XmlElement | undefined,
  ctx: ParseContext,
): string | undefined {
  if (!el) return undefined;
  const themeColor = attr(el, "themeColor");
  if (themeColor && ctx.theme) {
    const mapped = ctx.theme.colors.get(themeColor);
    if (mapped) {
      const tint = attr(el, "themeTint");
      const shade = attr(el, "themeShade");
      return applyTintShade(mapped, tint, shade);
    }
  }
  const val = attr(el, "val");
  if (!val || val === "auto") return val === "auto" ? "auto" : undefined;
  return "#" + val;
}

function applyTintShade(hex: string, tint?: string, shade?: string): string {
  let rgb = hexToRgb(hex);
  if (!rgb) return hex;
  if (tint) {
    const t = parseInt(tint, 16) / 255;
    rgb = rgb.map((c) => Math.round(c * t + 255 * (1 - t))) as [number, number, number];
  }
  if (shade) {
    const s = parseInt(shade, 16) / 255;
    rgb = rgb.map((c) => Math.round(c * s)) as [number, number, number];
  }
  return rgbToHex(rgb);
}

function hexToRgb(hex: string): [number, number, number] | null {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(hex);
  if (!m) return null;
  const v = parseInt(m[1], 16);
  return [(v >> 16) & 255, (v >> 8) & 255, v & 255];
}
function rgbToHex([r, g, b]: [number, number, number]): string {
  return "#" + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1);
}

/** w:shd → CSS fill color or undefined. */
export function parseShading(el: XmlElement | undefined, ctx: ParseContext): string | undefined {
  if (!el) return undefined;
  const fill = attr(el, "fill");
  const themeFill = attr(el, "themeFill");
  if (themeFill && ctx.theme) {
    const mapped = ctx.theme.colors.get(themeFill);
    if (mapped) return applyTintShade(mapped, attr(el, "themeFillTint"), attr(el, "themeFillShade"));
  }
  const val = attr(el, "val");
  const color = attr(el, "color");
  // val="pctNN": blend NN% of the pattern color over the fill (white when
  // auto). Word renders pct25 blue as a light blue tint, not solid navy.
  const pct = val ? /^pct(\d+)$/.exec(val) : null;
  if (pct && color && color !== "auto") {
    const frac = Math.min(parseInt(pct[1], 10), 100) / 100;
    const base = fill && fill !== "auto" && fill.length === 6 ? fill : "FFFFFF";
    const hex = (s: string) => [0, 2, 4].map((i) => parseInt(s.slice(i, i + 2), 16));
    const [br, bg, bb] = hex(base);
    const [cr, cg, cb] = hex(color.length === 6 ? color : "000000");
    const mix = (b: number, c: number) => Math.round(b + (c - b) * frac);
    return "#" + [mix(br, cr), mix(bg, cg), mix(bb, cb)].map((n) => n.toString(16).padStart(2, "0")).join("");
  }
  if (fill && fill !== "auto") return "#" + fill;
  if (val === "solid" && color && color !== "auto") return "#" + color;
  return undefined;
}

// ---------- borders ----------

const BORDER_STYLE_MAP: Record<string, BorderStyle> = {
  single: "single",
  thick: "thick",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dotDash: "dotDash",
  dotDotDash: "dotDotDash",
  triple: "triple",
  wave: "wave",
  none: "none",
  nil: "none",
};

export function parseBorder(el: XmlElement | undefined, ctx: ParseContext): Border | undefined {
  if (!el) return undefined;
  const val = attr(el, "val") ?? "single";
  const style = BORDER_STYLE_MAP[val] ?? "single";
  if (style === "none") return { style: "none", width: 0, color: "transparent", space: 0 };
  const sz = intAttr(el, "sz") ?? 4; // eighth-points
  const space = intAttr(el, "space") ?? 0; // points
  let color = "#000000";
  const themeColor = attr(el, "themeColor");
  const mapped = themeColor ? ctx.theme?.colors.get(themeColor) : undefined;
  if (mapped) color = mapped;
  else {
    const colorAttr = attr(el, "color");
    if (colorAttr && colorAttr !== "auto") color = "#" + colorAttr;
  }
  return {
    style,
    width: Math.max(eighthPtToPx(sz), 0.75),
    color,
    space: ptToPx(space),
  };
}

export function parseParagraphBorders(
  el: XmlElement | undefined,
  ctx: ParseContext,
): ParagraphBorders | undefined {
  if (!el) return undefined;
  const out: ParagraphBorders = {};
  for (const side of ["top", "bottom", "left", "right", "between"] as const) {
    const b = parseBorder(child(el, side), ctx);
    if (b) out[side] = b;
  }
  return Object.keys(out).length ? out : undefined;
}

// ---------- run properties ----------

export function parseRunProps(rPr: XmlElement | undefined, ctx: ParseContext): RunProps {
  const props: RunProps = {};
  if (!rPr) return props;

  const b = onOff(child(rPr, "b"));
  if (b !== undefined) props.bold = b;
  const i = onOff(child(rPr, "i"));
  if (i !== undefined) props.italic = i;
  const strike = onOff(child(rPr, "strike"));
  if (strike !== undefined) props.strike = strike;
  const dstrike = onOff(child(rPr, "dstrike"));
  if (dstrike !== undefined) props.doubleStrike = dstrike;
  const caps = onOff(child(rPr, "caps"));
  if (caps !== undefined) props.caps = caps;
  const smallCaps = onOff(child(rPr, "smallCaps"));
  if (smallCaps !== undefined) props.smallCaps = smallCaps;
  const vanish = onOff(child(rPr, "vanish"));
  if (vanish !== undefined) props.vanish = vanish;
  const outline = onOff(child(rPr, "outline"));
  if (outline !== undefined) props.outline = outline;
  const emboss = onOff(child(rPr, "emboss"));
  if (emboss !== undefined) props.emboss = emboss;
  const imprint = onOff(child(rPr, "imprint"));
  if (imprint !== undefined) props.imprint = imprint;

  const u = childVal(rPr, "u");
  if (u !== undefined) props.underline = u;

  const rFonts = child(rPr, "rFonts");
  if (rFonts) {
    const theme = ctx.theme;
    const themeFont = (kind: string | undefined) => {
      if (!kind || !theme) return undefined;
      if (kind === "majorBidi") return theme.majorBidiFont ?? theme.majorFont;
      if (kind === "minorBidi") return theme.minorBidiFont ?? theme.minorFont;
      return kind.startsWith("major") ? theme.majorFont : theme.minorFont;
    };
    let font = attr(rFonts, "ascii") ?? attr(rFonts, "hAnsi") ?? themeFont(attr(rFonts, "asciiTheme"));
    if (font) props.font = font;
    // East Asian font channel (used for CJK codepoints) and complex-script
    // channel (used for w:rtl runs). Word picks the channel per character.
    const ea = attr(rFonts, "eastAsia") ?? themeFont(attr(rFonts, "eastAsiaTheme"));
    if (ea) props.fontEastAsia = ea;
    const cs = attr(rFonts, "cs") ?? themeFont(attr(rFonts, "cstheme") ?? attr(rFonts, "csTheme"));
    if (cs) props.fontComplex = cs;
  }

  const rtl = onOff(child(rPr, "rtl"));
  if (rtl !== undefined) props.rtl = rtl;

  const sz = childVal(rPr, "sz");
  if (sz !== undefined) props.size = halfPtToPx(parseFloat(sz));

  const color = parseColor(child(rPr, "color"), ctx);
  if (color !== undefined) props.color = color;

  const highlight = childVal(rPr, "highlight");
  if (highlight && highlight !== "none") props.highlight = HIGHLIGHT_COLORS[highlight] ?? highlight;

  const shd = parseShading(child(rPr, "shd"), ctx);
  if (shd !== undefined) props.shading = shd;

  const vertAlign = childVal(rPr, "vertAlign");
  if (vertAlign === "superscript" || vertAlign === "subscript") props.verticalAlign = vertAlign;
  else if (vertAlign === "baseline") props.verticalAlign = "baseline";

  const spacing = intAttr(child(rPr, "spacing"), "val");
  if (spacing !== undefined) props.letterSpacing = twipsToPx(spacing);

  // w:w: horizontal character scaling in percent (Text scale 150% stretches
  // glyph advances and painted glyphs; 66% condenses).
  const wScale = intAttr(child(rPr, "w"), "val");
  if (wScale !== undefined && wScale > 0 && wScale !== 100) props.textScale = wScale / 100;

  // w:position: baseline shift in half-points, positive = raised. Word grows
  // the line box by the full shift (a +6pt raise adds exactly 6pt of pitch).
  const position = intAttr(child(rPr, "position"), "val");
  if (position !== undefined && position !== 0) props.raise = (position / 2) * (4 / 3);

  const rStyle = childVal(rPr, "rStyle");
  if (rStyle) props.styleId = rStyle;

  const lang = childVal(rPr, "lang");
  if (lang) props.lang = lang;

  return props;
}

// ---------- paragraph properties ----------

const ALIGN_MAP: Record<string, Alignment> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

export function parseParaProps(pPr: XmlElement | undefined, ctx: ParseContext): ParaProps {
  const props: ParaProps = {};
  if (!pPr) return props;

  const styleId = childVal(pPr, "pStyle");
  if (styleId) props.styleId = styleId;

  const bidi = onOff(child(pPr, "bidi"));
  if (bidi !== undefined) props.bidi = bidi;

  const jc = childVal(pPr, "jc");
  if (jc && ALIGN_MAP[jc]) props.alignment = ALIGN_MAP[jc];

  const ind = child(pPr, "ind");
  if (ind) {
    const left = intAttr(ind, "left") ?? intAttr(ind, "start");
    if (left !== undefined) props.indentLeft = twipsToPx(left);
    const right = intAttr(ind, "right") ?? intAttr(ind, "end");
    if (right !== undefined) props.indentRight = twipsToPx(right);
    // firstLine and hanging share ONE mutually-exclusive slot in a w:ind. When a
    // level specifies EITHER, it must clear an inherited value of the other, or a
    // parent style's hanging leaks past a child's firstLine="0" — half the hanging
    // then shifts a centered line left (gatech cover headings: CoverPageSingleSpace
    // hanging=360 under MainBodyHeadings firstLine=0 pushed the title 9pt left).
    // hanging wins when both are (invalidly) present, matching Word.
    const firstLine = intAttr(ind, "firstLine");
    const hanging = intAttr(ind, "hanging");
    if (hanging !== undefined) {
      props.indentHanging = twipsToPx(hanging);
      props.indentFirstLine = 0;
    } else if (firstLine !== undefined) {
      props.indentFirstLine = twipsToPx(firstLine);
      props.indentHanging = 0;
    }
  }

  const spacing = child(pPr, "spacing");
  if (spacing) {
    const before = intAttr(spacing, "before");
    if (before !== undefined) props.spacingBefore = twipsToPx(before);
    const after = intAttr(spacing, "after");
    if (after !== undefined) props.spacingAfter = twipsToPx(after);
    const onOffVal = (v: string | undefined): boolean | undefined =>
      v === undefined ? undefined : !(v === "0" || v === "false" || v === "off");
    const beforeAuto = onOffVal(attr(spacing, "beforeAutospacing"));
    if (beforeAuto !== undefined) props.beforeAutospacing = beforeAuto;
    const afterAuto = onOffVal(attr(spacing, "afterAutospacing"));
    if (afterAuto !== undefined) props.afterAutospacing = afterAuto;
    const line = intAttr(spacing, "line");
    const lineRule = attr(spacing, "lineRule") ?? "auto";
    if (line !== undefined) {
      const ls: LineSpacing =
        lineRule === "exact"
          ? { rule: "exact", value: twipsToPx(line) }
          : lineRule === "atLeast"
            ? { rule: "atLeast", value: twipsToPx(line) }
            : { rule: "auto", value: line / 240 };
      props.lineSpacing = ls;
    }
  }

  const contextual = onOff(child(pPr, "contextualSpacing"));
  if (contextual !== undefined) props.contextualSpacing = contextual;
  const snapToGrid = onOff(child(pPr, "snapToGrid"));
  if (snapToGrid !== undefined) props.snapToGrid = snapToGrid;
  const keepNext = onOff(child(pPr, "keepNext"));
  if (keepNext !== undefined) props.keepNext = keepNext;
  const keepLines = onOff(child(pPr, "keepLines"));
  if (keepLines !== undefined) props.keepLines = keepLines;
  const pageBreakBefore = onOff(child(pPr, "pageBreakBefore"));
  if (pageBreakBefore !== undefined) props.pageBreakBefore = pageBreakBefore;
  const widowControl = onOff(child(pPr, "widowControl"));
  if (widowControl !== undefined) props.widowControl = widowControl;
  const frame = child(pPr, "framePr");
  if (frame) {
    const dc = attr(frame, "dropCap");
    if (dc === "drop" || dc === "margin") {
      props.dropCap = {
        mode: dc,
        lines: intAttr(frame, "lines") ?? 3,
        hSpace: twipsToPx(intAttr(frame, "hSpace") ?? 0),
      };
    } else {
      // A general positioned text frame (w:framePr with a width and anchors):
      // the paragraph is lifted out of normal flow, placed at an absolute
      // location, and body text wraps around it (staging-frames). Only the
      // attributes actually present are emitted so the frame merges
      // attribute-by-attribute across the style cascade (ECMA-376: a framePr is
      // not a toggle, but a direct framePr specifying only h/x/y must keep the
      // style's width/anchor — IEEE authors' `w:h w:hRule=exact` over the
      // Authors style's centered full-width frame). The engine defaults the rest.
      const w = intAttr(frame, "w");
      const h = intAttr(frame, "h");
      const hRuleRaw = attr(frame, "hRule");
      const hAnchorRaw = attr(frame, "hAnchor");
      const vAnchorRaw = attr(frame, "vAnchor");
      const wrapRaw = attr(frame, "wrap");
      const xAlignRaw = attr(frame, "xAlign");
      const yAlignRaw = attr(frame, "yAlign");
      const xRaw = intAttr(frame, "x");
      const yRaw = intAttr(frame, "y");
      const hSpaceRaw = intAttr(frame, "hSpace");
      const vSpaceRaw = intAttr(frame, "vSpace");
      const f: NonNullable<ParaProps["frame"]> = {
        ...(w !== undefined && w > 0 ? { w: twipsToPx(w) } : {}),
        ...(h !== undefined ? { h: twipsToPx(h) } : {}),
        ...(hRuleRaw === "atLeast" || hRuleRaw === "exact" || hRuleRaw === "auto" ? { hRule: hRuleRaw } : {}),
        ...(xRaw !== undefined ? { x: twipsToPx(xRaw) } : {}),
        ...(yRaw !== undefined ? { y: twipsToPx(yRaw) } : {}),
        ...(hAnchorRaw === "page" || hAnchorRaw === "margin" || hAnchorRaw === "column" || hAnchorRaw === "text"
          ? { hAnchor: hAnchorRaw }
          : {}),
        ...(vAnchorRaw === "page" || vAnchorRaw === "margin" || vAnchorRaw === "paragraph" || vAnchorRaw === "text"
          ? { vAnchor: vAnchorRaw }
          : {}),
        ...(xAlignRaw ? { xAlign: xAlignRaw as never } : {}),
        ...(yAlignRaw ? { yAlign: yAlignRaw as never } : {}),
        ...(wrapRaw ? { wrap: wrapRaw as never } : {}),
        ...(hSpaceRaw !== undefined ? { hSpace: twipsToPx(hSpaceRaw) } : {}),
        ...(vSpaceRaw !== undefined ? { vSpace: twipsToPx(vSpaceRaw) } : {}),
      };
      if (Object.keys(f).length > 0) props.frame = f;
    }
  }

  const borders = parseParagraphBorders(child(pPr, "pBdr"), ctx);
  if (borders) props.borders = borders;

  const shd = parseShading(child(pPr, "shd"), ctx);
  if (shd !== undefined) props.shading = shd;

  const numPr = child(pPr, "numPr");
  if (numPr) {
    const numId = intAttr(child(numPr, "numId"), "val");
    const ilvl = intAttr(child(numPr, "ilvl"), "val");
    if (numId !== undefined) {
      props.numbering = numId === 0 ? null : { numId, ilvl: ilvl ?? 0 };
    } else if (ilvl !== undefined) {
      // ilvl with no numId: override only the level, keeping the numId the
      // style inherits (Heading3 basedOn Heading2 sets ilvl=2 -> "1.1.2").
      props.numberingLevelOverride = ilvl;
    }
  }

  const tabsEl = child(pPr, "tabs");
  if (tabsEl) {
    const tabs: TabStop[] = [];
    for (const t of children(tabsEl, "tab")) {
      const val = attr(t, "val") ?? "left";
      if (val === "clear") continue;
      const pos = intAttr(t, "pos");
      if (pos === undefined) continue;
      tabs.push({
        pos: twipsToPx(pos),
        align: (val === "center" || val === "right" || val === "decimal" || val === "bar"
          ? val
          : "left") as TabStop["align"],
        leader: (attr(t, "leader") as TabStop["leader"]) ?? "none",
      });
    }
    tabs.sort((a, b) => a.pos - b.pos);
    props.tabs = tabs;
  }

  const outline = intAttr(child(pPr, "outlineLvl"), "val");
  if (outline !== undefined) props.outlineLevel = outline;

  const markRPr = child(pPr, "rPr");
  if (markRPr) props.markRunProps = parseRunProps(markRPr, ctx);

  return props;
}

/** Shallow merge where defined values in `over` win. */
export function mergeRunProps(base: RunProps, over: RunProps): RunProps {
  const out: RunProps = { ...base };
  for (const key of Object.keys(over) as (keyof RunProps)[]) {
    const v = over[key];
    if (v !== undefined) (out as Record<string, unknown>)[key] = v;
  }
  return out;
}

export function mergeParaProps(base: ParaProps, over: ParaProps): ParaProps {
  const out: ParaProps = { ...base };
  for (const key of Object.keys(over) as (keyof ParaProps)[]) {
    const v = over[key];
    if (v !== undefined) (out as Record<string, unknown>)[key] = v;
  }
  if (base.markRunProps && over.markRunProps) {
    out.markRunProps = mergeRunProps(base.markRunProps, over.markRunProps);
  }
  // framePr merges attribute-by-attribute: a direct framePr that specifies only
  // some geometry inherits the rest from the style's framePr (IEEE authors).
  if (base.frame && over.frame) {
    out.frame = { ...base.frame, ...over.frame };
  }
  // Apply a level-only numPr override to the inherited numbering (keeping its
  // numId). Consume it once applied; keep it pending if numbering isn't set
  // yet so a later merge in the style chain can apply it.
  if (out.numberingLevelOverride !== undefined && out.numbering) {
    out.numbering = { ...out.numbering, ilvl: out.numberingLevelOverride };
    out.numberingLevelOverride = undefined;
  }
  return out;
}

export { localName };
