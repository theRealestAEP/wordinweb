import type * as HarfBuzz from "harfbuzzjs";
import { ShapedScript } from "./scripts.js";
import { Face } from "./sfnt.js";

/**
 * HarfBuzz (the harfbuzzjs WebAssembly build), loaded once per process on
 * demand: a render awaits loadShaper() before any text is refused or
 * shaped, so the shaper is there for the scripts that need one. When the
 * package is absent, those scripts are refused as unsupported_script.
 */
let hb: typeof HarfBuzz | null = null;
let loading: Promise<boolean> | undefined;

export function loadShaper(): Promise<boolean> {
  loading ??= import("harfbuzzjs").then(
    (module) => {
      hb = module;
      return true;
    },
    () => false,
  );
  return loading;
}

export function shaperLoaded(): boolean {
  return hb !== null;
}

/** One glyph HarfBuzz placed, in font units, in the order the glyphs are drawn left to right. */
export interface ShapedGlyph {
  gid: number;
  /** The UTF-16 offset in the shaped text of the cluster the glyph belongs to. */
  cluster: number;
  advance: number;
  xOffset: number;
  yOffset: number;
  /** A tatweel may be inserted before this glyph's cluster (Arabic elongation). */
  tatweel: boolean;
}

/** The HarfBuzz font of a Face, made on first use and kept with it. */
const fonts = new WeakMap<Face, HarfBuzz.Font>();
let buffer: HarfBuzz.Buffer | null = null;

function fontOf(face: Face): HarfBuzz.Font {
  if (!hb) throw new Error("the shaper is not loaded");
  let font = fonts.get(face);
  if (!font) {
    const hbFace = new hb.Face(new hb.Blob(face.sfnt), face.index);
    font = new hb.Font(hbFace);
    font.setScale(face.unitsPerEm, face.unitsPerEm);
    fonts.set(face, font);
  }
  return font;
}

/**
 * Shape `text` (one script, one direction) with a face. Pair kerning
 * follows `kerning`, as the plain path's does; the script's other features
 * are HarfBuzz's defaults. Default-ignorable characters (soft hyphens,
 * joiners) leave no glyph.
 */
export function shapeText(face: Face, text: string, script: ShapedScript, lang: string | undefined, kerning: boolean): ShapedGlyph[] {
  if (!hb) throw new Error("the shaper is not loaded");
  const font = fontOf(face);
  buffer ??= new hb.Buffer();
  buffer.reset();
  buffer.addText(text);
  buffer.setDirection(script.rtl ? hb.Direction.RTL : hb.Direction.LTR);
  buffer.setScript(script.tag);
  buffer.setLanguage(lang ?? script.lang);
  buffer.setFlags(hb.BufferFlag.REMOVE_DEFAULT_IGNORABLES | hb.BufferFlag.PRODUCE_SAFE_TO_INSERT_TATWEEL);
  hb.shape(font, buffer, kerning ? [] : [new hb.Feature("kern", 0)]);
  const infos = buffer.getGlyphInfos();
  const positions = buffer.getGlyphPositions();
  const out: ShapedGlyph[] = [];
  for (let i = 0; i < infos.length; i++) {
    const info = infos[i];
    const pos = positions[i];
    out.push({ gid: info.codepoint, cluster: info.cluster, advance: pos.xAdvance, xOffset: pos.xOffset, yOffset: pos.yOffset, tatweel: (info.flags & hb.GlyphFlag.SAFE_TO_INSERT_TATWEEL) !== 0 });
  }
  return out;
}
