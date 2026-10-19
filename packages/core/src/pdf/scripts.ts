/**
 * The scripts whose text a shaper must turn into glyphs (joining forms,
 * reordering, mark placement, cluster composition), each with the Unicode
 * blocks that hold it. Hebrew is here for its points: its letters alone
 * draw from the cmap (see FontSet.shape), as every script absent from this
 * table does.
 */
export interface ShapedScript {
  /** The script's name, for messages. */
  name: string;
  /** The ISO 15924 tag HarfBuzz takes. */
  tag: string;
  /** The OpenType script tags a font's GSUB or GPOS lists it under, the
   * preferred one first (the Indic "2" tags: the shaping model HarfBuzz
   * follows best). */
  otTags: string[];
  rtl: boolean;
  /** The language fontconfig is asked for a face in, and HarfBuzz shapes in when the run names none. */
  lang: string;
}

function script(name: string, tag: string, lang: string, rtl = false, ot?: string): ShapedScript {
  return { name, tag, otTags: ot ? [ot, tag.toLowerCase()] : [tag.toLowerCase()], rtl, lang };
}

const ARABIC = script("Arabic", "Arab", "ar", true);
const HEBREW = script("Hebrew", "Hebr", "he", true);
const DEVANAGARI = script("Devanagari", "Deva", "hi", false, "dev2");
const MYANMAR = script("Myanmar", "Mymr", "my", false, "mym2");
const KHMER = script("Khmer", "Khmr", "km");

const BLOCKS: [ShapedScript, from: number, to: number][] = [
  [HEBREW, 0x0590, 0x05ff],
  [ARABIC, 0x0600, 0x06ff],
  [script("Syriac", "Syrc", "syr", true), 0x0700, 0x074f],
  [ARABIC, 0x0750, 0x077f],
  [script("Thaana", "Thaa", "dv", true), 0x0780, 0x07bf],
  [script("N'Ko", "Nkoo", "nqo", true), 0x07c0, 0x07ff],
  [ARABIC, 0x0870, 0x08ff],
  [DEVANAGARI, 0x0900, 0x097f],
  [script("Bengali", "Beng", "bn", false, "bng2"), 0x0980, 0x09ff],
  [script("Gurmukhi", "Guru", "pa", false, "gur2"), 0x0a00, 0x0a7f],
  [script("Gujarati", "Gujr", "gu", false, "gjr2"), 0x0a80, 0x0aff],
  [script("Oriya", "Orya", "or", false, "ory2"), 0x0b00, 0x0b7f],
  [script("Tamil", "Taml", "ta", false, "tml2"), 0x0b80, 0x0bff],
  [script("Telugu", "Telu", "te", false, "tel2"), 0x0c00, 0x0c7f],
  [script("Kannada", "Knda", "kn", false, "knd2"), 0x0c80, 0x0cff],
  [script("Malayalam", "Mlym", "ml", false, "mlm2"), 0x0d00, 0x0d7f],
  [script("Sinhala", "Sinh", "si"), 0x0d80, 0x0dff],
  [script("Thai", "Thai", "th"), 0x0e00, 0x0e7f],
  [script("Lao", "Laoo", "lo", false, "lao"), 0x0e80, 0x0eff],
  [script("Tibetan", "Tibt", "bo"), 0x0f00, 0x0fff],
  [MYANMAR, 0x1000, 0x109f],
  [KHMER, 0x1780, 0x17ff],
  [script("Limbu", "Limb", "lif"), 0x1900, 0x194f],
  [script("Tai Le", "Tale", "tdd"), 0x1950, 0x197f],
  [script("New Tai Lue", "Talu", "khb"), 0x1980, 0x19df],
  [KHMER, 0x19e0, 0x19ff],
  [script("Buginese", "Bugi", "bug"), 0x1a00, 0x1a1f],
  [script("Tai Tham", "Lana", "nod"), 0x1a20, 0x1aaf],
  [script("Balinese", "Bali", "ban"), 0x1b00, 0x1b7f],
  [script("Sundanese", "Sund", "su"), 0x1b80, 0x1bbf],
  [script("Batak", "Batk", "bbc"), 0x1bc0, 0x1bff],
  [script("Lepcha", "Lepc", "lep"), 0x1c00, 0x1c4f],
  [DEVANAGARI, 0x1cd0, 0x1cff],
  [DEVANAGARI, 0xa8e0, 0xa8ff],
  [script("Javanese", "Java", "jv"), 0xa980, 0xa9df],
  [MYANMAR, 0xa9e0, 0xa9ff],
  [MYANMAR, 0xaa60, 0xaa7f],
  [script("Tai Viet", "Tavt", "blt"), 0xaa80, 0xaadf],
  [HEBREW, 0xfb1d, 0xfb4f],
  [ARABIC, 0xfb50, 0xfdff],
  [ARABIC, 0xfe70, 0xfeff],
];

/** The shaped script a code point belongs to, if any. */
export function shapedScriptOf(cp: number): ShapedScript | undefined {
  if (cp < 0x0590) return undefined;
  for (const [s, from, to] of BLOCKS) if (cp >= from && cp <= to) return s;
  return undefined;
}

/** Whether a code point is one of Hebrew's points, accents or marks: text with any shapes as Hebrew. */
export function isHebrewMark(cp: number): boolean {
  return (cp >= 0x0591 && cp <= 0x05bd) || cp === 0x05bf || cp === 0x05c1 || cp === 0x05c2 || cp === 0x05c4 || cp === 0x05c5 || cp === 0x05c7;
}

/**
 * Whether `text` can take the plain path when drawn: no code point in a
 * shaped script's block, none in the right-to-left blocks, no
 * right-to-left control. Latin, Greek, Cyrillic and East Asian text
 * qualifies, so a glyph per code point from the cmap draws it as before.
 */
export function plainText(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c < 0x0590) continue;
    if (c <= 0x1cff || (c >= 0xa8e0 && c <= 0xaadf) || (c >= 0xfb1d && c <= 0xfdff) || (c >= 0xfe70 && c <= 0xfeff) || c === 0x200f || c === 0x202b || c === 0x202e || c === 0x2067) return false;
  }
  return true;
}

/**
 * Whether `text` needs the shaper for its width: a shaped script's letter
 * in it. Order does not change a width, so text without one measures on
 * the plain path whatever its direction; nor do Hebrew points, which the
 * shaper places (zero advance) only when the text is drawn.
 */
export function needsShaping(text: string): boolean {
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    if (c < 0x0600) continue;
    if (c <= 0x1cff || (c >= 0xa8e0 && c <= 0xaadf) || (c >= 0xfb50 && c <= 0xfdff) || (c >= 0xfe70 && c <= 0xfeff)) {
      if (shapedScriptOf(c)) return true;
    }
  }
  return false;
}
