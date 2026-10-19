import { execFileSync } from "node:child_process";
import { existsSync, readdirSync, statSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, extname, join } from "node:path";
import { createRequire } from "node:module";
import { fontWidthScale, normalizeFamily, substituteFamily } from "../layout/measure.js";
import { FontSpec } from "../layout/types.js";
import { PdfRenderError, unsupportedScript } from "./errors.js";
import { Face, FaceInfo, scanFaces } from "./sfnt.js";

/**
 * Font resolution for the PDF renderer. A document's font name goes through
 * these rules, in order, and the first family a directory holds wins:
 *
 *   1. normalizeFamily: "Times New Roman Bold" is Times New Roman's bold
 *      face; Liberation faces are Word's names for them.
 *   2. The font's paintFamily (CJK: the Windows face), then the family.
 *   3. The layout's substitute for the family (Calibri -> Carlito, Cambria
 *      -> Caladea, Aptos -> system-ui ...; see measure.ts).
 *   4. Word's metric-compatible free substitutes: Arial -> Liberation Sans,
 *      Times New Roman -> Liberation Serif, Courier New -> Liberation Mono,
 *      and Helvetica -> Arial.
 *   5. When no directory holds any of those, fontconfig (fc-match), whose
 *      alias rules come with the installed font packages: Palatino
 *      Linotype -> P052 with fonts-urw-base35, SimSun -> Noto Serif CJK SC
 *      with fonts-noto-cjk. Its generic answer (the serif, sans-serif or
 *      monospace default) does not count, and a family that stands for an
 *      East Asian language gets fontconfig's face for that language.
 *   6. The fallback list, each family with the glyphs the ones before lack:
 *      Arial Unicode MS, DejaVu Sans, Noto Sans, Liberation Sans, Arial,
 *      Helvetica, Hiragino Sans GB, Apple Symbols, Symbol; then, for an
 *      East Asian character none has, fontconfig's face for its language.
 *   7. When none of those is installed either, the first family scanned.
 *
 * A family is missing (FontSet.missing) when only steps 6 and 7 drew it.
 *
 * A face the document embeds (FontSet.embed) comes before every rule: in
 * the asked style it is the family's face; when the document embeds the
 * family in other styles only, a directory's face in the asked style wins
 * and the embedded nearest face comes before the fallbacks.
 *
 * Within a family, the face with the asked bold and italic; when the family
 * lacks it, the nearest face and the missing style is synthesized (a fake
 * bold strokes the glyphs, a fake italic skews them). A glyph none of the
 * faces have draws as the primary face's .notdef.
 */
const PDF_SUBSTITUTES: Record<string, string> = {
  arial: "Liberation Sans",
  helvetica: "Arial",
  "times new roman": "Liberation Serif",
  "courier new": "Liberation Mono",
  courier: "Liberation Mono",
  "system-ui": "Helvetica Neue",
  "sans-serif": "Arial",
  serif: "Times New Roman",
  monospace: "Courier New",
};

const FALLBACK_FAMILIES = [
  "Arial Unicode MS",
  "DejaVu Sans",
  "Noto Sans",
  "Liberation Sans",
  "Arial",
  "Helvetica",
  "Hiragino Sans GB",
  "Apple Symbols",
  "Symbol",
];

const FONT_EXTENSIONS = new Set([".ttf", ".otf", ".ttc", ".woff"]);

/**
 * The directories a deployment adds to every render: the WORDINWEB_FONT_DIRS
 * environment variable, a list of paths separated as PATH is (":" here, ";"
 * on Windows). They are scanned ahead of the defaults, so a font file put
 * in one wins over the substitutes for its family.
 */
export function environmentFontDirs(): string[] {
  const value = process.env.WORDINWEB_FONT_DIRS ?? "";
  return value.split(process.platform === "win32" ? ";" : ":").map((dir) => dir.trim()).filter((dir) => dir && existsSync(dir));
}

/** The directories scanned when a renderer names none: the environment's, then the system's. */
export function defaultFontDirs(): string[] {
  const dirs = [
    ...environmentFontDirs(),
    "/System/Library/Fonts",
    "/Library/Fonts",
    join(homedir(), "Library/Fonts"),
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    join(homedir(), ".fonts"),
    join(homedir(), ".local/share/fonts"),
  ];
  // Carlito and Caladea as @fontsource ships them, when the package is near.
  const require = createRequire(import.meta.url);
  for (const pkg of ["@fontsource/carlito/files/carlito-latin-400-normal.woff", "@fontsource/caladea/files/caladea-latin-400-normal.woff"]) {
    try {
      dirs.push(dirname(require.resolve(pkg)));
    } catch {
      // Not installed: Calibri and Cambria fall through to the next rule.
    }
  }
  return dirs.filter((dir) => existsSync(dir));
}

/** Every face under a list of directories, by family name. */
export class FontRegistry {
  private readonly byFamily = new Map<string, FaceInfo[]>();

  constructor(readonly dirs: string[]) {
    for (const dir of dirs) this.scan(dir, 0);
  }

  private static shared = new Map<string, FontRegistry>();

  /** One registry per directory list, so a process scans each list once. */
  static for(dirs: string[]): FontRegistry {
    const key = dirs.join("\0");
    let registry = FontRegistry.shared.get(key);
    if (!registry) {
      registry = new FontRegistry(dirs);
      FontRegistry.shared.set(key, registry);
    }
    return registry;
  }

  private scan(dir: string, depth: number): void {
    let names: string[];
    try {
      names = readdirSync(dir);
    } catch {
      return;
    }
    for (const name of names.sort()) {
      const path = join(dir, name);
      let stat;
      try {
        stat = statSync(path);
      } catch {
        continue;
      }
      if (stat.isDirectory()) {
        if (depth < 4) this.scan(path, depth + 1);
        continue;
      }
      if (!FONT_EXTENSIONS.has(extname(name).toLowerCase())) continue;
      let faces: FaceInfo[];
      try {
        faces = scanFaces(path);
      } catch {
        continue;
      }
      for (const face of faces) {
        const key = face.family.toLowerCase();
        const list = this.byFamily.get(key);
        if (list) list.push(face);
        else this.byFamily.set(key, [face]);
      }
    }
    // A family split into Unicode subsets (@fontsource's Carlito) overlaps
    // on the basic Latin letters; the latin file goes first so Latin text
    // takes its glyphs from one file.
    for (const list of this.byFamily.values()) list.sort((a, b) => subsetRank(a.path) - subsetRank(b.path));
  }

  /** Some family, for text no asked family or fallback can show. */
  any(): FaceInfo[] | undefined {
    return this.byFamily.values().next().value;
  }

  /**
   * The faces of a family nearest the asked style: every file with the
   * exact style (a family split into Unicode subsets has several), else the
   * same slant, else the uprights, else whatever the family has.
   */
  faces(family: string, bold: boolean, italic: boolean): StyledFaces | undefined {
    const all = this.byFamily.get(family.toLowerCase());
    return all && nearestStyle(all, bold, italic);
  }
}

/** The faces of a family nearest a style, and the styles they must fake. */
export interface StyledFaces {
  faces: FaceInfo[];
  fakeBold: boolean;
  fakeItalic: boolean;
}

/** A family's faces nearest the asked style (see FontRegistry.faces). */
function nearestStyle(all: FaceInfo[], bold: boolean, italic: boolean): StyledFaces {
  const exact = all.filter((face) => face.bold === bold && face.italic === italic);
  if (exact.length > 0) return { faces: exact, fakeBold: false, fakeItalic: false };
  const slant = all.filter((face) => face.italic === italic && !face.bold);
  if (slant.length > 0) return { faces: slant, fakeBold: bold, fakeItalic: false };
  const upright = all.filter((face) => !face.italic && !face.bold);
  if (upright.length > 0) return { faces: upright, fakeBold: bold, fakeItalic: italic };
  return { faces: all, fakeBold: bold && !all[0].bold, fakeItalic: italic && !all[0].italic };
}

/**
 * The font file a document embeds, deobfuscated: ECMA-376 (Part 1 §17.8.1)
 * XORs the file's first 32 bytes with the w:fontKey GUID's 16 bytes, the
 * GUID's hex pairs as written but in reverse order (the last pair against
 * the first byte), the 16 repeated for the second 16 bytes. The rest of the
 * file is untouched.
 */
export function deobfuscateFont(bytes: Uint8Array, fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}-]/g, "");
  if (!/^[0-9a-fA-F]{32}$/.test(hex)) throw new Error(`malformed font key ${fontKey}`);
  const key = Buffer.from(hex, "hex");
  const out = bytes.slice();
  for (let i = 0; i < Math.min(32, out.length); i++) out[i] ^= key[15 - (i % 16)];
  return out;
}

function subsetRank(path: string): number {
  const name = path.toLowerCase();
  return /-latin-\d/.test(name) ? 0 : name.includes("-latin-ext-") ? 1 : 2;
}

/** A face in a resolved stack, with the styles it must fake. */
export interface StackFace {
  face: Face;
  fakeBold: boolean;
  fakeItalic: boolean;
}

/** The faces a FontSpec resolves to, first the primary. */
export interface FontStack {
  faces: StackFace[];
  /** The family that resolved (after substitution), for reports. */
  family: string;
  /** The East Asian language and style the asked family stands for, when it does. */
  eastAsian?: EastAsian;
}

interface EastAsian {
  lang: "ja" | "zh-cn" | "zh-tw" | "ko";
  /** A Song, Mincho, Batang or MingLiU face; else a Gothic, Hei or sans one. */
  serif: boolean;
}

/** The language and style an East Asian face name stands for: Mincho,
 * Gothic, Meiryo and Yu are Japanese; SimSun, SimHei, YaHei, Songti, Heiti
 * and PingFang SC simplified Chinese; JhengHei, MingLiU and PingFang TC
 * traditional; Batang, Gulim, Dotum and Malgun Korean. */
function eastAsianOf(family: string | undefined): EastAsian | undefined {
  if (!family) return undefined;
  const f = family.toLowerCase();
  const serif = /mincho|ming|song|simsun|batang|kai/.test(f);
  // Korean before Japanese: Malgun Gothic and Apple SD Gothic are Korean, MS Gothic and Yu Gothic Japanese.
  if (/batang|gulim|dotum|malgun|nanum|apple sd gothic/.test(f)) return { lang: "ko", serif };
  if (/jhenghei|mingliu|pingfang tc|hiragino sans tc/.test(f)) return { lang: "zh-tw", serif };
  if (/simsun|simhei|yahei|songti|heiti|pingfang|hiragino sans gb|stsong|stheiti|kaiti|fangsong/.test(f)) return { lang: "zh-cn", serif };
  if (/mincho|gothic|meiryo|^yu |hiragino sans$|hiragino kaku/.test(f)) return { lang: "ja", serif };
  return undefined;
}

/** The language a lone East Asian code point most likely belongs to. */
function eastAsianLangOf(cp: number): EastAsian["lang"] | undefined {
  if ((cp >= 0x3040 && cp <= 0x30ff) || (cp >= 0x31f0 && cp <= 0x31ff)) return "ja";
  if ((cp >= 0xac00 && cp <= 0xd7af) || (cp >= 0x1100 && cp <= 0x11ff) || (cp >= 0x3130 && cp <= 0x318f)) return "ko";
  if ((cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef) || (cp >= 0x20000 && cp <= 0x2ffff)) return "zh-cn";
  return undefined;
}

/**
 * fontconfig, asked where the rules found nothing: a face for a family
 * (its own alias rules decide, so a font package installed with its rules
 * works unchanged), or for an East Asian language. Each distinct question
 * is asked once per process; fontconfig's generic answers (what it gives
 * for serif, sans-serif, monospace or a family that cannot exist) count as
 * not found. Absent fc-match, every answer is null.
 */
export class Fontconfig {
  private available: boolean | null = null;
  private readonly answers = new Map<string, FaceInfo | null>();
  private readonly generic = new Map<string, Set<string>>();

  /** fontconfig's face for a family and style; null for its generic answer,
   * unless a rule aliased the family to that face. */
  match(family: string, bold: boolean, italic: boolean): FaceInfo | null {
    const style = `weight=${bold ? 200 : 80}:slant=${italic ? 100 : 0}`;
    const key = `${family.toLowerCase()}:${style}`;
    let answer = this.answers.get(key);
    if (answer !== undefined) return answer;
    answer = null;
    const found = this.ask(`${escape(family)}:${style}`);
    if (found && (!this.genericFiles(style).has(found.path) || this.aliased(family).includes(found.family.toLowerCase()))) answer = found;
    this.answers.set(key, answer);
    return answer;
  }

  /**
   * The families a rule put ahead of the generic chain for `family`: what
   * `fc-pattern -c` lists after config substitution, less the family itself,
   * up to where one of the generic chains (serif, sans-serif, monospace)
   * begins. Empty when fontconfig only classed the family as a generic.
   */
  private aliased(family: string): string[] {
    const chain = this.chain(escape(family)).filter((name) => name !== family.toLowerCase());
    let end = chain.length;
    for (const generic of ["serif", "sans-serif", "monospace"]) {
      const head = this.chain(generic)[0];
      const at = head ? chain.indexOf(head) : -1;
      if (at !== -1) end = Math.min(end, at);
    }
    return chain.slice(0, end);
  }

  private readonly chains = new Map<string, string[]>();

  private chain(pattern: string): string[] {
    let chain = this.chains.get(pattern);
    if (chain === undefined) {
      chain = this.available === false ? [] : this.run("fc-pattern", ["-c", "-f", "%{family}", pattern]).toLowerCase().split(",").map((name) => name.trim()).filter(Boolean);
      this.chains.set(pattern, chain);
    }
    return chain;
  }

  /** fontconfig's serif or sans face for an East Asian language and style. */
  matchLang(eastAsian: EastAsian, bold: boolean, italic: boolean): FaceInfo | null {
    const key = `${eastAsian.serif ? "serif" : "sans-serif"}:lang=${eastAsian.lang}:weight=${bold ? 200 : 80}:slant=${italic ? 100 : 0}`;
    let answer = this.answers.get(key);
    if (answer === undefined) {
      answer = this.ask(key);
      this.answers.set(key, answer);
    }
    return answer;
  }

  /** The files fontconfig answers with for the generic families and for a family that cannot exist. */
  private genericFiles(style: string): Set<string> {
    let files = this.generic.get(style);
    if (!files) {
      files = new Set();
      for (const family of ["sans-serif", "serif", "monospace", "wordinweb-no-such-family"]) {
        const found = this.ask(`${family}:${style}`);
        if (found) files.add(found.path);
      }
      this.generic.set(style, files);
    }
    return files;
  }

  private ask(pattern: string): FaceInfo | null {
    const out = this.run("fc-match", ["-f", "%{family[0]}\n%{file}\n%{index}\n%{weight}\n%{slant}", pattern]);
    const [family, file, index, weight, slant] = out.split("\n");
    if (!family || !file || !existsSync(file)) return null;
    return { path: file, index: Number(index) || 0, family, subfamily: "", postscriptName: family.replace(/\s+/g, ""), bold: Number(weight) >= 200, italic: Number(slant) > 0 };
  }

  /** A fontconfig tool's output, or "" when it is not installed (asked once). */
  private run(tool: string, args: string[]): string {
    if (this.available === false) return "";
    try {
      const out = execFileSync(tool, args, { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });
      this.available = true;
      return out;
    } catch {
      if (this.available === null) this.available = false;
      return "";
    }
  }
}

/** A family name in a fontconfig pattern: its separators escaped. */
function escape(family: string): string {
  return family.replace(/([\\:,-])/g, "\\$1");
}

/** The process's one fontconfig, so each question is asked once. */
export const fontconfig = new Fontconfig();

/** Glyphs of one face in a row, as the drawing layer shows them. */
export interface GlyphRun {
  face: StackFace;
  gids: number[];
  /** Each glyph's advance in px, kerning applied, letter spacing not. */
  advances: number[];
  /** The code point each glyph came from, for the ToUnicode map. */
  codePoints: number[];
}

/**
 * The faces a renderer loads, resolved per FontSpec and cached: the layout
 * measures with the same files the PDF embeds, so a line breaks where its
 * glyphs end.
 */
export class FontSet {
  private readonly registry: FontRegistry;
  private readonly loaded = new Map<string, Face>();
  private readonly stacks = new Map<string, FontStack>();
  /** The faces the document embeds, by the family it names (lower case). */
  private readonly embedded = new Map<string, FaceInfo[]>();
  /** Families the documents asked for that no directory holds. */
  readonly missing = new Set<string>();
  /** Code points no face in their stack has: each drew as an empty box. */
  readonly missingCodePoints = new Set<number>();

  constructor(dirs: string[] = defaultFontDirs()) {
    this.registry = FontRegistry.for(dirs);
  }

  /** Add a face the document embeds for `family` in a style, `bytes` its
   * plain sfnt (deobfuscateFont for an obfuscated one) and `part` the
   * package part it came from. Before any text is measured. */
  embed(family: string, bold: boolean, italic: boolean, bytes: Uint8Array, part: string): void {
    const face = new Face(part, 0, bytes);
    this.loaded.set(`${part}#0`, face);
    const info: FaceInfo = { path: part, index: 0, family, subfamily: face.subfamily, postscriptName: face.postscriptName, bold, italic };
    const key = family.toLowerCase();
    const list = this.embedded.get(key);
    if (list) list.push(info);
    else this.embedded.set(key, [info]);
  }

  /** The embedded faces nearest the style for the first of `families` the document embeds. */
  private embeddedFaces(families: string[], bold: boolean, italic: boolean): { family: string; found: StyledFaces } | undefined {
    for (const family of families) {
      const own = this.embedded.get(family.toLowerCase());
      if (own) return { family, found: nearestStyle(own, bold, italic) };
    }
    return undefined;
  }

  resolve(font: FontSpec): FontStack {
    const n = normalizeFamily(font.family, !!font.bold, !!font.italic);
    const key = `${n.bold ? "b" : ""}${n.italic ? "i" : ""}|${font.paintFamily ?? ""}|${n.family}`;
    let stack = this.stacks.get(key);
    if (stack) return stack;
    const candidates: string[] = [];
    const add = (family: string | undefined): void => {
      if (family && !candidates.some((c) => c.toLowerCase() === family.toLowerCase())) candidates.push(family);
    };
    add(font.paintFamily);
    add(n.family);
    add(substituteFamily(n.family));
    for (const family of [...candidates]) add(PDF_SUBSTITUTES[family.toLowerCase()]);
    const primaryCount = candidates.length;
    for (const family of FALLBACK_FAMILIES) add(family);
    const faces: StackFace[] = [];
    let resolved = "";
    // The document's own faces lead the stack: in the asked style ahead of
    // every directory; in another style only ahead of the fallbacks.
    const own = this.embeddedFaces(candidates.slice(0, primaryCount), n.bold, n.italic);
    const ownFaces = own ? own.found.faces.map((info): StackFace => ({ face: this.load(info), fakeBold: own.found.fakeBold, fakeItalic: own.found.fakeItalic })) : [];
    if (own && own.found.faces[0].bold === n.bold && own.found.faces[0].italic === n.italic) {
      resolved = own.family;
      faces.push(...ownFaces);
    }
    candidates.forEach((family, i) => {
      const found = this.registry.faces(family, n.bold, n.italic);
      if (!found) return;
      if (!resolved && i < primaryCount) resolved = family;
      for (const info of found.faces) faces.push({ face: this.load(info), fakeBold: found.fakeBold, fakeItalic: found.fakeItalic });
    });
    if (own && !resolved) {
      resolved = own.family;
      faces.unshift(...ownFaces);
    }
    const eastAsian = eastAsianOf(font.paintFamily) ?? eastAsianOf(n.family);
    if (!resolved) {
      // The rules found none of the asked families: fontconfig's substitute
      // for them, if it has one, else a face for the East Asian language
      // the family stands for; either leads the stack.
      let found: FaceInfo | null = null;
      for (const family of candidates.slice(0, primaryCount)) found ??= fontconfig.match(family, n.bold, n.italic);
      if (!found && eastAsian) found = fontconfig.matchLang(eastAsian, n.bold, n.italic);
      if (found) {
        resolved = found.family;
        faces.unshift({ face: this.load(found), fakeBold: n.bold && !found.bold, fakeItalic: n.italic && !found.italic });
      }
    }
    if (!resolved) this.missing.add(n.family);
    if (faces.length === 0) {
      // No asked family and no fallback: the first family the directories hold.
      const any = this.registry.any();
      if (!any) throw new Error(`no fonts in ${this.registry.dirs.join(", ")}`);
      for (const info of any) faces.push({ face: this.load(info), fakeBold: n.bold && !info.bold, fakeItalic: n.italic && !info.italic });
    }
    stack = { faces, family: resolved || faces[0].face.family, eastAsian };
    this.stacks.set(key, stack);
    return stack;
  }

  /** A face from fontconfig for the language a code point's script belongs
   * to (the stack's own language first), when the stack has no glyph for
   * it; the face joins the stack for the glyphs after it. */
  private langFace(stack: FontStack, cp: number): StackFace | null {
    const lang = eastAsianLangOf(cp);
    const eastAsian = stack.eastAsian ?? (lang ? { lang, serif: false } : undefined);
    if (!eastAsian) return null;
    const found = fontconfig.matchLang(eastAsian, false, false);
    if (!found) return null;
    const face = this.load(found);
    if (!face.hasGlyph(cp)) return null;
    const entry = { face, fakeBold: false, fakeItalic: false };
    stack.faces.push(entry);
    return entry;
  }

  private load(info: FaceInfo): Face {
    const key = `${info.path}#${info.index}`;
    let face = this.loaded.get(key);
    if (!face) {
      face = new Face(info.path, info.index);
      this.loaded.set(key, face);
    }
    return face;
  }

  /**
   * The glyphs of `text` in `font`, each code point from the first face of
   * the stack that has it, in runs per face. Pair kerning applies when the
   * run asked for it (FontSpec.kerning). A soft hyphen takes no glyph.
   *
   * Every string the renderer measures or draws comes through here, so
   * this is where text in a script with no shaper is refused for good:
   * render.ts refuses the document's text first, naming where it is.
   */
  shape(text: string, font: FontSpec): GlyphRun[] {
    const script = unsupportedScript(text);
    if (script) throw new PdfRenderError("unsupported_script", `${script} text "${text.trim().slice(0, 40)}" needs a shaper this renderer does not have`);
    const stack = this.resolve(font);
    const scale = font.size * fontWidthScale(font);
    const runs: GlyphRun[] = [];
    let run: GlyphRun | null = null;
    let previous = -1;
    for (const ch of text) {
      const cp = ch.codePointAt(0) as number;
      if (cp === 0xad) continue;
      // Tabs and line breaks have no glyph; the layout sizes them itself.
      const shown = cp === 9 || cp === 10 || cp === 13 ? 32 : cp;
      let chosen = stack.faces[0];
      let gid = 0;
      for (const candidate of stack.faces) {
        const id = candidate.face.glyphId(shown);
        if (id !== 0) {
          chosen = candidate;
          gid = id;
          break;
        }
      }
      if (gid === 0 && shown !== 32) {
        const extra = this.langFace(stack, shown);
        if (extra) {
          chosen = extra;
          gid = extra.face.glyphId(shown);
        } else this.missingCodePoints.add(cp);
      }
      if (!run || run.face !== chosen) {
        run = { face: chosen, gids: [], advances: [], codePoints: [] };
        runs.push(run);
        previous = -1;
      }
      const face = chosen.face;
      const unit = scale / face.unitsPerEm;
      if (font.kerning && previous !== -1) {
        const kern = face.kern(previous, gid);
        if (kern) run.advances[run.advances.length - 1] += kern * unit;
      }
      run.gids.push(gid);
      run.advances.push(face.advance(gid) * unit);
      run.codePoints.push(cp);
      previous = gid;
    }
    return runs;
  }
}
