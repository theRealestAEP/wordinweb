import type { LayoutResult } from "../layout/types.js";
import { metricCompatibleSubstitute, normalizeFamily } from "../layout/measure.js";

/** Families the browser always resolves; never worth a warning. */
const GENERIC = new Set([
  "serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui",
  "ui-serif", "ui-sans-serif", "ui-monospace",
]);

export interface MissingFont {
  /** The face the document asked for (first name in the item's stack). */
  family: string;
  /** A short sample of text that will render in a substitute. */
  sample: string;
}

/**
 * Report document-requested font faces the browser cannot actually render —
 * the cases where the page silently falls back to a substitute and may look
 * different from Word. Two independent checks per face:
 *
 * 1. Face availability: measure a Latin probe in `"<face>", monospace` vs
 *    bare monospace AND `"<face>", serif` vs bare serif. If the face is
 *    missing, each stack collapses to its generic and the two widths differ
 *    exactly like the generics do; if present, both stacks measure the face.
 * 2. Glyph coverage: even an available face can lack the document's script
 *    (Tahoma has no Lao). Measure the item's own text against the tofu width —
 *    when every glyph draws as .notdef in both the face and the generic, the
 *    text is falling through to a system fallback the layout never measured.
 *
 * Runs on the live DOM (canvas measurement), so call it after fonts settle
 * (document.fonts.ready) for stable answers.
 */
export function detectMissingFonts(layout: LayoutResult): MissingFont[] {
  if (typeof document === "undefined") return [];
  const canvas = document.createElement("canvas");
  const g = canvas.getContext("2d");
  if (!g) return [];

  // One representative text per primary face.
  const samples = new Map<string, string>();
  for (const sample of layout._fontSamples ?? []) {
    const requested = sample.font.family.split(",")[0].trim().replace(/^"|"$/g, "");
    const primary = normalizeFamily(requested, sample.font.bold, sample.font.italic).family;
    if (!primary || GENERIC.has(primary.toLowerCase())) continue;
    const existing = samples.get(primary);
    if (!existing || (!/[^\u0000-\u024f]/.test(existing) && /[^\u0000-\u024f]/.test(sample.text))) {
      samples.set(primary, sample.text);
    }
  }
  for (const page of layout.pages) {
    for (const item of page.items) {
      if (item.kind !== "text" || !item.text.trim()) continue;
      const requested = item.font.family.split(",")[0].trim().replace(/^"|"$/g, "");
      const primary = normalizeFamily(requested, item.font.bold, item.font.italic).family;
      if (!primary || GENERIC.has(primary.toLowerCase())) continue;
      const existing = samples.get(primary);
      // Prefer a sample with non-Latin content — that's where substitution shows.
      if (!existing || (!/[^\u0000-ɏ]/.test(existing) && /[^\u0000-ɏ]/.test(item.text))) {
        samples.set(primary, item.text.trim().slice(0, 40));
      }
    }
  }

  const missing: MissingFont[] = [];
  const probe = "mmmWWWiill178%&";
  for (const [family, sample] of samples) {
    const width = (font: string) => {
      g.font = font;
      return g.measureText(probe).width;
    };
    const mono = width("32px monospace");
    const serif = width("32px serif");
    const viaMono = width(`32px "${family}", monospace`);
    const viaSerif = width(`32px "${family}", serif`);
    const faceMissing = viaMono === mono && viaSerif === serif && mono !== serif;
    if (faceMissing) {
      // Carlito/Caladea carry Calibri/Cambria's exact advances: with one of
      // them loaded the page breaks lines where Word does, which is the whole
      // point of the warning — so a loaded substitute counts as the face.
      const substitute = metricCompatibleSubstitute(family);
      const substitutePresent = !!substitute &&
        !(width(`32px "${substitute}", monospace`) === mono && width(`32px "${substitute}", serif`) === serif);
      if (!substitutePresent) missing.push({ family, sample });
      continue;
    }
    // Coverage check on the document's own sample: compare the sample's width
    // in the face against the same string with every char replaced by U+0000
    // (canvas draws unsupported chars via font fallback, so a face lacking the
    // script measures IDENTICALLY through fallback in both stacks — flag only
    // clear tofu, where the face itself would draw .notdef boxes).
    const nonLatin = sample.replace(/[\u0000-ɏ]+/g, "");
    if (nonLatin.length >= 3) {
      g.font = `32px "${family}"`;
      const inFace = g.measureText(nonLatin).width;
      g.font = "32px monospace";
      const inMono = g.measureText(nonLatin).width;
      g.font = "32px serif";
      const inSerif = g.measureText(nonLatin).width;
      // All three equal => every glyph came from the same system fallback,
      // i.e. the face contributed nothing for this script.
      if (inFace === inMono && inFace === inSerif) missing.push({ family, sample: nonLatin.slice(0, 20) });
    }
  }
  return missing;
}
