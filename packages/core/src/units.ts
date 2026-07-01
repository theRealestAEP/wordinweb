/**
 * OOXML unit conversions. Everything in the layout engine is normalized to
 * CSS pixels at 96 dpi so the DOM renderer can position content 1:1.
 *
 *  - twip: 1/20 point (page sizes, margins, indents, spacing)
 *  - half-point: font sizes (w:sz)
 *  - eighth-point: border widths (w:sz on borders)
 *  - EMU: 914400 per inch (drawing/image extents)
 */

export const PX_PER_PT = 96 / 72; // 4/3
export const TWIPS_PER_PT = 20;
export const EMU_PER_PX = 914400 / 96; // 9525

export function twipsToPx(twips: number): number {
  return (twips / TWIPS_PER_PT) * PX_PER_PT;
}

export function ptToPx(pt: number): number {
  return pt * PX_PER_PT;
}

export function halfPtToPx(halfPt: number): number {
  return (halfPt / 2) * PX_PER_PT;
}

export function eighthPtToPx(eighthPt: number): number {
  return (eighthPt / 8) * PX_PER_PT;
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToTwips(px: number): number {
  return (px / PX_PER_PT) * TWIPS_PER_PT;
}
