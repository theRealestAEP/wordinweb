import { GlyphRun, ShapedRun } from "./fonts.js";
import { EmbeddedFont, utf16Hex } from "./resources.js";
import { num } from "./writer.js";

/**
 * Show a shaped run at each glyph's pen position plus its HarfBuzz offsets.
 * Distances are in unscaled px; the caller's text matrix maps them to its
 * coordinate space, with positive y offsets pointing above the baseline.
 * Returns the pen advance before `stretch` (the caller's PDF Tz).
 * Spacing goes after each cluster, or into stretched kashida at the places
 * HarfBuzz marked. Each cluster's ActualText covers its glyphs and kashida.
 */
export function showShaped(ops: string[], run: GlyphRun, shaped: ShapedRun, font: EmbeddedFont, size: number, spacing: number, stretch: number, widthScale: number, tm: (pen: number, yOffset: number) => string): number {
  const face = run.face.face;
  const unit = size / face.unitsPerEm;
  const places = shaped.tatweelAfter.filter(Boolean).length;
  const tatweel = spacing > 0 && places > 0 ? face.glyphId(0x0640) : 0;
  const tatweelWidth = tatweel ? face.advance(tatweel) * unit : 0;
  const kashida = tatweelWidth > 0 ? (spacing * shaped.text.length) / places : 0;
  let pen = 0;
  /** Where the next glyph shows unless shifted, while a TJ is open. */
  let at = 0;
  let open = false;
  let parts: string[] = [];
  let hex = "";
  const flush = (): void => {
    if (hex) parts.push(`<${hex}>`);
    hex = "";
    if (parts.length) ops.push(`[${parts.join(" ")}] TJ`);
    parts = [];
    open = false;
  };
  for (let i = 0; i < run.gids.length; i++) {
    // Text extractors reverse RTL text from its visual order, including
    // the characters within an ActualText span.
    const unicode = shaped.rtl ? [...shaped.unicode[i]].reverse().join("") : shaped.unicode[i];
    if (i === 0 || shaped.units[i - 1] > 0) ops.push(`/Span << /ActualText <feff${utf16Hex(unicode)}> >> BDC`);
    const gid = run.gids[i];
    const cid = font.cid(gid, unicode).toString(16).padStart(4, "0");
    const xOffset = shaped.xOffsets[i] / widthScale;
    const yOffset = shaped.yOffsets[i];
    if (yOffset !== 0) {
      flush();
      ops.push(tm(pen + xOffset, yOffset), `<${cid}> Tj`);
    } else {
      if (!open) {
        ops.push(tm(pen + xOffset, 0));
        at = pen + xOffset;
        open = true;
      }
      const shift = pen + xOffset - at;
      if (Math.abs(shift) > 0.001) {
        if (hex) parts.push(`<${hex}>`);
        hex = "";
        parts.push(num((-shift / size) * 1000));
      }
      hex += cid;
      at = pen + xOffset + face.advance(gid) * unit;
    }
    pen += run.advances[i] / widthScale;
    if (kashida && shaped.tatweelAfter[i]) {
      flush();
      ops.push(`${num(stretch * 100 * (kashida / tatweelWidth))} Tz`, tm(pen, 0), `<${font.cid(tatweel, 0).toString(16).padStart(4, "0")}> Tj`, `${num(stretch * 100)} Tz`);
      pen += kashida;
    } else if (!kashida) pen += spacing * shaped.units[i];
    if (shaped.units[i] > 0) {
      flush();
      ops.push("EMC");
    }
  }
  flush();
  return pen;
}

