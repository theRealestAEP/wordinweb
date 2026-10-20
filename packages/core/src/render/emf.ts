/**
 * EMF records rewritten into ones emf-converter draws ([MS-EMF] 2.3).
 *
 * A BITBLT or STRETCHBLT with no source bitmap and the PATCOPY raster
 * operation fills its destination rectangle with the current brush (a
 * PatBlt). Word writes page backgrounds, bands and gradient steps this way;
 * emf-converter only draws the blits that carry a bitmap. Each such record
 * becomes BEGINPATH, a POLYGON of the rectangle, ENDPATH and FILLPATH,
 * which fill with the current brush and draw no outline.
 *
 * emf-converter reads an EXTCREATEPEN's style, width and color 8 bytes
 * early (from the bitmap fields before the LogPenEx), so every such pen
 * drew black and 1 unit wide, a null pen included. Each becomes the
 * CREATEPEN of the same style, width and color; a pen whose brush is
 * BS_NULL becomes a PS_NULL pen. Hatched and pattern pens draw solid.
 *
 * A pen's width is in logical units, but emf-converter strokes with it as
 * canvas pixels. Each pen's width is scaled here by the viewport extent
 * over the window extent as they stand where the pen is made, the mapping
 * emf-converter applies to coordinates (it keeps no extents across
 * SaveDC/RestoreDC, and neither does this).
 *
 * An EMF that also carries EMF+ records is left as it is: emf-converter
 * draws its GDI records as well, without their world transform, so drawn
 * right they would land over the EMF+ picture at the wrong scale.
 */

const EMR_HEADER = 1;
const EMR_POLYGON = 3;
const EMR_SETWINDOWEXTEX = 9;
const EMR_SETVIEWPORTEXTEX = 11;
const EMR_EOF = 14;
const EMR_SCALEVIEWPORTEXTEX = 31;
const EMR_SCALEWINDOWEXTEX = 32;
const EMR_CREATEPEN = 38;
const EMR_COMMENT = 70;
const EMFPLUS = 0x2b464d45;
const EMR_BEGINPATH = 59;
const EMR_ENDPATH = 60;
const EMR_FILLPATH = 62;
const EMR_BITBLT = 76;
const EMR_STRETCHBLT = 77;
const EMR_EXTCREATEPEN = 95;
const BS_NULL = 1;
const PS_NULL = 5;
const PATCOPY = 0x00f00021;

function record(type: number, size: number): { bytes: Uint8Array; view: DataView } {
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, type, true);
  view.setUint32(4, size, true);
  return { bytes, view };
}

/** The records that fill the rectangle with the current brush. */
function fillRect(left: number, top: number, right: number, bottom: number): Uint8Array[] {
  const polygon = record(EMR_POLYGON, 60);
  const corners = [left, top, right, top, right, bottom, left, bottom];
  [left, top, right, bottom].forEach((value, i) => polygon.view.setInt32(8 + i * 4, value, true));
  polygon.view.setUint32(24, 4, true);
  corners.forEach((value, i) => polygon.view.setInt32(28 + i * 4, value, true));
  const fill = record(EMR_FILLPATH, 24);
  [left, top, right, bottom].forEach((value, i) => fill.view.setInt32(8 + i * 4, value, true));
  return [record(EMR_BEGINPATH, 8).bytes, polygon.bytes, record(EMR_ENDPATH, 8).bytes, fill.bytes];
}

/** A CREATEPEN record. */
function createPen(index: number, style: number, width: number, color: number): Uint8Array {
  const pen = record(EMR_CREATEPEN, 28);
  pen.view.setUint32(8, index, true);
  pen.view.setUint32(12, style, true);
  pen.view.setInt32(16, width, true);
  pen.view.setUint32(24, color, true);
  return pen.bytes;
}

/** Whether an EMR_COMMENT carries EMF+ records ([MS-EMFPLUS] 2.3.3: identifier "EMF+"). */
function hasEmfPlus(view: DataView): boolean {
  let offset = 0;
  while (offset + 16 <= view.byteLength) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (type === EMR_COMMENT && view.getUint32(offset + 12, true) === EMFPLUS) return true;
    if (size < 8 || type === EMR_EOF) return false;
    offset += size;
  }
  return false;
}

/** The EMF with the records above rewritten; the same bytes when it has none. */
export function prepareEmf(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 88 || view.getUint32(0, true) !== EMR_HEADER || hasEmfPlus(view)) return bytes;
  // emf-converter starts both extents at the bounds' width: one unit, one pixel.
  let windowExt = view.getInt32(16, true) - view.getInt32(8, true);
  let viewportExt = windowExt;
  const scaled = (width: number): number => (windowExt ? Math.round(Math.abs((width * viewportExt) / windowExt)) : width);
  const parts: Uint8Array[] = [];
  let changed = false;
  let records = 0;
  let offset = 0;
  while (offset + 8 <= bytes.length) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) break;
    let replacement: Uint8Array[] | null = null;
    if (type === EMR_SETWINDOWEXTEX && size >= 16) windowExt = view.getInt32(offset + 8, true);
    else if (type === EMR_SETVIEWPORTEXTEX && size >= 16) viewportExt = view.getInt32(offset + 8, true);
    else if ((type === EMR_SCALEWINDOWEXTEX || type === EMR_SCALEVIEWPORTEXTEX) && size >= 24 && view.getInt32(offset + 12, true) !== 0) {
      const ext = Math.round(((type === EMR_SCALEWINDOWEXTEX ? windowExt : viewportExt) * view.getInt32(offset + 8, true)) / view.getInt32(offset + 12, true));
      if (type === EMR_SCALEWINDOWEXTEX) windowExt = ext;
      else viewportExt = ext;
    } else if ((type === EMR_BITBLT || type === EMR_STRETCHBLT) && size >= 100 && view.getUint32(offset + 40, true) === PATCOPY && view.getUint32(offset + 96, true) === 0) {
      // xDest, yDest, cxDest, cyDest, dwRop after the bounds; cbBitsSrc is the last field both records share.
      const x = view.getInt32(offset + 24, true);
      const y = view.getInt32(offset + 28, true);
      replacement = fillRect(x, y, x + view.getInt32(offset + 32, true), y + view.getInt32(offset + 36, true));
    } else if (type === EMR_CREATEPEN && size >= 28) {
      const width = view.getInt32(offset + 16, true);
      if (scaled(width) !== width) replacement = [createPen(view.getUint32(offset + 8, true), view.getUint32(offset + 12, true), scaled(width), view.getUint32(offset + 24, true))];
    } else if (type === EMR_EXTCREATEPEN && size >= 48) {
      // ihPen, the bitmap fields, then LogPenEx: PenStyle, Width, BrushStyle, Color.
      const style = view.getUint32(offset + 36, true) === BS_NULL ? PS_NULL : view.getUint32(offset + 28, true);
      replacement = [createPen(view.getUint32(offset + 8, true), style, scaled(view.getInt32(offset + 32, true)), view.getUint32(offset + 40, true))];
    }
    if (replacement) {
      parts.push(...replacement);
      records += replacement.length;
      changed = true;
    } else {
      parts.push(bytes.subarray(offset, offset + size));
      records++;
    }
    offset += size;
    if (type === EMR_EOF) break;
  }
  if (!changed) return bytes;
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  const outView = new DataView(out.buffer);
  outView.setUint32(48, total, true);
  outView.setUint32(52, records, true);
  return out;
}
