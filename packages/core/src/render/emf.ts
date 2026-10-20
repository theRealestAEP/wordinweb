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
 */

const EMR_HEADER = 1;
const EMR_POLYGON = 3;
const EMR_EOF = 14;
const EMR_CREATEPEN = 38;
const EMR_EXTCREATEPEN = 95;
const BS_NULL = 1;
const PS_NULL = 5;
const EMR_BEGINPATH = 59;
const EMR_ENDPATH = 60;
const EMR_FILLPATH = 62;
const EMR_BITBLT = 76;
const EMR_STRETCHBLT = 77;
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

/** The EMF with its pattern blits made fills; the same bytes when it has none. */
export function prepareEmf(bytes: Uint8Array): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 88 || view.getUint32(0, true) !== EMR_HEADER) return bytes;
  const parts: Uint8Array[] = [];
  let changed = false;
  let records = 0;
  let offset = 0;
  while (offset + 8 <= bytes.length) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) break;
    // xDest, yDest, cxDest, cyDest, dwRop follow the bounds; cbBitsSrc is the
    // last field both records share.
    if ((type === EMR_BITBLT || type === EMR_STRETCHBLT) && size >= 100 && view.getUint32(offset + 40, true) === PATCOPY && view.getUint32(offset + 96, true) === 0) {
      const x = view.getInt32(offset + 24, true);
      const y = view.getInt32(offset + 28, true);
      const fill = fillRect(x, y, x + view.getInt32(offset + 32, true), y + view.getInt32(offset + 36, true));
      parts.push(...fill);
      records += fill.length;
      changed = true;
    } else if (type === EMR_EXTCREATEPEN && size >= 48) {
      const pen = record(EMR_CREATEPEN, 28);
      const brushNull = view.getUint32(offset + 36, true) === BS_NULL;
      pen.view.setUint32(8, view.getUint32(offset + 8, true), true);
      pen.view.setUint32(12, brushNull ? PS_NULL : view.getUint32(offset + 28, true), true);
      pen.view.setUint32(16, view.getUint32(offset + 32, true), true);
      pen.view.setUint32(24, view.getUint32(offset + 40, true), true);
      parts.push(pen.bytes);
      records++;
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
