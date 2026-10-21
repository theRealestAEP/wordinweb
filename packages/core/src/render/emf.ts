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
 * An EXTCREATEFONTINDIRECTW's height and width are also in logical units,
 * but emf-converter uses the height as canvas pixels. Each font's height
 * is scaled by the y extents and its nonzero width by the x extents where
 * it is made. The height keeps its sign: negative for character height,
 * positive for cell height ([MS-EMF] 2.2.13, 2.3.7.8).
 *
 * When a window extent precedes any viewport extent, an explicit 1 x 1
 * viewport extent supplies the device context's default. emf-converter
 * otherwise uses the canvas size, multiplying coordinates and font sizes.
 *
 * EXTSELECTCLIPRGN rectangles in device coordinates (the flight-plan map
 * uses its viewport rectangle) go through emf-converter's logical mapping.
 * Each rectangle is rewritten through the inverse mapping, using the
 * current origins and extents, so the converter clips at its device position.
 * Before any mapping record, its bounds-relative device mapping already
 * places these rectangles correctly and they are kept as they are.
 * [MS-EMF] 2.3.2.2 calls RgnData logical units; these files store device units.
 *
 * An EMF that also carries EMF+ records gets none of the rewrites above
 * (emf-converter draws its EMF+ records, with their world transform), and
 * three of its own ([MS-EMFPLUS] 1.3.1, 2.1.1.2, 2.3.3.2, 2.3.4.6):
 *
 * - EMF+ Dual (the EmfPlusHeader's flag 0x0001): the EMF records draw the
 *   picture a second time, for a player that reads no EMF+. emf-converter
 *   played both, without the EMF records' world transform: labels twice,
 *   shapes at the wrong scale. Only the EMF records an EMF+ player plays
 *   are kept: those after an EmfPlusGetDC, up to the next EMF+ record.
 * - A DrawDriverString's glyph positions follow its glyphs directly, 2
 *   bytes off a 4-byte boundary when the glyph count is odd; emf-converter
 *   reads them from the next boundary, misread floats (a text origin of
 *   1e+32). Two bytes of padding go after such a record's glyphs.
 * - A bitmap image object's BitmapDataType is 0 for pixels and 1 for a
 *   compressed image (a PNG or a JPEG); emf-converter reads 1 as pixels and
 *   2 as compressed, so drew neither. Each is set to the value it reads.
 *   An object continued over several records is left as it is.
 */

const EMR_HEADER = 1;
const EMR_POLYGON = 3;
const EMR_SETWINDOWEXTEX = 9;
const EMR_SETWINDOWORGEX = 10;
const EMR_SETVIEWPORTEXTEX = 11;
const EMR_SETVIEWPORTORGEX = 12;
const EMR_EOF = 14;
const EMR_SETMAPMODE = 17;
const EMR_SCALEVIEWPORTEXTEX = 31;
const EMR_SCALEWINDOWEXTEX = 32;
const EMR_CREATEPEN = 38;
const EMR_COMMENT = 70;
const EMFPLUS = 0x2b464d45;
const EMFPLUS_HEADER = 0x4001;
const EMFPLUS_GETDC = 0x4004;
const EMFPLUS_OBJECT = 0x4008;
const EMFPLUS_DRAWDRIVERSTRING = 0x4036;
const EMFPLUS_DUAL = 0x0001;
const EMFPLUS_CONTINUED = 0x8000;
const OBJECT_IMAGE = 5;
const IMAGE_BITMAP = 1;
const GRAPHICS_VERSION_SIGNATURE = 0xdbc01;
const EMR_BEGINPATH = 59;
const EMR_ENDPATH = 60;
const EMR_FILLPATH = 62;
const EMR_EXTSELECTCLIPRGN = 75;
const EMR_BITBLT = 76;
const EMR_STRETCHBLT = 77;
const EMR_EXTCREATEFONTINDIRECTW = 82;
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
  if (bytes.length < 88 || view.getUint32(0, true) !== EMR_HEADER) return bytes;
  if (hasEmfPlus(view)) return prepareEmfPlus(bytes, view);
  // emf-converter starts the extents at the bounds' size: one unit, one pixel.
  let windowExtX = view.getInt32(16, true) - view.getInt32(8, true);
  let windowExtY = view.getInt32(20, true) - view.getInt32(12, true);
  let viewportExtX = windowExtX;
  let viewportExtY = windowExtY;
  let windowOrgX = view.getInt32(8, true);
  let windowOrgY = view.getInt32(12, true);
  let viewportOrgX = 0;
  let viewportOrgY = 0;
  let viewportExtentSet = false;
  let mapping = false;
  const scaled = (value: number, windowExt = windowExtX, viewportExt = viewportExtX): number => (windowExt ? Math.round(Math.abs((value * viewportExt) / windowExt)) : value);
  const parts: Uint8Array[] = [];
  let changed = false;
  let records = 0;
  let offset = 0;
  while (offset + 8 <= bytes.length) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) break;
    let replacement: Uint8Array[] | null = null;
    if (!viewportExtentSet && ((type === EMR_SETWINDOWEXTEX && size >= 16) || (type === EMR_SCALEWINDOWEXTEX && size >= 24))) {
      const extent = record(EMR_SETVIEWPORTEXTEX, 16);
      extent.view.setInt32(8, 1, true);
      extent.view.setInt32(12, 1, true);
      parts.push(extent.bytes);
      records++;
      changed = true;
      viewportExtX = viewportExtY = 1;
      viewportExtentSet = true;
    }
    if (type === EMR_SETWINDOWEXTEX && size >= 16) {
      mapping = true;
      windowExtX = view.getInt32(offset + 8, true);
      windowExtY = view.getInt32(offset + 12, true);
    } else if (type === EMR_SETVIEWPORTEXTEX && size >= 16) {
      mapping = true;
      viewportExtentSet = true;
      viewportExtX = view.getInt32(offset + 8, true);
      viewportExtY = view.getInt32(offset + 12, true);
    } else if (type === EMR_SETWINDOWORGEX && size >= 16) {
      mapping = true;
      windowOrgX = view.getInt32(offset + 8, true);
      windowOrgY = view.getInt32(offset + 12, true);
    } else if (type === EMR_SETVIEWPORTORGEX && size >= 16) {
      mapping = true;
      viewportOrgX = view.getInt32(offset + 8, true);
      viewportOrgY = view.getInt32(offset + 12, true);
    } else if (type === EMR_SETMAPMODE && size >= 12) {
      const mode = view.getUint32(offset + 8, true);
      if (mode === 7 || mode === 8) mapping = true;
    } else if ((type === EMR_SCALEWINDOWEXTEX || type === EMR_SCALEVIEWPORTEXTEX) && size >= 24) {
      mapping = true;
      if (type === EMR_SCALEVIEWPORTEXTEX) viewportExtentSet = true;
      if (view.getInt32(offset + 12, true) !== 0) {
        const ext = Math.round(((type === EMR_SCALEWINDOWEXTEX ? windowExtX : viewportExtX) * view.getInt32(offset + 8, true)) / view.getInt32(offset + 12, true));
        if (type === EMR_SCALEWINDOWEXTEX) windowExtX = ext;
        else viewportExtX = ext;
      }
      if (view.getInt32(offset + 20, true) !== 0) {
        const ext = Math.round(((type === EMR_SCALEWINDOWEXTEX ? windowExtY : viewportExtY) * view.getInt32(offset + 16, true)) / view.getInt32(offset + 20, true));
        if (type === EMR_SCALEWINDOWEXTEX) windowExtY = ext;
        else viewportExtY = ext;
      }
    } else if (type === EMR_EXTSELECTCLIPRGN && mapping && size >= 48 && view.getUint32(offset + 8, true) >= 32) {
      const region = bytes.slice(offset, offset + size);
      const out = new DataView(region.buffer);
      const count = view.getUint32(offset + 24, true);
      const end = Math.min(size, 16 + view.getUint32(offset + 8, true), 48 + count * 16);
      // RegionDataHeader.Bounds, then every RectL in RegionData.Data.
      for (let at = 32; at + 16 <= end; at += 16) {
        for (let edge = 0; edge < 4; edge++) {
          const x = edge % 2 === 0;
          const device = out.getInt32(at + edge * 4, true);
          const logical = x
            ? ((device - viewportOrgX) * (windowExtX || 1)) / (viewportExtX || 1) + windowOrgX
            : ((device - viewportOrgY) * (windowExtY || 1)) / (viewportExtY || 1) + windowOrgY;
          out.setInt32(at + edge * 4, Math.round(logical), true);
        }
      }
      replacement = [region];
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
    } else if (type === EMR_EXTCREATEFONTINDIRECTW && size >= 104) {
      // ihFonts, then LogFont: Height, Width, and the remaining attributes.
      const height = view.getInt32(offset + 12, true);
      const width = view.getInt32(offset + 16, true);
      const scaledHeight = Math.sign(height) * scaled(Math.abs(height), windowExtY, viewportExtY);
      const scaledWidth = Math.sign(width) * scaled(Math.abs(width));
      if (scaledHeight !== height || scaledWidth !== width) {
        const font = bytes.slice(offset, offset + size);
        const out = new DataView(font.buffer);
        out.setInt32(12, scaledHeight, true);
        out.setInt32(16, scaledWidth, true);
        replacement = [font];
      }
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
  return changed ? assemble(parts, records) : bytes;
}

/** An EMF with EMF+ records, as an EMF+ player plays it (see above); the same bytes when nothing changes. */
function prepareEmfPlus(bytes: Uint8Array, view: DataView): Uint8Array {
  const parts: Uint8Array[] = [];
  let changed = false;
  let dual = false;
  /** Whether the last EMF+ record was an EmfPlusGetDC. */
  let getDC = false;
  let offset = 0;
  while (offset + 8 <= bytes.length) {
    const type = view.getUint32(offset, true);
    const size = view.getUint32(offset + 4, true);
    if (size < 8 || offset + size > bytes.length) break;
    const record = bytes.subarray(offset, offset + size);
    if (type === EMR_COMMENT && size >= 16 && view.getUint32(offset + 12, true) === EMFPLUS) {
      const plus = emfPlusRecords(record);
      for (const r of plus) if (r.type === EMFPLUS_HEADER) dual = (r.flags & EMFPLUS_DUAL) !== 0;
      if (plus.length > 0) getDC = plus[plus.length - 1].type === EMFPLUS_GETDC;
      const rewritten = plus.map(rewritePlusRecord);
      if (rewritten.some((r, i) => r !== plus[i].bytes)) {
        parts.push(emfPlusComment(rewritten));
        changed = true;
      } else {
        parts.push(record);
      }
    } else if (dual && !getDC && type !== EMR_HEADER && type !== EMR_EOF && type !== EMR_COMMENT) {
      changed = true;
    } else {
      parts.push(record);
    }
    offset += size;
    if (type === EMR_EOF) break;
  }
  return changed ? assemble(parts, parts.length) : bytes;
}

interface EmfPlusRecord {
  type: number;
  flags: number;
  /** The record's bytes in the EMR_COMMENT, its 12-byte header first. */
  bytes: Uint8Array;
}

/** The EMF+ records in an EMR_COMMENT: its data after the identifier ([MS-EMFPLUS] 2.3.3). */
function emfPlusRecords(comment: Uint8Array): EmfPlusRecord[] {
  const view = new DataView(comment.buffer, comment.byteOffset, comment.byteLength);
  const end = Math.min(comment.length, 12 + view.getUint32(8, true));
  const records: EmfPlusRecord[] = [];
  for (let at = 16; at + 12 <= end; ) {
    const size = view.getUint32(at + 4, true);
    if (size < 12 || at + size > end) break;
    records.push({ type: view.getUint16(at, true), flags: view.getUint16(at + 2, true), bytes: comment.subarray(at, at + size) });
    at += size;
  }
  return records;
}

/** An EMF+ record as emf-converter reads it (see above); the same bytes when it reads it right. */
function rewritePlusRecord({ type, flags, bytes }: EmfPlusRecord): Uint8Array {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (type === EMFPLUS_DRAWDRIVERSTRING && bytes.length >= 28) {
    // BrushId, DriverStringOptionsFlags, MatrixPresent, GlyphCount, Glyphs, GlyphPos, TransformMatrix.
    const count = view.getUint32(24, true);
    const glyphsEnd = 28 + count * 2;
    const end = 12 + view.getUint32(8, true);
    if (count % 2 === 0 || glyphsEnd > end || end > bytes.length) return bytes;
    // The record keeps a size of a multiple of 4: the converter finds the
    // next records' glyph positions from their offsets in the file.
    const padded = new Uint8Array(Math.ceil((end + 2) / 4) * 4);
    padded.set(bytes.subarray(0, glyphsEnd));
    padded.set(bytes.subarray(glyphsEnd, end), glyphsEnd + 2);
    const out = new DataView(padded.buffer);
    out.setUint32(4, padded.length, true);
    out.setUint32(8, padded.length - 12, true);
    return padded;
  }
  if (type === EMFPLUS_OBJECT && ((flags >> 8) & 0x7f) === OBJECT_IMAGE && !(flags & EMFPLUS_CONTINUED) && bytes.length >= 40) {
    // Version, ImageDataType, then a bitmap's Width, Height, Stride, PixelFormat, BitmapDataType.
    // The version's signature tells an object from the last part of a continued one.
    const dataType = view.getUint32(36, true);
    if (view.getUint32(12, true) >>> 12 !== GRAPHICS_VERSION_SIGNATURE || view.getUint32(16, true) !== IMAGE_BITMAP || dataType > 1) return bytes;
    const copy = bytes.slice();
    new DataView(copy.buffer).setUint32(36, dataType + 1, true);
    return copy;
  }
  return bytes;
}

/** An EMR_COMMENT that holds these EMF+ records. */
function emfPlusComment(records: Uint8Array[]): Uint8Array {
  const dataSize = 4 + records.reduce((sum, record) => sum + record.length, 0);
  const out = new Uint8Array(12 + Math.ceil(dataSize / 4) * 4);
  const view = new DataView(out.buffer);
  view.setUint32(0, EMR_COMMENT, true);
  view.setUint32(4, out.length, true);
  view.setUint32(8, dataSize, true);
  view.setUint32(12, EMFPLUS, true);
  let at = 16;
  for (const record of records) {
    out.set(record, at);
    at += record.length;
  }
  return out;
}

/** The records as one EMF, its header's size and record count set to theirs. */
function assemble(parts: Uint8Array[], records: number): Uint8Array {
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
