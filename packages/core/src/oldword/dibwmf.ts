// A device-independent bitmap as a Windows metafile of one StretchDIB
// record: the renderer draws a metafile picture, and a bitmap kept this way
// keeps its pixels as they are.

/** The DIB (BITMAPINFOHEADER or BITMAPCOREHEADER, colors, pixels) as a WMF, or undefined when its header is not one. */
export function dibToWmf(dib: Uint8Array): Uint8Array | undefined {
  if (dib.length < 12) return undefined;
  const v = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const header = v.getUint32(0, true);
  let width: number;
  let height: number;
  if (header === 12) {
    width = v.getUint16(4, true);
    height = v.getUint16(6, true);
  } else if (header >= 40 && dib.length >= 40) {
    width = v.getInt32(4, true);
    height = Math.abs(v.getInt32(8, true));
  } else return undefined;
  if (width <= 0 || height <= 0 || width > 0x7fff || height > 0x7fff) return undefined;
  const padded = dib.length + (dib.length & 1);
  const records: number[][] = [
    [0x020b, 0, 0], // SetWindowOrg
    [0x020c, height, width], // SetWindowExt
  ];
  const stretchWords = 3 + 2 + 1 + 8 + padded / 2;
  const total = 9 + records.reduce((sum, r) => sum + 2 + r.length, 0) + stretchWords + 3;
  const out = new Uint8Array(total * 2);
  const o = new DataView(out.buffer);
  o.setUint16(0, 1, true); // in memory
  o.setUint16(2, 9, true); // header words
  o.setUint16(4, 0x300, true);
  o.setUint32(6, total, true);
  o.setUint16(10, 0, true);
  o.setUint32(12, stretchWords, true);
  let at = 18;
  for (const [fn, ...params] of records) {
    o.setUint32(at, 2 + 1 + params.length, true);
    o.setUint16(at + 4, fn, true);
    params.forEach((p, i) => o.setInt16(at + 6 + 2 * i, p, true));
    at += 2 * (3 + params.length);
  }
  // StretchDIB: SRCCOPY, RGB colors, the whole bitmap onto the whole window.
  o.setUint32(at, stretchWords, true);
  o.setUint16(at + 4, 0x0f43, true);
  o.setUint32(at + 6, 0x00cc0020, true);
  o.setUint16(at + 10, 0, true);
  [height, width, 0, 0, height, width, 0, 0].forEach((p, i) => o.setInt16(at + 12 + 2 * i, p, true));
  out.set(dib, at + 28);
  at += 2 * stretchWords;
  o.setUint32(at, 3, true); // EOF
  o.setUint16(at + 4, 0, true);
  return out;
}
