// Little-endian readers over the raw bytes, and the string tables of [MS-DOC] 2.2.4.

export function u8(b: Uint8Array, at: number): number {
  return b[at];
}

export function u16(b: Uint8Array, at: number): number {
  return b[at] | (b[at + 1] << 8);
}

export function i16(b: Uint8Array, at: number): number {
  return (u16(b, at) << 16) >> 16;
}

export function u32(b: Uint8Array, at: number): number {
  return (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24)) >>> 0;
}

export function i32(b: Uint8Array, at: number): number {
  return b[at] | (b[at + 1] << 8) | (b[at + 2] << 16) | (b[at + 3] << 24);
}

export function utf16(b: Uint8Array, at: number, chars: number): string {
  let out = "";
  for (let i = 0; i < chars && at + 2 * i + 1 < b.length; i++) out += String.fromCharCode(u16(b, at + 2 * i));
  return out;
}

const cp1252 = new TextDecoder("windows-1252");

/** An STTB ([MS-DOC] 2.2.4): its strings and, per string, its extra data. */
export function parseSttb(b: Uint8Array): { strings: string[]; extra: Uint8Array[] } {
  const strings: string[] = [];
  const extra: Uint8Array[] = [];
  if (b.length < 4) return { strings, extra };
  const wide = u16(b, 0) === 0xffff;
  let at = wide ? 2 : 0;
  const count = u16(b, at);
  const cbExtra = u16(b, at + 2);
  at += 4;
  for (let i = 0; i < count && at < b.length; i++) {
    const cch = wide ? u16(b, at) : b[at];
    at += wide ? 2 : 1;
    if (wide) {
      strings.push(utf16(b, at, cch));
      at += cch * 2;
    } else {
      strings.push(cp1252.decode(b.subarray(at, at + cch)));
      at += cch;
    }
    extra.push(b.subarray(at, at + cbExtra));
    at += cbExtra;
  }
  return { strings, extra };
}

/** A PLC ([MS-DOC] 2.2.2) with `cbData` bytes per element: its CPs and its data elements. */
export function parsePlc(b: Uint8Array, cbData: number): { cps: number[]; data: Uint8Array[] } {
  const cps: number[] = [];
  const data: Uint8Array[] = [];
  if (b.length < 4) return { cps, data };
  const n = Math.floor((b.length - 4) / (4 + cbData));
  for (let i = 0; i <= n; i++) cps.push(u32(b, i * 4));
  const base = (n + 1) * 4;
  for (let i = 0; i < n; i++) data.push(b.subarray(base + i * cbData, base + (i + 1) * cbData));
  return { cps, data };
}

/** A DTTM ([MS-DOC] 2.9.65) as an ISO date, or undefined for the zero date. */
export function dttmToIso(v: number): string | undefined {
  const minute = v & 0x3f;
  const hour = (v >>> 6) & 0x1f;
  const day = (v >>> 11) & 0x1f;
  const month = (v >>> 16) & 0x0f;
  const year = 1900 + ((v >>> 20) & 0x1ff);
  if (day === 0 || month === 0) return undefined;
  const two = (n: number): string => String(n).padStart(2, "0");
  return `${year}-${two(month)}-${two(day)}T${two(hour)}:${two(minute)}:00Z`;
}
