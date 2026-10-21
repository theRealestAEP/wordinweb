const decoder = new TextDecoder("windows-1252");
// https://encoding.spec.whatwg.org/index-windows-1252.txt, bytes 0x80-0x9F.
const C1 = "\u20ac\u0081\u201a\u0192\u201e\u2026\u2020\u2021\u02c6\u2030\u0160\u2039\u0152\u008d\u017d\u008f\u0090\u2018\u2019\u201c\u201d\u2022\u2013\u2014\u02dc\u2122\u0161\u203a\u0153\u009d\u017e\u0178";

/** Node 22.20's decoder leaves Windows-1252 punctuation as C1 controls. */
export const windows1252 = {
  decode(bytes: Uint8Array): string {
    return decoder.decode(bytes).replace(/[\u0080-\u009f]/g, (char) => C1[char.charCodeAt(0) - 0x80]);
  },
};
