// A legacy form field's data ([MS-DOC] 2.9.78 FFData), kept in the Data
// stream where the field's begin character points (sprmCPicLocation, with
// sprmCFData), written as the w:ffData of its w:fldChar: a text input, a
// checkbox (whose state is drawn from it) or a dropdown (whose selected
// entry is the field's only result).
import { parseSttb, u16, u32, utf16 } from "./bytes.js";
import { escapeAttr } from "./sprm.js";

/** The w:ffData element for the form field whose NilPICFAndBinData is at `offset`, or "" when there is none. */
export function ffDataXml(data: Uint8Array | null, offset: number): string {
  // NilPICFAndBinData: lcb (4), cbHeader (2, 0x44), then the FFData.
  if (!data || offset < 0 || offset + 6 > data.length) return "";
  const lcb = u32(data, offset);
  const cbHeader = u16(data, offset + 4);
  const end = Math.min(data.length, offset + lcb);
  let at = offset + cbHeader;
  if (cbHeader !== 0x44 || at + 10 > end || u32(data, at) !== 0xffffffff) return "";
  const bits = u16(data, at + 4);
  const cch = u16(data, at + 6);
  const hps = u16(data, at + 8);
  at += 10;
  const type = bits & 3;
  const res = (bits >>> 2) & 0x1f;
  const iSize = (bits >>> 10) & 1;
  const typeTxt = (bits >>> 11) & 7;
  // An Xstz: a character count, the characters, a zero terminator.
  const xstz = (): string => {
    if (at + 2 > end) return "";
    const n = u16(data, at);
    const s = at + 2 + n * 2 <= end ? utf16(data, at + 2, n) : "";
    at += 2 + n * 2 + 2;
    return s;
  };
  const name = xstz();
  const textDef = type === 0 ? xstz() : "";
  let wDef = 0;
  if (type === 1 || type === 2) {
    wDef = at + 2 <= end ? u16(data, at) : 0;
    at += 2;
  }
  const format = xstz();
  const help = xstz();
  const status = xstz();
  const entry = xstz();
  const exit = xstz();
  const entries = type === 2 && at < end ? parseSttb(data.subarray(at, end)).strings : [];

  let out = `<w:ffData><w:name w:val="${escapeAttr(name)}"/><w:enabled/><w:calcOnExit w:val="${(bits >>> 14) & 1}"/>`;
  if (entry) out += `<w:entryMacro w:val="${escapeAttr(entry)}"/>`;
  if (exit) out += `<w:exitMacro w:val="${escapeAttr(exit)}"/>`;
  if (help) out += `<w:helpText w:type="text" w:val="${escapeAttr(help)}"/>`;
  if (status) out += `<w:statusText w:type="text" w:val="${escapeAttr(status)}"/>`;
  if (type === 1) {
    // iRes 25 is an undefined state, which reads as unchecked.
    const checked = res === 25 ? 0 : res;
    out += `<w:checkBox>${iSize ? `<w:size w:val="${hps}"/>` : "<w:sizeAuto/>"}<w:default w:val="${wDef}"/>${checked ? "<w:checked/>" : '<w:checked w:val="0"/>'}</w:checkBox>`;
  } else if (type === 2) {
    out += `<w:ddList>${res !== 25 ? `<w:result w:val="${res}"/>` : ""}${wDef ? `<w:default w:val="${wDef}"/>` : ""}${entries.map((e) => `<w:listEntry w:val="${escapeAttr(e)}"/>`).join("")}</w:ddList>`;
  } else {
    const types = ["regular", "number", "date", "currentDate", "currentTime", "calculated"];
    let input = "";
    if (typeTxt) input += `<w:type w:val="${types[typeTxt] ?? "regular"}"/>`;
    if (textDef) input += `<w:default w:val="${escapeAttr(textDef)}"/>`;
    if (cch) input += `<w:maxLength w:val="${cch}"/>`;
    if (format) input += `<w:format w:val="${escapeAttr(format)}"/>`;
    out += input ? `<w:textInput>${input}</w:textInput>` : "<w:textInput/>";
  }
  return out + "</w:ffData>";
}
