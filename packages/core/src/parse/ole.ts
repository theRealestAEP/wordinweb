import * as CFB from "cfb";

const PACKAGE_CLSID = "0C00030000000000C000000000000046";
// MS-OLEDS 3.1: embedded-object \1Ole stream = Version, Flags,
// LinkUpdateOption, Reserved (four little-endian uint32 values).
const OLE_MARKER = new Uint8Array([1, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function safeFilename(value: string): string {
  return value.split(/[\\/]/).pop()?.replace(/[\u0000-\u001f]/g, "") || "embedded-file.bin";
}

function writeUint16(out: Uint8Array, offset: number, value: number): number {
  new DataView(out.buffer).setUint16(offset, value, true);
  return offset + 2;
}

function writeUint32(out: Uint8Array, offset: number, value: number): number {
  new DataView(out.buffer).setUint32(offset, value, true);
  return offset + 4;
}

/** Wrap arbitrary file bytes in the OLE Package container Word activates. */
export function buildOlePackage(data: Uint8Array, requestedFilename: string): Uint8Array {
  const filename = safeFilename(requestedFilename);
  const label = encoder.encode(filename);
  const command = encoder.encode(filename);
  const payloadSize = 2 + label.length + 1 + label.length + 1 + 2 + 2 + 4 + command.length + 1 + 4 + data.length + 2;
  const native = new Uint8Array(payloadSize + 4);
  let offset = writeUint32(native, 0, payloadSize);
  offset = writeUint16(native, offset, 2);
  native.set(label, offset); offset += label.length + 1;
  native.set(label, offset); offset += label.length + 1;
  offset = writeUint16(native, offset, 0);
  offset = writeUint16(native, offset, 3);
  offset = writeUint32(native, offset, command.length + 1);
  native.set(command, offset); offset += command.length + 1;
  offset = writeUint32(native, offset, data.length);
  native.set(data, offset); offset += data.length;
  writeUint16(native, offset, 0);

  const cfb = CFB.utils.cfb_new({ CLSID: PACKAGE_CLSID });
  CFB.utils.cfb_del(cfb, "\u0001Sh33tJ5");
  CFB.utils.cfb_add(cfb, "\u0001Ole", OLE_MARKER);
  CFB.utils.cfb_add(cfb, "\u0001Ole10Native", native);
  return new Uint8Array(CFB.write(cfb, { type: "array", fileType: "cfb" }));
}

function readAsciiZ(bytes: Uint8Array, offset: number): { value: string; next: number } | null {
  const end = bytes.indexOf(0, offset);
  if (end === -1) return null;
  return { value: decoder.decode(bytes.subarray(offset, end)), next: end + 1 };
}

/** Recover the original file stored in a Word OLE Package container. */
export function extractOlePackage(bytes: Uint8Array): { filename: string; data: Uint8Array } | null {
  try {
    const cfb = CFB.read(bytes, { type: "array" });
    const entry = CFB.find(cfb, "\u0001Ole10Native");
    if (!entry?.content) return null;
    const native = new Uint8Array(entry.content);
    const view = new DataView(native.buffer, native.byteOffset, native.byteLength);
    if (native.length < 16) return null;
    let offset = 6;
    const label = readAsciiZ(native, offset);
    if (!label) return null;
    offset = label.next;
    const filename = readAsciiZ(native, offset);
    if (!filename) return null;
    offset = filename.next + 4;
    if (offset + 4 > native.length) return null;
    const commandLength = view.getUint32(offset, true);
    offset += 4 + commandLength;
    if (offset + 4 > native.length) return null;
    const dataSize = view.getUint32(offset, true);
    offset += 4;
    if (offset + dataSize > native.length) return null;
    return {
      filename: safeFilename(filename.value || label.value),
      data: native.slice(offset, offset + dataSize),
    };
  } catch {
    return null;
  }
}
