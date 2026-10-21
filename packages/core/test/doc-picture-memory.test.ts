import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { deflateSync } from "node:zlib";
import * as CFB from "cfb";
import { unzipSync } from "fflate";
import { expect, it } from "vitest";
import { docPackageFiles, docToDocx } from "../src/doc/hooks.js";
import { decodeBlip, REC } from "../src/doc/officeart.js";
import { openStreamingDocument } from "../src/stream/body.js";
import { bytesSource } from "../src/stream/source.js";
import { wmfLines } from "./pictures.js";

/** Word 97 text and CHPX runs pointing at distinct compressed WMF BLIPs. */
function pictureDoc(count: number, pictureBytes: number): Uint8Array {
  const record = (type: number, body: Uint8Array, verInst = 0): Buffer => {
    const header = Buffer.alloc(8);
    header.writeUInt16LE(verInst);
    header.writeUInt16LE(type, 2);
    header.writeUInt32LE(body.length, 4);
    return Buffer.concat([header, body]);
  };
  const pictures: Buffer[] = [];
  const offsets: number[] = [];
  let dataSize = 0;
  for (let i = 0; i < count; i++) {
    const lines = wmfLines();
    const image = Buffer.alloc(pictureBytes);
    image.set(lines.subarray(0, -6));
    // An ignored WMF record pads a valid, visible line drawing.
    const padding = lines.length - 6;
    image.writeUInt32LE((image.length - padding - 6) / 2, padding);
    image.writeUInt16LE(0x0626, padding + 4); // META_ESCAPE
    image[padding + 10] = i;
    image.writeUInt32LE(3, image.length - 6); // META_EOF
    image.writeUInt32LE(image.length / 2, 6);
    const compressed = deflateSync(image);
    const header = Buffer.alloc(50); // UID + OfficeArtMetafileHeader
    header.writeUInt32LE(image.length, 16);
    header.writeInt32LE(914400, 36);
    header.writeInt32LE(457200, 40);
    header.writeUInt32LE(compressed.length, 44);
    const blip = record(REC.BlipWMF, Buffer.concat([header, compressed]));
    const shape = record(REC.SpContainer, Buffer.alloc(0), 0x0f);
    const picf = Buffer.alloc(68);
    picf.writeUInt32LE(picf.length + shape.length + blip.length);
    picf.writeUInt16LE(68, 4);
    picf.writeUInt16LE(0x64, 6);
    picf.writeUInt16LE(1440, 28);
    picf.writeUInt16LE(720, 30);
    picf.writeUInt16LE(1000, 32);
    picf.writeUInt16LE(1000, 34);
    offsets.push(dataSize);
    const picture = Buffer.concat([picf, shape, blip]);
    pictures.push(picture);
    dataSize += picture.length;
  }

  const text = `Before\r${"\x01\r".repeat(count)}After\r`;
  const word = Buffer.alloc(2048);
  word.writeUInt16LE(0xa5ec);
  word.writeUInt16LE(0xc1, 2);
  word.writeUInt16LE(14, 32); // csw
  word.writeUInt16LE(22, 62); // cslw
  word.writeUInt32LE(text.length, 76); // ccpText
  word.writeUInt16LE(34, 152); // fc/lcb pair count
  word.write(text, 1024, "latin1");
  const table = Buffer.alloc(33);
  table[0] = 2; // Clx Pcdt, one compressed piece
  table.writeUInt32LE(16, 1);
  table.writeUInt32LE(text.length, 9);
  table.writeUInt32LE(0x40000000 | 2048, 15);
  const pair = (index: number, at: number, size: number): void => {
    word.writeUInt32LE(at, 154 + index * 8);
    word.writeUInt32LE(size, 158 + index * 8);
  };
  pair(33, 0, 21); // Clx
  pair(12, 21, 12); // PlcBteChpx
  table.writeUInt32LE(1024, 21);
  table.writeUInt32LE(1024 + text.length, 25);
  table.writeUInt32LE(3, 29); // CHPX page at 1536
  const page = word.subarray(1536);
  const starts = [1024, ...offsets.map((_, i) => 1031 + 2 * i), 1031 + 2 * count, 1024 + text.length];
  starts.forEach((fc, i) => page.writeUInt32LE(fc, i * 4));
  const runs = starts.length - 1;
  page[511] = runs;
  for (let i = 0; i < count; i++) {
    const at = 176 + i * 8;
    page[(runs + 1) * 4 + i + 1] = at / 2;
    page[at] = 6;
    page.writeUInt16LE(0x6a03, at + 1); // sprmCPicLocation
    page.writeUInt32LE(offsets[i], at + 3);
  }
  const cfb = CFB.utils.cfb_new();
  CFB.utils.cfb_add(cfb, "WordDocument", word);
  CFB.utils.cfb_add(cfb, "0Table", table);
  CFB.utils.cfb_add(cfb, "Data", Buffer.concat(pictures));
  return CFB.write(cfb, { type: "buffer" });
}

it("preserves every package part when media is zipped during conversion", () => {
  const doc = pictureDoc(2, 65536);
  const expected = docPackageFiles(doc);
  expect(Object.keys(expected).filter(name => name.startsWith("word/media/"))).toHaveLength(2);
  expect(unzipSync(docToDocx(doc))).toEqual(expected);
  expect(docToDocx(doc)).toEqual(docToDocx(doc));
});

it("keeps public media reads independent across converted pictures", async () => {
  const doc = pictureDoc(2, 65536);
  const expected = docPackageFiles(doc);
  const opened = await openStreamingDocument(bytesSource(doc));
  const names = Object.keys(expected).filter(name => name.startsWith("word/media/"));
  const images = await Promise.all(names.map(name => opened.media(name)));
  images.forEach((bytes, i) => expect(bytes).toEqual(expected[names[i]]));
});

it("keeps borrowed pictures independent between open documents", async () => {
  const doc = pictureDoc(2, 65536);
  const expected = docPackageFiles(doc);
  const first = await openStreamingDocument(bytesSource(doc), { borrowMedia: true });
  const second = await openStreamingDocument(bytesSource(doc), { borrowMedia: true });
  const names = Object.keys(expected).filter(name => name.startsWith("word/media/"));
  const image = await first.media(names[0]);
  await second.media(names[1]);
  expect(image).toEqual(expected[names[0]]);
});

it("preserves metafile truncation with inaccurate expanded sizes and reports damaged zlib", () => {
  const payload = Buffer.alloc(128 << 10, 123);
  const compressed = deflateSync(payload, { level: 0 });
  const body = Buffer.alloc(50 + compressed.length);
  body.writeUInt32LE(compressed.length, 44);
  body.set(compressed, 50);
  const rec = { start: 0, length: body.length, instance: 0, type: REC.BlipWMF, ver: 0 };
  for (const size of [1, 0, payload.length, 0x7fffffff]) {
    body.writeUInt32LE(size, 16);
    const image = decodeBlip(body, rec)!;
    expect(image.data).toEqual(new Uint8Array(payload.subarray(0, size)));
    expect(image.borrowData!()).toEqual(new Uint8Array(payload.subarray(0, size)));
  }
  body[50] = 0;
  expect(decodeBlip(body, rec)).toBeNull();
});

it("converts and renders 30 compressed 4 MiB pictures within 160 MiB of warm RSS", async () => {
  const dir = mkdtempSync(join(tmpdir(), "doc-picture-memory-"));
  try {
    writeFileSync(join(dir, "control.doc"), pictureDoc(1, 65536));
    writeFileSync(join(dir, "pictures.doc"), pictureDoc(30, 4 << 20));
    const viteNode = fileURLToPath(new URL("../../../node_modules/.bin/vite-node", import.meta.url));
    const child = fileURLToPath(new URL("./doc-picture-memory-child.ts", import.meta.url));
    const outcome = await new Promise<{ rssRiseMB: number; pages: number; unsupportedImages: string[]; images: number }>((resolve, reject) => {
      const run = spawn(process.execPath, ["--max-old-space-size=1024", "--expose-gc", viteNode, child, join(dir, "control.doc"), join(dir, "pictures.doc")], { stdio: ["ignore", "pipe", "pipe"] });
      let out = "";
      let err = "";
      run.stdout.on("data", chunk => { out += chunk; });
      run.stderr.on("data", chunk => { err += chunk; });
      run.on("error", reject);
      run.on("close", status => status === 0 ? resolve(JSON.parse(out.trim().split("\n").pop()!)) : reject(new Error(`exit ${status}: ${err}`)));
    });
    console.info("30 x 4 MiB compressed pictures:", outcome);
    expect(outcome.rssRiseMB).toBeLessThan(160);
    expect(outcome.pages).toBeGreaterThan(1);
    expect(outcome.unsupportedImages).toEqual([]);
    expect(outcome.images).toBe(30);
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}, 60_000);
