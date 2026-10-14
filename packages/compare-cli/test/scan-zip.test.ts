import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import { scanBody, type BodyChild } from "../src/scan.js";
import { ZipReader, ZipWriter } from "../src/zip.js";

const dir = mkdtempSync(join(tmpdir(), "scan-zip-test-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));
let files = 0;

function scan(xml: string): { children: string[]; body: string } {
  const path = join(dir, `${files++}.xml`);
  writeFileSync(path, xml);
  const found: BodyChild[] = [];
  const bounds = scanBody(path, (c) => found.push(c));
  const bytes = readFileSync(path);
  return {
    children: found.map((c) => bytes.toString("utf8", c.start, c.end)),
    body: bytes.toString("utf8", bounds.contentStart, bounds.contentEnd),
  };
}

const doc = (body: string): string =>
  `<?xml version="1.0"?><w:document xmlns:w="w"><w:body>${body}</w:body></w:document>`;

describe("scanBody", () => {
  it("finds each child of the body, skipping what is not an element", () => {
    const body = `<w:p a="x>y"><w:r><w:t>a</w:t></w:r></w:p><!-- <w:p>no</w:p> --><w:tbl/><?pi x?><w:sectPr><w:pgSz/></w:sectPr>`;
    const { children } = scan(doc(body));
    expect(children).toEqual([`<w:p a="x>y"><w:r><w:t>a</w:t></w:r></w:p>`, `<w:tbl/>`, `<w:sectPr><w:pgSz/></w:sectPr>`]);
  });

  it("reads blocks across chunk boundaries", () => {
    const paragraphs = Array.from({ length: 30000 }, (_, i) => `<w:p><w:r><w:t>Paragraph ${i} ${"x".repeat(i % 97)}</w:t></w:r></w:p>`);
    const { children } = scan(doc(paragraphs.join("")));
    expect(children).toEqual(paragraphs);
  });

  it("an empty body has no children", () => {
    expect(scan(doc("")).children).toEqual([]);
    expect(scan(`<w:document><w:body/></w:document>`).children).toEqual([]);
  });

  it("fails on what it cannot read, and does not hang", () => {
    expect(() => scan(doc(`<w:p/><!-- never closed`))).toThrow(/unterminated comment/);
    expect(() => scan(doc(`<w:p/><![CDATA[ never closed`))).toThrow(/unterminated CDATA/);
    expect(() => scan(doc(`<w:p/><?pi never closed`))).toThrow(/unterminated processing instruction/);
    expect(() => scan(`<w:document><w:body><w:p a="never closed></w:body></w:document>`)).toThrow(/unterminated start tag/);
    expect(() => scan(`<!DOCTYPE x><w:document><w:body/></w:document>`)).toThrow(/document type/);
    expect(() => scan(`<w:document><w:body><w:p/>`)).toThrow(/not closed/);
    expect(() => scan(`<w:document></w:document>`)).toThrow(/no w:body/);
  });
});

describe("ZipWriter and ZipReader", () => {
  it("write entries that read back, deflated and copied raw", async () => {
    const big = Buffer.from("word ".repeat(200000));
    const path = join(dir, "a.zip");
    const writer = new ZipWriter(path);
    writer.addBuffer("a.xml", Buffer.from("<a/>"));
    writeFileSync(join(dir, "big.txt"), big);
    await writer.addFile("dir/big.txt", join(dir, "big.txt"));
    writer.addBuffer("empty", Buffer.alloc(0));
    writer.close();

    const reader = ZipReader.open(path);
    expect([...reader.entries.keys()]).toEqual(["a.xml", "dir/big.txt", "empty"]);
    expect(reader.read(reader.entries.get("a.xml")!).toString()).toBe("<a/>");
    expect(reader.read(reader.entries.get("dir/big.txt")!).equals(big)).toBe(true);
    await reader.extract(reader.entries.get("dir/big.txt")!, join(dir, "out.txt"));
    expect(readFileSync(join(dir, "out.txt")).equals(big)).toBe(true);

    const copy = join(dir, "b.zip");
    const second = new ZipWriter(copy);
    reader.copyInto(second, reader.entries.get("dir/big.txt")!, "renamed.txt");
    second.close();
    reader.close();
    const copied = ZipReader.open(copy);
    expect(copied.read(copied.entries.get("renamed.txt")!).equals(big)).toBe(true);
    copied.close();
  });

  it("rejects a file that is not a zip", () => {
    writeFileSync(join(dir, "not.zip"), "hello");
    expect(() => ZipReader.open(join(dir, "not.zip"))).toThrow(/not a \.docx file/);
  });
});
