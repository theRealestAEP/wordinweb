import { execFileSync } from "node:child_process";
import { existsSync, mkdtempSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { Face } from "../src/pdf/sfnt.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx } from "./helpers.js";

/**
 * Adding fonts: a directory given as `extraFontDirs`, with `docx-pdf --fonts
 * DIR`, or in the WORDINWEB_FONT_DIRS environment variable is scanned ahead
 * of the default directories, so a font file dropped in it wins over the
 * substitutes for its family. The family here is made up: a copy of Carlito
 * renamed in its name table, which no default directory holds.
 */
const FAMILY = "Zorbaq7";
const dir = mkdtempSync(join(tmpdir(), "wordinweb-fonts-"));
const docx = makeDocx({
  "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="${FAMILY}" w:hAnsi="${FAMILY}"/></w:rPr><w:t>Set in a made-up family.</w:t></w:r></w:p></w:body></w:document>`,
});

class Collect extends Writable {
  readonly chunks: Buffer[] = [];
  _write(chunk: Buffer, _encoding: string, callback: () => void): void {
    this.chunks.push(chunk);
    callback();
  }
}

beforeAll(() => {
  // Carlito's plain sfnt with "Carlito" replaced by a name of the same length in its name table, both encodings.
  const carlito = createRequire(import.meta.url).resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff");
  const sfnt = Buffer.from(new Face(carlito).sfnt);
  for (const [from, to] of [
    [Buffer.from("Carlito", "latin1"), Buffer.from(FAMILY, "latin1")],
    [Buffer.from("Carlito", "utf16le").swap16(), Buffer.from(FAMILY, "utf16le").swap16()],
  ]) {
    let at = sfnt.indexOf(from);
    while (at !== -1) {
      to.copy(sfnt, at);
      at = sfnt.indexOf(from, at + from.length);
    }
  }
  writeFileSync(join(dir, "zorbaq7.ttf"), sfnt);
});

afterAll(() => {
  delete process.env.WORDINWEB_FONT_DIRS;
});

async function render(options: Parameters<typeof renderDocxToPdf>[2]): Promise<{ fonts: string[]; missingFonts: string[]; pdf: string }> {
  const out = new Collect();
  const result = await renderDocxToPdf(bytesSource(docx), out, options);
  return { fonts: result.fonts, missingFonts: result.missingFonts, pdf: Buffer.concat(out.chunks).toString("latin1") };
}

describe("font directories", () => {
  it("reports the made-up family missing when no directory holds it", async () => {
    const { missingFonts, fonts } = await render({});
    expect(missingFonts).toEqual([FAMILY]);
    expect(fonts.some((f) => f.startsWith(FAMILY))).toBe(false);
  });

  it("embeds a face from an extra directory, ahead of the defaults", async () => {
    const { missingFonts, fonts, pdf } = await render({ extraFontDirs: [dir] });
    expect(missingFonts).toEqual([]);
    expect(fonts).toContain(`${FAMILY} Regular`);
    expect(pdf).toContain(`/BaseFont /`);
    expect(pdf).toMatch(new RegExp(`/BaseFont /\\w+\\+${FAMILY}`));
  });

  it("scans an extra directory ahead of fontDirs too", async () => {
    const { missingFonts, fonts } = await render({ fontDirs: [], extraFontDirs: [dir] });
    expect(missingFonts).toEqual([]);
    expect(fonts).toContain(`${FAMILY} Regular`);
  });

  it("reads WORDINWEB_FONT_DIRS as extra directories", async () => {
    process.env.WORDINWEB_FONT_DIRS = `/no/such/dir:${dir}`;
    try {
      const { missingFonts, fonts } = await render({});
      expect(missingFonts).toEqual([]);
      expect(fonts).toContain(`${FAMILY} Regular`);
    } finally {
      delete process.env.WORDINWEB_FONT_DIRS;
    }
  });

  const cli = join(import.meta.dirname, "../dist/pdf/cli.js");
  it.skipIf(!existsSync(cli))("docx-pdf --fonts DIR scans the directory", () => {
    const input = join(dir, "doc.docx");
    writeFileSync(input, docx);
    const line = execFileSync("node", [cli, input, join(dir, "out.pdf"), "--fonts", dir], { encoding: "utf8" });
    const result = JSON.parse(line.trim().split("\n").pop() as string) as { fonts: string[]; missingFonts: string[] };
    expect(result.missingFonts).toEqual([]);
    expect(result.fonts).toContain(`${FAMILY} Regular`);
  });
});
