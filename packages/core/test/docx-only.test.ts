import { build } from "esbuild";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { DocxDocument as FullDocument } from "../src/docx.js";
import { DocxDocument } from "../src/docx-only.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";
import { strictPackage } from "./strict-helpers.js";

describe("DOCX-only entry", () => {
  it.each([false, true])("loads and clones DOCX packages (Strict: %s)", (strict) => {
    const bytes = strict ? strictPackage(true) : makeDocx({ "word/document.xml": wrapDocument(p("hello")) });
    const full = FullDocument.load(bytes);
    const doc = DocxDocument.load(bytes);
    expect(doc.saveFiles()).toEqual(full.saveFiles());
    expect(doc.clone().saveFiles()).toEqual(full.clone().saveFiles());
    expect(full).toBeInstanceOf(FullDocument);
    expect(full.clone()).toBeInstanceOf(FullDocument);
    expect(FullDocument.fromFiles(full.saveFiles())).toBeInstanceOf(FullDocument);
  });

  it("excludes binary Word converters from a browser bundle", async () => {
    const result = await build({
      entryPoints: [fileURLToPath(new URL("../src/docx-only.ts", import.meta.url))],
      bundle: true,
      platform: "browser",
      format: "esm",
      target: "es2020",
      write: false,
      metafile: true,
      minify: true,
    });
    const inputs = Object.keys(result.metafile!.inputs);
    expect(inputs.filter((name) => /\/core\/src\/(doc|oldword)\//.test(name))).toEqual([]);
    expect(inputs.some((name) => name.endsWith("docx-document.ts"))).toBe(true);
  });
});
