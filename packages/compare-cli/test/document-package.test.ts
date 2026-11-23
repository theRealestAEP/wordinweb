import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterAll, describe, expect, it } from "vitest";
import { DocxDocument } from "../../core/src/docx.js";
import { compareDocuments } from "../../core/src/edit/compare/index.js";
import { AUTHOR, DATE, canonicalBody, para } from "../../core/test/compare-helpers.js";
import { makeDocx, wrapDocument } from "../../core/test/helpers.js";
import { parseXml, serializeXml } from "../../core/src/xml.js";
import { compareFiles } from "../src/compare.js";

const dir = mkdtempSync(join(tmpdir(), "compare-document-package-"));
afterAll(() => rmSync(dir, { recursive: true, force: true }));

const documentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const formats = [
  ["dotx", "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"],
  ["docm", "application/vnd.ms-word.document.macroEnabled.main+xml"],
  ["dotm", "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"],
] as const;
const macroTypes = [
  "application/vnd.ms-office.vbaProject",
  "application/vnd.ms-word.vbaData+xml",
  "application/vnd.ms-word.keyMapCustomizations+xml",
  "application/vnd.ms-word.attachedToolbars",
];
const macroParts = ["word/vbaProject.bin", "word/vbaData.xml", "word/customizations.xml", "word/attachedToolbars.bin"];
const macroRelationships = ["vbaProject", "wordVbaData", "keyMapCustomizations", "attachedToolbars"];
const relNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
const keptRelationships = `<Relationship Id="keep" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>`;
const options = { author: AUTHOR, date: DATE };

function input(body: string, mainType = documentType): Uint8Array {
  const files = unzipSync(makeDocx({
    "word/document.xml": wrapDocument(body),
    "word/_rels/document.xml.rels": `<Relationships xmlns="${relNamespace}">${keptRelationships}</Relationships>`,
    "customXml/item1.xml": "<kept>Preserve this part.</kept>",
  }));
  let types = strFromU8(files["[Content_Types].xml"]).replace(documentType, mainType);
  types = types.replace("</Types>", `<Override PartName="/word/keep.bin" ContentType="application/octet-stream"/></Types>`);
  files["word/keep.bin"] = new Uint8Array([0, 1, 2, 255]);
  if (mainType.includes("macroEnabled")) {
    // Exercise both Default and Override declarations for each macro type.
    types = types.replace("</Types>", macroTypes.map((type, i) =>
      `<Default Extension="${["bin", "vbaData", "keyMap", "toolbar"][i]}" ContentType="${type}"/>` +
      `<Override PartName="/${macroParts[i]}" ContentType="${type}"/>`,
    ).join("") + `</Types>`);
    files["word/_rels/document.xml.rels"] = strToU8(`<Relationships xmlns="${relNamespace}">${keptRelationships}${macroRelationships.map((type, i) =>
      `<Relationship Id="macro${i}" Type="http://schemas.microsoft.com/office/2006/relationships/${type}" Target="${macroParts[i].slice(5)}"/>`,
    ).join("")}</Relationships>`);
    files["word/vbaProject.bin"] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]);
    files["word/vbaData.xml"] = strToU8("<vbaData/>");
    files["word/customizations.xml"] = strToU8("<keyMapCustomizations/>");
    files["word/attachedToolbars.bin"] = new Uint8Array([1, 2, 3]);
    files["word/_rels/vbaProject.bin.rels"] = strToU8(`<Relationships xmlns="${relNamespace}"><Relationship Id="data" Type="http://schemas.microsoft.com/office/2006/relationships/wordVbaData" Target="vbaData.xml"/></Relationships>`);
  }
  files["[Content_Types].xml"] = strToU8(types);
  return zipSync(files);
}

function checkPackage(bytes: Uint8Array, reference: Uint8Array): void {
  const files = unzipSync(bytes);
  const expected = unzipSync(reference);
  const types = parseXml(strFromU8(files["[Content_Types].xml"]));
  expect(types.children.find((c) => c.attrs.PartName === "/word/document.xml")?.attrs.ContentType).toBe(documentType);
  for (const type of macroTypes) expect(types.children.some((c) => c.attrs.ContentType === type)).toBe(false);
  for (const part of [...macroParts, "word/_rels/vbaProject.bin.rels"]) expect(files[part]).toBeUndefined();
  const rels = strFromU8(files["word/_rels/document.xml.rels"]);
  for (const type of macroRelationships) expect(rels).not.toContain(type);
  expect(serializeXml(parseXml(rels))).toBe(serializeXml(parseXml(strFromU8(expected["word/_rels/document.xml.rels"]))));
  expect(Object.keys(files).sort()).toEqual(Object.keys(expected).sort());
  for (const name of Object.keys(expected)) {
    if (name === "[Content_Types].xml" || name === "word/_rels/document.xml.rels") {
      expect(serializeXml(parseXml(strFromU8(files[name])))).toBe(serializeXml(parseXml(strFromU8(expected[name]))));
    } else {
      expect(files[name], name).toEqual(expected[name]);
    }
  }
  expect(canonicalBody(DocxDocument.load(bytes))).toBe(canonicalBody(DocxDocument.load(reference)));
}

const originalBody = para("The quick brown fox jumps.");
const revisedBody = para("The quick brown cat jumps.");

describe.each(formats)("plain document comparison from .%s", (extension, mainType) => {
  it("compareFiles normalizes the result and printable package", async () => {
    const run = async (mainType: string, suffix: string): Promise<Uint8Array[]> => {
      const inputExtension = mainType === documentType ? "docx" : extension;
      const paths = ["original", "revised", "result", "printable"].map((name) => join(dir, `${extension}-${suffix}-${name}.${name === "original" || name === "revised" ? inputExtension : "docx"}`));
      writeFileSync(paths[0], input(originalBody, mainType));
      writeFileSync(paths[1], input(revisedBody, mainType));
      await compareFiles(paths[0], paths[1], paths[2], { ...options, printable: { path: paths[3] } });
      return paths.slice(2).map((path) => new Uint8Array(readFileSync(path)));
    };
    const expected = await run(documentType, "docx");
    const actual = await run(mainType, "source");
    actual.forEach((bytes, i) => checkPackage(bytes, expected[i]));
  });

  it.each([false, true])("compareDocuments normalizes the saved result (reuseRevised: %s)", async (reuseRevised) => {
    const original = DocxDocument.load(input(originalBody, mainType));
    const revised = DocxDocument.load(input(revisedBody, mainType));
    const originalFiles = original.saveFiles();
    const revisedFiles = revised.saveFiles();
    const result = compareDocuments(original, revised, { ...options, reuseRevised });
    const expected = compareDocuments(DocxDocument.load(input(originalBody)), DocxDocument.load(input(revisedBody)), options).save();
    checkPackage(result.save(), expected);
    checkPackage(await result.saveAsync(), expected);
    expect(original.saveFiles()).toEqual(originalFiles);
    if (!reuseRevised) expect(revised.saveFiles()).toEqual(revisedFiles);
  });
});
