import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { makeDocx, makeDocxWithMedia, p, wrapDocument } from "./helpers.js";
import { toStrictPackage } from "./strict-helpers.js";

let hasPoppler = false;
try {
  execFileSync("pdftotext", ["-v"], { stdio: "ignore" });
  execFileSync("pdfimages", ["-v"], { stdio: "ignore" });
  hasPoppler = true;
} catch {
  // PDF text and image inspection needs Poppler.
}

async function render(docx: Uint8Array): Promise<{ text: string; images: number }> {
  const chunks: Buffer[] = [];
  const out = new Writable({
    write(chunk: Buffer, _encoding, callback) {
      chunks.push(chunk);
      callback();
    },
  });
  await renderDocxToPdf(bytesSource(docx), out);
  const bytes = Buffer.concat(chunks);
  const dir = mkdtempSync(join(tmpdir(), "wordinweb-strict-"));
  try {
    const path = join(dir, "doc.pdf");
    writeFileSync(path, bytes);
    return {
      text: execFileSync("pdftotext", ["-layout", path, "-"], { encoding: "utf8" }),
      images: execFileSync("pdfimages", ["-list", path], { encoding: "utf8" }).split("\n").filter((line) => /^\s*\d+\s+\d+\s+image\s/.test(line)).length,
    };
  } finally {
    rmSync(dir, { recursive: true });
  }
}

describe.skipIf(!hasPoppler)("Strict PDF rendering", () => {
  it("preserves all 21 cover words and the text box's picture", async () => {
    const words = "This cover page contains twenty one words inside an anchored DrawingML group text box alongside its own small embedded color picture";
    const image = '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="Cover picture"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Cover picture"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="rIdPicture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
    const body = '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp14:pctPosHOffset>2300</wp14:pctPosHOffset></wp:positionH>' +
      '<mc:AlternateContent><mc:Choice Requires="wp14"><wp:positionV relativeFrom="page"><wp14:pctPosVOffset>30000</wp14:pctPosVOffset></wp:positionV></mc:Choice>' +
      '<mc:Fallback><wp:positionV relativeFrom="page"><wp:posOffset>3017520</wp:posOffset></wp:positionV></mc:Fallback></mc:AlternateContent>' +
      '<wp:extent cx="5486400" cy="2743200"/><wp:wrapNone/><wp:docPr id="1" name="Cover group"/><wp:cNvGraphicFramePr/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="2743200"/><a:chOff x="0" y="0"/><a:chExt cx="5486400" cy="2743200"/></a:xfrm></wpg:grpSpPr>' +
      '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></wps:spPr>' +
      `<wps:txbx><w:txbxContent>${p(words)}<w:p>${image}</w:p></w:txbxContent></wps:txbx><wps:bodyPr anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>` +
      '</wpg:wgp></a:graphicData></a:graphic>' +
      '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>94100</wp14:pctWidth></wp14:sizeRelH>' +
      '<wp14:sizeRelV relativeFrom="page"><wp14:pctHeight>36300</wp14:pctHeight></wp14:sizeRelV></wp:anchor></w:drawing></w:r></w:p>';
    const document = wrapDocument(body).replace('<w:document ', '<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ');
    const docx = makeDocxWithMedia({
      "word/document.xml": document,
      "word/_rels/document.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>',
    }, { "word/media/picture.png": Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAFgQIAffwlNAAAAABJRU5ErkJggg==", "base64") });
    const transitional = await render(docx);
    expect(words.split(" ")).toHaveLength(21);
    expect(transitional.text.trim().split(/\s+/)).toEqual(words.split(" "));
    expect(transitional.images).toBe(1);
    expect(await render(toStrictPackage(docx))).toEqual(transitional);
  }, 60_000);

  it("preserves all 32 ruby characters and vertical table text", async () => {
    // The six clusters in the Transitional original; the supplied Strict file omits them.
    const clusters = [["漢字", "かんじ"], ["東京", "とうきょう"], ["日本語", "にほんご"], ["愛", "あい"], ["平和", "へいわ"], ["未来", "みらい"]];
    const rubyText = clusters.flat().join("");
    const ruby = clusters.map(([base, annotation]) => '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="distributeSpace"/><w:hps w:val="8"/><w:hpsRaise w:val="11"/><w:hpsBaseText w:val="22"/></w:rubyPr>' +
      `<w:rt><w:r><w:rPr><w:sz w:val="8"/></w:rPr><w:t>${annotation}</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>${base}</w:t></w:r></w:rubyBase></w:ruby></w:r>`).join("");
    const cells = [["tbRl", "縦書文章"], ["btLr", "文字表記"]];
    const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3200"/><w:gridCol w:w="3200"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="3200"/></w:trPr>' +
      cells.map(([direction, text]) => `<w:tc><w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:textDirection w:val="${direction}"/></w:tcPr>${p(text)}</w:tc>`).join("") + '</w:tr></w:tbl>';
    const docx = makeDocx({ "word/document.xml": wrapDocument(`<w:p>${ruby}</w:p>${table}<w:sectPr><w:textDirection w:val="tbRl"/></w:sectPr>`) });
    const transitional = await render(docx);
    expect([...rubyText]).toHaveLength(32);
    expect([...transitional.text.replace(/\s/g, "")].sort()).toEqual([...(rubyText + cells.map(([, text]) => text).join(""))].sort());
    expect(transitional.images).toBe(0);
    expect(await render(toStrictPackage(docx))).toEqual(transitional);
  }, 60_000);
});
