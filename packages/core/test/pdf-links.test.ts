import { Writable } from "node:stream";
import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { hyperlinkTarget } from "../src/parse/document.js";
import { linkAction, pdfName } from "../src/pdf/links.js";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx } from "./helpers.js";

const R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function field(instr: string, result: string): string {
  return (
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    result +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r>`
  );
}

function run(text: string): string {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

/**
 * Page 1: an external w:hyperlink, a w:hyperlink to the bookmark "last" on
 * page 2, a HYPERLINK field to a URL, a HYPERLINK \l field to the bookmark,
 * a javascript: w:hyperlink, and a w:hyperlink long enough to wrap onto a
 * second line. Page 2: the bookmark, on its second paragraph.
 */
function linksDocx(): Uint8Array {
  const long = Array.from({ length: 30 }, (_, i) => `word${i}`).join(" ");
  const body =
    `<w:p>${run("See ")}<w:hyperlink r:id="rIdA">${run("the site")}</w:hyperlink>${run(" and ")}<w:hyperlink w:anchor="last">${run("the end")}</w:hyperlink>${run(".")}</w:p>` +
    `<w:p>${field('HYPERLINK "https://example.com/field"', run("field link"))}${run(" then ")}${field('HYPERLINK \\l "last"', run("field anchor"))}</w:p>` +
    `<w:p><w:hyperlink r:id="rIdJ">${run("not a link")}</w:hyperlink></w:p>` +
    `<w:p><w:hyperlink r:id="rIdA">${run(long)}</w:hyperlink></w:p>` +
    `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
    `<w:p>${run("Page two opens here.")}</w:p>` +
    `<w:p><w:bookmarkStart w:id="1" w:name="last"/>${run("The last words.")}<w:bookmarkEnd w:id="1"/></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440" w:header="720" w:footer="720"/></w:sectPr>`;
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ${W_NS} ${R_NS}><w:body>${body}</w:body></w:document>`;
  const rels = `<?xml version="1.0"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdA" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a" TargetMode="External"/>
<Relationship Id="rIdJ" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
</Relationships>`;
  return makeDocx({ "word/document.xml": document, "word/_rels/document.xml.rels": rels });
}

async function render(docx: Uint8Array): Promise<string> {
  const chunks: Buffer[] = [];
  const out = new Writable({
    write(chunk: Buffer, _encoding, callback) {
      chunks.push(chunk);
      callback();
    },
  });
  await renderDocxToPdf(bytesSource(docx), out);
  return Buffer.concat(chunks).toString("latin1");
}

/** The link annotations in a PDF, each with its page (from 1) and rect. */
function links(pdf: string): { page: number; rect: number[]; body: string }[] {
  const objects = new Map<number, string>();
  for (const m of pdf.matchAll(/(\d+) 0 obj\n([^]*?)\nendobj/g)) objects.set(Number(m[1]), m[2]);
  const pages = [...pdf.matchAll(/(\d+) 0 obj\n<< \/Type \/Page [^]*?endobj/g)];
  const out: { page: number; rect: number[]; body: string }[] = [];
  pages.forEach((m, index) => {
    const annots = /\/Annots \[([^\]]*)\]/.exec(m[0])?.[1] ?? "";
    for (const ref of annots.matchAll(/(\d+) 0 R/g)) {
      const body = objects.get(Number(ref[1])) ?? "";
      const rect = /\/Rect \[([^\]]*)\]/.exec(body)![1].split(" ").map(Number);
      out.push({ page: index + 1, rect, body });
    }
  });
  return out;
}

describe("HYPERLINK fields", () => {
  it("parse their URL and \\l bookmark, past the switches that take arguments", () => {
    expect(hyperlinkTarget('HYPERLINK "https://example.com/x" \\o "Tip here" \\t "_blank"')).toEqual({ href: "https://example.com/x" });
    expect(hyperlinkTarget(' HYPERLINK \\l "_Toc1" ')).toEqual({ anchor: "_Toc1" });
    expect(hyperlinkTarget('HYPERLINK "https://example.com/x" \\l frag')).toEqual({ href: "https://example.com/x", anchor: "frag" });
    expect(hyperlinkTarget("HYPERLINK https://example.com/bare")).toEqual({ href: "https://example.com/bare" });
    expect(hyperlinkTarget("PAGEREF last \\h")).toBeUndefined();
    expect(hyperlinkTarget("HYPERLINK")).toBeUndefined();
  });

  it("become hyperlink children whose runs are the field's result", () => {
    const doc = DocxDocument.load(linksDocx());
    const para = doc.sections[0].blocks[1];
    if (para.type !== "paragraph") throw new Error("paragraph expected");
    const links = para.children.filter((child) => child.type === "hyperlink");
    expect(links.map((link) => [link.href, link.anchor, link.runs.map((r) => r.content.map((c) => (c.kind === "text" ? c.text : c.kind)).join("")).join("|")])).toEqual([
      ["https://example.com/field", undefined, "field link"],
      [undefined, "last", "field anchor"],
    ]);
  });
});

describe("link annotations", () => {
  it("open safe URLs, go to bookmarks, skip unsafe schemes, and name destinations", () => {
    const bookmarks = new Map([["last", 1], ["a b/c", 0]]);
    expect(linkAction("https://example.com/a?x=1&y=2", bookmarks)).toBe("/A << /S /URI /URI (https://example.com/a?x=1&y=2) >>");
    expect(linkAction("mailto:someone@example.com", bookmarks)).toBe("/A << /S /URI /URI (mailto:someone@example.com) >>");
    expect(linkAction("https://example.com/ü (1)", bookmarks)).toBe("/A << /S /URI /URI (https://example.com/%C3%BC%20\\(1\\)) >>");
    expect(linkAction("javascript:alert(1)", bookmarks)).toBeUndefined();
    expect(linkAction("#last", bookmarks)).toBe("/Dest /last");
    expect(linkAction("#a b/c", bookmarks)).toBe("/Dest /a#20b#2fc");
    expect(linkAction("#missing", bookmarks)).toBeUndefined();
    expect(pdfName("_Toc12#(x)")).toBe("/_Toc12#23#28x#29");
  });

  it("cover each line of a link's text, with the page, and go to the bookmark's line on its page", async () => {
    const pdf = await render(linksDocx());
    const found = links(pdf);
    const byTarget = (needle: string) => found.filter((link) => link.body.includes(needle));
    // Page 1: "the site" and the wrapped link share the URL: one box on the
    // first line, one per line of the wrapped paragraph, each on its own line.
    const siteLinks = byTarget("/URI (https://example.com/a)");
    expect(siteLinks.length).toBeGreaterThanOrEqual(3);
    expect(siteLinks.every((link) => link.page === 1)).toBe(true);
    const tops = siteLinks.map((link) => link.rect[3]);
    expect(new Set(tops.map((top) => top.toFixed(1))).size).toBe(siteLinks.length);
    // The wrapped link's boxes follow one another down the page; the first spans the column.
    const [first, second] = siteLinks.slice(1);
    expect(first.rect[1]).toBeGreaterThan(second.rect[3] - 1);
    expect(first.rect[2] - first.rect[0]).toBeGreaterThan(300);
    expect(byTarget("/URI (https://example.com/field)")).toHaveLength(1);
    expect(byTarget("/Dest /last")).toHaveLength(2);
    expect(found.some((link) => link.body.includes("javascript"))).toBe(false);
    expect(found).toHaveLength(siteLinks.length + 3);
    // Every box is inside the page and as tall as a line.
    for (const link of found) {
      const [x0, y0, x1, y1] = link.rect;
      expect(x0).toBeGreaterThan(70);
      expect(x1).toBeLessThan(550);
      expect(y1 - y0).toBeGreaterThan(8);
      expect(y1 - y0).toBeLessThan(30);
      expect(link.body).toContain("/Border [0 0 0]");
    }
    // The destination: the second page's object, at the top of the bookmark's line (below the page's first paragraph).
    const pageObjects = [...pdf.matchAll(/(\d+) 0 obj\n<< \/Type \/Page /g)].map((m) => Number(m[1]));
    expect(pageObjects).toHaveLength(2);
    const dests = /\/Dests (\d+) 0 R/.exec(pdf);
    expect(dests).not.toBeNull();
    const dict = new RegExp(`${dests![1]} 0 obj\\n<< (.*) >>`).exec(pdf)![1];
    const last = /\/last \[(\d+) 0 R \/XYZ null ([\d.]+) null\]/.exec(dict)!;
    expect(Number(last[1])).toBe(pageObjects[1]);
    // The body starts 72pt below the 792pt page's top; the bookmark's line is the second.
    const top = Number(last[2]);
    expect(top).toBeLessThan(720 - 10);
    expect(top).toBeGreaterThan(720 - 40);
  });
});
