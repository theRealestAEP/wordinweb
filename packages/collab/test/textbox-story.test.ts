import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, serializeXml, localName, type XmlElement } from "@wordinweb/core";
import { DocumentSession } from "../src/session.js";

/**
 * Text-box story editing over the wire (checkpoint: the "I can't change the
 * number column on pleading paper" report).
 *
 * Pleading paper does NOT draw its margin numbers with w:lnNumType. Word's
 * pleading templates put a text box in the HEADER and type the numbers into
 * it, so the gutter is a w:txbxContent story — 28 numbered paragraphs riding
 * in header1.xml. The real Word-authored fixture used to diagnose this
 * (wordinweb-parity parity/pleading-anon.docx) has no lnNumType anywhere.
 *
 * The defect was the exact sibling of the header/footer one pinned in
 * header-footer.test.ts, one level deeper. Story paragraphs carry stable ids
 * (they ride in doc.editableRoots(), which is what StableIds walks) and the
 * editor lets a caret into them, but the apply's run index was built by
 * walking section/header/footer BLOCKS — and a text box's story hangs off the
 * RUN that carries the drawing (run.content anchor -> shape.blocks), which
 * that walk never descends into. So every story run was unresolvable,
 * resolveCaret returned null, and every text intent aimed at the gutter was a
 * clean reject on the server and on every peer — while the originating
 * client, which applies locally before emitting, kept the edit and then had
 * it rolled back by the confirmed state. The user sees keystrokes vanish.
 *
 * The markup below is copied VERBATIM out of the real fixture's header1.xml
 * (trimmed to three numbers). It is Word's own output, not something this
 * project's writer produced, so the test cannot pass by agreeing with our own
 * serializer — which a synthetic text box would have done.
 */

const STORY_PARA = (n: string) =>
  `<w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/><w:jc w:val="right"/></w:pPr><w:r><w:t>${n}</w:t></w:r></w:p>`;

const PLEADING_HEADER =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ` +
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:o="urn:schemas-microsoft-com:office:office" ` +
  `xmlns:v="urn:schemas-microsoft-com:vml" ` +
  `xmlns:w10="urn:schemas-microsoft-com:office:word"><w:p><w:r><w:pict>` +
  `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">` +
  `<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
  `<v:shape id="LineNumbers" type="#_x0000_t202" style="position:absolute;margin-left:-47.15pt;margin-top:0;` +
  `width:36pt;height:669.6pt;z-index:251657216;visibility:visible;mso-position-horizontal:absolute;` +
  `mso-position-horizontal-relative:margin;mso-position-vertical:absolute;` +
  `mso-position-vertical-relative:margin;v-text-anchor:top" stroked="f">` +
  `<v:textbox inset="0,0,0,0"><w:txbxContent>` +
  STORY_PARA("1") + STORY_PARA("2") + STORY_PARA("3") +
  `</w:txbxContent></v:textbox></v:shape>` +
  `<v:line id="MarginRule" style="position:absolute;z-index:3" from="48pt,24pt" to="48pt,624pt" ` +
  `strokecolor="#AA0000" strokeweight="1pt"><v:stroke dashstyle="solid"/></v:line>` +
  `</w:pict></w:r></w:p></w:hdr>`;

const DOCUMENT = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:body><w:p><w:r><w:t>Complaint body</w:t></w:r></w:p>` +
  `<w:sectPr><w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdH"/>` +
  `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;

function pleadingDoc(): DocxDocument {
  return DocxDocument.load(zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/_rels/document.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`),
    "word/document.xml": strToU8(DOCUMENT),
    "word/header1.xml": strToU8(PLEADING_HEADER),
  }));
}

/** The header part root. */
function hdrRoot(doc: DocxDocument): XmlElement {
  const root = doc.editableRoots().find((r) => localName(r.name) === "hdr");
  if (!root) throw new Error("no header part");
  return root;
}

/** The digits actually inside w:txbxContent, in order — the number column as
 * the saved file carries it. Reading the whole header would also match text
 * that leaked OUTSIDE the box, which is a different bug. */
function gutter(doc: DocxDocument): string[] {
  const xml = serializeXml(hdrRoot(doc));
  const story = xml.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/)?.[0] ?? "";
  // `<w:t[^>]*>` would also match <w:txbxContent> — require a space or the
  // closing bracket right after the tag name.
  return [...story.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => m[1]);
}

/** Number of paragraphs inside the text box — one per printed line number. */
function storyParagraphCount(doc: DocxDocument): number {
  const xml = serializeXml(hdrRoot(doc));
  const story = xml.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/)?.[0] ?? "";
  return [...story.matchAll(/<w:p(?:\s[^>]*)?>/g)].length;
}

/** Stable ids of the paragraph + run holding the Nth number in the gutter. */
function gutterAddress(s: DocumentSession, index: number): { blockId: number; runId: number } {
  const paras: XmlElement[] = [];
  const walk = (el: XmlElement, inStory: boolean): void => {
    const name = localName(el.name);
    const story = inStory || name === "txbxContent";
    if (story && name === "p") paras.push(el);
    for (const c of el.children) walk(c, story);
  };
  walk(hdrRoot(s.doc), false);
  const para = paras[index];
  if (!para) throw new Error(`no story paragraph ${index}`);
  const run = para.children.find((c) => localName(c.name) === "r");
  if (!run) throw new Error("story paragraph has no run");
  const blockId = s.ids.idOf(para);
  const runId = s.ids.idOf(run);
  // Addressability was never the defect — these ids have always existed.
  // Asserting it here keeps the test honest about WHICH half broke.
  if (blockId === undefined || runId === undefined) throw new Error("story is not id-tracked");
  return { blockId, runId };
}

function gutterDrawingRunId(s: DocumentSession): number {
  let result: number | undefined;
  const containsGutter = (el: XmlElement): boolean =>
    (localName(el.name) === "shape" && el.attrs.id === "LineNumbers") || el.children.some(containsGutter);
  const walk = (el: XmlElement): void => {
    if (localName(el.name) === "r" && containsGutter(el)) result = s.ids.idOf(el);
    for (const child of el.children) walk(child);
  };
  walk(hdrRoot(s.doc));
  if (result === undefined) throw new Error("gutter carrier run is not id-tracked");
  return result;
}

describe("pleading number column (header text-box story) is editable over the wire", () => {
  it("the story paragraphs and runs carry stable ids", () => {
    const s = new DocumentSession(pleadingDoc());
    expect(gutter(s.doc)).toEqual(["1", "2", "3"]);
    expect(gutterAddress(s, 0).runId).toEqual(expect.any(Number));
  });

  it("insertText into the gutter APPLIES (it used to reject on every replica)", () => {
    const s = new DocumentSession(pleadingDoc());
    const at = { ...gutterAddress(s, 1), offset: 0 };
    const e = s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at, text: "9" });
    expect(e.kind).toBe("applied");
    expect(gutter(s.doc)).toEqual(["1", "92", "3"]);
  });

  it("deleteText in the gutter applies too (the whole text path shared one index)", () => {
    const s = new DocumentSession(pleadingDoc());
    const addr = gutterAddress(s, 2);
    const e = s.submit({
      kind: "deleteText", clientId: "a", clientSeq: 1, base: 0,
      blockId: addr.blockId, runId: addr.runId, start: 0, end: 1,
    });
    expect(e.kind).toBe("applied");
    // Emptying the only w:t drops that element, so the third number is gone —
    // but its paragraph stays, which is what keeps the column's line spacing.
    expect(gutter(s.doc)).toEqual(["1", "2"]);
    expect(storyParagraphCount(s.doc)).toBe(3);
  });

  it("the edit lands INSIDE the text box, leaving the body and the shape intact", () => {
    const s = new DocumentSession(pleadingDoc());
    const at = { ...gutterAddress(s, 0), offset: 0 };
    expect(s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at, text: "A" }).kind).toBe("applied");
    const saved = DocxDocument.load(s.doc.save());
    const header = saved.pkg.text("word/header1.xml") ?? "";
    // Still Word's own shape, still the same story container.
    expect(header).toContain('id="LineNumbers"');
    expect(header).toContain("margin-left:-47.15pt");
    expect(header.match(/<w:txbxContent>[\s\S]*?<\/w:txbxContent>/)?.[0]).toContain("<w:t>A1</w:t>");
    expect(saved.pkg.text("word/document.xml") ?? "").toContain("Complaint body");
    expect(gutter(saved)).toEqual(["A1", "2", "3"]);
  });

  it("two replicas of the same intent converge byte-for-byte (the fork this closes)", () => {
    // The defect was ASYMMETRIC — the originator applied locally while every
    // peer rejected — so identity of two independent applies is the check
    // that actually sees it.
    const a = new DocumentSession(pleadingDoc());
    const b = new DocumentSession(pleadingDoc());
    const intent = (s: DocumentSession) =>
      ({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at: { ...gutterAddress(s, 0), offset: 0 }, text: "7" }) as const;
    expect(a.submit(intent(a)).kind).toBe("applied");
    expect(b.submit(intent(b)).kind).toBe("applied");
    expect(gutter(a.doc)).toEqual(["71", "2", "3"]);
    expect(serializeXml(hdrRoot(a.doc))).toBe(serializeXml(hdrRoot(b.doc)));
    expect(Buffer.from(a.doc.save())).toEqual(Buffer.from(b.doc.save()));
  });

  it("moves, resizes, and rotates the VML gutter through drawing intents", () => {
    const a = new DocumentSession(pleadingDoc());
    const b = new DocumentSession(pleadingDoc());
    const intents = [
      { kind: "setFloatingPagePosition", clientId: "a", clientSeq: 1, base: 0, runId: gutterDrawingRunId(a), xPx: 120, yPx: 240 },
      { kind: "setDrawingRotation", clientId: "a", clientSeq: 2, base: 1, runId: gutterDrawingRunId(a), degrees: 45 },
      { kind: "resizeDrawing", clientId: "a", clientSeq: 3, base: 2, runId: gutterDrawingRunId(a), widthPx: 72, heightPx: 900 },
    ] as const;
    const bIntents = intents.map((intent) => ({ ...intent, runId: gutterDrawingRunId(b) }));

    expect(a.submit(intents[0]).kind).toBe("applied");
    expect(a.submit(intents[1]).kind).toBe("applied");
    expect(a.submit(intents[2]).kind).toBe("applied");
    expect(b.submit(bIntents[0]).kind).toBe("applied");
    expect(b.submit(bIntents[1]).kind).toBe("applied");
    expect(b.submit(bIntents[2]).kind).toBe("applied");

    const header = serializeXml(hdrRoot(a.doc));
    expect(header).toContain("margin-left:90pt");
    expect(header).toContain("margin-top:180pt");
    expect(header).toContain("mso-position-horizontal-relative:page");
    expect(header).toContain("rotation:45");
    expect(header).toContain("width:54pt");
    expect(header).toContain("height:675pt");
    expect(header).toBe(serializeXml(hdrRoot(b.doc)));
  });

  it("edits the exact VML object when one run carries a text box and a line", () => {
    const s = new DocumentSession(pleadingDoc());
    const parsed = s.doc.headers.get("rIdH")?.blocks[0];
    if (!parsed || parsed.type !== "paragraph" || parsed.children[0].type !== "run") throw new Error("header run missing");
    expect(parsed.children[0].content.map((content) => content.kind)).toEqual(["anchor", "anchor"]);
    const e = s.submit({
      kind: "setDrawingLineStyle", clientId: "a", clientSeq: 1, base: 0,
      runId: gutterDrawingRunId(s), objectIndex: 1, color: "156082", widthPx: 3, dash: "dotted",
    });
    expect(e.kind).toBe("applied");
    const header = serializeXml(hdrRoot(s.doc));
    expect(header).toContain('id="LineNumbers"');
    expect(header).toContain('stroked="f"');
    expect(header).toContain('id="MarginRule"');
    expect(header).toContain('strokecolor="156082"');
    expect(header).toContain('dashstyle="dot"');
  });

  it("an out-of-range offset in the gutter is still a clean reject", () => {
    // The fix widens what RESOLVES; it must not widen what is ACCEPTED.
    const s = new DocumentSession(pleadingDoc());
    const at = { ...gutterAddress(s, 0), offset: 99 };
    expect(s.submit({ kind: "insertText", clientId: "a", clientSeq: 1, base: 0, at, text: "x" }).kind).toBe("rejected");
    expect(gutter(s.doc)).toEqual(["1", "2", "3"]);
  });
});
