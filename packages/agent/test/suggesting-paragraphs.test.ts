import { describe, expect, it } from "vitest";
import { DocxDocument, acceptAllRevisions, localName, rejectAllRevisions, serializeXml, type XmlElement } from "@wordinweb/core";
import { AgentDocument, LocalDocumentSession } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";

/**
 * The LikeOffice AI panel path: an agent connected to a LocalDocumentSession
 * over a blank document, suggesting mode on, `suggest: true` injected into
 * the operations. Creating real paragraphs must work on every route a model
 * would take (this was the "editor wouldn't let me create new paragraphs"
 * bug: a suggested insertion moves its text into a new w:ins run, so a
 * follow-up splitParagraph addressed the dead run and the whole patch was
 * rejected).
 */

const BLANK = body(`<w:p><w:r><w:t></w:t></w:r></w:p>`);

function connect(xml = BLANK) {
  const session = new LocalDocumentSession(makeDocx(xml));
  const agent = AgentDocument.connect(session, { provenance: { author: "AI", now: () => "2026-08-06T00:00:00Z" } });
  const tool = (name: string) => agent.tools().find((candidate) => candidate.name === name)!;
  return { session, agent, tool };
}

function paragraphTexts(doc: DocxDocument): string[] {
  const out: string[] = [];
  const textOf = (element: XmlElement): string =>
    (localName(element.name) === "t" ? element.text : "") + element.children.map(textOf).join("");
  const visit = (element: XmlElement): void => {
    if (localName(element.name) === "p") {
      out.push(textOf(element));
      return;
    }
    for (const child of element.children) visit(child);
  };
  visit(doc.docRoot);
  return out;
}

function copyOf(doc: DocxDocument): DocxDocument {
  return DocxDocument.load(doc.save());
}

describe("paragraph creation while suggesting", () => {
  it("authors a heading and two body paragraphs through word_document_edit", async () => {
    const { session, agent, tool } = connect();
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    await tool("word_document_edit").execute({
      revision: agent.revision,
      operations: [
        {
          kind: "insertText",
          at: { blockRef: paragraph.ref, runRef: paragraph.runs[0].ref, offset: 0 },
          text: "Declaration\nWhen in the course of human events.\nWe hold these truths.",
          suggest: true,
        },
        { kind: "formatParagraph", blockRef: paragraph.ref, styleId: "Heading1", suggest: true },
      ],
    });

    // Three real paragraphs, all of it tracked: text inside w:ins, the two
    // introduced paragraph marks carried as inserted glyphs.
    expect(paragraphTexts(session.doc)).toEqual(["Declaration", "When in the course of human events.", "We hold these truths."]);
    expect(serializeXml(session.doc.docRoot)).toContain("<w:ins ");

    const accepted = copyOf(session.doc);
    acceptAllRevisions(accepted);
    expect(paragraphTexts(accepted)).toEqual(["Declaration", "When in the course of human events.", "We hold these truths."]);
    const acceptedXml = serializeXml(accepted.docRoot);
    expect(acceptedXml).not.toContain("<w:ins ");
    expect(acceptedXml).toContain('w:val="Heading1"');

    const rejected = copyOf(session.doc);
    rejectAllRevisions(rejected);
    expect(paragraphTexts(rejected)).toEqual([""]);
    expect(serializeXml(rejected.docRoot)).not.toContain("Heading1");
  });

  it("creates real paragraphs from multi-line patch newText, heading over body", async () => {
    const { session, agent, tool } = connect();
    const projection = agent.project({ mode: "md" });
    const result = await tool("word_document_patch").execute({
      revision: projection.revision,
      mode: "md",
      edits: [{ startLine: 1, endLine: 1, newText: "# Declaration\nWhen in the course of human events.\nWe hold these truths." }],
      suggest: true,
    }) as { operations: string[] };
    expect(result.operations).toEqual(["splitParagraph", "splitParagraph", "insertText", "insertText", "insertText", "formatParagraph"]);
    expect(paragraphTexts(session.doc)).toEqual(["Declaration", "When in the course of human events.", "We hold these truths."]);

    const accepted = copyOf(session.doc);
    acceptAllRevisions(accepted);
    expect(paragraphTexts(accepted)).toEqual(["Declaration", "When in the course of human events.", "We hold these truths."]);
    // The heading landed on the first paragraph only.
    const chunks = serializeXml(accepted.docRoot).split("</w:p>");
    expect(chunks[0]).toContain('w:val="Heading1"');
    expect(chunks[1]).not.toContain('w:val="Heading1"');

    const rejected = copyOf(session.doc);
    rejectAllRevisions(rejected);
    expect(paragraphTexts(rejected)).toEqual([""]);
  });

  it("splits an existing paragraph with suggest through word_document_edit", async () => {
    const { session, agent, tool } = connect(body(`<w:p><w:r><w:t>HeadingBody</w:t></w:r></w:p>`));
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    await tool("word_document_edit").execute({
      revision: agent.revision,
      operations: [
        { kind: "splitParagraph", at: { blockRef: paragraph.ref, runRef: paragraph.runs[0].ref, offset: 7 }, suggest: true },
      ],
    });
    expect(paragraphTexts(session.doc)).toEqual(["Heading", "Body"]);
    const rejected = copyOf(session.doc);
    rejectAllRevisions(rejected);
    expect(paragraphTexts(rejected)).toEqual(["HeadingBody"]);
  });

  it("keeps pre-transaction run refs valid across a suggested insertion in one transaction", async () => {
    const { session, agent, tool } = connect(body(`<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>`));
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    const run = paragraph.runs[0];
    await tool("word_document_edit").execute({
      revision: agent.revision,
      operations: [
        { kind: "insertText", at: { blockRef: paragraph.ref, runRef: run.ref, offset: 11 }, text: " again", suggest: true },
        { kind: "formatRange", blockRef: paragraph.ref, runRef: run.ref, start: 0, end: 5, patch: { bold: true }, suggest: true },
      ],
    });
    expect(paragraphTexts(session.doc)).toEqual(["Hello world again"]);
    expect(serializeXml(session.doc.docRoot)).toContain("<w:b/>");
  });

  it("names the working alternative when a splitParagraph offset cannot resolve", async () => {
    const { agent, tool } = connect();
    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    await expect(tool("word_document_edit").execute({
      revision: agent.revision,
      operations: [
        { kind: "insertText", at: { blockRef: paragraph.ref, runRef: paragraph.runs[0].ref, offset: 0 }, text: "HeadingBody", suggest: true },
        { kind: "splitParagraph", at: { blockRef: paragraph.ref, runRef: paragraph.runs[0].ref, offset: 7 }, suggest: true },
      ],
    })).rejects.toThrow('insertText whose text contains "\\n"');
  });
});
