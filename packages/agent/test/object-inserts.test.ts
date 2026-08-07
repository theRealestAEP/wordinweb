import { describe, expect, it } from "vitest";
import { DocxDocument } from "@wordinweb/core";
import { AgentDocument } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";

// Every object insert the ribbon offers, through word_document_edit's execute
// path (AgentDocument.edit) on an EXISTING loaded document — not a fresh
// create() and not compose. Each insert is a hand-written collab intent
// (intents.ts), already replicated and validated at the collab layer; this
// file locks the agent-facing half: the payload the schema advertises is
// accepted, the XML lands, the save round-trips, and a bad reference rejects
// without mutating anything.

const PNG_1x1 = Uint8Array.from(Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
  "base64",
));

function loadedAgent(): { agent: AgentDocument; runRef: string } {
  const bytes = makeDocx(body('<w:p><w:r><w:t>Existing document text</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'));
  const agent = AgentDocument.load(bytes);
  const read = agent.inspect({ kind: "read", story: "body" });
  if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("no paragraph");
  return { agent, runRef: read.blocks[0].runs[0].ref };
}

describe("object inserts through AgentDocument.edit on an existing document", () => {
  it("insertShape (textBox preset)", async () => {
    const { agent, runRef } = loadedAgent();
    const result = await agent.edit({ revision: agent.revision, operations: [{ kind: "insertShape", runRef, preset: "textBox", text: "Boxed" }] });
    expect(result.status).toBe("applied");
    const saved = DocxDocument.load(agent.save());
    expect(saved.pkg.text("word/document.xml")).toContain("wps:wsp");
    expect(saved.pkg.text("word/document.xml")).toContain("Boxed");
  });

  it("insertMath (equation)", async () => {
    const { agent, runRef } = loadedAgent();
    await agent.edit({ revision: agent.revision, operations: [{ kind: "insertMath", runRef, mathText: "x^2 + y^2 = z^2" }] });
    expect(DocxDocument.load(agent.save()).pkg.text("word/document.xml")).toContain("m:oMath");
  });

  it("insertChart", async () => {
    const { agent, runRef } = loadedAgent();
    await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertChart", runRef, chart: { type: "column", title: "T", categories: ["Q1", "Q2"], series: [{ name: "S", values: [1, 2] }] } }],
    });
    const saved = DocxDocument.load(agent.save());
    const chartPart = saved.pkg.names().find((name) => /word\/charts\/chart\d+\.xml/.test(name));
    expect(chartPart).toBeTruthy();
  });

  it("insertSmartArt", async () => {
    const { agent, runRef } = loadedAgent();
    await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertSmartArt", runRef, smartArt: { layout: "cycle", items: ["Plan", "Do", "Check"] } }],
    });
    const saved = DocxDocument.load(agent.save());
    expect(saved.pkg.names().some((name) => name.startsWith("word/diagrams/"))).toBe(true);
  });

  it("insertImage from a registered asset (local path)", async () => {
    const { agent, runRef } = loadedAgent();
    const assetRef = agent.addAsset(PNG_1x1, "image/png");
    await agent.edit({ revision: agent.revision, operations: [{ kind: "insertImage", runRef, assetRef, widthPx: 24, heightPx: 24 }] });
    const saved = DocxDocument.load(agent.save());
    expect(saved.pkg.names().some((name) => name.startsWith("word/media/"))).toBe(true);
    expect(saved.pkg.text("word/document.xml")).toContain("w:drawing");
  });

  it("insertFootnote", async () => {
    const { agent, runRef } = loadedAgent();
    await agent.edit({ revision: agent.revision, operations: [{ kind: "insertFootnote", runRef, text: "A footnote body" }] });
    const saved = DocxDocument.load(agent.save());
    expect(saved.pkg.text("word/footnotes.xml")).toContain("A footnote body");
  });

  it("insertBookmark and insertBreak and insertSectionBreak", async () => {
    const { agent, runRef } = loadedAgent();
    await agent.edit({ revision: agent.revision, operations: [{ kind: "insertBookmark", runRef, name: "Anchor1" }] });
    await agent.edit({ revision: agent.revision, operations: [{ kind: "insertBreak", runRef, breakKind: "page" }] });
    const read = agent.inspect({ kind: "read", story: "body" });
    if (!("blocks" in read) || read.blocks[1].type !== "paragraph") throw new Error("no paragraph");
    const freshRef = read.blocks[1].runs[0].ref;
    await agent.edit({ revision: agent.revision, operations: [{ kind: "insertSectionBreak", runRef: freshRef, breakType: "nextPage" }] });
    const xml = DocxDocument.load(agent.save()).pkg.text("word/document.xml");
    expect(xml).toContain('w:name="Anchor1"');
    expect(xml).toContain('w:type="page"');
    expect(xml).toContain("w:sectPr");
  });

  it("rejects a bad run reference cleanly without mutating", async () => {
    const { agent } = loadedAgent();
    const before = agent.save();
    await expect(
      agent.edit({ revision: agent.revision, operations: [{ kind: "insertShape", runRef: "run:99999", preset: "rectangle" }] }),
    ).rejects.toThrow();
    expect(Buffer.from(agent.save()).equals(Buffer.from(before))).toBe(true);
  });
});
