import { describe, expect, it } from "vitest";
import { DocxDocument } from "@wordinweb/core";
import { AgentDocument } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";

// A write that throws applies nothing: edit, patch, and compose each build a
// trial clone and adopt it only after the whole request succeeds. The error
// text has to say that. It used to end "Re-inspect and send the remaining
// operations in a new transaction", which reads as though the operations
// before the failing one had landed. A model that believes it re-sends only
// the rest and loses the earlier ones, or re-inspects, sees an object it
// cannot read, and inserts a second copy.

const CHART = {
  type: "column" as const,
  title: "Revenue",
  categories: ["Q1", "Q2", "Q3", "Q4"],
  series: [{ name: "Revenue", values: [10, 12, 9, 15] }],
};

/** One empty paragraph — the blank document the object-insert benchmark task
 * starts from, which is where this failure was found. */
function blank(): { agent: AgentDocument; blockRef: string; runRef: string } {
  const bytes = makeDocx(body('<w:p><w:r><w:t></w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'));
  const agent = AgentDocument.load(bytes);
  const read = agent.inspect({ kind: "read", story: "body" });
  if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("no paragraph");
  return { agent, blockRef: read.blocks[0].ref, runRef: read.blocks[0].runs[0].ref };
}

/** The state an assertion can compare: the markup, the parts, the revision. */
function snapshot(agent: AgentDocument): { xml: string; parts: string[]; revision: string } {
  const saved = DocxDocument.load(agent.save());
  return { xml: saved.pkg.text("word/document.xml") ?? "", parts: saved.pkg.names().sort(), revision: agent.revision };
}

function chartParts(agent: AgentDocument): string[] {
  return DocxDocument.load(agent.save()).pkg.names().filter((name) => name.startsWith("word/charts/"));
}

/** An insertChart that succeeds on its own, followed by a splitParagraph the
 * engine cannot place. The second operation is what throws, so the first one
 * is the operation a caller would wrongly believe survived. */
function chartThenBadSplit(agent: AgentDocument, blockRef: string, runRef: string): Promise<unknown> {
  return agent.edit({
    revision: agent.revision,
    operations: [
      { kind: "insertChart", runRef, chart: CHART },
      { kind: "splitParagraph", at: { blockRef, runRef, offset: 5 }, suggest: true },
    ],
  });
}

describe("a failed write says the document is unchanged", () => {
  it("edit: the message states nothing applied and never asks for the remaining operations", async () => {
    const { agent, blockRef, runRef } = blank();
    const error = await chartThenBadSplit(agent, blockRef, runRef).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(Error);
    const message = (error as Error).message;
    expect(message).toContain("NOTHING was applied");
    expect(message).toContain("the document is unchanged");
    expect(message).toContain("no earlier operation in this transaction landed either");
    expect(message).toContain("Re-inspect the document, then send every operation again");
    expect(message).not.toContain("remaining operations");
    expect(message).not.toContain("new transaction");
    // The diagnosis and the working alternative survive alongside the truth.
    expect(message).toContain("splitParagraph could not apply");
    expect(message).toContain('insertText whose text contains "\\n"');
  });

  it("edit: the failed transaction mutates nothing", async () => {
    // The same insertChart on its own does land, so an absent chart part
    // after the failure is the rollback and not an inert operation.
    const alone = blank();
    await alone.agent.edit({ revision: alone.agent.revision, operations: [{ kind: "insertChart", runRef: alone.runRef, chart: CHART }] });
    expect(chartParts(alone.agent).length).toBeGreaterThan(0);

    const { agent, blockRef, runRef } = blank();
    const before = snapshot(agent);
    await chartThenBadSplit(agent, blockRef, runRef).catch(() => {});
    expect(snapshot(agent)).toEqual(before);
    expect(chartParts(agent)).toEqual([]);
    expect(agent.revision).toBe("0");
  });

  it("edit: a rejection before any operation runs carries the same truth", async () => {
    const { agent } = blank();
    await expect(agent.edit({ revision: "stale", operations: [{ kind: "acceptAllRevisions" }] }))
      .rejects.toThrow("NOTHING was applied");
  });

  it("patch: the message speaks of hunks and re-projecting, not of operations", async () => {
    const bytes = makeDocx(body('<w:p><w:r><w:t>Existing document text</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'));
    const agent = AgentDocument.load(bytes);
    const projection = agent.project({ story: "body", mode: "md" });
    const before = snapshot(agent);
    const error = await agent.patch({
      revision: projection.revision,
      story: "body",
      mode: "md",
      edits: [{ startLine: 1, endLine: 99, newText: "rewritten" }],
    }).catch((e: unknown) => e);
    const message = (error as Error).message;
    expect(message).toContain("NOTHING was applied");
    expect(message).toContain("no earlier hunk in this patch landed either");
    expect(message).toContain("Re-project the story, then send every edit again");
    expect(message).not.toContain("send every operation again");
    expect(snapshot(agent)).toEqual(before);
  });

  it("compose: the message says nothing was created and asks for the whole request", async () => {
    const agent = AgentDocument.create();
    const before = snapshot(agent);
    const error = await agent.compose({
      revision: agent.revision,
      body: [{ type: "table", rows: [["A", "B"], ["C"]] }],
    }).catch((e: unknown) => e);
    const message = (error as Error).message;
    expect(message).toContain("NOTHING was created");
    expect(message).toContain("Send the whole compose request again");
    expect(snapshot(agent)).toEqual(before);
  });
});
