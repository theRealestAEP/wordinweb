import { describe, expect, it } from "vitest";
import { DocxDocument } from "@wordinweb/core";
import { applyIntentScoped, resyncScope, type Intent } from "@wordinweb/collab/client";
import { AgentDocument, DOCMD_ATOMS } from "../src/index.js";
import type { AgentCollaborativeTarget, AgentProjectResult } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";

/** A brief with a heading, a mixed-formatting paragraph, a bullet list, a
 * table, and an equation — one of every projection shape. */
async function brief(): Promise<AgentDocument> {
  const agent = AgentDocument.create();
  await agent.compose({
    revision: agent.revision,
    body: [
      { type: "heading", level: 1, text: "Findings" },
      { type: "paragraph", runs: [{ text: "Decision: ", bold: true }, { text: "adopt the managed platform." }] },
      { type: "paragraph", text: "Latency", list: "bullet" },
      { type: "paragraph", text: "Cost", list: "bullet" },
      { type: "table", headerRows: 1, rows: [["Option", "Score"], ["Managed", "9"]] },
      { type: "equation", mathText: "E=mc^2" },
      { type: "heading", level: 2, text: "Next steps" },
      { type: "paragraph", text: "Schedule the review." },
    ],
  });
  return agent;
}

function documentXml(agent: AgentDocument): string {
  return DocxDocument.load(agent.save()).pkg.text("word/document.xml");
}

function identityEdits(projection: AgentProjectResult) {
  return projection.text.split("\n")
    .map((newText, index) => ({ startLine: index + 1, endLine: index + 1, newText }))
    .filter((_, index) => projection.anchors[index].role === "paragraph");
}

function lineOf(projection: AgentProjectResult, text: string): number {
  const index = projection.text.split("\n").indexOf(text);
  if (index < 0) throw new Error(`the projection has no line "${text}"`);
  return index + 1;
}

describe("DocMD projection", () => {
  it("renders headings, lists, tables, and equations deterministically", async () => {
    const agent = await brief();
    const md = agent.project({ mode: "md" });
    expect(md.text).toBe([
      "# Findings",
      "Decision: adopt the managed platform.",
      "- Latency",
      "- Cost",
      "",
      "| Option | Score |",
      "| --- | --- |",
      "| Managed | 9 |",
      "",
      "$E=mc^(2)$",
      "## Next steps",
      "Schedule the review.",
    ].join("\n"));
    expect(md).toMatchObject({ revision: "1", story: "body", mode: "md", lines: 12, truncated: false });

    const again = agent.project({ mode: "md" });
    expect(again.text).toBe(md.text);
    expect(JSON.stringify(again.anchors)).toBe(JSON.stringify(md.anchors));

    // text mode is the document as a text file: cell paragraphs are lines of
    // their own and every non-text atom is one character.
    const text = agent.project({ mode: "text" });
    expect(text.text).toBe([
      "Findings",
      "Decision: adopt the managed platform.",
      "Latency",
      "Cost",
      "Option",
      "Score",
      "Managed",
      "9",
      DOCMD_ATOMS.math,
      "Next steps",
      "Schedule the review.",
    ].join("\n"));

    expect(agent.project({ mode: "outline" }).text).toBe("# Findings\n## Next steps");
  });

  it("maps every projection line to the block and runs it came from", async () => {
    const agent = await brief();
    const md = agent.project({ mode: "md" });
    expect(md.anchors.map((anchor) => anchor.role)).toEqual([
      "paragraph", "paragraph", "paragraph", "paragraph",
      "structure", "table", "table", "table", "structure",
      "paragraph", "paragraph", "paragraph",
    ]);
    expect(md.anchors[0]).toMatchObject({ line: 1, blockRef: expect.stringMatching(/^block:/), marker: 2, editable: true });
    // The mixed paragraph keeps one segment per run, and both are editable
    // because both came verbatim from a w:t.
    expect(md.anchors[1].segments).toEqual([
      { start: 0, end: 10, runRef: expect.stringMatching(/^run:/), wireStart: 0, wireEnd: 10, editable: true },
      { start: 10, end: 37, runRef: expect.stringMatching(/^run:/), wireStart: 0, wireEnd: 27, editable: true },
    ]);
    // The equation is one opaque atom: it occupies columns 0-10 and no column
    // inside it is addressable, though the empty run after it still is.
    const equation = md.anchors[9];
    expect(equation.segments[0]).toEqual({ start: 0, end: 10, runRef: "", wireStart: 0, wireEnd: 0, editable: false });
    expect(equation.segments.filter((segment) => segment.editable).every((segment) => segment.start === segment.end)).toBe(true);
    // Table and separator lines carry identity for context but no segments.
    expect(md.anchors[5]).toMatchObject({ role: "table", blockRef: expect.stringMatching(/^block:/), editable: false, segments: [] });
    expect(md.anchors[4]).toMatchObject({ role: "structure", editable: false, segments: [] });
  });

  it("anchors the same runs and wire offsets in text and md modes", async () => {
    const agent = await brief();
    const md = agent.project({ mode: "md" });
    const text = agent.project({ mode: "text" });
    const editableByBlock = (projection: AgentProjectResult) => new Map(projection.anchors
      .filter((anchor) => anchor.role === "paragraph" && anchor.blockRef)
      .map((anchor) => [anchor.blockRef!, anchor.segments
        .filter((segment) => segment.editable)
        .map((segment) => `${segment.runRef}:${segment.wireStart}-${segment.wireEnd}`)]));

    const mdBlocks = editableByBlock(md);
    const textBlocks = editableByBlock(text);
    // md mode adds markers and skips table cells, but wherever both modes
    // project a paragraph they address exactly the same wire spans.
    const shared = [...mdBlocks.keys()].filter((ref) => textBlocks.has(ref));
    expect(shared.length).toBeGreaterThan(3);
    for (const ref of shared) expect(textBlocks.get(ref), ref).toEqual(mdBlocks.get(ref));
    // The marker is the only column shift md mode introduces.
    for (const anchor of md.anchors) {
      if (anchor.role !== "paragraph") continue;
      for (const segment of anchor.segments) expect(segment.start).toBeGreaterThanOrEqual(anchor.marker);
    }
  });

  it("windows a long story by cursor and stamps every window with the revision", async () => {
    const agent = AgentDocument.create();
    await agent.compose({
      revision: agent.revision,
      body: Array.from({ length: 60 }, (_, index) => ({ type: "paragraph" as const, text: `Paragraph ${index + 1}` })),
    });

    const whole = agent.project({ mode: "text" });
    expect(whole.truncated).toBe(false);
    expect(whole.lines).toBe(60);

    const pages: string[] = [];
    let cursor: { value: string } | undefined;
    let windows = 0;
    do {
      const page = agent.project({ mode: "text", cursor, maxBlocks: 25 });
      expect(page.revision).toBe(agent.revision);
      expect(page.window.cursor).toBe(cursor?.value ?? null);
      pages.push(page.text);
      cursor = page.next;
      windows++;
    } while (cursor);
    expect(windows).toBe(3);
    expect(pages.join("\n")).toBe(whole.text);
    expect(() => agent.project({ mode: "text", cursor: { value: "docmd:99:0" } })).toThrow("stale cursor");
  });

  it("renders an image as a live object reference in md mode", async () => {
    const agent = AgentDocument.create();
    const logo = agent.addAsset(Uint8Array.from(Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC",
      "base64",
    )), "image/png");
    await agent.compose({
      revision: agent.revision,
      body: [{ type: "image", assetRef: logo, widthPx: 96, heightPx: 24, alt: "Company mark" }],
    });
    const md = agent.project({ mode: "md" });
    const ref = /!\[Company mark\]\((object:\d+:\d+)\)/.exec(md.text)?.[1];
    expect(ref, md.text).toBeTruthy();
    // The reference in the projection is the one the object tools accept.
    expect(agent.inspect({ kind: "object", ref: ref! })).toMatchObject({ type: "image", detail: { width: 96, height: 24 } });
    expect(agent.project({ mode: "text" }).text).toContain(DOCMD_ATOMS.object);
  });

  it("escapes paragraph text that would otherwise read as a markdown marker", async () => {
    const agent = AgentDocument.create();
    await agent.compose({
      revision: agent.revision,
      body: [
        // Numbered contract headings and literal dashes are ordinary text, not
        // list structure, and a patch must not mistake one for the other.
        { type: "paragraph", text: "3. TERM AND TERMINATION" },
        { type: "paragraph", text: "- not a bullet" },
        { type: "paragraph", text: "# not a heading" },
        { type: "paragraph", text: "Ordinary text" },
        { type: "paragraph", text: "Real bullet", list: "bullet" },
      ],
    });
    const md = agent.project({ mode: "md" });
    expect(md.text).toBe("\\3. TERM AND TERMINATION\n\\- not a bullet\n\\# not a heading\nOrdinary text\n- Real bullet");
    expect(md.anchors.map((anchor) => anchor.marker)).toEqual([1, 1, 1, 0, 2]);
    // text mode carries no markdown, so it carries no escapes either.
    expect(agent.project({ mode: "text" }).text).toBe("3. TERM AND TERMINATION\n- not a bullet\n# not a heading\nOrdinary text\nReal bullet");

    const identity = await agent.patch({ revision: md.revision, mode: "md", edits: identityEdits(md) });
    expect(identity.operations).toEqual([]);
    // The escape is structure, so editing behind it edits the document text.
    const edited = await agent.patch({
      revision: md.revision,
      mode: "md",
      edits: [{ startLine: 1, endLine: 1, newText: "\\3. TERM, TERMINATION, AND RENEWAL" }],
    });
    expect(edited.projection.text.split("\n")[0]).toBe("\\3. TERM, TERMINATION, AND RENEWAL");
    expect(agent.project({ mode: "text" }).text.split("\n")[0]).toBe("3. TERM, TERMINATION, AND RENEWAL");
  });

  it("projects header and footer stories under the ids the overview reports", async () => {
    const agent = AgentDocument.create();
    await agent.compose({
      revision: agent.revision,
      body: [{ type: "paragraph", text: "Body" }],
      header: [{ type: "paragraph", text: "Confidential" }],
      footer: [{ type: "pageNumber", fieldKind: "page" }],
    });
    const overview = agent.inspect({ kind: "overview" });
    if (!("stories" in overview)) throw new Error("overview missing");
    const header = overview.stories.find((story) => story.kind === "header")!;
    const footer = overview.stories.find((story) => story.kind === "footer")!;
    expect(agent.project({ story: header.id, mode: "text" }).text).toBe("Confidential");
    // A field is one atom in text mode and its instruction name in md mode.
    expect(agent.project({ story: footer.id, mode: "text" }).text).toBe(DOCMD_ATOMS.field);
    expect(agent.project({ story: footer.id, mode: "md" }).text).toBe("{{PAGE}}");
    expect(() => agent.project({ story: "nope" })).toThrow("Unknown story");
  });
});

describe("DocMD patch", () => {
  it("is a fixed point under an identity patch", async () => {
    const agent = await brief();
    const before = agent.project({ mode: "md" });
    const result = await agent.patch({ revision: before.revision, mode: "md", edits: identityEdits(before) });
    expect(result).toMatchObject({ revision: before.revision, status: "applied", operations: [] });
    expect(result.projection.text).toBe(before.text);
    expect(JSON.stringify(agent.project({ mode: "md" }).anchors)).toBe(JSON.stringify(before.anchors));
    expect(agent.revision).toBe(before.revision);
  });

  it("edits text inside a styled run without touching its formatting", async () => {
    const agent = await brief();
    const before = agent.project({ mode: "md" });
    const boldBefore = documentXml(agent).match(/<w:b\/>/g)?.length;
    const line = lineOf(before, "Decision: adopt the managed platform.");
    const result = await agent.patch({
      revision: before.revision,
      mode: "md",
      edits: [{ startLine: line, endLine: line, newText: "Decision: adopt the custom platform." }],
    });
    expect(result.operations).toEqual(["deleteText", "insertText"]);
    expect(result.projection.text.split("\n")[line - 1]).toBe("Decision: adopt the custom platform.");

    const xml = documentXml(agent);
    // The bold lead-in is untouched, and the rewritten words stayed in the
    // plain run rather than inheriting the bold.
    expect(xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Decision: </w:t></w:r>');
    expect(xml).toContain('<w:r><w:t xml:space="preserve">adopt the custom platform.</w:t></w:r>');
    expect(xml.match(/<w:b\/>/g)?.length).toBe(boldBefore);
  });

  it("rewrites a line that differs inside an astral character", async () => {
    const agent = AgentDocument.create();
    await agent.compose({ revision: agent.revision, body: [{ type: "paragraph", text: "score \u{10000} today" }] });
    const projection = agent.project({ mode: "text" });
    const result = await agent.patch({
      revision: projection.revision,
      mode: "text",
      edits: [{ startLine: 1, endLine: 1, newText: "score \u{10400} today" }],
    });
    expect(result.projection.text).toBe("score \u{10400} today");
    expect([...result.projection.text]).toHaveLength(13);
  });

  it("splits, merges, and retypes paragraph structure through existing intents", async () => {
    const split = await brief();
    let projection = split.project({ mode: "md" });
    let line = lineOf(projection, "Schedule the review.");
    let result = await split.patch({
      revision: projection.revision,
      mode: "md",
      edits: [{ startLine: line, endLine: line, newText: "Schedule the review.\nName an owner." }],
    });
    expect(result.operations).toEqual(["insertText", "splitParagraph"]);
    expect(result.projection.text.split("\n").slice(-2)).toEqual(["Schedule the review.", "Name an owner."]);

    // A pure split rewrites no text at all, so every run keeps its formatting.
    const pure = await brief();
    projection = pure.project({ mode: "md" });
    line = lineOf(projection, "Decision: adopt the managed platform.");
    result = await pure.patch({
      revision: projection.revision,
      mode: "md",
      edits: [{ startLine: line, endLine: line, newText: "Decision: \nadopt the managed platform." }],
    });
    expect(result.operations).toEqual(["splitParagraph"]);
    expect(documentXml(pure)).toContain('<w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Decision: </w:t>');

    // A bullet that splits keeps its numbering, and the new line repeats the
    // marker the projection showed.
    const bullets = await brief();
    projection = bullets.project({ mode: "md" });
    line = lineOf(projection, "- Latency");
    result = await bullets.patch({
      revision: projection.revision,
      mode: "md",
      edits: [{ startLine: line, endLine: line, newText: "- Latency\n- Throughput" }],
    });
    expect(result.projection.text.split("\n").slice(2, 5)).toEqual(["- Latency", "- Throughput", "- Cost"]);

    // Fewer lines than paragraphs merges the remainder away.
    const merge = await brief();
    projection = merge.project({ mode: "md" });
    line = lineOf(projection, "- Latency");
    result = await merge.patch({
      revision: projection.revision,
      mode: "md",
      edits: [{ startLine: line, endLine: line + 1, newText: "- Latency and cost" }],
    });
    expect(result.operations).toEqual(["mergeParagraph", "deleteText", "insertText"]);
    expect(result.projection.text.split("\n").slice(2, 4)).toEqual(["- Latency and cost", ""]);

    // Markers are structure: changing one compiles to the paragraph intent.
    const restyle = await brief();
    projection = restyle.project({ mode: "md" });
    result = await restyle.patch({
      revision: projection.revision,
      mode: "md",
      edits: [
        { startLine: lineOf(projection, "# Findings"), endLine: lineOf(projection, "# Findings"), newText: "## Findings" },
        { startLine: lineOf(projection, "- Cost"), endLine: lineOf(projection, "- Cost"), newText: "Cost" },
      ],
    });
    expect(result.operations).toEqual(["setListType", "formatParagraph"]);
    expect(result.projection.text.split("\n").slice(0, 4)).toEqual(["## Findings", "Decision: adopt the managed platform.", "- Latency", "Cost"]);
  });

  it("refuses hunks that leave editable paragraph text", async () => {
    const agent = await brief();
    const projection = agent.project({ mode: "md" });
    const request = (startLine: number, endLine: number, newText: string) =>
      agent.patch({ revision: projection.revision, mode: "md", edits: [{ startLine, endLine, newText }] });

    const table = projection.anchors.find((anchor) => anchor.role === "table")!.line;
    await expect(request(table, table, "| a | b |")).rejects.toThrow("is table content and cannot be patched");
    const structure = projection.anchors.find((anchor) => anchor.role === "structure")!.line;
    await expect(request(structure, structure, "x")).rejects.toThrow("is structure content and cannot be patched");
    const math = lineOf(projection, "$E=mc^(2)$");
    await expect(request(math, math, "rewritten")).rejects.toThrow("across a non-text atom");
    await expect(request(1, 99, "x")).rejects.toThrow("outside the projected");
    await expect(request(1, 1, "# A\n## B")).rejects.toThrow("repeat the first line's structural marker");
    await expect(agent.patch({ revision: projection.revision, mode: "md" })).rejects.toThrow("either edits or diff");
    // Nothing above reached the document.
    expect(agent.revision).toBe(projection.revision);
    expect(agent.project({ mode: "md" }).text).toBe(projection.text);
  });

  it("guards each touched block by fingerprint and lets concurrent edits elsewhere through", async () => {
    const agent = await brief();
    const projection = agent.project({ mode: "md" });
    const target = lineOf(projection, "Schedule the review.");

    // An edit somewhere else moves the revision on but leaves the hunk's block
    // untouched, so the patch still applies against the older revision.
    const heading = projection.anchors[0];
    await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertText", at: { blockRef: heading.blockRef!, runRef: heading.segments[0].runRef, offset: 0 }, text: "Key " }],
    });
    expect(agent.revision).not.toBe(projection.revision);
    const result = await agent.patch({
      revision: projection.revision,
      mode: "md",
      edits: [{ startLine: target, endLine: target, newText: "Schedule the design review." }],
    });
    expect(result.projection.text).toContain("Schedule the design review.");
    expect(result.projection.text).toContain("# Key Findings");

    // Once the hunk's own block changes underneath it, the same patch is stale.
    const second = agent.project({ mode: "md" });
    const line = lineOf(second, "Schedule the design review.");
    await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertText", at: { blockRef: second.anchors[line - 1].blockRef!, runRef: second.anchors[line - 1].segments[0].runRef, offset: 0 }, text: "Soon: " }],
    });
    await expect(agent.patch({
      revision: second.revision,
      mode: "md",
      edits: [{ startLine: line, endLine: line, newText: "Schedule it." }],
    })).rejects.toThrow("stale");
    await expect(agent.patch({ revision: "gone", mode: "md", edits: [{ startLine: 1, endLine: 1, newText: "x" }] })).rejects.toThrow("stale");
  });

  it("applies a unified diff and refuses one written against other text", async () => {
    const agent = await brief();
    const projection = agent.project({ mode: "md" });
    const result = await agent.patch({
      revision: projection.revision,
      mode: "md",
      diff: "@@ -12,1 +12,2 @@\n-Schedule the review.\n+Schedule the review.\n+Name an owner.\n",
    });
    expect(result.operations).toEqual(["insertText", "splitParagraph"]);
    expect(result.projection.text.split("\n").slice(-2)).toEqual(["Schedule the review.", "Name an owner."]);

    const next = agent.project({ mode: "md" });
    await expect(agent.patch({ revision: next.revision, mode: "md", diff: "@@ -1,1 +1,1 @@\n-# Summary\n+# Overview\n" }))
      .rejects.toThrow('but the projection has "# Findings"');
    await expect(agent.patch({ revision: next.revision, mode: "md", diff: "no hunks here" })).rejects.toThrow("no @@ hunks");
  });

  it("records a suggested hunk as a tracked change", async () => {
    const inserted = await brief();
    let projection = inserted.project({ mode: "md" });
    let line = lineOf(projection, "Schedule the review.");
    let result = await inserted.patch({
      revision: projection.revision,
      mode: "md",
      suggest: true,
      edits: [{ startLine: line, endLine: line, newText: "Schedule the design review." }],
    });
    expect(result.operations).toEqual(["insertText"]);
    expect(documentXml(inserted)).toMatch(/<w:ins\b/);

    const removed = await brief();
    projection = removed.project({ mode: "md" });
    line = lineOf(projection, "Decision: adopt the managed platform.");
    result = await removed.patch({
      revision: projection.revision,
      mode: "md",
      suggest: true,
      edits: [{ startLine: line, endLine: line, newText: "Decision: adopt the platform." }],
    });
    expect(result.operations).toEqual(["suggestRevision"]);
    expect(documentXml(removed)).toMatch(/<w:del\b/);

    // A tracked deletion re-splits the run the insertion would target, so the
    // two halves of a replacement have to arrive as separate patches.
    const both = await brief();
    projection = both.project({ mode: "md" });
    line = lineOf(projection, "Decision: adopt the managed platform.");
    await expect(both.patch({
      revision: projection.revision,
      mode: "md",
      suggest: true,
      edits: [{ startLine: line, endLine: line, newText: "Decision: adopt the custom stack." }],
    })).rejects.toThrow("add text or remove text, not both");
  });

  it("submits a patch through a collaborative target", async () => {
    const doc = DocxDocument.load(makeDocx(body("<w:p><w:r><w:t>Shared draft</w:t></w:r></w:p>")));
    const ids = doc.enableStableIds();
    let revision = 0;
    let nextId = 100_000;
    const target: AgentCollaborativeTarget = {
      getDocument: () => doc,
      getRevision: () => revision,
      allocateIds: (count) => Array.from({ length: count }, () => nextId++),
      submit: (operation) => {
        const result = applyIntentScoped(doc, ids, { ...operation, clientId: "agent", clientSeq: revision + 1, base: revision } as Intent);
        if (!result.applied) throw new Error("shared apply failed");
        resyncScope(doc, ids, result);
        revision++;
      },
      getConnectionState: () => "live",
    };
    const agent = AgentDocument.connect(target);
    const projection = agent.project({ mode: "text" });
    expect(projection.text).toBe("Shared draft");
    const result = await agent.patch({
      revision: projection.revision,
      mode: "text",
      edits: [{ startLine: 1, endLine: 1, newText: "Shared final draft" }],
    });
    expect(result).toMatchObject({ status: "submitted", operations: ["insertText"], connection: "live" });
    expect(result.projection.text).toBe("Shared final draft");
    expect(agent.project({ mode: "text" }).revision).toBe("1");
  });

  it("reaches table cell text through the text-mode projection", async () => {
    const agent = await brief();
    const projection = agent.project({ mode: "text" });
    const line = lineOf(projection, "Managed");
    const result = await agent.patch({
      revision: projection.revision,
      mode: "text",
      edits: [{ startLine: line, endLine: line, newText: "Managed cloud" }],
    });
    expect(result.operations).toEqual(["insertText"]);
    expect(agent.project({ mode: "md" }).text).toContain("| Managed cloud | 9 |");
  });
});
