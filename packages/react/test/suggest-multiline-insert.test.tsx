// @vitest-environment jsdom
//
// Regression: an agent insertText whose text carries embedded line breaks
// ("\n" between paragraphs, or "\v") must never land those control characters
// in a painted text span. The layout engine measures the raw string on canvas
// (a "\n" advances ~zero px) while the DOM painter renders spans with
// `white-space: pre` — the browser breaks the span at the "\n" and paints the
// remainder one line down at the span's left edge, on top of the engine-placed
// spans of the next line. That is the "ghost text painted over other text"
// bug from the LikeOffice AI panel (suggesting mode, multi-paragraph insert).
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { blankDocxBytes } from "@wordinweb/server";
import { AgentDocument, LocalDocumentSession, localDocumentViewBinding } from "../../agent/src/index.js";
import { DocxView, useAgentDocumentSession, type DocxViewApi } from "../src/index.js";

async function tick(): Promise<void> {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 10)); });
}

async function insertMultiline(text: string): Promise<{ container: HTMLDivElement; session: LocalDocumentSession; api: DocxViewApi; cleanup: () => Promise<void> }> {
  const session = new LocalDocumentSession(blankDocxBytes());
  const agent = AgentDocument.connect(session, { provenance: { author: "AI" } });
  const binding = localDocumentViewBinding(session);
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  let api: DocxViewApi | null = null;

  function Host() {
    const view = useAgentDocumentSession(binding);
    return createElement(DocxView, { ...view, editable: true, onReady: (value: DocxViewApi) => { api = value; } });
  }

  await act(async () => { root.render(createElement(Host)); });
  for (let attempt = 0; attempt < 30 && !api; attempt++) await tick();
  expect(api).toBeTruthy();
  api!.setSuggesting(true, "AI");

  const read = agent.inspect({ kind: "read" });
  if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
  const paragraph = read.blocks[0];
  const run = paragraph.runs[0];
  await act(async () => {
    const result = await agent.edit({
      revision: agent.revision,
      operations: [{ kind: "insertText", at: { blockRef: paragraph.ref, runRef: run.ref, offset: 0 }, text, suggest: true }],
    });
    expect(result.connection).toBe("local");
  });
  // Wait for the coalesced renderSignal repaint to land.
  for (let attempt = 0; attempt < 30; attempt++) await tick();

  return {
    container,
    session,
    api: api!,
    cleanup: async () => {
      await act(async () => { root.unmount(); });
      container.remove();
    },
  };
}

/** Painted text spans that carry a raw line-break control character. Any such
 * span paints a second visual line inside itself (white-space: pre) that the
 * engine never laid out — overlapping ghost text in a real browser. */
function brokenSpans(container: HTMLElement): string[] {
  return Array.from(container.querySelectorAll<HTMLElement>(".dxw-page span"))
    .map((span) => span.textContent ?? "")
    .filter((text) => /[\n\v\r]/.test(text));
}

const TEXT =
  "When in the Course of human events, it becomes necessary for one people to dissolve the political bands which have connected them with another, and to assume among the powers of the earth, the separate and equal station to which the Laws of Nature and of Nature's God entitle them, a decent respect to the opinions of mankind requires that they should declare the causes which impel them to the separation.";

describe("suggested multi-paragraph agent insert", () => {
  it('splits "\\n" text into paragraphs and paints no span with a raw line break', async () => {
    const { container, session, api, cleanup } = await insertMultiline(`Declaration of Independence\n\n${TEXT}`);

    expect(container.textContent).toContain("Declaration of Independence");
    expect(container.textContent).toContain("to the separation.");
    expect(brokenSpans(container)).toEqual([]);
    // The split must happen in the model (real paragraphs), not by scrubbing
    // control characters at paint time.
    const paragraphs = session.doc.sections[0]?.blocks.filter((block) => block.type === "paragraph") ?? [];
    expect(paragraphs.length).toBe(3);
    const texts = paragraphs.map((block) =>
      block.type === "paragraph"
        ? block.children.flatMap((child) => (child.type === "run" ? child.content : [])).map((c) => (c.kind === "text" ? c.text : "")).join("")
        : "");
    expect(texts[0]).toBe("Declaration of Independence");
    expect(texts[1]).toBe("");
    expect(texts[2]).toBe(TEXT);
    for (const text of texts) expect(text).not.toMatch(/[\n\v\r]/);
    // Every inserted paragraph mark must be tracked: rejecting the whole
    // suggestion restores the blank single-paragraph document (a missed
    // intermediate mark would strand an extra paragraph).
    await act(async () => { expect(api.rejectAllRevisions()).toBeGreaterThan(0); });
    const after = session.doc.sections[0]?.blocks.filter((block) => block.type === "paragraph") ?? [];
    expect(after.length).toBe(1);
    await cleanup();
  });

  it('treats "\\v" as a paragraph break too', async () => {
    const { container, session, cleanup } = await insertMultiline("First line\vSecond line");

    expect(container.textContent).toContain("First line");
    expect(container.textContent).toContain("Second line");
    expect(brokenSpans(container)).toEqual([]);
    const paragraphs = session.doc.sections[0]?.blocks.filter((block) => block.type === "paragraph") ?? [];
    expect(paragraphs.length).toBe(2);
    await cleanup();
  });
});
