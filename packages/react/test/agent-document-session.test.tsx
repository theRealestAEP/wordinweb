// @vitest-environment jsdom
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { describe, expect, it } from "vitest";
import { blankDocxBytes } from "@wordinweb/server";
import { AgentDocument, LocalDocumentSession, localDocumentViewBinding } from "../../agent/src/index.js";
import { DocxView, useAgentDocumentSession, type DocxViewApi } from "../src/index.js";

async function tick(): Promise<void> {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, 10)); });
}

describe("agent and solo browser session", () => {
  it("lets DocxView and AgentDocument edit one canonical document", async () => {
    const session = new LocalDocumentSession(blankDocxBytes());
    const agent = AgentDocument.connect(session);
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

    const read = agent.inspect({ kind: "read" });
    if (!("blocks" in read) || read.blocks[0].type !== "paragraph") throw new Error("missing paragraph");
    const paragraph = read.blocks[0];
    const run = paragraph.runs[0];
    await act(async () => {
      const result = await agent.edit({
        revision: agent.revision,
        operations: [{ kind: "insertText", at: { blockRef: paragraph.ref, runRef: run.ref, offset: 0 }, text: "Shared with the agent" }],
      });
      expect(result.connection).toBe("local");
    });
    for (let attempt = 0; attempt < 30 && !container.textContent?.includes("Shared with the agent"); attempt++) await tick();
    expect(container.textContent).toContain("Shared with the agent");

    const blockId = Number(paragraph.ref.slice("block:".length));
    const runId = Number(run.ref.slice("run:".length));
    expect(api!.setCaretFromEncoded({ blockId, runId, offset: "Shared with the agent".length })).toBe(true);
    await act(async () => { expect(api!.insertEquation("x^2")).toBe(true); });
    const overview = agent.inspect({ kind: "overview" });
    expect("components" in overview && overview.components.math).toBe(1);

    await act(async () => { root.unmount(); });
    container.remove();
  });
});
