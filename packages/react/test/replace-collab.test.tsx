// @vitest-environment jsdom
/**
 * #112 — replaceAll/replaceCurrent in a COLLAB mount ride the wire.
 *
 * The mock submitOp records intents WITHOUT applying them, so the assertion
 * "the local document did not change" is exactly the invariant that was
 * broken: the old replaceAll mutated locally and emitted nothing (a silent
 * permanent fork); the fixed one emits the compiled intents and mutates
 * nothing outside the canonical apply.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { serializeXml, type DocxDocument } from "@wordinweb/core";

const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">Body cat one</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">Second cat here</w:t></w:r></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
})();

async function tick(ms = 5) {
  await act(async () => {
    await new Promise<void>((r) => setTimeout(r, ms));
  });
}

async function mountCollab() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const intents: Array<{ kind: string } & Record<string, unknown>> = [];
  let nextId = 900_000;
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(
      createElement(DocxView, {
        source: FIXTURE,
        editable: true,
        onReady: (api: DocxViewApi) => { seen.api = api; },
        onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
        collab: {
          submit: (intent: never) => { intents.push(intent); },
          submitOp: (intent: { kind: string } & Record<string, unknown>) => { intents.push(intent); },
          allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
        },
      }),
    );
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  return {
    intents,
    api: () => seen.api!,
    doc: () => seen.doc!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

describe("replace on the wire in collab mounts (#112)", () => {
  it("replaceAll emits deleteText/insertText per match and leaves the local doc to the canonical apply", async () => {
    const m = await mountCollab();
    const before = serializeXml(m.doc().docRoot);
    let result: ReturnType<DocxViewApi["replaceAll"]> | undefined;
    await act(async () => { result = m.api().replaceAll("cat", "dog"); });
    expect(result!.total).toBe(2);
    expect(result!.byStory).toEqual({ body: 2 });
    expect(m.intents.map((i) => i.kind)).toEqual(["deleteText", "insertText", "deleteText", "insertText"]);
    // Back-to-front: the first emitted delete strikes the LAST match.
    const inserts = m.intents.filter((i) => i.kind === "insertText") as Array<{ at: { blockId: number } }>;
    expect(inserts[0].at.blockId).toBeGreaterThan(inserts[1].at.blockId);
    // The mock does not apply, and the api must not mutate on its own.
    expect(serializeXml(m.doc().docRoot)).toBe(before);
    await m.unmount();
  });

  it("replaceAll in suggesting mode emits the strike-then-insert pair with one author/date", async () => {
    const m = await mountCollab();
    await act(async () => { m.api().setSuggesting(true, "Reviewer"); });
    await act(async () => { m.api().replaceAll("cat", "dog"); });
    expect(m.intents.map((i) => i.kind)).toEqual(["suggestRevision", "insertText", "suggestRevision", "insertText"]);
    const metas = m.intents.map((i) => JSON.stringify(i.suggest));
    expect(new Set(metas).size).toBe(1);
    expect((m.intents[0].suggest as { author: string }).author).toBe("Reviewer");
    await m.unmount();
  });

  it("replaceCurrent emits the same compiled shape for the selected match", async () => {
    const m = await mountCollab();
    await act(async () => { m.api().find("cat"); });
    let remaining = -1;
    await act(async () => { remaining = m.api().replaceCurrent("dog"); });
    expect(remaining).toBe(1);
    expect(m.intents.map((i) => i.kind)).toEqual(["deleteText", "insertText"]);
    await m.unmount();
  });
});
