// @vitest-environment jsdom
/**
 * SUGGESTING MODE REACHES THE TABLE COMMANDS (task #44).
 *
 * Every run and paragraph command in the api draws `editor.suggestionMeta()`
 * once per gesture and threads it into BOTH the local core call and the
 * emitted collab intent (collab plan doc 05 rule a: the same author+date on
 * every replica). The two table entry points did not: `tableOp` called
 * `applyTableOp` with no meta and emitted an intent with no `suggest`, and
 * `runTableOperation` — which serves all eight registered table-property
 * operations — did the same. A reviewer restyling a table in suggesting mode
 * got an applied change with nothing to accept or reject, while core, the
 * wire schema and the agent had carried the payload all along.
 *
 * Two mounts, because the api takes exactly one path per gesture: outside a
 * room the local mutation runs and no intent is emitted; in a room the intent
 * is emitted and the local mutation is skipped (collabOp never falls through,
 * so a replica cannot fork). So the local mount pins the recorded revision and
 * the collab mount pins the wire payload.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { zipSync, strToU8 } from "fflate";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { serializeXml, type DocxDocument, type EditorIntent } from "@wordinweb/core";

const AUTHOR = "Reviewer A";
const ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;

function cell(text: string): string {
  return `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
}

/** A 2x2 table whose cells already carry a tcW, so a recorded tcPrChange
 * holds real previous properties rather than an empty shell. */
const FIXTURE = (() => {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">Tables parity</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
    `<w:tr>${cell("r0c0")}${cell("r0c1")}</w:tr>` +
    `<w:tr>${cell("r1c0")}${cell("r1c1")}</w:tr></w:tbl>` +
    `<w:p><w:r><w:t xml:space="preserve">After the table</w:t></w:r></w:p>`;
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${body}</w:body></w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(`<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
    "_rels/.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
    "word/document.xml": strToU8(xml),
  });
})();

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

/** `collab` supplied mounts in a room (intents emitted, local path skipped);
 * omitted mounts standalone (local mutation, no intents). */
async function mount(inRoom: boolean) {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const intents: ({ kind: string } & Record<string, unknown>)[] = [];
  let nextId = 900_000;
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: FIXTURE,
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
      ...(inRoom
        ? {
            collab: {
              submit: (intent: EditorIntent) => { intents.push(intent as unknown as { kind: string } & Record<string, unknown>); },
              submitOp: (intent: { kind: string } & Record<string, unknown>) => { intents.push(intent); },
              allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
            },
          }
        : {}),
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const api = seen.api!;
  api.setSuggesting(true, AUTHOR);
  // Put the caret in the first cell. find() leaves a selection; the table
  // commands address the caret, so collapse it the way a user would.
  const press = async (key: string) => {
    const target = (container.contains(document.activeElement) ? (document.activeElement as HTMLElement) : container.querySelector("textarea")) ?? container;
    await act(async () => {
      target.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }));
      await new Promise((r) => setTimeout(r, 2));
    });
  };
  api.find("r0c0");
  await tick();
  await press("ArrowRight");
  expect(api.getTableCellFill(), "caret is inside the table").not.toBe(undefined);
  return {
    api,
    intents,
    xml: () => serializeXml(seen.doc!.docRoot as never, true),
    unmount: async () => { await act(async () => root.unmount()); container.remove(); },
  };
}

/** The w:author/w:date on the first change record of the given kind. */
function recorded(xml: string, change: "tcPrChange" | "tblPrChange"): { author: string; date: string } | null {
  const m = new RegExp(`<w:${change}[^>]*>`).exec(xml);
  if (!m) return null;
  const author = /w:author="([^"]*)"/.exec(m[0]);
  const date = /w:date="([^"]*)"/.exec(m[0]);
  return author && date ? { author: author[1], date: date[1] } : null;
}

describe("table commands in suggesting mode", () => {
  it("records a tcPrChange for a cell shading, with the drawn author", async () => {
    const m = await mount(false);
    m.api.tableOp({ kind: "cellShading", fill: "#FFFF00" });
    await tick();

    const xml = m.xml();
    const change = recorded(xml, "tcPrChange");
    expect(change?.author).toBe(AUTHOR);
    expect(change?.date).toMatch(ISO);
    // The shading is applied AND what it replaced is recorded: the tcW the
    // cell already carried is inside the record, not lost.
    expect(xml).toContain(`w:fill="FFFF00"`);
    expect(xml).toMatch(/<w:tcPrChange[^>]*><w:tcPr><w:tcW[^>]*\/><\/w:tcPr><\/w:tcPrChange>/);
    expect(m.api.revisionCount()).toBe(1);
    await m.unmount();
  });

  it("records a tblPrChange for table borders, with the drawn author", async () => {
    const m = await mount(false);
    m.api.setTableBorders("table", ["top", "bottom"], { style: "single", sz: 8, color: "auto" });
    await tick();

    const change = recorded(m.xml(), "tblPrChange");
    expect(change?.author).toBe(AUTHOR);
    expect(change?.date).toMatch(ISO);
    expect(m.api.revisionCount()).toBe(1);
    await m.unmount();
  });

  it("carries the suggest payload on the emitted tableOp and setTableBorders intents", async () => {
    const m = await mount(true);
    m.api.tableOp({ kind: "cellShading", fill: "#FFFF00" });
    await tick();
    m.api.setTableBorders("table", ["top", "bottom"], { style: "single", sz: 8, color: "auto" });
    await tick();

    const shading = m.intents.find((i) => i.kind === "tableOp");
    const borders = m.intents.find((i) => i.kind === "setTableBorders");
    for (const intent of [shading, borders]) {
      expect(intent, "intent emitted").toBeTruthy();
      const suggest = intent!.suggest as { author: string; date: string } | undefined;
      expect(suggest?.author).toBe(AUTHOR);
      expect(suggest?.date).toMatch(ISO);
    }
    // In a room the local mutation is skipped, so the record every replica
    // writes comes from the payload above and from nowhere else.
    expect(m.xml()).not.toContain("PrChange");
    await m.unmount();
  });
});
