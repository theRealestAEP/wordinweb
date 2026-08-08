// @vitest-environment jsdom
/**
 * INVARIANT C — the insert-command half of the capability matrix.
 *
 * WHY THIS FILE EXISTS SEPARATELY. capability-matrix.test.tsx audits commands
 * run on a SELECTED OBJECT, and its collab mount deliberately withholds
 * `submitOp` so the fixture is built identically in both modes (see the
 * comment on its `mount`). That is the right call for what it audits — and it
 * means every api INSERT command's collab path sits outside its INVARIANT B.
 * `insertImage` shipped with NO collab branch at all: in a room it mutated the
 * local document and emitted nothing, forking silently, and the whole matrix
 * stayed green through it.
 *
 *   INVARIANT C  in a collab mount, every insert command either EMITS at
 *                least one intent, or its toolbar feature is GATED OFF so no
 *                control exists to press.
 *
 * An absent button is honest. A present button that mutates locally without
 * emitting is a fork factory. A command that neither mutates nor emits is an
 * honest no-op and passes — unaddressable positions are allowed to decline.
 *
 * NO RESTATEMENT, so the audit cannot drift from the UI:
 *   - the command list is `INSERT_COMMANDS` (react/src/toolbar.tsx), exported
 *     next to `ToolbarFeature` the way core exports SELECTED_OBJECT_COMMANDS;
 *   - whether a command is offered in a room is read from the REAL
 *     `COLLAB_TOOLBAR_DEFAULTS`, not a copy — flip a gate there and this audit
 *     follows automatically;
 *   - COMPLETENESS is asserted against the api surface itself, so a newly
 *     added insert command fails here until it is enumerated.
 */
import { describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import { DocxView, INSERT_COMMANDS, type DocxViewApi } from "../src/index.js";
import { COLLAB_TOOLBAR_DEFAULTS } from "../src/collab.js";
import { blankDocxBytes } from "@wordinweb/server";
import { serializeXml, type DocxDocument, type EditorIntent } from "@wordinweb/core";

// ---------------------------------------------------------------- jsdom gaps
const glob = globalThis as unknown as Record<string, unknown>;
glob.createImageBitmap ??= async () => ({ width: 64, height: 48, close() {} });
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:insert-commands";
  URL.revokeObjectURL = () => {};
}
// A 3D model would lazy-import @google/model-viewer, whose WebGL renderer
// throws in jsdom; claiming the tag keeps the import from happening (same
// shim capability-matrix uses).
if (typeof customElements !== "undefined" && !customElements.get("model-viewer")) {
  customElements.define("model-viewer", class extends HTMLElement {});
}
// The shared setup's fake 2D context covers text measurement; a couple of
// inserts paint a poster/preview, so widen it here rather than globally.
{
  const proto = HTMLCanvasElement.prototype as unknown as { getContext: () => unknown };
  const base = proto.getContext;
  proto.getContext = function (this: HTMLCanvasElement) {
    const ctx = (base as () => Record<string, unknown>).call(this) ?? {};
    for (const fn of ["fillRect", "strokeRect", "closePath", "moveTo", "lineTo", "arc", "stroke", "rect", "clip", "createImageData", "putImageData", "getImageData", "measureText"]) {
      if (typeof ctx[fn] !== "function") ctx[fn] = () => ({ width: 0, data: new Uint8ClampedArray(4) });
    }
    return ctx;
  } as never;
  // jsdom never invokes the toBlob callback (no real canvas backend), which
  // HANGS every command that builds a poster image. A stand-in blob keeps
  // those commands drivable, so they are genuinely audited rather than
  // excused — the audit is about emission, not about poster pixels.
  (HTMLCanvasElement.prototype as unknown as { toBlob: unknown }).toBlob =
    function (cb: (b: Blob) => void) { cb(new Blob([new Uint8Array([137, 80, 78, 71])], { type: "image/png" })); };
}

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3, 4]);
const GLB = new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]);

async function tick(ms = 5) { await act(async () => { await new Promise<void>((r) => setTimeout(r, ms)); }); }

/**
 * A collab mount with a REAL submit hook — the difference from the capability
 * matrix's mount, and the whole point. `submitOp` is supplied, so the api's
 * insert commands take their collab path (`collabOp`) rather than falling back
 * to the local mutation; `submitOp` applies the intent through the same
 * canonical apply the server runs, which is also what makes the resulting
 * document trustworthy to compare.
 */
async function mountCollab() {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const intents: EditorIntent[] = [];
  let nextId = 900_000;
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(),
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
      collab: {
        submit: (intent: EditorIntent) => { intents.push(intent); },
        submitOp: (intent: { kind: string } & Record<string, unknown>) => {
          intents.push(intent as unknown as EditorIntent);
        },
        allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
      },
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  const selectAll = async () => {
    const target = container.querySelector("textarea") ?? container;
    await act(async () => {
      target.dispatchEvent(new KeyboardEvent("keydown", { key: "a", ctrlKey: true, bubbles: true, cancelable: true }));
    });
    await tick();
  };
  const click = async () => {
    const page = container.querySelector<HTMLElement>(".dxw-page")!;
    const span = page.querySelector("span") ?? page;
    await act(async () => {
      const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
      span.dispatchEvent(new MouseEvent("mousedown", opts));
      span.dispatchEvent(new MouseEvent("mouseup", opts));
    });
    await tick();
  };
  /** The document's whole editable surface + its media parts: the oracle for
   * "did this mutate?". Media parts are included because a media-registering
   * insert changes the package, not just the XML. */
  const snapshot = (): string =>
    seen.doc!.editableRoots().map((r) => serializeXml(r)).join("|") +
    "||" + seen.doc!.pkg.names().filter((n) => n.startsWith("word/media/")).join(",");
  return {
    container, intents, click, selectAll, snapshot,
    api: () => seen.api!,
    unmount: async () => { await act(async () => { root.unmount(); }); },
  };
}

/** How to drive each command headlessly. Args are arbitrary-but-valid; the
 * audit is about EMISSION, not about what each command produces. */
const INVOKE: Record<string, (api: DocxViewApi) => unknown> = {
  insertTable: (api) => api.insertTable(2, 2),
  // The caret sits in a body paragraph, not a table cell, so this declines —
  // the honest no-op branch of the fork rule, audited on purpose.
  insertTableFormula: (api) => api.insertTableFormula("SUM(ABOVE)"),
  insertImage: (api) => api.insertImage(new Blob([PNG], { type: "image/png" })),
  insertScreenshot: (api) => api.insertScreenshot(),
  insertModel3D: (api) => api.insertModel3D(new Blob([GLB]), new Blob([PNG], { type: "image/png" })),
  insertOnlineVideo: (api) => api.insertOnlineVideo("https://example.com/v"),
  insertEmbeddedObject: (api) => api.insertEmbeddedObject(new Blob([PNG]), "a.docx"),
  insertChart: (api) => api.insertChart({ type: "column", categories: ["a", "b"], series: [{ name: "s", values: [1, 2] }] }),
  insertSmartArt: (api) => api.insertSmartArt({ layout: "process", items: ["One", "Two"] }),
  insertShape: (api) => api.insertShape("rectangle", ""),
  insertWordArt: (api) => api.insertWordArt("WA", "plain"),
  insertEquation: (api) => api.insertEquation("a+b"),
  insertSymbol: (api) => api.insertSymbol("§"),
  insertPageNumber: (api) => api.insertPageNumber("page"),
  insertPageNumberPosition: (api) => api.insertPageNumberPosition("top", "center"),
  insertField: (api) => api.insertField("AUTHOR", "me"),
  insertToc: (api) => api.insertToc(),
  // An explicit entry, because the blank mount has no selection to mark.
  addIndexEntry: (api) => api.addIndexEntry("Widgets:assembly"),
  insertIndex: (api) => api.insertIndex(),
  // A blank room has no sources part, so the citation is the honest no-op
  // path in collab (emits nothing, mutates nothing) — allowed by the rule.
  insertCitation: (api) => api.insertCitation("Doe03"),
  insertBibliography: (api) => api.insertBibliography(),
  insertDateTime: (api) => api.insertDateTime("date"),
  insertCrossReference: (api) => api.insertCrossReference("Anchor1", "page"),
  // A blank document lists no headings/captions, so this declines — the
  // honest no-op branch of the fork rule, audited on purpose.
  insertCrossRefToTarget: (api) => { api.listCrossRefTargets(); return api.insertCrossRefToTarget("0", "page"); },
  insertCaption: (api) => api.insertCaption("Figure", "audit", "below"),
  insertBreak: (api) => api.insertBreak("page"),
  insertBlankPage: (api) => api.insertBlankPage(),
  insertCoverPage: (api) => api.insertCoverPage({ title: "T", subtitle: "S", author: "A", date: "2026" } as never),
  insertWatermark: (api) => api.insertWatermark({ text: "DRAFT" }),
  insertHeaderFooterPreset: (api) => api.insertHeaderFooterPreset("header", "blank"),
  addComment: (api) => api.addComment("note"),
  addFootnote: (api) => api.addFootnote("note"),
  addEndnote: (api) => api.addEndnote("note"),
  addBookmark: (api) => api.addBookmark("Anchor1"),
};

/** Offered in a room? Read from the REAL defaults, never restated. */
const offeredInCollab = (feature: string): boolean =>
  (COLLAB_TOOLBAR_DEFAULTS as Record<string, boolean | undefined>)[feature] !== false;

/**
 * Commands that cannot be DRIVEN headlessly, with the reason. Reported as
 * skipped rather than silently dropped (capability-matrix's idiom for a
 * fixture it cannot build). Their gate is still asserted — what cannot be
 * exercised here is the command body, not the offer.
 */
const NOT_HEADLESS: Record<string, string> = {
  insertScreenshot: "needs navigator.mediaDevices.getDisplayMedia, which jsdom does not implement",
};

/** Commands that need a SELECTION rather than just a caret. */
const NEEDS_SELECTION = new Set(["addComment", "addBookmark"]);

/**
 * GATED commands that still mutate the local document when an app calls the
 * api METHOD directly. Not reachable from the UI — the toolbar offers no
 * control — so they do not violate invariant C, which is a statement about
 * what a user can press. Recorded explicitly (the same idiom invariant A uses
 * for offered-but-inert) because each one is a landmine for the day someone
 * un-gates it: un-gating without wiring is precisely the insertImage bug.
 *
 * DELIBERATELY EMPTY. It held insertModel3D, insertOnlineVideo and
 * insertEmbeddedObject, and "not reachable from the UI" turned out to be the
 * wrong bar: the api is public, so an embedder calling the method directly
 * forked the room with no toolbar involved. All three now refuse in collab,
 * so the entries would assert a landmine that no longer exists.
 *
 * The map stays because the mechanism is the point — the next gated-but-
 * mutating command must be recorded here with a reason, or wired.
 */
const GATED_BUT_MUTATES: Record<string, string> = {};

describe("INVARIANT C — insert commands in collab either EMIT or are ABSENT", () => {
  it("the enumeration covers the api's whole insert surface", async () => {
    // Mechanical completeness: a newly added insert command fails HERE until
    // someone enumerates it, which is what stops the audit drifting from the
    // api the toolbar drives.
    const ed = await mountCollab();
    const apiNames = Object.keys(ed.api()).filter((k) => /^(insert|add)[A-Z]/.test(k));
    const enumerated = new Set(INSERT_COMMANDS.map((c) => c.command));
    const missing = apiNames.filter((n) => !enumerated.has(n));
    expect(missing, `api insert commands missing from INSERT_COMMANDS: ${missing.join(", ")}`).toEqual([]);
    // …and nothing enumerated that the api doesn't have (a rename would
    // otherwise leave a silently-unaudited entry behind).
    const stale = [...enumerated].filter((c) => !apiNames.includes(c));
    expect(stale, `INSERT_COMMANDS entries with no api method: ${stale.join(", ")}`).toEqual([]);
    const undriven = [...enumerated].filter((c) => typeof INVOKE[c] !== "function");
    expect(undriven, `enumerated but not driven by this audit: ${undriven.join(", ")}`).toEqual([]);
    await ed.unmount();
  });

  /** Filled by the per-command cases; read by the non-vacuity check below. */
  const emittedBy: string[] = [];

  for (const { command, feature } of INSERT_COMMANDS) {
    const offered = offeredInCollab(feature);
    const cannotDrive = NOT_HEADLESS[command];
    it(`${command} (${feature}, ${offered ? "offered" : "gated"})${cannotDrive ? " [not driven headlessly]" : ""}`, async () => {
      if (cannotDrive) {
        // Still a real assertion: a command this audit cannot exercise must
        // not be OFFERED in a room, or it would be un-auditable and pressable
        // at the same time.
        expect(offered, `${command} cannot be driven headlessly (${cannotDrive}) yet is offered in collab — gate it or make it drivable`).toBe(false);
        return;
      }
      const ed = await mountCollab();
      try {
        await ed.click();
        if (NEEDS_SELECTION.has(command)) await ed.selectAll();
        const before = ed.snapshot();
        ed.intents.length = 0; // a selection can legitimately emit nothing; start clean
        // Bounded: a command that never settles (one awaiting a network or a
        // device) must fail ITS OWN case, not wedge the act() queue and take
        // every later case down with it — which is exactly what happened
        // before this race was added.
        let timer: ReturnType<typeof setTimeout> | undefined;
        const hung = new Promise<"hung">((r) => { timer = setTimeout(() => r("hung"), 4000); });
        const outcome = await act(async () =>
          Promise.race([Promise.resolve(INVOKE[command](ed.api())).then(() => "done" as const), hung]));
        clearTimeout(timer);
        expect(outcome, `${command} did not settle in 4s — give it a NOT_HEADLESS reason if it needs a browser`).toBe("done");
        await tick(20);
        const mutated = ed.snapshot() !== before;
        const emitted = ed.intents.length > 0;
        if (emitted) emittedBy.push(command);

        if (offered) {
        // THE FORK RULE. Doing nothing is honest (an unaddressable position
        // may decline); changing this replica's document while telling nobody
        // is the bug this invariant exists for.
          expect(
            !mutated || emitted,
            `${command} is offered in collab and MUTATED the document without emitting an intent — the room forks silently`,
          ).toBe(true);
        } else if (mutated) {
          // Unreachable from the UI, but it must be a KNOWN landmine, not a
          // new one nobody noticed.
          expect(
            GATED_BUT_MUTATES[command],
            `${command} is gated off yet mutates locally, and is not in GATED_BUT_MUTATES — add it with a reason, or wire it`,
          ).toBeTruthy();
        }
      } finally {
        await ed.unmount();
      }
    });
  }

  it("is not vacuous: the commands that CAN act here actually emitted", async () => {
    // Without this, every command silently no-opping would pass the fork rule
    // and the audit would prove nothing. These are the ones a caret alone is
    // enough for, so they must genuinely ride the wire.
    const mustEmit = ["insertTable", "insertEquation", "insertShape", "insertChart", "insertSmartArt", "insertBreak", "insertBlankPage", "insertPageNumber", "insertToc"];
    const silent = mustEmit.filter((c) => !emittedBy.includes(c));
    expect(silent, `these should have emitted an intent but did not: ${silent.join(", ")}`).toEqual([]);
  });
});
