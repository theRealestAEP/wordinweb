// @vitest-environment jsdom
/**
 * CAPABILITY MATRIX for the selected-object command surface.
 *
 * Two bug classes kept escaping the suite, both invisible to per-feature tests
 * because each one is a statement about the WHOLE surface, not about any one
 * command:
 *
 *   1. OFFERED-BUT-INERT — the Layout ribbon offered "Rotate" for SmartArt,
 *      but setDrawingRotation no-ops on a graphic frame, so the button did
 *      nothing. A user found it, not a test.
 *   2. MUTATES-BUT-DOESN'T-EMIT — in collab mode a command changed the local
 *      document without emitting an intent, so the room forked. Each instance
 *      (drawing drag, table resize, Cmd+Enter, selection delete, ...) was
 *      found live and fixed piecemeal.
 *
 * This file turns both into standing invariants over every (object kind ×
 * command) pair:
 *
 *   INVARIANT A  offered ⇒ the command DOES something (mutates the document
 *                or opens an editor affordance), unless the pair is listed in
 *                EXCEPTIONS with a human-readable reason.
 *   INVARIANT B  in collab mode, a command that MUTATED the document emitted
 *                at least one intent. Mutating silently is a fork; not
 *                mutating at all is an honest no-op and passes.
 *
 * "Offered" is not re-stated here — it comes from core's
 * `availableObjectCommands`, the same function both toolbar sites render
 * from, so the UI and this audit cannot drift apart.
 *
 * ─── HOW TO EXTEND ────────────────────────────────────────────────────────
 *
 * Add an OBJECT KIND: append a FIXTURES entry with an `insert` that builds the
 * object through the public api and a `select` that clicks it (the object's
 * rendered element; see `selectDrawing` / `selectImageBinding`). Nothing else
 * changes — the matrix picks it up. If a kind cannot be built or selected
 * headlessly, give it `skip: "<reason>"`; it is reported as SKIPPED rather
 * than silently dropped.
 *
 * Add a COMMAND: add it to `SelectedObjectCommand` and to
 * `SELECTED_OBJECT_COMMANDS` in packages/core/src/edit/editor.ts, and decide
 * in `availableObjectCommands` which kinds it is offered for. The matrix then
 * covers it automatically, and THE RULE APPLIES: a new command must either
 * work everywhere it is offered, or ship with a gate (don't offer it there)
 * or an EXCEPTIONS entry explaining why the inert offer is intentional.
 *
 * Add a DIALOG: commands that open one are answered by `answerDialog` below,
 * which drives the real dialog DOM (deterministic — the dialogs are plain
 * jsdom forms). Give a new dialog a case there keyed on its title. No module
 * mocking is involved: core is consumed as a bundle, so the dialog module is
 * not an interceptable boundary from here, and driving the real form is both
 * simpler and a truer test of the post-dialog path.
 *
 * SCOPE: the two ribbon surfaces. The floating overlay bar the editor draws
 * next to a selected object (editor.ts `imageToolbar`) has its own gating and
 * a couple of controls that are not SelectedObjectCommands at all (WordArt
 * "Edit text", image "Replace"); it is not audited here.
 */
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot, type Root } from "react-dom/client";
import { DocxView, type DocxViewApi } from "../src/index.js";
import { blankDocxBytes } from "@wordinweb/server";
import {
  SELECTED_OBJECT_COMMANDS,
  availableObjectCommands,
  serializeXml,
  setModel3DRotation,
  type DocxDocument,
  type EditorIntent,
  type SelectedObjectCommand,
  type SelectedObjectContext,
  type SelectedObjectKind,
} from "@wordinweb/core";

// ---------------------------------------------------------------- jsdom gaps
const glob = globalThis as unknown as Record<string, unknown>;
// insertImage measures the bitmap; jsdom has no image decoder.
glob.createImageBitmap ??= async () => ({ width: 64, height: 48, close() {} });
if (typeof URL.createObjectURL !== "function") {
  URL.createObjectURL = () => "blob:capability-matrix";
  URL.revokeObjectURL = () => {};
}
// A 3D model makes DocxView lazy-import @google/model-viewer, whose WebGL
// renderer throws inside jsdom. Claiming the tag name keeps the import from
// happening at all (DocxView checks customElements first); the 3D fixture
// still renders its poster image, which is what selection targets.
if (typeof customElements !== "undefined" && !customElements.get("model-viewer")) {
  customElements.define("model-viewer", class extends HTMLElement {});
}

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
const GLB = new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]);

// ------------------------------------------------------------------ verdicts
type Effect = "DOC-MUTATION" | "UI-INTERACTION" | "INERT" | "SKIPPED";

interface Cell {
  fixture: string;
  kind: SelectedObjectKind | "?";
  command: SelectedObjectCommand;
  offered: boolean;
  effect: Effect;
  /** Collab mode: did the document change, and how many intents were sent. */
  collabMutated?: boolean;
  collabIntents?: number;
  collabIntentKinds?: string;
  collabEffect?: Effect;
  /** Set when the command only acts after a precondition (e.g. must be
   * floating first) — a state-dependent offer, reported but not failed. */
  needs?: string;
  needsKind?: Precondition["kind"];
  note?: string;
}

const matrix: Cell[] = [];

/**
 * Offered pairs that legitimately do nothing. Every entry needs a reason a
 * human can act on. An offered-but-inert pair NOT listed here fails the suite.
 *
 * Adding an entry should feel expensive — the cheap fix is usually to stop
 * offering the command for that kind (a gate in `availableObjectCommands`).
 * The one entry below is a MODE, not a mutation, which is the only shape this
 * audit cannot drive: it measures a command by whether the document changed.
 */
const EXCEPTIONS: { fixture: string; command: SelectedObjectCommand; reason: string }[] = [
  {
    fixture: "image",
    command: "crop",
    reason:
      "Crop is a mode toggle: it paints the crop frame and waits for a handle drag, " +
      "and the drag is what writes a:srcRect. Nothing is mutated by the command itself, " +
      "so there is no document change for this audit to see. The write path is covered " +
      "by core's setImageCrop tests and collab's setCrop intent tests.",
  },
  // The Autofit menu is a three-way state choice, and every shape this matrix
  // can insert starts in the "none" state (insertShapeAt and insertWordArtAt
  // both author a:noAutofit). Picking the state a shape is already in writes
  // the same bytes, so it reads as inert here; the other two members of the
  // same menu mutate on the same fixtures, which is what proves the control
  // is live.
  ...(["shape", "wordArt"] as const).map((fixture) => ({
    fixture,
    command: "autofitNone" as SelectedObjectCommand,
    reason:
      "Re-selecting the autofit mode the shape already has. Both fixtures are " +
      "inserted with a:noAutofit, so this writes the bytes that are already there; " +
      "autofitResizeShape and autofitShrinkText mutate the same fixtures.",
  })),
];

// -------------------------------------------------------------------- mounts
async function tick(ms = 4) {
  await act(async () => { await new Promise<void>((resolve) => setTimeout(resolve, ms)); });
}

interface Mounted {
  container: HTMLElement;
  root: Root;
  api: DocxViewApi;
  doc: DocxDocument;
  intents: EditorIntent[];
  unmount: () => Promise<void>;
}

/**
 * Mount an editable DocxView on the blank document.
 *
 * `collab: true` supplies only the intent sink (`submit`) and id allocator —
 * deliberately NOT `submitOp`, so the api's insert commands still take their
 * local path and the fixture is built the same way in both modes. What the
 * collab flag actually turns on is the editor host's `onIntent`, which is the
 * seam the emission invariant audits.
 */
async function mount(collab: boolean): Promise<Mounted> {
  const container = document.createElement("div");
  document.body.appendChild(container);
  const root = createRoot(container);
  const intents: EditorIntent[] = [];
  let nextId = 500_000;
  const seen: { api: DocxViewApi | null; doc: DocxDocument | null } = { api: null, doc: null };
  await act(async () => {
    root.render(createElement(DocxView, {
      source: blankDocxBytes(),
      editable: true,
      onReady: (api: DocxViewApi) => { seen.api = api; },
      onLoad: (info: { document: DocxDocument }) => { seen.doc = info.document; },
      collab: collab
        ? {
          submit: (intent: EditorIntent) => { intents.push(intent); },
          allocIds: (n: number) => Array.from({ length: n }, () => nextId++),
        }
        : undefined,
    }));
  });
  for (let i = 0; i < 30 && !container.querySelector(".dxw-page"); i++) await tick();
  expect(container.querySelector(".dxw-page")).toBeTruthy();
  expect(seen.api).toBeTruthy();
  expect(seen.doc).toBeTruthy();
  // Collab mode is "onIntent set AND the doc carries stable ids" — the second
  // half is what a real session gets from the replica.
  if (collab) seen.doc!.enableStableIds();

  // Place a caret so the insert commands have an anchor.
  const page = container.querySelector<HTMLElement>(".dxw-page")!;
  const target = page.querySelector("span") ?? page;
  await act(async () => {
    const opts = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
    target.dispatchEvent(new MouseEvent("mousedown", opts));
    target.dispatchEvent(new MouseEvent("mouseup", opts));
  });
  await tick();

  return {
    container,
    root,
    api: seen.api!,
    doc: seen.doc!,
    intents,
    unmount: async () => {
      await act(async () => { root.unmount(); });
      container.remove();
    },
  };
}

/**
 * The mutation oracle: the serialized editable stories plus a digest of every
 * package part.
 *
 * The stories alone are not enough — SmartArt fills and chart data live in
 * their own parts (word/diagrams/*, word/charts/*), which a story-only oracle
 * reports as "nothing happened". That false INERT is exactly the blind spot
 * this suite exists to remove.
 */
function documentXml(doc: DocxDocument): string {
  const stories = doc.editableRoots().map((root) => serializeXml(root)).join(" ");
  const parts = doc.pkg.raw();
  const digest = Object.keys(parts).sort().map((path) => {
    const bytes = parts[path];
    let hash = 2166136261;
    for (let i = 0; i < bytes.length; i++) hash = Math.imul(hash ^ bytes[i], 16777619);
    return `${path}:${bytes.length}:${hash >>> 0}`;
  }).join("|");
  return `${stories}\n${digest}`;
}

async function clickElement(el: HTMLElement) {
  await act(async () => {
    const opts = { bubbles: true, cancelable: true, clientX: 5, clientY: 5, button: 0, detail: 1 };
    el.dispatchEvent(new MouseEvent("mousedown", opts));
    el.dispatchEvent(new MouseEvent("mouseup", opts));
    document.dispatchEvent(new MouseEvent("mouseup", opts));
  });
  await tick();
}

/** Select the most recently inserted DrawingML object (shape/line/chart/…). */
async function selectDrawing(m: Mounted): Promise<boolean> {
  const drawings = [...m.container.querySelectorAll<HTMLElement>("[data-dxw-drawing]")];
  const el = drawings[drawings.length - 1];
  if (!el) return false;
  await clickElement(el);
  return !!m.api.getSelectedObjectContext();
}

/** Select an image/3D-model binding (the positioned wrapper around the <img>). */
async function selectImageBinding(m: Mounted): Promise<boolean> {
  const items = [...m.container.querySelectorAll<HTMLElement>('[data-dxw-item-kind="image"]')];
  const wrapper = items[items.length - 1];
  if (!wrapper) return false;
  // A 3D model's wrapper also holds the viewer element, which swallows the
  // pointer; the poster <img> is the binding the editor hit-tests against.
  const el = wrapper.querySelector<HTMLElement>("img") ?? wrapper;
  await clickElement(el);
  return !!m.api.getSelectedObjectContext();
}

/** Select a single SmartArt node (a different context: canEditText is true). */
async function selectSmartArtNode(m: Mounted): Promise<boolean> {
  const node = m.container.querySelector<HTMLElement>("[data-dxw-smart-art-node]");
  if (!node) return false;
  await clickElement(node);
  return m.api.getSelectedObjectContext()?.smartArtNodeSelected === true;
}

// ------------------------------------------------------------------ fixtures
interface Fixture {
  name: string;
  insert: (m: Mounted) => Promise<unknown>;
  select: (m: Mounted) => Promise<boolean>;
  skip?: string;
}

const CHART_DATA = {
  kind: "bar" as const,
  categories: ["Q1", "Q2"],
  series: [{ name: "Revenue", values: [3, 5] }],
};
const SMART_ART_DATA = { layout: "cycle" as const, items: ["Plan", "Build", "Ship"] };

const FIXTURES: Fixture[] = [
  { name: "shape", insert: async (m) => m.api.insertShape("rectangle", "Box"), select: selectDrawing },
  { name: "line", insert: async (m) => m.api.insertShape("line"), select: selectDrawing },
  // WordArt is a drawing whose context kind is "shape" but whose text is NOT
  // an editable textbox story — the pair that makes canEditText matter.
  { name: "wordArt", insert: async (m) => m.api.insertWordArt("Header", "plain"), select: selectDrawing },
  { name: "chart", insert: async (m) => m.api.insertChart(CHART_DATA), select: selectDrawing },
  { name: "smartArt", insert: async (m) => m.api.insertSmartArt(SMART_ART_DATA), select: selectDrawing },
  {
    name: "smartArt(node)",
    insert: async (m) => m.api.insertSmartArt(SMART_ART_DATA),
    select: selectSmartArtNode,
  },
  {
    name: "image",
    insert: async (m) => m.api.insertImage(new Blob([PNG], { type: "image/png" })),
    select: selectImageBinding,
  },
  {
    name: "model3d",
    insert: async (m) => m.api.insertModel3D(new Blob([GLB]), new Blob([PNG], { type: "image/png" })),
    select: selectImageBinding,
  },
];

// ------------------------------------------------------------------- dialogs
/**
 * Answer whatever modal the command opened, with fixed values, by driving the
 * real dialog DOM. Returns the dialog's title, or null when none is open.
 *
 * The values are chosen to differ from every fixture's current state, so a
 * command that works produces a visible mutation.
 */
async function answerDialog(): Promise<string | null> {
  const backdrop = document.querySelector<HTMLElement>(".dxw-input-dialog-backdrop");
  if (!backdrop) return null;
  const form = backdrop.querySelector("form")!;
  const title = form.querySelector("strong")?.textContent ?? "";
  const setValue = (input: HTMLInputElement | HTMLSelectElement, value: string) => {
    input.value = value;
    input.dispatchEvent(new Event("input", { bubbles: true }));
  };
  if (backdrop.dataset.dxwNumberPairDialog !== undefined) {
    const [first, second] = [...form.querySelectorAll<HTMLInputElement>('input[type="number"]')];
    // Exact size vs page position: distinct values so neither can coincide
    // with a fixture's existing geometry (which would look INERT).
    const pair = title === "Exact size" ? ["200", "100"] : ["137", "89"];
    setValue(first, pair[0]);
    setValue(second, pair[1]);
  } else if (backdrop.dataset.dxwColorDialog !== undefined) {
    setValue(form.querySelector<HTMLInputElement>('input[type="text"]')!, "#FF0000");
  } else if (backdrop.dataset.dxwLineStyleDialog !== undefined) {
    setValue(form.querySelector<HTMLInputElement>('input[type="text"]')!, "#00FF00");
    setValue(form.querySelector<HTMLInputElement>('input[type="number"]')!, "2");
    setValue(form.querySelector<HTMLSelectElement>("select")!, "dashed");
  } else {
    // Plain text dialog: rotation degrees, or alternative text.
    const field = form.querySelector<HTMLInputElement>("input, textarea")!;
    setValue(field, title === "Rotation" ? "45" : "Matrix alt text");
  }
  await act(async () => {
    form.dispatchEvent(new Event("submit", { bubbles: true, cancelable: true }));
  });
  await tick();
  return title;
}

function closeStrayDialogs() {
  for (const el of document.querySelectorAll(".dxw-input-dialog-backdrop")) el.remove();
}

// ------------------------------------------------------------- one matrix run
interface Outcome {
  effect: Effect;
  mutated: boolean;
  intents: EditorIntent[];
  /** The live context at selection time — what the UI would have gated on. */
  context: SelectedObjectContext | null;
  note?: string;
}

/**
 * Some commands are state SETTERS: applying one whose state already holds
 * changes nothing (a floating shape asked to float), and stacking order only
 * exists for a floating anchor at all. Those are honest no-ops, not dead
 * controls — so before calling such a command inert, re-run it from a state
 * where it can act. Kind-level deadness (SmartArt rotate) survives this and
 * still fails, which is the point.
 */
interface Precondition {
  label: string;
  /** "idempotent": the command's state simply already held (not a finding).
   * "precondition": the command CANNOT act in the object's default state —
   * a real, reportable gap in the offered surface. */
  kind: "idempotent" | "precondition";
  run: (m: Mounted) => Promise<void>;
}

function viaCommand(step: SelectedObjectCommand, kind: Precondition["kind"]): Precondition {
  return {
    label: step,
    kind,
    run: async (m) => {
      await act(async () => { m.api.runSelectedObjectCommand(step); });
      await tick();
    },
  };
}

function precondition(fixture: Fixture, command: SelectedObjectCommand): Precondition | null {
  // Asking for the wrap mode the object is already in changes nothing.
  if (command.startsWith("wrap")) {
    return viaCommand(command === "wrapSquare" ? "wrapInline" : "wrapSquare", "idempotent");
  }
  // Stacking order lives on a floating anchor; an inline object has none.
  if (command === "bringForward" || command === "sendBackward") return viaCommand("wrapSquare", "precondition");
  // A freshly inserted model already sits at zero rotation, so "Reset 3D" has
  // nothing to reset. Rotating it is a viewer drag in the browser (WebGL, not
  // reachable headlessly), so the state is set through core directly.
  if (command === "reset3d" && fixture.name === "model3d") {
    return {
      label: "a non-zero 3D rotation",
      kind: "precondition",
      run: async (m) => {
        setModel3DRotation(m.doc, m.doc.editableRoots()[0], { x: 30, y: 40, z: 50 });
      },
    };
  }
  return null;
}

/**
 * Run one command against a freshly built, freshly selected object and
 * classify what it did. A fresh mount per cell keeps cells independent — a
 * command that deletes or re-wraps its object cannot color the next one.
 */
async function runCell(
  fixture: Fixture,
  command: SelectedObjectCommand,
  collab: boolean,
  prepare: Precondition | null = null,
): Promise<Outcome> {
  const m = await mount(collab);
  try {
    await act(async () => { await fixture.insert(m); });
    await tick();
    if (!(await fixture.select(m))) {
      return { effect: "SKIPPED", mutated: false, intents: [], context: null, note: "could not select" };
    }
    if (prepare) {
      await prepare.run(m);
      // A preparation step can rebuild the page DOM and drop the selection.
      if (!m.api.hasSelectedObject() && !(await fixture.select(m))) {
        return { effect: "SKIPPED", mutated: false, intents: [], context: null, note: "lost selection while preparing" };
      }
    }
    const context = m.api.getSelectedObjectContext();
    const before = documentXml(m.doc);
    m.intents.length = 0;
    let returned = false;
    await act(async () => { returned = m.api.runSelectedObjectCommand(command); });
    await tick();

    // A dialog is an editor affordance; answering it lets the command's real
    // post-dialog effect show up in the mutation oracle.
    const dialogTitle = await answerDialog();
    const inlineEditor = !!m.container.querySelector("[data-dxw-smart-art-node-editor]");
    await tick();

    const mutated = documentXml(m.doc) !== before;
    const intents = [...m.intents];
    if (mutated) {
      return { effect: "DOC-MUTATION", mutated, intents, context, note: dialogTitle ? `via "${dialogTitle}" dialog` : undefined };
    }
    if (dialogTitle) return { effect: "UI-INTERACTION", mutated, intents, context, note: `"${dialogTitle}" dialog, no change` };
    if (inlineEditor) return { effect: "UI-INTERACTION", mutated, intents, context, note: "inline node text editor" };

    // Text-entry affordances (editText on a textbox shape) leave no marker of
    // their own: the object is deselected and the caret moves into the
    // object's story. Prove the affordance is LIVE the only way that matters —
    // a keystroke now lands in the document.
    if (returned && !m.api.hasSelectedObject()) {
      const textarea = m.container.querySelector<HTMLElement>("textarea") ?? m.container;
      await act(async () => {
        textarea.dispatchEvent(new KeyboardEvent("keydown", { key: "Z", bubbles: true, cancelable: true }));
      });
      await tick();
      if (documentXml(m.doc) !== before) {
        return { effect: "UI-INTERACTION", mutated: false, intents, context, note: "text entry opened (typing lands)" };
      }
    }
    return { effect: "INERT", mutated, intents, context, note: returned ? "returned true, did nothing" : undefined };
  } finally {
    closeStrayDialogs();
    await m.unmount();
  }
}

/** Kind reported by a fixture, for the matrix's kind column. */
async function fixtureKind(fixture: Fixture): Promise<SelectedObjectKind | "?"> {
  const m = await mount(false);
  try {
    await act(async () => { await fixture.insert(m); });
    await tick();
    if (!(await fixture.select(m))) return "?";
    return m.api.getSelectedObjectContext()?.kind ?? "?";
  } finally {
    await m.unmount();
  }
}

// ------------------------------------------------------------------- the runs
const kinds = new Map<string, SelectedObjectKind | "?">();

describe("capability matrix: the selected-object command surface", () => {
  beforeAll(async () => {
    for (const fixture of FIXTURES) kinds.set(fixture.name, fixture.skip ? "?" : await fixtureKind(fixture));
  }, 60_000);

  for (const fixture of FIXTURES) {
    it(`${fixture.name}: every offered command does something`, async () => {
      if (fixture.skip) {
        for (const command of SELECTED_OBJECT_COMMANDS) {
          matrix.push({ fixture: fixture.name, kind: "?", command, offered: false, effect: "SKIPPED", note: fixture.skip });
        }
        return;
      }
      const kind = kinds.get(fixture.name)!;
      const inertOffers: string[] = [];
      for (const command of SELECTED_OBJECT_COMMANDS) {
        let local = await runCell(fixture, command, false);
        let collab = await runCell(fixture, command, true);
        // A state setter looks inert when its state already holds. Re-run it
        // from a state where it can act before calling it dead.
        const prepare = local.effect === "INERT" ? precondition(fixture, command) : null;
        let needs: string | undefined;
        if (prepare) {
          const prepped = await runCell(fixture, command, false, prepare);
          if (prepped.effect !== "INERT" && prepped.effect !== "SKIPPED") {
            needs = prepare.label;
            local = prepped;
            collab = await runCell(fixture, command, true, prepare);
          }
        }
        const needsKind = needs ? prepare!.kind : undefined;
        // Offered is read from the LIVE context the UI itself would gate on.
        const offered = !!local.context && availableObjectCommands(local.context).includes(command);
        matrix.push({
          fixture: fixture.name,
          kind,
          command,
          offered,
          effect: local.effect,
          collabMutated: collab.mutated,
          collabIntents: collab.intents.length,
          collabIntentKinds: collab.intents.map((i) => (i as { kind: string }).kind).join(","),
          collabEffect: collab.effect,
          needs,
          needsKind,
          note: local.note,
        });
        const excused = EXCEPTIONS.some((e) => e.fixture === fixture.name && e.command === command);
        if (offered && local.effect === "INERT" && !excused) {
          inertOffers.push(`${command}${local.note ? ` (${local.note})` : ""}`);
        }
      }
      // INVARIANT A: nothing the UI offers may be a dead control.
      expect(inertOffers, `${fixture.name}: offered but INERT — hide it, fix it, or add an EXCEPTIONS entry`).toEqual([]);
    }, 120_000);
  }

  it("the SmartArt/chart rotate gate is load-bearing", () => {
    // The bug that started this file: Layout offered "Rotate" for SmartArt,
    // and setDrawingRotation no-ops on a graphic frame. The gate now hides
    // it. This pins BOTH halves — the offer is withheld AND the command is
    // genuinely inert there — so deleting the gate reopens the bug and
    // INVARIANT A fails instead of the change sailing through review.
    for (const fixture of ["smartArt", "chart"]) {
      const cell = matrix.find((c) => c.fixture === fixture && c.command === "rotate");
      expect(cell, `${fixture}/rotate missing from the matrix`).toBeTruthy();
      expect(cell!.offered, `${fixture} must not offer rotate`).toBe(false);
      expect(cell!.effect, `${fixture} rotate must be inert (that is why it is hidden)`).toBe("INERT");
    }
  });

  it("availableObjectCommands encodes the offered set the toolbars render", () => {
    const of = (kind: SelectedObjectKind, canEditText = false) =>
      availableObjectCommands({ kind, canEditText });
    expect(of("shape")).toContain("outline");
    expect(of("shape")).not.toContain("lineStyle");
    expect(of("line")).toContain("lineStyle");
    expect(of("line")).not.toContain("outline");
    expect(of("smartArt")).toContain("fill");
    expect(of("chart")).not.toContain("fill");
    expect(of("image")).toContain("altText");
    expect(of("model3d")).toContain("reset3d");
    expect(of("image")).not.toContain("reset3d");
    expect(of("shape")).not.toContain("editText");
    expect(of("shape", true)).toContain("editText");
    // The "arrange" toolbar feature flag is the only host-level input.
    expect(availableObjectCommands({ kind: "shape", canEditText: false }, { arrange: false }))
      .not.toContain("bringForward");
    // Every command in the union is reachable from some kind — a command no
    // kind offers is dead code, not a feature.
    const everywhere = new Set(
      (["shape", "line", "smartArt", "chart", "image", "model3d"] as SelectedObjectKind[])
        .flatMap((kind) => of(kind, true)),
    );
    expect([...SELECTED_OBJECT_COMMANDS].filter((command) => !everywhere.has(command))).toEqual([]);
  });

  it("collab: every mutation emitted at least one intent (INVARIANT B)", () => {
    const silent = matrix
      .filter((cell) => cell.collabMutated && (cell.collabIntents ?? 0) === 0)
      .map((cell) => `${cell.fixture}/${cell.command}`);
    expect(silent, "mutated the document in collab mode without emitting an intent — the room forks").toEqual([]);
  });

  afterAll(() => {
    const pad = (s: string, n: number) => s.padEnd(n);
    const lines: string[] = [];
    lines.push("");
    lines.push("CAPABILITY MATRIX (offered? / local effect / collab: mutated:intents)");
    lines.push(`${pad("fixture", 15)}${pad("kind", 10)}${pad("command", 17)}${pad("offered", 9)}${pad("effect", 16)}collab`);
    for (const cell of matrix) {
      const collab = cell.effect === "SKIPPED"
        ? "-"
        : `${cell.collabMutated ? "mutated" : "no-op"}:${cell.collabIntents ?? 0}${cell.collabIntentKinds ? ` (${cell.collabIntentKinds})` : ""}`;
      lines.push(
        pad(cell.fixture, 15) + pad(String(cell.kind), 10) + pad(cell.command, 17) +
        pad(cell.offered ? "yes" : "no", 9) + pad(cell.effect, 16) + collab +
        (cell.needs ? `  // needs ${cell.needs} first` : "") +
        (cell.note ? `  // ${cell.note}` : ""),
      );
    }
    const hidden = matrix.filter((c) => !c.offered && c.effect !== "INERT" && c.effect !== "SKIPPED");
    lines.push("");
    lines.push("HIDDEN-BUT-WORKING (not offered, yet does something — candidate features):");
    for (const cell of hidden) lines.push(`  ${cell.fixture}/${cell.command} -> ${cell.effect}`);
    if (hidden.length === 0) lines.push("  (none)");
    const stateDependent = matrix.filter((c) => c.offered && c.needsKind === "precondition");
    lines.push("");
    lines.push("STATE-DEPENDENT OFFERS (offered, but a no-op in the object's default state):");
    for (const cell of stateDependent) lines.push(`  ${cell.fixture}/${cell.command} -> needs ${cell.needs}`);
    if (stateDependent.length === 0) lines.push("  (none)");
    const collabDead = matrix.filter((c) => c.offered && c.effect !== "INERT" && c.effect !== "SKIPPED" && c.collabEffect === "INERT");
    lines.push("");
    lines.push("COLLAB-ONLY DEAD CONTROLS (offered and works locally, honest no-op in a room):");
    for (const cell of collabDead) lines.push(`  ${cell.fixture}/${cell.command}`);
    if (collabDead.length === 0) lines.push("  (none)");
    // eslint-disable-next-line no-console
    console.log(lines.join("\n"));
  });
});
