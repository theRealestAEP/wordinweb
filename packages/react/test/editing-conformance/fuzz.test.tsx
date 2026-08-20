// @vitest-environment jsdom
/**
 * GESTURE FUZZER — random gesture sequences against invariants that hold
 * REGARDLESS of contract details:
 *
 *   I1  no event listener throws;
 *   I2  the document passes structural lint after every gesture;
 *   I3  editing is never lost: whenever neither an encoded caret nor a
 *       selection is reported, an arrow-key collapse or a typed character
 *       must still act on the document (hard violation only when typing goes
 *       nowhere — the reporting-only variant is a counted known finding);
 *   I4  the document saves and reloads cleanly (checked periodically);
 *   I5  undo pressed repeatedly returns the document byte-identical to the
 *       pre-sequence state — within EditHistory's designed 200-entry depth;
 *       a longer sequence may legitimately stop at the undo horizon, but
 *       redo must always replay byte-exactly back to the post-sequence
 *       state. BOTH mounts: the session mount replays local history too
 *       since the G1 fix (EditorHost.onLocalHistory).
 *
 * Deterministic: mulberry32 PRNG, the seed printed with every run and every
 * violation. Violations in KNOWN families (FUZZ_ALLOWED) are counted and
 * reported, not failed; anything else is MINIMIZED (ddmin over the recorded
 * gesture script) and thrown with its seed + minimal repro script.
 *
 * Opt-in scale: FUZZ_STEPS (steps per seed, default 120 as a smoke run) and
 * FUZZ_SEEDS (comma list, default "1,2"). Example:
 *   FUZZ_STEPS=2500 FUZZ_SEEDS=1,2,3,4,5 vitest run test/editing-conformance/fuzz.test.tsx
 */
import { describe, expect, it } from "vitest";
import { act } from "react";
import { mountEditor, structuralLint, saveReloadProblems, type Editor, type Mode } from "./harness.js";
import { kitchenSinkDoc } from "./fixtures.js";

const STEPS = Number(process.env.FUZZ_STEPS ?? 120);
const SEEDS = (process.env.FUZZ_SEEDS ?? "1,2").split(",").map((s) => Number(s.trim()));
const MODES: Mode[] = process.env.FUZZ_MODE ? [process.env.FUZZ_MODE as Mode] : ["session", "local"];
const SAVE_CHECK_EVERY = 25;

/** mulberry32 — matches the PRNG the existing stress tests use. */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

type Gesture =
  | { k: "key"; key: string; meta?: boolean; shift?: boolean }
  | { k: "copy" }
  | { k: "cut" }
  | { k: "paste" }
  | { k: "pastePlain"; text: string }
  | { k: "pasteHtml"; html: string; plain: string }
  | { k: "click"; idx: number };

const LETTERS = "abcdefgh ij";
const PLAIN_SNIPPETS = ["zz", "one\ntwo", "  spaced  ", "line1\nline2\nline3"];
const HTML_SNIPPETS: { html: string; plain: string }[] = [
  { html: "<b>bold</b>", plain: "bold" },
  { html: "<p>pp1</p><p>pp2</p>", plain: "pp1\npp2" },
  { html: "plain <i>ital</i> mix", plain: "plain ital mix" },
];

function nextGesture(rand: () => number, docLen: number): Gesture {
  const r = rand();
  // Weighted vocabulary. Paste-family gestures are skipped on huge documents
  // (deterministic: depends only on deterministic state) to stop exponential
  // growth from copy-all + paste loops.
  if (r < 0.34) return { k: "key", key: LETTERS[Math.floor(rand() * LETTERS.length)] };
  if (r < 0.46) return { k: "key", key: "Backspace" };
  if (r < 0.52) return { k: "key", key: "Delete" };
  if (r < 0.60) return { k: "key", key: "Enter" };
  if (r < 0.63) return { k: "key", key: "Enter", shift: true };
  if (r < 0.70) {
    const key = rand() < 0.5 ? "ArrowLeft" : "ArrowRight";
    return { k: "key", key, shift: rand() < 0.35 };
  }
  if (r < 0.74) {
    const key = rand() < 0.5 ? "ArrowUp" : "ArrowDown";
    return { k: "key", key, shift: rand() < 0.2 };
  }
  if (r < 0.77) return { k: "key", key: rand() < 0.5 ? "Home" : "End" };
  if (r < 0.80) return { k: "key", key: "Tab", shift: rand() < 0.4 };
  if (r < 0.82) return { k: "key", key: "a", meta: true };
  if (r < 0.85) return { k: "key", key: ["b", "i", "u"][Math.floor(rand() * 3)], meta: true };
  if (r < 0.87) return { k: "key", key: "z", meta: true };
  if (r < 0.88) return { k: "key", key: "z", meta: true, shift: true };
  if (r < 0.91) return { k: "click", idx: Math.floor(rand() * 40) };
  if (r < 0.93) return { k: "copy" };
  if (r < 0.945) return { k: "cut" };
  if (docLen > 20_000) return { k: "key", key: "Backspace" };
  if (r < 0.965) return { k: "paste" };
  if (r < 0.985) return { k: "pastePlain", text: PLAIN_SNIPPETS[Math.floor(rand() * PLAIN_SNIPPETS.length)] };
  return { k: "pasteHtml", ...HTML_SNIPPETS[Math.floor(rand() * HTML_SNIPPETS.length)] };
}

async function applyGesture(ed: Editor, g: Gesture): Promise<void> {
  switch (g.k) {
    case "key":
      await ed.press(g.key, { meta: g.meta, shift: g.shift });
      return;
    case "copy":
      await ed.copy();
      return;
    case "cut":
      await ed.cut();
      return;
    case "paste":
      await ed.paste();
      return;
    case "pastePlain":
      await ed.paste({ "text/plain": g.text });
      return;
    case "pasteHtml":
      await ed.paste({ "text/html": g.html, "text/plain": g.plain });
      return;
    case "click": {
      const page = ed.container.querySelector<HTMLElement>(".dxw-page");
      const spans = page ? [...page.querySelectorAll("span")] : [];
      const el = spans.length > 0 ? spans[g.idx % spans.length] : page;
      if (!el) return;
      await act(async () => {
        const o = { bubbles: true, cancelable: true, clientX: 10, clientY: 10, button: 0 };
        el.dispatchEvent(new MouseEvent("mousedown", o));
        el.dispatchEvent(new MouseEvent("mouseup", o));
        await new Promise((r) => setTimeout(r, 1));
      });
      return;
    }
  }
}

/**
 * Violation families the suite already pins as KNOWN (matrix KNOWN_GAPS or
 * the findings inventory): counted here, never failed — every other
 * violation fails with a minimized repro.
 */
const FUZZ_ALLOWED: { match: string; family: string }[] = [
  {
    match: "zero-length w:t stranded",
    family: "G13 (residual) — char-wise deletes can empty ONE w:t beside populated runs; the merge half is FIXED (mergeParagraphBackward drops placeholder runs)",
  },
  // G15 (the Shift+Arrow-into-empty-paragraph wedge, formerly the "I3" family
  // here) is FIXED — zero-segment selections collapse via selection points and
  // delete their covered paragraph mark. Deliberately NOT allowed anymore: a
  // recurrence of any I3 resolvability loss must fail loudly with a repro.
  // G16 (keystroke after undo of type-over-selection threw "run.content is
  // not iterable") is FIXED — checkboxStateElement tolerates unbound
  // placeholder caret runs. Deliberately NOT allowed here anymore: a
  // recurrence must fail loudly with a minimized repro.
  {
    match: "w:r holding only properties",
    family:
      "G17 — undo of format-over-selection then Enter strands a properties-only w:r shell (KNOWN_GAPS undo.format-selection-strands-run; fuzz seeds 6+7)",
  },
];

interface Violation {
  step: number;
  problems: string[];
}

function classify(problems: string[]): { allowed: Map<string, number>; hard: string[] } {
  const allowed = new Map<string, number>();
  const hard: string[] = [];
  for (const p of problems) {
    const hit = FUZZ_ALLOWED.find((a) => p.includes(a.match));
    if (hit) allowed.set(hit.family, (allowed.get(hit.family) ?? 0) + 1);
    else hard.push(p);
  }
  return { allowed, hard };
}

/**
 * I3 — resolvability with a recovery ladder. When neither a caret nor a
 * selection is reported, press ArrowLeft (the collapse gesture); if still
 * nothing, type a probe character: if even that changes nothing, editing is
 * truly lost (hard). If typing lands, only the REPORTING is broken (soft,
 * the known F-CARET family). The recovery presses intentionally act on the
 * editor — they are deterministic given the state, in the driver and in
 * replay alike.
 */
async function checkResolvable(ed: Editor): Promise<string | null> {
  const resolvable = (): boolean => ed.api.getEncodedCaret() !== null || ed.api.getSelectionFormat() !== null;
  if (ed.mode !== "session") return null; // encoded caret exists only with stable ids
  if (resolvable()) return null;
  await ed.press("ArrowLeft");
  if (resolvable()) return null;
  const before = ed.xml();
  await ed.press("k");
  if (ed.xml() === before) return "I3 caret/selection lost: arrows cannot collapse and typing changes nothing";
  return "I3-soft caret/selection unreported (typing still lands)";
}

/** EditHistory's designed depth (core/src/edit/history.ts `limit = 200`,
 * oldest entries evicted on push). Byte-identity to the PRE-SEQUENCE
 * document is only promised while the whole sequence fits inside this
 * horizon; past it, undo legitimately stops at the oldest surviving
 * checkpoint — but redo must STILL replay byte-exactly to the final state. */
const UNDO_DEPTH = 200;

/** Undo until byte-identical to `goal`, the stack empties, or `cap` presses
 * (consecutive identical states are legal — caret-only checkpoints).
 * Returns the number of presses made. */
async function undoTo(ed: Editor, goal: string, cap: number): Promise<number> {
  let i = 0;
  while (ed.xml() !== goal && ed.api.canUndo() && i < cap) {
    await ed.undo();
    i++;
  }
  return i;
}

/** Drive one full fuzz round (generation or replay). Returns the executed
 * script, the first hard violation, and counts of allowed-family hits. */
async function drive(
  mode: Mode,
  source: { rand: () => number; steps: number } | { replay: Gesture[] },
): Promise<{ script: Gesture[]; violation: Violation | null; allowedHits: Map<string, number> }> {
  const ed = await mountEditor(kitchenSinkDoc(), mode);
  const script: Gesture[] = [];
  const allowedHits = new Map<string, number>();
  const bump = (m: Map<string, number>): void => {
    for (const [f, n] of m) allowedHits.set(f, (allowedHits.get(f) ?? 0) + n);
  };
  try {
    await ed.caretAt("alpha", "end");
    const base = ed.xml();
    const total = "replay" in source ? source.replay.length : source.steps;
    for (let i = 0; i < total; i++) {
      const g = "replay" in source ? source.replay[i] : nextGesture(source.rand, ed.text().length);
      script.push(g);
      await applyGesture(ed, g);
      const problems: string[] = [];
      while (ed.errors.length > 0) problems.push(`I1 listener threw: ${String(ed.errors.shift())}`);
      for (const p of structuralLint(ed.doc())) problems.push(`I2 lint: ${p}`);
      const i3 = await checkResolvable(ed);
      if (i3) problems.push(i3);
      if ((i + 1) % SAVE_CHECK_EVERY === 0 || i === total - 1) {
        for (const p of saveReloadProblems(ed)) problems.push(`I4 save/reload: ${p}`);
      }
      const { allowed, hard } = classify(problems);
      bump(allowed);
      if (hard.length > 0) return { script, violation: { step: i, problems: hard }, allowedHits };
    }
    // I5 — the undo/redo ladder. Both mounts: since the G1 fix the session
    // mount replays local history exactly like the local one.
    const final = ed.xml();
    const undos = await undoTo(ed, base, script.length * 2 + 20);
    if (ed.xml() !== base) {
      // Short of base with entries left, or exhausted well inside the
      // designed horizon: the stack lost state it should still hold.
      if (ed.api.canUndo() || undos < UNDO_DEPTH - 20) {
        return {
          script,
          violation: {
            step: script.length,
            problems: [`I5 undo (x${undos}, canUndo=${ed.api.canUndo()}) never returns the pre-sequence document (byte-compare)`],
          },
          allowedHits,
        };
      }
      // Depth-bounded unwind: the sequence outran the 200-entry cap — by
      // design, counted (not failed); redo-exactness below still applies.
      allowedHits.set(
        "I5-depth — sequence exceeded EditHistory's designed 200-entry undo depth; unwound to the horizon (redo-exactness still enforced)",
        (allowedHits.get(
          "I5-depth — sequence exceeded EditHistory's designed 200-entry undo depth; unwound to the horizon (redo-exactness still enforced)",
        ) ?? 0) + 1,
      );
    }
    for (let i = 0; i < undos; i++) await ed.redo();
    if (ed.xml() !== final) {
      return {
        script,
        violation: { step: script.length, problems: [`I5 redo (x${undos}) did not return to the post-sequence state`] },
        allowedHits,
      };
    }
    return { script, violation: null, allowedHits };
  } finally {
    await ed.unmount();
  }
}

/** ddmin-style minimizer: drop chunks while the replay still hard-fails.
 * Wall-clock bounded so a violation deep in a long sweep still reports a
 * (possibly non-minimal) repro instead of stalling the run. */
async function minimize(mode: Mode, script: Gesture[], budgetMs = 10 * 60_000): Promise<Gesture[]> {
  const deadline = Date.now() + budgetMs;
  const fails = async (s: Gesture[]): Promise<boolean> => (await drive(mode, { replay: s })).violation !== null;
  let current = script;
  // A violation usually needs only recent context: try aggressive tail-only
  // prefixes first — cheap wins before ddmin proper.
  for (const keep of [200, 50, 12]) {
    if (current.length > keep && (await fails(current.slice(-keep)))) current = current.slice(-keep);
  }
  let chunk = Math.max(1, Math.floor(current.length / 2));
  while (chunk >= 1 && Date.now() < deadline) {
    let removed = false;
    for (let start = 0; start + chunk <= current.length && Date.now() < deadline; start += chunk) {
      const candidate = [...current.slice(0, start), ...current.slice(start + chunk)];
      if (candidate.length > 0 && (await fails(candidate))) {
        current = candidate;
        removed = true;
        start -= chunk;
      }
    }
    if (!removed) {
      if (chunk === 1) break;
      chunk = Math.floor(chunk / 2);
    }
  }
  return current;
}

describe(`gesture fuzzer (${STEPS} steps x seeds [${SEEDS.join(",")}] x modes [${MODES.join(",")}])`, () => {
  for (const mode of MODES) {
    for (const seed of SEEDS) {
      it(`seed ${seed} (${mode}) holds all invariants`, async () => {
        const { script, violation, allowedHits } = await drive(mode, { rand: rng(seed), steps: STEPS });
        if (allowedHits.size > 0) {
          console.error(
            `[fuzz seed ${seed} ${mode}] known-family hits (allowed, counted): ` +
              [...allowedHits.entries()].map(([f, n]) => `${n}x {${f}}`).join("; "),
          );
        }
        if (violation) {
          console.error(`[fuzz seed ${seed} ${mode}] VIOLATION at step ${violation.step}; minimizing…`);
          const minimal = await minimize(mode, script.slice(0, violation.step + 1));
          const replay = await drive(mode, { replay: minimal });
          const detail =
            `FUZZ VIOLATION seed=${seed} mode=${mode} step=${violation.step}\n` +
            `problems: ${violation.problems.join(" | ")}\n` +
            `minimal repro (${minimal.length} gestures): ${JSON.stringify(minimal)}\n` +
            `minimal replay problems: ${JSON.stringify(replay.violation?.problems ?? [])}`;
          console.error(detail);
          expect.fail(detail);
        }
      }, 3_600_000);
    }
  }
});
