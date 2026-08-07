# Interactive-editing conformance — failure inventory

Produced by the suite in this directory at engine tip `2bcb71c` (branch
`edit-conformance`). Two mounts are exercised everywhere:

- **session** — `DocxView` + `LocalDocumentSession` via
  `useAgentDocumentSession`, byte-for-byte the likeoffice desktop mount
  (`apps/desktop/src/renderer/src/App.tsx`). The editor runs collab-gated
  (`doc.stableIds` set, intents emitted).
- **local** — solo `DocxView editable` (local `EditHistory`).

All gestures are dispatched as real `KeyboardEvent` / `ClipboardEvent` /
`MouseEvent` against the editor's own listeners. The matrix holds 78 cases x
2 modes plus auto-generated undo/redo companions; every divergence is a
`KNOWN_GAPS` entry, enforced in both directions (an unlisted failure fails
the suite; a listed entry that starts passing also fails the suite). The
fuzzer ran 50,000 steps (2,500 x seeds 1-10 x both modes, mulberry32,
deterministic) plus the default smoke run.

## The three known bugs (excluded from ranking — harness validation)

- **Paste undo dead** — caught, and generalized: in the session mount
  **every** undo/redo is dead, not just paste. `applyHistory` routes to
  `onCollabUndo` past the collab gate (`core/src/edit/editor.ts` ~6354) and
  `LocalDocumentSession` wires no `undoLast`, so Cmd+Z / Cmd+Shift+Z decline
  for all 63 mutating session cases (gap `session/undoredo.*`), and the
  fuzzer's I5(session) invariant confirms undo never changes the document.
- **Plain paste** — the contract is pinned by five `paste.plain-*` cases
  (inline at caret, multiline split, into empty paragraph, over selection,
  in table cell). At this engine tip they PASS in both mounts, so the
  reported app symptom lives above the engine (app clipboard wiring); any
  engine regression now fails these cases.
- **Empty list item deletion** — caught twice: `backspace.list-item-empty`
  (engine hard-deletes the item where Word removes the bullet first — gap
  family `*/backspace.list-*`) and `backspace.list-item-empty-then-char`
  (the follow-up merge strands a zero-length `w:t`, a structural-lint
  violation — same family, and fuzz family G13).

## Ranked findings

Severity ranks user impact of the interactive loop. "Layer" is the suspected
owning code, from reading the gesture path (`core/src/edit/editor.ts` unless
noted).

**F1 - CRITICAL - G15 — keyboard selection wedges the editor at an empty
paragraph.** `Shift+ArrowRight` at the end of a paragraph whose next
paragraph is empty produces a zero-width focus selection: no caret and no
selection are reported, arrow keys cannot collapse it, and every subsequent
keystroke is swallowed until a mouse click. Both mounts. Found by the fuzzer
(seed 2 minimal repro: `Enter, ArrowLeft, Shift+ArrowRight`; seed 1:
`ArrowRight x2, Delete x2, Shift+ArrowRight`); pinned as
`select.shift-arrow-empty-para`. Crossing into a NONEMPTY paragraph works.
Fuzz observed the wedged state ~800-1,100 step-observations per session
seed. Layer: `moveFocus`/`stepPoint` + selection-segment derivation for
empty-paragraph placeholder runs.

**F2 - HIGH - G14 — Enter with an active selection is a complete no-op.**
Word deletes the selection and splits at that point. Both mounts
(`enter.over-selection`). Layer: `splitParagraph` declines when a range
selection is active instead of collapsing it first.

**F3 - HIGH - G8 — deleting/typing over an across-paragraph selection keeps
the paragraph boundary.** Word merges the surviving halves into one
paragraph; the engine deletes only the characters
(`backspace.across-para-selection`, `type.over-across-para-selection`). Both
mounts. Layer: `removeSelectedText` multi-paragraph path (no merge step).

**F4 - HIGH - G5/G6 — rich paste is block-level and shatters inline HTML.**
A one-paragraph rich copy pastes as its own paragraph plus a stranded EMPTY
paragraph instead of merging inline (`paste.rich-single-inline`,
`paste.rich-multi`, `paste.cutpaste-roundtrip`), and external HTML like
`<b>H1</b> plus <i>H2</i>` becomes one paragraph PER inline chunk
(`paste.html-inline-runs`). Both mounts. Layer:
`core/src/edit/clipboard.ts` `htmlClipboardBlocks` (inline elements become
blocks) + `pasteBlocks` (always splices whole paragraphs).

**F5 - MEDIUM-HIGH - G7 — session-only: select-all + delete/type leaves one
empty paragraph per former paragraph.** Word (and the LOCAL mount, which
passes) collapse to a single paragraph — a session/local behavior fork in
the same engine (`selectall.delete`, `selectall.type`,
`selectall.delete-with-table`). Layer: the collab-gated multi-block removal
path lacks the block merge the local path has.

**F6 - MEDIUM - G13 — paragraph merges strand zero-length runs.** Removing
an empty paragraph (or an empty list item, then a character) merges its
placeholder `<w:t/>` into the neighbor, leaving `<w:t xml:space="preserve"/>`
beside real content. Dominant fuzz signal: tens of thousands of
step-observations across the seed-runs (a stranded run persists and is
re-observed each step). Visible consequence beyond hygiene: `Enter` then
`Backspace` does NOT return the document to its prior bytes (undo eventually
does, after unwinding both steps). Layer: paragraph merge in
`deleteContents`/`mergeParagraphBackward` (`core/src/edit/blocks.ts`).

**F7 - MEDIUM - G2 — Shift+Enter is a paragraph split, not a soft line
break.** No `w:br` is ever produced; a list item splits into two items.
(`shiftenter.para-middle`, `.list-item`, `.table-cell`.) Layer: `onKeyDown`
Enter branch has no `shiftKey` case and no line-break insert.

**F8 - MEDIUM - G4 — no pending format at a collapsed caret.** Cmd+B/I/U
without a selection does nothing; the next typed character is unformatted
(`format.caret-bold-then-type`, `format.caret-italic-then-type`). Layer:
`onKeyDown` forwards the shortcut only `if (hasSelection())`; there is no
pending-format state in the typing path.

**F9 - MEDIUM - G9 — Backspace at list-item start skips Word's
bullet-removal step.** The engine merges/deletes the paragraph immediately;
Word first converts the item to a plain paragraph
(`backspace.list-item-start`, `backspace.list-item-empty`). Layer:
`deleteContents(-1)` at paragraph start does not consult list context
(`core/src/edit/lists.ts`).

**F10 - MEDIUM-LOW - G10 — Enter after a heading inherits the heading
style.** Word applies the style's next-style (Normal); the new paragraph
stays Heading1 (`enter.heading-end`). Layer: `splitParagraph` clones `pPr`
wholesale; no next-style lookup against the styles part.

**F11 - LOW-MEDIUM - G11 — Shift+Tab on a level-0 list item is a no-op.**
Word promotes the item out of the list (`shifttab.list-item-level0`). Layer:
`setListLevel(-1)` bottoms out instead of removing `numPr`.

**F12 - LOW-MEDIUM - G3 — Tab in a body paragraph is a no-op.** Word inserts
a tab character (or steps first-line indent at paragraph start)
(`tab.plain-paragraph`). Layer: `onKeyDown` Tab falls through when the caret
is in neither a table nor a list.

**F13 - LOW - G12 — cross-cell selection deletes character-wise.** Word
switches to whole-cell selection and clears the selected cells. The hard
safety half — table structure always survives — PASSES
(`backspace.across-cell-structure`). Layer: `edit/selection.ts` has no
cell-granular selection mode.

## What held (worth knowing)

- No listener ever threw across the fuzz sweep + the matrix (I1).
- Every post-gesture document saved and reloaded byte-stably (I4) — no case
  or fuzz step produced an unloadable .docx.
- LOCAL undo/redo is byte-exact for every matrix case and every fuzz seed
  once the ladder unwinds multi-checkpoint gestures (I5).
- Table integrity: no gesture sequence ever merged/destroyed cells or rows
  outside explicit table ops; Tab nav/append-row and cell-local editing all
  conform.
- Doc-start/doc-end guards, paragraph merge on Backspace, Tab list
  indent/outdent (level >= 1), Enter-exits-empty-list-item, tracked-change
  markers under plain editing: all conform to Word.

## Fuzzer runbook

Default (CI smoke): 120 steps x seeds {1,2} x both modes, green with known
families counted. Scale: `FUZZ_STEPS=2500 FUZZ_SEEDS=1,...,10` (see
fuzz.test.tsx header). Any new violation is ddmin-minimized (wall-clock
bounded) and fails with `seed`, `mode`, `step`, and the minimal gesture
script as JSON.
