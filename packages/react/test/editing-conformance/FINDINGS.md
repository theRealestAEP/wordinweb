# Interactive-editing conformance — failure inventory

> **Editor-residuals wave addendum (branch `editor-residuals`).** The last
> ledgered findings are FIXED and the KNOWN_GAPS table is EMPTY:
> F13/G12 (a selection whose endpoints sit in different cells of one table
> is now cell-granular — delete clears the whole covered cells through the
> ordinary deleteText/merge machinery; the across-cell-structure case's
> incidental "r0c0 text survives" assertion was amended to the uncovered
> cell), F5/G7's table residual (a fully selected table now deletes AS A
> BLOCK on the session mount via the canonical tableOp("deleteTable")
> mutation + intent), and F15/G17 (root-caused to a STALE CARET PATH: the
> checkpoint after applyRunFormat's split fell back to a pre-split
> structural path, undo resolved it to an rPr, and Enter split a paragraph
> around it — fixed by rebinding the caret in selectRanges plus a
> text-leaf guard in history's restoreCaret, not by lint-scrubbing). The
> separator-awareness residual recorded below is fixed too: Backspace/
> Delete adjacent to a soft break or tab delete the SEPARATOR itself
> (applyDeleteSeparator — applyInsertSeparator's exact inverse, a one-unit
> deleteSeparator wire intent; a w:del around the separator's run slice in
> suggesting mode), pinned by backspace.soft-break, delete.soft-break, and
> backspace.tab-char. Sweep results for this wave are at the bottom of
> this file.

> **Keyboard-contract wave addendum (branch `keyboard-contracts`).** The
> following findings below are FIXED and un-ledgered (see known-gaps.ts for
> the per-fix notes): F1/G15 (the Shift+Arrow empty-paragraph wedge — plus a
> previously masked caret loss on forward-Delete in an empty paragraph),
> F2/G14 (Enter over a selection), F6/G13's merge half and G18
> (mergeParagraphBackward no longer strands placeholder runs), F7/G2
> (Shift+Enter now inserts a w:br via the new insertSeparator intent),
> F8/G4 (pending Cmd+B/I/U at a collapsed caret), F10/G10 (next-style on
> Enter at a styled paragraph end), F11/G11 (Shift+Tab level-0 unbullet),
> F12/G3 (Tab types a w:tab in body paragraphs; Word's first-line-indent
> nuance at the exact paragraph start remains a recorded divergence).
> Residuals recorded honestly: Backspace/Delete are not separator-aware
> (a soft break is stepped past, not deleted), and char-wise deletes can
> still empty one w:t beside populated runs (the surviving G13 fuzz
> family). The fuzzer's allowed "I3" family is deleted — any caret/
> selection resolvability loss now fails loudly — and deleting it paid
> off immediately: the first full sweep surfaced three UNDO-path caret
> losses the family had been masking (a stale selection kept after a
> history install, and no caret restored when the checkpoint had been
> taken mid-selection), fixed by applyHistory's post-replay
> reconciliation (clear the selection, land a caret) plus an attachment
> guard on the boundary-selection fallback. Sweep results for this wave
> are at the bottom of this file.

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
deterministic) plus the default smoke run. Final suite shape: 80 cases x 2
modes + companions = 321 matrix tests, 36 KNOWN_GAPS entries covering 115
case results.

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

Severity ranks user impact of the interactive loop; F-numbers are discovery
order, the list is severity order. "Layer" is the suspected owning code, from
reading the gesture path (`core/src/edit/editor.ts` unless noted).

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

**F14 - CRITICAL (crash) - G16 — keystroke after undo of
type-over-selection throws.** In the local mount (real undo), undoing a
character typed over a keyboard selection leaves a caret/run binding whose
shape the next keystroke cannot consume: the keydown listener throws
`TypeError: run.content is not iterable` and the gesture is lost. Found by
the fuzzer, two independent seeds; minimal repros (kitchen-sink fixture,
caret after "alpha"): seed 5 — `pasteHtml <b>bold</b>, Backspace, Home,
Shift+ArrowLeft, type f, Cmd+Z, type f`; seed 6 — `Shift+Enter, Backspace,
Backspace, Delete, Shift+ArrowLeft, type i, Cmd+Z, Backspace`. Pinned as
`undo.crash-after-type-over-selection`. The session mount is immune only
because its undo is dead (G1) — fixing G1 without fixing this imports the
crash into likeoffice. Layer: history restore -> caret rebinding
(`applyHistory` / `positionCaret` run binding).

**F15 - MEDIUM-HIGH - G17 — undo of format-over-selection then Enter strands
a properties-only run.** Local mount: `Enter, type j, Shift+ArrowLeft,
Cmd+B, type e, Cmd+Z, Enter` leaves a `w:r` holding only `rPr` (structural
lint violation; fuzz seeds 6 and 7). In the session mount the same sequence
strands a zero-length `w:t` instead (G13 family). Pinned as
`undo.format-selection-strands-run`. Layer: history text-leaf restore vs
runs created by `applyRunFormat` splitting.

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

## Sweep summary (2,500 steps x seeds 1-10 x mode)

- session: all 10 seeds green; known families counted — G13 strands
  (~30k-56k step-observations per seed once a strand persists) and G15
  wedge states (~640-1,100 per seed). No crash, no save/reload failure.
- local: all 10 seeds green after G16/G17 were pinned; G16 crash hit in
  seeds 5, 6, 9; G17 strand in seeds 6, 7, 9; every seed exceeded the
  200-entry undo depth exactly once and redo replayed byte-exactly from the
  horizon every time.

## What held (worth knowing)

- No listener ever threw across the fuzz sweep + the matrix (I1) — except
  the G16 crash family above, which is pinned and counted.
- Every post-gesture document saved and reloaded byte-stably (I4) — no case
  or fuzz step produced an unloadable .docx.
- LOCAL undo/redo is byte-exact for every matrix case and every fuzz seed
  once the ladder unwinds multi-checkpoint gestures (I5). One spec discovery
  (not a defect): `EditHistory` keeps a designed 200-entry undo depth
  (`core/src/edit/history.ts` `limit = 200`, oldest evicted) — a 2,500-step
  sequence therefore unwinds only to the horizon. The I5 invariant respects
  the horizon, and REDO from it still replays byte-exactly to the final
  state on every seed.
- Table integrity: no gesture sequence ever merged/destroyed cells or rows
  outside explicit table ops; Tab nav/append-row and cell-local editing all
  conform.
- Doc-start/doc-end guards, paragraph merge on Backspace, Tab list
  indent/outdent (level >= 1), Enter-exits-empty-list-item, tracked-change
  markers under plain editing: all conform to Word.

## Fuzzer runbook

Default (CI smoke): 120 steps x seeds {1,2} x both modes, green with known
families counted. The undo invariant honors EditHistory's designed
200-entry depth: within it, undo must return the pre-sequence bytes; past
it, undo may stop at the horizon but redo must still replay byte-exactly. Scale: `FUZZ_STEPS=2500 FUZZ_SEEDS=1,...,10` (see
fuzz.test.tsx header). Any new violation is ddmin-minimized (wall-clock
bounded) and fails with `seed`, `mode`, `step`, and the minimal gesture
script as JSON.

## Keyboard-contract wave sweep (branch `keyboard-contracts`, 2,500 steps x seeds 1-10 x both modes)

All 20 seed-runs green — 50,000 steps, zero hard violations. The blanket
"I3" allowance is gone, so a wedge or any caret/selection resolvability
loss now fails the sweep outright; the first run at this scale proved the
point by surfacing three undo-path caret losses (seeds 3/4/6, session),
fixed before this final run (applyHistory reconciliation + the
boundary-selection attachment guard). Remaining allowed families:

- G13 residual (char-wise deletes emptying ONE w:t beside populated
  runs): 237-3,205 step-observations per seed-run — down from the
  baseline's ~30k-56k per seed now that paragraph merges drop placeholder
  runs instead of stranding them.
- I5-depth: exactly one per seed-run (2,500 steps exceed the designed
  200-entry undo horizon; redo replayed byte-exactly every time).
- G17 (properties-only w:r after undo of format-over-selection): zero
  hits this sweep; its matrix cases remain ledgered.

No listener threw, every periodic save/reload round-tripped, and undo
returned the pre-sequence bytes (or the horizon, redo-exact) on every
seed in both mounts.
