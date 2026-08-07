/**
 * KNOWN_GAPS — the explicit, countable ledger of Word-contract divergences.
 *
 * Every entry keeps its matrix test GREEN while the engine diverges from the
 * Word contract the test encodes. The runner enforces hygiene in both
 * directions: a test that fails without an entry here fails the suite, and an
 * entry whose test now PASSES also fails the suite (remove the entry — the
 * gap is fixed). So this table is always the exact current divergence count.
 *
 * Keys are `${mode}/${caseId}` (mode: "session" = the likeoffice
 * LocalDocumentSession mount, "local" = solo DocxView). A trailing `*` makes
 * the key a prefix that covers a family of cases.
 */
export const KNOWN_GAPS: Record<string, string> = {
  // G1 — FIXED: undo/redo were DEAD in the session mount (applyHistory routed
  // to an unwired onCollabUndo hook). The single-process session now replays
  // the local history stack (EditorHost.onLocalHistory, set by
  // useAgentDocumentSession via collab.noteLocalHistory), so every
  // session/undoredo.* companion passes.

  // G2 — FIXED: Shift+Enter inserts a w:br soft line break in-paragraph via
  // applyInsertSeparator (in-run w:t split; insertSeparator intent — a
  // one-unit wire insert). Word-contract residual, recorded honestly:
  // Backspace/Delete are not yet separator-aware, so the caret steps PAST a
  // soft break and deletes the neighboring character instead of the break
  // itself; the break disappears only with its paragraph.

  // G3 — Tab in a plain body paragraph is a silent no-op; Word inserts a tab
  // character (or first-line indent at paragraph start).
  "session/tab.plain-paragraph": "Tab in a body paragraph is a no-op (Word inserts a tab character)",
  "local/tab.plain-paragraph": "Tab in a body paragraph is a no-op (Word inserts a tab character)",

  // G4 — FIXED: Cmd+B/I/U at a collapsed caret queue a pending toggle
  // (editor.pendingFormat) that the next insertText consumes by formatting
  // the typed range through the host's shortcut path (canonical
  // formatRun/formatRange intents, suggest-aware). Any other gesture —
  // arrows, clicks, deletes — clears the queue, like Word.

  // G5 — rich (text/html) paste is always BLOCK-level: a one-paragraph rich
  // paste splits the target paragraph, inserts its own paragraph, and strands
  // an empty paragraph; Word merges a one-paragraph paste inline at the caret.
  "session/paste.rich-*": "rich paste lands as separate paragraphs plus a stranded empty paragraph (Word pastes a single-paragraph copy inline)",
  "local/paste.rich-*": "rich paste lands as separate paragraphs plus a stranded empty paragraph (Word pastes a single-paragraph copy inline)",
  "session/paste.cutpaste-roundtrip": "cut+paste of an inline snippet lands as its own paragraph (Word restores it inline)",
  "local/paste.cutpaste-roundtrip": "cut+paste of an inline snippet lands as its own paragraph (Word restores it inline)",

  // G6 — external HTML with inline markup (<b>x</b> plus <i>y</i>) shatters
  // into one paragraph PER inline element; Word keeps one paragraph with
  // formatted runs.
  "session/paste.html-inline-runs": "inline HTML formatting chunks each become their own paragraph on paste",
  "local/paste.html-inline-runs": "inline HTML formatting chunks each become their own paragraph on paste",

  // G7 (residual after the 734b6a8 spanning-merge fix) — select-all + delete
  // in a document WITH A TABLE leaves the table standing on the session
  // mount; Word removes it and leaves one empty paragraph.
  "session/selectall.delete-with-table": "select-all + delete leaves the table standing (Word removes it, leaving one empty paragraph)",

  // G18 — FIXED with G13: mergeParagraphBackward leaves the merged
  // paragraph's blank placeholder runs behind, so the unbullet/merge path
  // strands nothing.

  // G8 — deleting a selection that spans a paragraph boundary deletes the
  // characters but KEEPS the boundary; Word merges the two paragraphs.

  // G9 — Backspace at the start of a list item deletes/merges the paragraph
  // immediately; Word first removes the bullet (the paragraph survives as
  // plain text) and only a second Backspace merges. The empty-item variant is
  // the reported "undeletable/mishandled empty list item" family.

  // G10 — Enter at the end of a heading carries Heading1 onto the new
  // paragraph; Word applies the style's next-style (Normal).
  "session/enter.heading-end": "the paragraph created after a heading keeps the heading style (Word switches to Normal)",
  "local/enter.heading-end": "the paragraph created after a heading keeps the heading style (Word switches to Normal)",

  // G11 — Shift+Tab on a level-0 list item is a no-op; Word promotes the item
  // out of the list into a body paragraph.
  "session/shifttab.list-item-level0": "Shift+Tab at list level 0 is a no-op (Word converts the item to a body paragraph)",
  "local/shifttab.list-item-level0": "Shift+Tab at list level 0 is a no-op (Word converts the item to a body paragraph)",

  // G13 — FIXED (merge half): mergeParagraphBackward drops the merged
  // paragraph's blank placeholder runs (rPr + zero-length w:t) instead of
  // splicing them beside real content. Char-wise deletes that empty ONE w:t
  // beside populated runs can still strand (fuzz family stays counted).

  // G14 — FIXED: Enter with an active selection now composes
  // removeSelectedText (canonical delete + intents) with the ordinary split,
  // exactly like typing over a selection.

  // G15 — FIXED: the Shift+Arrow-into-empty-paragraph wedge. A selection with
  // no character segments now still collapses (arrow keys use the selection
  // POINTS) and still deletes its covered paragraph mark (removeSelectedText
  // derives the spanned paragraphs from the selection endpoints), so typing
  // over it merges the paragraphs and lands the character — Word's contract.

  // G12 — a selection extended across a table-cell boundary deletes
  // character-wise; Word switches to whole-cell selection and clears the
  // selected cells' contents (structure survives in both — that hard
  // invariant is asserted separately and passes).
  "session/backspace.across-cell-word-semantics": "cross-cell delete is character-wise (Word clears whole selected cells)",
  "local/backspace.across-cell-word-semantics": "cross-cell delete is character-wise (Word clears whole selected cells)",

  // G16 — FIXED: the keystroke after undoing a type-over-keyboard-selection
  // threw "TypeError: run.content is not iterable" (fuzzer, seeds 5 and 6).
  // checkboxStateElement now tolerates the editor's unbound placeholder
  // caret runs, so the keystroke lands instead of throwing — in both mounts
  // (the session mount exercises this since its undo came alive, G1).
  // Regression tests: packages/agent/test/local-undo.test.ts (both seeds).

  // G17 — undo of a format-over-selection then Enter strands structure
  // (fuzzer, seeds 6 and 7): a properties-only w:r shell in the local mount;
  // the session mount (undo now live, G1) strands the same shell.
  "local/undo.format-selection-strands-run": "undo of Cmd+B-over-selection then Enter strands a properties-only w:r shell",
  "session/undo.format-selection-strands-run": "undo of Cmd+B-over-selection then Enter strands a properties-only w:r shell",
};

/** Match a case id against the table (exact first, then prefix entries). */
export function gapFor(id: string): { key: string; reason: string } | undefined {
  if (KNOWN_GAPS[id]) return { key: id, reason: KNOWN_GAPS[id] };
  for (const [k, v] of Object.entries(KNOWN_GAPS)) {
    if (k.endsWith("*") && id.startsWith(k.slice(0, -1))) return { key: k, reason: v };
  }
  return undefined;
}
