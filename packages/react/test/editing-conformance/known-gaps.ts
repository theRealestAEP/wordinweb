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

  // G2 — Shift+Enter splits the paragraph exactly like Enter; Word inserts a
  // soft line break (w:br) and keeps one paragraph.
  "session/shiftenter.*": "Shift+Enter splits the paragraph instead of inserting a w:br soft line break",
  "local/shiftenter.*": "Shift+Enter splits the paragraph instead of inserting a w:br soft line break",

  // G3 — Tab in a plain body paragraph is a silent no-op; Word inserts a tab
  // character (or first-line indent at paragraph start).
  "session/tab.plain-paragraph": "Tab in a body paragraph is a no-op (Word inserts a tab character)",
  "local/tab.plain-paragraph": "Tab in a body paragraph is a no-op (Word inserts a tab character)",

  // G4 — Cmd+B/I/U at a collapsed caret set no pending format: the next typed
  // character comes out unformatted. Word toggles formatting for what is
  // typed next. (onKeyDown only forwards the shortcut when hasSelection().)
  "session/format.caret-*": "Cmd+B/I/U at a collapsed caret is a no-op (Word makes the next typed text formatted)",
  "local/format.caret-*": "Cmd+B/I/U at a collapsed caret is a no-op (Word makes the next typed text formatted)",

  // G5 — FIXED: rich (text/html and internal OOXML-flavor) paste now follows
  // Word's fragment semantics — a one-paragraph paste lands INLINE at the
  // caret (runs join the target paragraph, formatting preserved), a
  // multi-paragraph paste joins its first fragment at the caret and its last
  // fragment to the moved tail — instead of always splicing whole blocks and
  // stranding an empty paragraph (editor.ts pasteBlocks/insertInlineSpans).

  // G6 — FIXED: external HTML with inline markup (<b>x</b> plus <i>y</i>)
  // now folds consecutive inline nodes into ONE paragraph of formatted runs
  // (clipboard.ts appendHtmlBlocks) instead of one paragraph per element.

  // G7 (residual after the 734b6a8 spanning-merge fix) — select-all + delete
  // in a document WITH A TABLE leaves the table standing on the session
  // mount; Word removes it and leaves one empty paragraph.
  "session/selectall.delete-with-table": "select-all + delete leaves the table standing (Word removes it, leaving one empty paragraph)",

  // G18 (introduced/exposed by 734b6a8) — deleting an empty list item and
  // then a character strands a zero-length run where the unbullet/merge
  // rebind happened. Structural-lint class; visible only in the XML.
  "session/backspace.list-item-empty-then-char": "the unbullet/merge path strands a zero-length run",
  "local/backspace.list-item-empty-then-char": "the unbullet/merge path strands a zero-length run",

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

  // G13 — removing an empty paragraph (Backspace inside it, or Backspace at
  // the following paragraph's start) merges the empty paragraph's placeholder
  // run into the neighbor, stranding a zero-length w:t beside real content —
  // a structural-lint violation, not just a cosmetic one.
  "session/backspace.empty-paragraph": "deleting an empty paragraph strands its zero-length placeholder w:t inside the merged paragraph",
  "local/backspace.empty-paragraph": "deleting an empty paragraph strands its zero-length placeholder w:t inside the merged paragraph",

  // G14 — Enter with an active selection is a complete no-op; Word deletes
  // the selection and splits the paragraph at that point.
  "session/enter.over-selection": "Enter with an active selection is a no-op (Word deletes the selection, then splits)",
  "local/enter.over-selection": "Enter with an active selection is a no-op (Word deletes the selection, then splits)",

  // G15 — THE WEDGE (found by the fuzzer, seeds 1 and 2): Shift+ArrowRight at
  // the end of a paragraph whose NEXT paragraph is empty produces a zero-width
  // focus selection with no caret — after it, no caret or selection is
  // reported, arrow keys cannot collapse, and every typed character is
  // swallowed until a mouse click. Crossing into a NONEMPTY paragraph works
  // (select.shift-arrow-mid-text and the across-para cases pass).
  "session/select.shift-arrow-empty-para": "Shift+Arrow at a paragraph end before an empty paragraph wedges the editor (no caret/selection, typing dead until a mouse click)",
  "local/select.shift-arrow-empty-para": "Shift+Arrow at a paragraph end before an empty paragraph wedges the editor (no caret/selection, typing dead until a mouse click)",

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
