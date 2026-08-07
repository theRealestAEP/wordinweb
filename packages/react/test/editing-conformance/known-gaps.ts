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
  // G1 — undo/redo are DEAD in the session mount: applyHistory routes to
  // onCollabUndo (editor.ts ~6354), and LocalDocumentSession wires no undoLast
  // hook, so Cmd+Z / Cmd+Shift+Z decline for EVERY gesture (this generalizes
  // the reported "paste undo dead" bug — being fixed by a sibling change).
  "session/undoredo.*": "undo/redo decline entirely in the session mount (no undoLast hook wired)",

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

  // G7 — SESSION ONLY: select-all + delete/type leaves every emptied
  // paragraph in place (three empty w:p from a three-paragraph body); Word —
  // and the LOCAL mount — collapse the whole selection to a single paragraph.
  // The collab-gated removal path lacks the multi-block merge the local path
  // has.
  "session/selectall.*": "session mount only: select-all + delete/type leaves one empty paragraph per former paragraph (local mount and Word leave exactly one)",

  // G8 — deleting a selection that spans a paragraph boundary deletes the
  // characters but KEEPS the boundary; Word merges the two paragraphs.
  "session/backspace.across-para-selection": "deleting an across-paragraph selection keeps the paragraph boundary (Word merges)",
  "local/backspace.across-para-selection": "deleting an across-paragraph selection keeps the paragraph boundary (Word merges)",
  "session/type.over-across-para-selection": "typing over an across-paragraph selection keeps the paragraph boundary (Word merges)",
  "local/type.over-across-para-selection": "typing over an across-paragraph selection keeps the paragraph boundary (Word merges)",

  // G9 — Backspace at the start of a list item deletes/merges the paragraph
  // immediately; Word first removes the bullet (the paragraph survives as
  // plain text) and only a second Backspace merges. The empty-item variant is
  // the reported "undeletable/mishandled empty list item" family.
  "session/backspace.list-*": "Backspace at list-item start merges/deletes at once (Word removes the bullet first, keeping the paragraph)",
  "local/backspace.list-*": "Backspace at list-item start merges/deletes at once (Word removes the bullet first, keeping the paragraph)",

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
};

/** Match a case id against the table (exact first, then prefix entries). */
export function gapFor(id: string): { key: string; reason: string } | undefined {
  if (KNOWN_GAPS[id]) return { key: id, reason: KNOWN_GAPS[id] };
  for (const [k, v] of Object.entries(KNOWN_GAPS)) {
    if (k.endsWith("*") && id.startsWith(k.slice(0, -1))) return { key: k, reason: v };
  }
  return undefined;
}
