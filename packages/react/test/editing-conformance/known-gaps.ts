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
  // one-unit wire insert). The separator-awareness residual is fixed too:
  // Backspace/Delete adjacent to a soft break or tab now delete the
  // SEPARATOR itself via applyDeleteSeparator (the insert's inverse: the
  // element leaves the run and the split w:t halves rejoin — insert+delete
  // round-trips byte-exactly), riding the deleteSeparator intent (a one-unit
  // wire delete); suggesting mode records a w:del around the separator's own
  // run slice (deleteSuggestedSeparator). Pinned by backspace.soft-break,
  // delete.soft-break, and backspace.tab-char.

  // G3 — FIXED (with a documented approximation): Tab in a plain body
  // paragraph inserts a w:tab character via the insertSeparator machinery.
  // Word nuance kept as a divergence, honestly: at the EXACT paragraph start
  // Word steps the first-line indent instead of inserting a tab character;
  // the engine inserts the tab character there too.

  // G4 — FIXED: Cmd+B/I/U at a collapsed caret queue a pending toggle
  // (editor.pendingFormat) that the next insertText consumes by formatting
  // the typed range through the host's shortcut path (canonical
  // formatRun/formatRange intents, suggest-aware). Any other gesture —
  // arrows, clicks, deletes — clears the queue, like Word.

  // G5 — FIXED: rich (text/html and internal OOXML-flavor) paste now follows
  // Word's fragment semantics — a one-paragraph paste lands INLINE at the
  // caret (runs join the target paragraph, formatting preserved), a
  // multi-paragraph paste joins its first fragment at the caret and its last
  // fragment to the moved tail — instead of always splicing whole blocks and
  // stranding an empty paragraph (editor.ts pasteBlocks/insertInlineSpans).

  // G6 — FIXED: external HTML with inline markup (<b>x</b> plus <i>y</i>)
  // now folds consecutive inline nodes into ONE paragraph of formatted runs
  // (clipboard.ts appendHtmlBlocks) instead of one paragraph per element.

  // G7 — FIXED (residual included): the spanning merge (734b6a8) closed the
  // session/local fork for plain paragraphs, and a fully selected table now
  // deletes AS A BLOCK on the session mount too — the canonical
  // tableOp("deleteTable") mutation + intent (the same apply every replica
  // runs), composed with the ordinary deleteText + mergeParagraph intents,
  // leaves one empty paragraph exactly like Word.

  // G18 — FIXED with G13: mergeParagraphBackward leaves the merged
  // paragraph's blank placeholder runs behind, so the unbullet/merge path
  // strands nothing.

  // G8 — deleting a selection that spans a paragraph boundary deletes the
  // characters but KEEPS the boundary; Word merges the two paragraphs.

  // G9 — Backspace at the start of a list item deletes/merges the paragraph
  // immediately; Word first removes the bullet (the paragraph survives as
  // plain text) and only a second Backspace merges. The empty-item variant is
  // the reported "undeletable/mishandled empty list item" family.

  // G10 — FIXED: applySplitParagraph applies the style's next-style (w:next;
  // a heading without one defaults to Normal, other styles continue
  // themselves) to the empty paragraph an end-of-paragraph Enter creates —
  // in the shared mutation, so every replica derives the same style.

  // G11 — FIXED: Shift+Tab on a level-0 list item promotes it out of the
  // list into a body paragraph — the same canonical unbullet Backspace at a
  // list-item start uses (setListType(null) + intent; tracked w:pPrChange in
  // suggesting mode). Deeper levels still step one level per press.

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

  // G12 — FIXED: a selection whose endpoints sit in different cells of one
  // table is cell-granular (Word's model): delete/typing-over expands to the
  // rectangular cell range and CLEARS the covered cells' contents — via the
  // ordinary per-range deletes (deleteText intents) plus in-cell paragraph
  // merges, so structure, ids, and replicas all survive identically.

  // G16 — FIXED: the keystroke after undoing a type-over-keyboard-selection
  // threw "TypeError: run.content is not iterable" (fuzzer, seeds 5 and 6).
  // checkboxStateElement now tolerates the editor's unbound placeholder
  // caret runs, so the keystroke lands instead of throwing — in both mounts
  // (the session mount exercises this since its undo came alive, G1).
  // Regression tests: packages/agent/test/local-undo.test.ts (both seeds).

  // G17 — FIXED (fuzzer, seeds 6 and 7): the strand was a STALE CARET PATH,
  // not a lint problem. applyRunFormat's split detached the caret element,
  // so the next checkpoint's caret ref fell back to a structural path from
  // the pre-split tree shape; undo resolved that path to a DIFFERENT element
  // (an rPr) and the following Enter split the paragraph around it —
  // stranding a properties-only w:r. Two-sided fix: selectRanges rebinds the
  // caret into the post-format tree (the record side), and history's
  // restoreCaret refuses to bind anything but a text leaf (the restore side
  // — a stale path can no longer corrupt a structural gesture).
};

/** Match a case id against the table (exact first, then prefix entries). */
export function gapFor(id: string): { key: string; reason: string } | undefined {
  if (KNOWN_GAPS[id]) return { key: id, reason: KNOWN_GAPS[id] };
  for (const [k, v] of Object.entries(KNOWN_GAPS)) {
    if (k.endsWith("*") && id.startsWith(k.slice(0, -1))) return { key: k, reason: v };
  }
  return undefined;
}
