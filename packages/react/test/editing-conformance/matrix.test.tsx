// @vitest-environment jsdom
/**
 * INTERACTIVE-EDITING CONFORMANCE MATRIX (gesture x context).
 *
 * Each case pins a Word contract: place the caret/selection with the gesture
 * path, perform ONE gesture (real KeyboardEvent / ClipboardEvent dispatch),
 * then assert the contract plus three blanket invariants — structural lint,
 * save/reload round-trip, and no listener throw. Every case runs in BOTH
 * mounts ("session" = the likeoffice LocalDocumentSession mount, "local" =
 * solo DocxView) and, when the gesture changed the document, gets an
 * auto-generated undo/redo companion (Cmd+Z restores the pre-gesture bytes,
 * Cmd+Shift+Z reapplies).
 *
 * Divergences are NEVER encoded as the expectation: the test keeps the Word
 * contract and the divergence lives in KNOWN_GAPS (known-gaps.ts), which the
 * runner enforces in both directions — see that file.
 */
import { afterAll, describe, expect, it } from "vitest";
import { mountEditor, structuralLint, saveReloadProblems, type Editor, type Mode } from "./harness.js";
import { gapFor, KNOWN_GAPS } from "./known-gaps.js";
import {
  plainDoc,
  emptyParaDoc,
  listDoc,
  emptyListItemDoc,
  tableDoc,
  headingDoc,
  trackedDoc,
  kitchenSinkDoc,
} from "./fixtures.js";

interface CaseDef {
  id: string;
  contract: string;
  fixture: () => Uint8Array;
  run: (ed: Editor) => Promise<void>;
  expectAfter: (ed: Editor) => void | Promise<void>;
}

const count = (s: string, re: RegExp): number => (s.match(re) ?? []).length;
const paras = (ed: Editor): number => count(ed.xml(), /<w:p[ >/]/g);
const numPrs = (ed: Editor): number => count(ed.xml(), /<w:numPr>/g);
const rows = (ed: Editor): number => count(ed.xml(), /<w:tr[ >]/g);
const cells = (ed: Editor): number => count(ed.xml(), /<w:tc[ >]/g);

const intoEmptyPara = async (ed: Editor): Promise<void> => {
  await ed.caretAt("above empty", "end");
  await ed.press("ArrowRight");
};
const intoEmptyListItem = async (ed: Editor): Promise<void> => {
  await ed.caretAt("item only", "end");
  await ed.press("ArrowRight");
};
const intoEmptyCell = async (ed: Editor): Promise<void> => {
  await ed.caretAt("cell r1c0", "end");
  await ed.press("Tab");
};
const selectAcrossParas = async (ed: Editor): Promise<void> => {
  // From the start of "one" (end of paragraph 1), extend over "one", the
  // paragraph boundary, and "br" of paragraph 2.
  await ed.caretAt("one", "start");
  for (let i = 0; i < 6; i++) await ed.press("ArrowRight", { shift: true });
};

// ---------------------------------------------------------------------------
// The matrix.
// ---------------------------------------------------------------------------
const CASES: CaseDef[] = [
  // ----- type ---------------------------------------------------------------
  {
    id: "type.para-middle",
    contract: "A typed character inserts at the caret inside a word",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.type("X");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alphaX one\nbravo two\ncharlie three"),
  },
  {
    id: "type.doc-start",
    contract: "Typing at the very start of the document prepends to the first paragraph",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "start");
      await ed.type("X");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("Xalpha one\nbravo two\ncharlie three"),
  },
  {
    id: "type.para-end",
    contract: "Typing at the end of a paragraph appends to it",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha one", "end");
      await ed.type("X");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha oneX\nbravo two\ncharlie three"),
  },
  {
    id: "type.empty-paragraph",
    contract: "Typing in an empty paragraph fills it without touching neighbors",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await intoEmptyPara(ed);
      await ed.type("Q");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("above empty\nQ\nbelow empty"),
  },
  {
    id: "type.list-item",
    contract: "Typing in a list item keeps the item's numbering",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item middle", "end");
      await ed.type("X");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("item middleX");
      expect(numPrs(ed)).toBe(3);
    },
  },
  {
    id: "type.table-cell-end",
    contract: "Typing at the end of a cell stays inside that cell",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      await ed.type("X");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("cell r0c0X");
      expect(cells(ed)).toBe(4);
    },
  },
  {
    id: "type.table-cell-empty",
    contract: "Typing in an empty cell fills that cell only",
    fixture: tableDoc,
    run: async (ed) => {
      await intoEmptyCell(ed);
      await ed.type("E");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before table\ncell r0c0\ncell r0c1\ncell r1c0\nE\nafter table");
      expect(cells(ed)).toBe(4);
    },
  },
  {
    id: "type.heading",
    contract: "Typing in a heading keeps the heading style",
    fixture: headingDoc,
    run: async (ed) => {
      await ed.caretAt("Title Here", "end");
      await ed.type("X");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("Title HereX");
      expect(ed.xml()).toContain(`w:pStyle w:val="Heading1"`);
    },
  },
  {
    id: "type.tracked-present",
    contract: "With tracking off, typing beside existing revisions inserts plainly and keeps the revision markers intact",
    fixture: trackedDoc,
    run: async (ed) => {
      await ed.caretAt("mid", "end");
      await ed.type("W");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("midW");
      expect(ed.xml()).toContain("<w:ins ");
      expect(ed.xml()).toContain("<w:del ");
    },
  },
  {
    id: "type.over-selection",
    contract: "Typing over a selection replaces it",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.type("R");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nR two\ncharlie three"),
  },
  {
    id: "type.over-across-para-selection",
    contract: "Typing over an across-paragraph selection replaces it and merges the paragraphs",
    fixture: plainDoc,
    run: async (ed) => {
      await selectAcrossParas(ed);
      await ed.type("R");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha Ravo two\ncharlie three"),
  },

  // ----- Backspace ----------------------------------------------------------
  {
    id: "backspace.para-middle",
    contract: "Backspace deletes the character before the caret",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alph one\nbravo two\ncharlie three"),
  },
  {
    id: "backspace.para-start",
    contract: "Backspace at a paragraph start merges it into the previous paragraph",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("bravo", "start");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha onebravo two\ncharlie three"),
  },
  {
    id: "backspace.doc-start",
    contract: "Backspace at the very start of the document is a no-op",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "start");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nbravo two\ncharlie three"),
  },
  {
    id: "backspace.empty-paragraph",
    contract: "Backspace in an empty paragraph removes it",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await intoEmptyPara(ed);
      await ed.press("Backspace");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("above empty\nbelow empty"),
  },
  {
    id: "backspace.list-item-start",
    contract: "Backspace at the start of a nonempty list item removes the bullet first; the paragraph text survives in place",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item first", "start");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before list\nitem first\nitem middle\nitem last\nafter list");
      expect(numPrs(ed)).toBe(2); // the first item lost its bullet, nothing merged
    },
  },
  {
    id: "backspace.list-item-empty",
    contract: "Backspace on an empty list item removes the bullet first; the empty paragraph survives",
    fixture: emptyListItemDoc,
    run: async (ed) => {
      await intoEmptyListItem(ed);
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before list\nitem only\n\nafter list");
      expect(numPrs(ed)).toBe(1);
    },
  },
  {
    id: "backspace.list-item-empty-then-char",
    contract: "Deleting an empty list item and then a character strands no empty runs",
    fixture: emptyListItemDoc,
    run: async (ed) => {
      await intoEmptyListItem(ed);
      await ed.press("Backspace");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      // Word (bullet-first contract): second Backspace merges the emptied
      // paragraph; nothing else changes. Lint (run by the harness on every
      // case) is the load-bearing half of this case.
      expect(ed.text()).toContain("item only");
    },
  },
  {
    id: "backspace.table-cell-start",
    contract: "Backspace at the start of a table cell is a no-op (cells never merge from typing gestures)",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c1", "start");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before table\ncell r0c0\ncell r0c1\ncell r1c0\n\nafter table");
      expect(cells(ed)).toBe(4);
    },
  },
  {
    id: "backspace.heading-start",
    contract: "Backspace at a heading start merges it into the previous paragraph, which keeps its own (plain) style",
    fixture: headingDoc,
    run: async (ed) => {
      await ed.caretAt("Title Here", "start");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("intro textTitle Here\nbody text");
      expect(ed.xml()).not.toContain("Heading1");
    },
  },
  {
    id: "backspace.selection",
    contract: "Backspace deletes the selection",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo ");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\ntwo\ncharlie three"),
  },
  {
    id: "backspace.across-para-selection",
    contract: "Deleting an across-paragraph selection merges the surviving halves into one paragraph",
    fixture: plainDoc,
    run: async (ed) => {
      await selectAcrossParas(ed);
      await ed.press("Backspace");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha avo two\ncharlie three"),
  },
  {
    id: "backspace.across-cell-structure",
    contract: "A delete over a selection that crosses a cell boundary never breaks table structure",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      for (let i = 0; i < 3; i++) await ed.press("ArrowRight", { shift: true });
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(rows(ed)).toBe(2);
      expect(cells(ed)).toBe(4);
      // Word semantics (G12): the covered cells' CONTENT is cleared, so the
      // uncovered cell's text is the survival witness here. (The original
      // "cell r0c0 text survives" assertion encoded the char-wise divergence,
      // not the contract — authored while G12 was ledgered.)
      expect(ed.text()).toContain("cell r1c0");
    },
  },
  {
    id: "backspace.across-cell-word-semantics",
    contract: "A selection crossing a cell boundary selects whole cells; delete clears those cells' contents entirely",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      for (let i = 0; i < 3; i++) await ed.press("ArrowRight", { shift: true });
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before table\n\n\ncell r1c0\n\nafter table");
    },
  },
  {
    id: "backspace.tracked-into-ins",
    contract: "Backspace inside a tracked insertion (tracking off) edits its text and keeps the w:ins marker",
    fixture: trackedDoc,
    run: async (ed) => {
      await ed.caretAt("inserted text", "end");
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("inserted tex ");
      expect(ed.xml()).toContain("<w:ins ");
    },
  },

  // ----- Delete (forward) ---------------------------------------------------
  {
    id: "delete.para-middle",
    contract: "Delete removes the character after the caret",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.press("Delete");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alphaone\nbravo two\ncharlie three"),
  },
  {
    id: "delete.para-end",
    contract: "Delete at a paragraph end merges the next paragraph up",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha one", "end");
      await ed.press("Delete");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha onebravo two\ncharlie three"),
  },
  {
    id: "delete.doc-end",
    contract: "Delete at the very end of the document is a no-op",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("charlie three", "end");
      await ed.press("Delete");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nbravo two\ncharlie three"),
  },
  {
    id: "delete.empty-paragraph",
    contract: "Delete in an empty paragraph removes the paragraph mark, pulling the next paragraph up",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await intoEmptyPara(ed);
      await ed.press("Delete");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("above empty\nbelow empty"),
  },
  {
    id: "delete.table-cell-end",
    contract: "Delete at the end of a table cell is a no-op (never pulls the next cell in)",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      await ed.press("Delete");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before table\ncell r0c0\ncell r0c1\ncell r1c0\n\nafter table");
      expect(cells(ed)).toBe(4);
    },
  },
  {
    id: "delete.before-table",
    contract: "Delete at the end of the paragraph before a table never swallows the table",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("before table", "end");
      await ed.press("Delete");
    },
    expectAfter: (ed) => {
      expect(rows(ed)).toBe(2);
      expect(cells(ed)).toBe(4);
      expect(ed.text()).toContain("cell r0c0");
    },
  },
  {
    id: "delete.list-item-end-last",
    contract: "Delete at the end of the last list item merges the following paragraph into the item, which keeps its numbering",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item last", "end");
      await ed.press("Delete");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before list\nitem first\nitem middle\nitem lastafter list");
      expect(numPrs(ed)).toBe(3);
    },
  },

  // ----- Enter --------------------------------------------------------------
  {
    id: "enter.para-middle",
    contract: "Enter splits the paragraph at the caret",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.press("Enter");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha\n one\nbravo two\ncharlie three"),
  },
  {
    id: "enter.para-start",
    contract: "Enter at a paragraph start inserts an empty paragraph above",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("bravo", "start");
      await ed.press("Enter");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("alpha one\n\nbravo two\ncharlie three");
      expect(paras(ed)).toBe(4);
    },
  },
  {
    id: "enter.para-end",
    contract: "Enter at a paragraph end creates a new empty paragraph below and the caret lands in it",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha one", "end");
      await ed.press("Enter");
      await ed.type("n");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nn\nbravo two\ncharlie three"),
  },
  {
    id: "enter.empty-paragraph",
    contract: "Enter in an empty paragraph adds another empty paragraph",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await intoEmptyPara(ed);
      await ed.press("Enter");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("above empty\n\n\nbelow empty"),
  },
  {
    id: "enter.list-item-middle",
    contract: "Enter inside a list item splits it into two list items",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item mid", "end");
      await ed.press("Enter");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before list\nitem first\nitem mid\ndle\nitem last\nafter list");
      expect(numPrs(ed)).toBe(4);
    },
  },
  {
    id: "enter.list-item-end",
    contract: "Enter at the end of a list item creates a new empty list item",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item last", "end");
      await ed.press("Enter");
    },
    expectAfter: (ed) => expect(numPrs(ed)).toBe(4),
  },
  {
    id: "enter.list-item-empty-exits",
    contract: "Enter on an empty list item exits the list (removes numbering, keeps the empty paragraph)",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item last", "end");
      await ed.press("Enter"); // creates an empty item
      await ed.press("Enter"); // Word: exits the list
    },
    expectAfter: (ed) => expect(numPrs(ed)).toBe(3),
  },
  {
    id: "enter.table-cell",
    contract: "Enter inside a table cell adds a paragraph inside that cell (never a new row)",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      await ed.press("Enter");
      await ed.type("p2");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("cell r0c0\np2");
      expect(rows(ed)).toBe(2);
      expect(cells(ed)).toBe(4);
    },
  },
  {
    id: "enter.heading-end",
    contract: "Enter at the end of a heading starts a NORMAL paragraph (the heading's next-style), not another heading",
    fixture: headingDoc,
    run: async (ed) => {
      await ed.caretAt("Title Here", "end");
      await ed.press("Enter");
      await ed.type("x");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("intro text\nTitle Here\nx\nbody text");
      expect(count(ed.xml(), /Heading1/g)).toBe(1);
    },
  },
  {
    id: "enter.over-selection",
    contract: "Enter over a selection deletes the selection, then splits at that point",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.press("Enter");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("alpha one\n\n two\ncharlie three");
    },
  },

  // ----- Shift+Enter (soft line break) --------------------------------------
  {
    id: "shiftenter.para-middle",
    contract: "Shift+Enter inserts a soft line break (w:br) and keeps one paragraph",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.press("Enter", { shift: true });
    },
    expectAfter: (ed) => {
      expect(paras(ed)).toBe(3);
      expect(ed.xml()).toContain("<w:br");
    },
  },
  {
    id: "shiftenter.list-item",
    contract: "Shift+Enter in a list item breaks the line inside ONE item (no new bullet)",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item mid", "end");
      await ed.press("Enter", { shift: true });
    },
    expectAfter: (ed) => {
      expect(numPrs(ed)).toBe(3);
      expect(ed.xml()).toContain("<w:br");
    },
  },
  {
    id: "shiftenter.table-cell",
    contract: "Shift+Enter in a table cell breaks the line inside the cell paragraph",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      await ed.press("Enter", { shift: true });
    },
    expectAfter: (ed) => {
      expect(ed.xml()).toContain("<w:br");
      expect(count(ed.xml(), /<w:p[ >/]/g)).toBe(paras(ed)); // tautology guard replaced below
      expect(cells(ed)).toBe(4);
    },
  },

  // ----- Tab / Shift+Tab ----------------------------------------------------
  {
    id: "tab.plain-paragraph",
    contract: "Tab in a body paragraph inserts a tab character",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("bravo", "start");
      await ed.press("Tab");
    },
    expectAfter: (ed) => expect(ed.xml()).toContain("<w:tab"),
  },
  {
    id: "tab.list-item",
    contract: "Tab in a list item demotes it one level",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item middle", "end");
      await ed.press("Tab");
    },
    expectAfter: (ed) => expect(ed.xml()).toContain(`<w:ilvl w:val="1"/>`),
  },
  {
    id: "shifttab.list-item-indented",
    contract: "Shift+Tab promotes an indented list item back one level",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item middle", "end");
      await ed.press("Tab");
      await ed.press("Tab", { shift: true });
    },
    expectAfter: (ed) => {
      expect(ed.xml()).not.toContain(`<w:ilvl w:val="1"/>`);
      expect(numPrs(ed)).toBe(3);
    },
  },
  {
    id: "shifttab.list-item-level0",
    contract: "Shift+Tab on a level-0 list item promotes it OUT of the list into a body paragraph",
    fixture: listDoc,
    run: async (ed) => {
      await ed.caretAt("item middle", "end");
      await ed.press("Tab", { shift: true });
    },
    expectAfter: (ed) => expect(numPrs(ed)).toBe(2),
  },
  {
    id: "tab.table-next-cell",
    contract: "Tab in a table moves the caret to the next cell",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      await ed.press("Tab");
      await ed.type("T");
    },
    expectAfter: (ed) => expect(ed.text()).toContain("Tcell r0c1"),
  },
  {
    id: "shifttab.table-prev-cell",
    contract: "Shift+Tab in a table moves the caret to the previous cell",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c1", "end");
      await ed.press("Tab", { shift: true });
      await ed.type("S");
    },
    expectAfter: (ed) => expect(ed.text()).toContain("Scell r0c0"),
  },
  {
    id: "tab.table-last-cell-appends-row",
    contract: "Tab in the last cell appends a new table row",
    fixture: tableDoc,
    run: async (ed) => {
      await intoEmptyCell(ed); // r1c1, the last cell
      await ed.press("Tab");
    },
    expectAfter: (ed) => expect(rows(ed)).toBe(3),
  },
  {
    id: "shifttab.table-first-cell",
    contract: "Shift+Tab in the first cell of a table is a no-op",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "start");
      await ed.press("Tab", { shift: true });
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("before table\ncell r0c0\ncell r0c1\ncell r1c0\n\nafter table");
      expect(rows(ed)).toBe(2);
    },
  },

  // ----- Cmd+B / Cmd+I / Cmd+U ----------------------------------------------
  {
    id: "format.bold-selection",
    contract: "Cmd+B bolds the selection",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.press("b", { meta: true });
    },
    expectAfter: (ed) => expect(ed.xml()).toMatch(/<w:b\/>|<w:b [^>]*\/>/),
  },
  {
    id: "format.italic-selection",
    contract: "Cmd+I italicizes the selection",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.press("i", { meta: true });
    },
    expectAfter: (ed) => expect(ed.xml()).toMatch(/<w:i\/>|<w:i [^>]*\/>/),
  },
  {
    id: "format.underline-selection",
    contract: "Cmd+U underlines the selection",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.press("u", { meta: true });
    },
    expectAfter: (ed) => expect(ed.xml()).toMatch(/<w:u[ /]/),
  },
  {
    id: "format.bold-toggle-off",
    contract: "Cmd+B twice returns the selection to unbold",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.press("b", { meta: true });
      await ed.select("bravo");
      await ed.press("b", { meta: true });
    },
    expectAfter: async (ed) => {
      await ed.select("bravo");
      expect(ed.api.getSelectionFormat()?.bold ?? false).toBe(false);
    },
  },
  {
    id: "format.bold-across-para-selection",
    contract: "Cmd+B over an across-paragraph selection bolds both halves",
    fixture: plainDoc,
    run: async (ed) => {
      await selectAcrossParas(ed);
      await ed.press("b", { meta: true });
    },
    expectAfter: (ed) => expect(count(ed.xml(), /<w:b\/>/g)).toBe(2),
  },
  {
    id: "format.caret-bold-then-type",
    contract: "Cmd+B at a collapsed caret makes the NEXT typed text bold (Word's pending format)",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("charlie", "end");
      await ed.press("b", { meta: true });
      await ed.type("bb");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("charliebb");
      expect(ed.xml()).toMatch(/<w:b\/>|<w:b [^>]*\/>/);
    },
  },
  {
    id: "format.caret-italic-then-type",
    contract: "Cmd+I at a collapsed caret makes the NEXT typed text italic",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("charlie", "end");
      await ed.press("i", { meta: true });
      await ed.type("ii");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("charlieii");
      expect(ed.xml()).toMatch(/<w:i\/>|<w:i [^>]*\/>/);
    },
  },

  // ----- clipboard ----------------------------------------------------------
  {
    id: "copy.selection",
    contract: "Copy fills the clipboard (plain + rich) and leaves the document unchanged",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.copy();
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("alpha one\nbravo two\ncharlie three");
      expect(ed.clipboard.get("text/plain")).toBe("bravo");
      expect(ed.clipboard.get("text/html")).toContain("bravo");
    },
  },
  {
    id: "copy.across-para",
    contract: "Copying an across-paragraph selection yields plain text with a newline at the paragraph boundary",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "start");
      for (let i = 0; i < 14; i++) await ed.press("ArrowRight", { shift: true });
      await ed.copy();
    },
    expectAfter: (ed) => expect(ed.clipboard.get("text/plain")).toBe("alpha one\nbrav"),
  },
  {
    id: "copy.no-selection",
    contract: "Copy with a collapsed caret is a no-op",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.copy();
    },
    expectAfter: (ed) => {
      expect(ed.clipboard.size).toBe(0);
      expect(ed.text()).toBe("alpha one\nbravo two\ncharlie three");
    },
  },
  {
    id: "cut.selection",
    contract: "Cut removes the selection and fills the clipboard",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo ");
      await ed.cut();
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("alpha one\ntwo\ncharlie three");
      expect(ed.clipboard.get("text/plain")).toBe("bravo ");
    },
  },
  {
    id: "paste.plain-inline",
    contract: "Pasting plain text inserts it inline at the caret",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("bravo", "end");
      await ed.paste({ "text/plain": "PLAIN" });
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nbravoPLAIN two\ncharlie three"),
  },
  {
    id: "paste.plain-multiline",
    contract: "Pasting multi-line plain text splits into paragraphs, first line inline at the caret",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("bravo", "end");
      await ed.paste({ "text/plain": "L1\nL2" });
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nbravoL1\nL2 two\ncharlie three"),
  },
  {
    id: "paste.plain-empty-paragraph",
    contract: "Pasting plain text into an empty paragraph fills it",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await intoEmptyPara(ed);
      await ed.paste({ "text/plain": "Q" });
    },
    expectAfter: (ed) => expect(ed.text()).toBe("above empty\nQ\nbelow empty"),
  },
  {
    id: "paste.plain-into-selection",
    contract: "Pasting over a selection replaces it",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.paste({ "text/plain": "R" });
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nR two\ncharlie three"),
  },
  {
    id: "paste.plain-table-cell",
    contract: "Pasting plain text in a table cell stays inside the cell",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.caretAt("cell r0c0", "end");
      await ed.paste({ "text/plain": "P" });
    },
    expectAfter: (ed) => {
      expect(ed.text()).toContain("cell r0c0P");
      expect(cells(ed)).toBe(4);
    },
  },
  {
    id: "paste.rich-single-inline",
    contract: "Pasting a one-paragraph rich copy inserts it INLINE at the caret (no new paragraph, no stray empty paragraph)",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo");
      await ed.copy();
      await ed.caretAt("charlie", "start");
      await ed.paste();
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\nbravo two\nbravocharlie three"),
  },
  {
    id: "paste.rich-multi",
    contract: "Pasting a multi-paragraph rich copy merges its first chunk inline and its last chunk with the remainder",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "start");
      for (let i = 0; i < 14; i++) await ed.press("ArrowRight", { shift: true });
      await ed.copy();
      await ed.caretAt("charlie three", "end");
      await ed.paste();
    },
    expectAfter: (ed) =>
      expect(ed.text()).toBe("alpha one\nbravo two\ncharlie threealpha one\nbrav"),
  },
  {
    id: "paste.cutpaste-roundtrip",
    contract: "Cut text pasted at a new caret lands inline there",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.select("bravo ");
      await ed.cut();
      await ed.caretAt("charlie", "start");
      await ed.paste();
    },
    expectAfter: (ed) => expect(ed.text()).toBe("alpha one\ntwo\nbravo charlie three"),
  },
  {
    id: "paste.html-inline-runs",
    contract: "Pasting external HTML with inline formatting keeps one paragraph with formatted runs",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.caretAt("bravo", "end");
      await ed.paste({ "text/html": "<b>H1</b> plus <i>H2</i>", "text/plain": "H1 plus H2" });
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("alpha one\nbravoH1 plus H2 two\ncharlie three");
      expect(ed.xml()).toMatch(/<w:b\/>/);
    },
  },

  // ----- keyboard selection state -------------------------------------------
  {
    id: "select.shift-arrow-empty-para",
    contract:
      "Shift+ArrowRight at a paragraph end extends the selection into the following empty paragraph; the selection stays live and typing replaces it",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await ed.caretAt("above empty", "end");
      await ed.press("ArrowRight", { shift: true });
      await ed.press("k");
    },
    expectAfter: (ed) => expect(ed.text()).toContain("k"),
  },
  {
    id: "select.shift-arrow-mid-text",
    contract: "Shift+ArrowRight inside text selects one character; typing replaces it",
    fixture: emptyParaDoc,
    run: async (ed) => {
      await ed.caretAt("above", "end");
      await ed.press("ArrowRight", { shift: true });
      await ed.press("k");
    },
    expectAfter: (ed) => expect(ed.text()).toBe("abovekempty\n\nbelow empty"),
  },

  // ----- undo interaction sequences (fuzzer-found, local history) -----------
  {
    id: "undo.crash-after-type-over-selection",
    contract:
      "Undo after typing over a keyboard selection, then typing again, never throws (fuzz seeds 5+6; minimized: pasteHtml bold, Backspace, Home, Shift+ArrowLeft, type, Cmd+Z, type)",
    fixture: kitchenSinkDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.paste({ "text/html": "<b>bold</b>", "text/plain": "bold" });
      await ed.press("Backspace");
      await ed.press("Home");
      await ed.press("ArrowLeft", { shift: true });
      await ed.press("f");
      await ed.undo();
      await ed.press("f");
    },
    // The blanket checks (listener no-throw, lint, save/reload) carry this
    // case; the crash surfaces through ed.errors.
    expectAfter: () => {},
  },
  {
    id: "undo.format-selection-strands-run",
    contract:
      "Undo of a format-over-selection, then Enter, strands no properties-only run (fuzz seeds 6+7; minimized: Enter, type, Shift+ArrowLeft, Cmd+B, type, Cmd+Z, Enter)",
    fixture: kitchenSinkDoc,
    run: async (ed) => {
      await ed.caretAt("alpha", "end");
      await ed.press("Enter");
      await ed.press("j");
      await ed.press("ArrowLeft", { shift: true });
      await ed.press("b", { meta: true });
      await ed.press("e");
      await ed.undo();
      await ed.press("Enter");
    },
    // The blanket structural-lint check carries this case.
    expectAfter: () => {},
  },

  // ----- select-all ---------------------------------------------------------
  {
    id: "selectall.delete",
    contract: "Select-all + Backspace leaves exactly one empty paragraph",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.press("a", { meta: true });
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("");
      expect(paras(ed)).toBe(1);
    },
  },
  {
    id: "selectall.type",
    contract: "Select-all + typing replaces the whole document with the typed character in one paragraph",
    fixture: plainDoc,
    run: async (ed) => {
      await ed.press("a", { meta: true });
      await ed.type("Z");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("Z");
      expect(paras(ed)).toBe(1);
    },
  },
  {
    id: "selectall.delete-with-table",
    contract: "Select-all + Backspace in a document with a table removes the table too, leaving one empty paragraph",
    fixture: tableDoc,
    run: async (ed) => {
      await ed.press("a", { meta: true });
      await ed.press("Backspace");
    },
    expectAfter: (ed) => {
      expect(ed.text()).toBe("");
      expect(count(ed.xml(), /<w:tbl[ >]/g)).toBe(0);
    },
  },
];

// Guard against the accidental tautology left in shiftenter.table-cell.
const shiftEnterTable = CASES.find((c) => c.id === "shiftenter.table-cell")!;
shiftEnterTable.expectAfter = (ed) => {
  expect(ed.xml()).toContain("<w:br");
  expect(cells(ed)).toBe(4);
  // Unchanged paragraph count: before + 4 cell paras + after. (The guard
  // originally said 7 while its own comment said "unchanged" — authored
  // while the whole shiftenter family was ledgered, so it had never run
  // green; 6 is the fixture's actual count.)
  expect(paras(ed)).toBe(6);
};

// ---------------------------------------------------------------------------
// Runner + KNOWN_GAPS enforcement.
// ---------------------------------------------------------------------------
const gapHits = new Map<string, string[]>(); // gap key -> case ids that hit it

async function collectProblems(ed: Editor, c: CaseDef): Promise<string[]> {
  const problems: string[] = [];
  try {
    await c.run(ed);
    await c.expectAfter(ed);
  } catch (e) {
    problems.push(`contract: ${(e as Error).message.split("\n").slice(0, 4).join(" ")}`);
  }
  for (const p of structuralLint(ed.doc())) problems.push(`lint: ${p}`);
  for (const p of saveReloadProblems(ed)) problems.push(`save/reload: ${p}`);
  for (const err of ed.errors) problems.push(`listener threw: ${String(err)}`);
  return problems;
}

function settle(id: string, contract: string, problems: string[]): void {
  const gap = gapFor(id);
  if (problems.length > 0) {
    if (gap) {
      gapHits.set(gap.key, [...(gapHits.get(gap.key) ?? []), id]);
      return; // divergence is known, counted, and explicit — suite stays green
    }
    throw new Error(`UNLISTED DIVERGENCE ${id}\n  contract: ${contract}\n  ${problems.join("\n  ")}`);
  }
  if (gap) throw new Error(`KNOWN_GAP entry '${gap.key}' now PASSES for ${id} — remove it: ${gap.reason}`);
}

for (const mode of ["session", "local"] as Mode[]) {
  describe(`conformance matrix (${mode})`, () => {
    for (const c of CASES) {
      it(`${c.id} — ${c.contract}`, async () => {
        const ed = await mountEditor(c.fixture(), mode);
        try {
          settle(`${mode}/${c.id}`, c.contract, await collectProblems(ed, c));
        } finally {
          await ed.unmount();
        }
      }, 20000);
    }
  });

  describe(`undo/redo companions (${mode})`, () => {
    for (const c of CASES) {
      it(`undoredo.${c.id}`, async () => {
        const ed = await mountEditor(c.fixture(), mode);
        try {
          const base = ed.xml();
          try {
            await c.run(ed);
          } catch {
            return; // the base case reports this; nothing meaningful to undo
          }
          const after = ed.xml();
          if (after === base) return; // gesture was a no-op; nothing to undo
          const problems: string[] = [];
          // A case may perform several history steps (place + gesture + verify
          // typing), so undo UNTIL the pre-gesture document returns (bounded),
          // then redo the same number of times.
          let undos = 0;
          while (ed.xml() !== base && undos < 10) {
            await ed.undo();
            undos++;
          }
          if (ed.xml() !== base) {
            problems.push(`undo (x${undos}) did not restore the pre-gesture document`);
          } else {
            for (let i = 0; i < undos; i++) await ed.redo();
            if (ed.xml() !== after) problems.push(`redo (x${undos}) did not reapply the gesture`);
          }
          for (const err of ed.errors) problems.push(`listener threw: ${String(err)}`);
          settle(`${mode}/undoredo.${c.id}`, `undo restores / redo reapplies: ${c.contract}`, problems);
        } finally {
          await ed.unmount();
        }
      }, 20000);
    }
  });
}

describe("KNOWN_GAPS hygiene", () => {
  it("every KNOWN_GAPS entry was exercised by at least one failing case", () => {
    const unexercised = Object.keys(KNOWN_GAPS).filter((k) => !gapHits.has(k));
    expect(unexercised, `stale KNOWN_GAPS entries (no case hits them): ${unexercised.join(", ")}`).toEqual([]);
  });

  afterAll(() => {
    const total = [...gapHits.values()].reduce((n, ids) => n + ids.length, 0);
    console.error(
      `\n[editing-conformance] matrix: ${CASES.length} cases x 2 modes (+ undo/redo companions); ` +
        `${gapHits.size} known gaps covering ${total} case results:\n` +
        [...gapHits.entries()].map(([k, ids]) => `  ${k}  (${ids.length}) — ${KNOWN_GAPS[k]}`).join("\n"),
    );
  });
});
