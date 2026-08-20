/**
 * #146: every popover and dialog the toolbar can open from a plain document,
 * named by the tab it lives on and the tooltip its control carries.
 *
 * The tooltip is the key because it is what the user reads and what the bar
 * already guarantees is unique per control. Nothing here names a CSS class or
 * a `data-` hook of the panel: the harness finds the panel by watching the
 * document, so adding a surface to this table is a one-line change and a
 * surface that is renamed fails loudly instead of silently going untested.
 */

export type SurfaceKind =
  /** A dropdown list of choices (`ToolbarMenuSelect`, including `ActionMenu`). */
  | "menu"
  /** A panel of controls the user fills in. */
  | "form"
  /** A gallery of pickable tiles or swatches. */
  | "gallery";

export interface Surface {
  /** Ribbon tab the control sits on. */
  tab: "home" | "insert" | "draw" | "layout" | "review";
  /** The control's tooltip — how the harness finds it. */
  tip: string;
  kind: SurfaceKind;
}

export const SURFACES: Surface[] = [
  // ---- Home -------------------------------------------------------------
  { tab: "home", tip: "Help and keyboard shortcuts (Ctrl+/)", kind: "form" },
  { tab: "home", tip: "Paragraph style", kind: "menu" },
  { tab: "home", tip: "Styles pane: create, change and remove this document's styles", kind: "form" },
  { tab: "home", tip: "Font", kind: "menu" },
  { tab: "home", tip: "Font size", kind: "menu" },
  { tab: "home", tip: "Text Effects and Typography", kind: "form" },
  { tab: "home", tip: "Change case", kind: "menu" },
  { tab: "home", tip: "Text color", kind: "gallery" },
  { tab: "home", tip: "Highlight color", kind: "gallery" },
  { tab: "home", tip: "Tab stops", kind: "form" },
  { tab: "home", tip: "Line & paragraph spacing", kind: "menu" },
  { tab: "home", tip: "Multilevel list gallery", kind: "menu" },
  { tab: "home", tip: "Paragraph borders and shading", kind: "menu" },

  // ---- Insert -----------------------------------------------------------
  { tab: "insert", tip: "Insert cover page", kind: "gallery" },
  { tab: "insert", tip: "Table", kind: "gallery" },
  { tab: "insert", tip: "Insert or edit SmartArt", kind: "form" },
  { tab: "insert", tip: "Insert or edit chart", kind: "form" },
  { tab: "insert", tip: "Insert online video", kind: "form" },
  { tab: "insert", tip: "Insert shape", kind: "form" },
  { tab: "insert", tip: "Insert or edit divider", kind: "form" },
  { tab: "insert", tip: "Insert text box", kind: "form" },
  { tab: "insert", tip: "Insert WordArt", kind: "form" },
  { tab: "insert", tip: "Insert link", kind: "form" },
  { tab: "insert", tip: "Add comment (select text first)", kind: "form" },
  { tab: "insert", tip: "Insert footnote (at the caret)", kind: "form" },
  { tab: "insert", tip: "Insert endnote (at the caret)", kind: "form" },
  { tab: "insert", tip: "Footnote and endnote options", kind: "form" },
  { tab: "insert", tip: "Insert bookmark", kind: "form" },
  { tab: "insert", tip: "Insert cross-reference", kind: "form" },
  { tab: "insert", tip: "Insert a caption for the selected object or table", kind: "form" },
  { tab: "insert", tip: "Edit the repeating header or footer", kind: "menu" },
  { tab: "insert", tip: "Watermark", kind: "form" },
  { tab: "insert", tip: "Insert a page number and choose its format", kind: "menu" },
  { tab: "insert", tip: "Insert a page, column or section break at the caret", kind: "menu" },
  { tab: "insert", tip: "Insert an automatically updating date or time", kind: "menu" },
  { tab: "insert", tip: "Insert a Word field", kind: "menu" },
  { tab: "insert", tip: "Insert or update a table of contents", kind: "menu" },
  { tab: "insert", tip: "Citations and bibliography", kind: "form" },
  { tab: "insert", tip: "Quick Parts: save and reuse content", kind: "form" },
  { tab: "insert", tip: "Insert equation", kind: "form" },
  { tab: "insert", tip: "Insert advanced symbol", kind: "gallery" },
  { tab: "insert", tip: "Drop cap", kind: "menu" },

  // ---- Draw -------------------------------------------------------------
  { tab: "draw", tip: "Pen color", kind: "gallery" },
  { tab: "draw", tip: "Pen width", kind: "menu" },

  // ---- Layout -----------------------------------------------------------
  { tab: "layout", tip: "Apply layout changes to", kind: "menu" },
  { tab: "layout", tip: "Margins", kind: "menu" },
  { tab: "layout", tip: "Orientation", kind: "menu" },
  { tab: "layout", tip: "Size", kind: "menu" },
  { tab: "layout", tip: "Columns & divider", kind: "menu" },
  { tab: "layout", tip: "Page border", kind: "menu" },
  { tab: "layout", tip: "Line numbers", kind: "menu" },
  { tab: "layout", tip: "Hyphenation", kind: "menu" },

  // ---- Review -----------------------------------------------------------
  { tab: "review", tip: "Accept tracked changes", kind: "menu" },
  { tab: "review", tip: "Reject tracked changes", kind: "menu" },
  { tab: "review", tip: "Find & replace", kind: "form" },
];
