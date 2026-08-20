/**
 * The editor's keyboard shortcut table — one declaration, two consumers.
 *
 * `DocxEditor.onKeyDown` matches an event against this table and dispatches
 * the command it names; the React HelpGuide renders the same table as its
 * shortcuts reference. The sheet therefore cannot drift from what the keys
 * actually do, which is the whole reason the table exists as data instead of
 * a chain of `if (e.metaKey && e.key === …)`.
 *
 * Convention notes:
 *  - `mod` is Cmd on Apple platforms and Ctrl everywhere else, matched as
 *    `metaKey || ctrlKey` — the same test the editor has always used, which
 *    also lets a test drive an Apple binding with ctrlKey under jsdom.
 *  - The bindings follow Word where Word has one (alignment, headings,
 *    lists, super/subscript, track changes), and Google Docs only where Word
 *    has nothing to copy (indent in/out, clear formatting).
 */

/** A command the editor itself carries out. */
export type EditorLocalCommand =
  | "bold"
  | "italic"
  | "underline"
  | "selectAll"
  | "undo"
  | "redo"
  | "pageBreak"
  | "columnBreak"
  | "heading1"
  | "heading2"
  | "heading3"
  | "heading4"
  | "heading5"
  | "heading6"
  | "bodyText"
  | "lineStart"
  | "lineEnd"
  | "paragraphUp"
  | "paragraphDown"
  | "documentStart"
  | "documentEnd"
  | "pageUp"
  | "pageDown";

/** A command the host (the React view) carries out through its public API,
 * so the shortcut runs exactly the code path the toolbar button runs. */
export type HostCommand =
  | "link"
  | "comment"
  | "nextComment"
  | "previousComment"
  | "trackChanges"
  | "alignLeft"
  | "alignCenter"
  | "alignRight"
  | "justify"
  | "bullet"
  | "number"
  | "indentIn"
  | "indentOut"
  | "lineSpacingSingle"
  | "lineSpacingOneAndHalf"
  | "lineSpacingDouble"
  | "strikethrough"
  | "superscript"
  | "subscript"
  | "clearFormatting"
  | "copyFormatting"
  | "pasteFormatting"
  | "growFont"
  | "shrinkFont"
  | "goToPage"
  /**
   * Not a shortcut of its own: Tab in a table's last cell adds a row, and the
   * row has to ride the wire, so the editor hands it to the host instead of
   * appending it privately.
   */
  | "tableRowBelow";

export type EditorCommand = EditorLocalCommand | HostCommand;

export interface KeyCombo {
  /** KeyboardEvent.key, letters lowercase. */
  key: string;
  /** Cmd on Apple, Ctrl elsewhere (matched as metaKey || ctrlKey). */
  mod?: true;
  shift?: true;
  alt?: true;
  /** Bind only on this platform family (used where Word itself differs). */
  platform?: "apple" | "other";
  /**
   * Accepted but not printed. Two kinds of alias exist: the shifted legend a
   * layout actually reports (⇧⌘8 arrives as "*"), and the Shift-extend twin
   * of a navigation binding, which the detail line describes instead.
   */
  alias?: true;
}

export type ShortcutGroup =
  | "History and selection"
  | "Text formatting"
  | "Paragraphs and lists"
  | "Insert and structure"
  | "Review"
  | "Navigation";

export interface EditorShortcut {
  command: EditorCommand;
  /** Help-sheet label. */
  label: string;
  group: ShortcutGroup;
  /** Accepted combos; the first is the one the help sheet leads with. */
  combos: KeyCombo[];
  /** Extra line under the label in the help sheet. */
  detail?: string;
}

const modKey = (key: string): KeyCombo => ({ key, mod: true });
const modShift = (key: string): KeyCombo => ({ key, mod: true, shift: true });
const modAlt = (key: string): KeyCombo => ({ key, mod: true, alt: true });
/** The same combo, accepted but not printed (see KeyCombo.alias). */
const also = (combo: KeyCombo): KeyCombo => ({ ...combo, alias: true });

const HEADING_LABELS = ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6"];

export const EDITOR_SHORTCUTS: EditorShortcut[] = [
  // ---------- History and selection ----------
  {
    command: "undo",
    label: "Undo",
    group: "History and selection",
    combos: [modKey("z")],
  },
  {
    command: "redo",
    label: "Redo",
    group: "History and selection",
    combos: [modShift("z"), { key: "y", mod: true, platform: "other" }],
  },
  {
    command: "selectAll",
    label: "Select all in the active story",
    group: "History and selection",
    combos: [modKey("a")],
  },

  // ---------- Text formatting ----------
  {
    command: "bold",
    label: "Bold",
    group: "Text formatting",
    combos: [modKey("b")],
    detail: "With no selection, the next text you type is bold.",
  },
  {
    command: "italic",
    label: "Italic",
    group: "Text formatting",
    combos: [modKey("i")],
    detail: "With no selection, the next text you type is italic.",
  },
  {
    command: "underline",
    label: "Underline",
    group: "Text formatting",
    combos: [modKey("u")],
    detail: "With no selection, the next text you type is underlined.",
  },
  {
    command: "strikethrough",
    label: "Strikethrough",
    group: "Text formatting",
    combos: [modShift("x")],
    detail: "Requires selected text.",
  },
  {
    command: "superscript",
    label: "Superscript",
    group: "Text formatting",
    combos: [modShift("="), also(modShift("+"))],
    detail: "Requires selected text. Press again to return to the baseline.",
  },
  {
    command: "subscript",
    label: "Subscript",
    group: "Text formatting",
    combos: [modKey("=")],
    detail: "Requires selected text. Press again to return to the baseline.",
  },
  {
    command: "growFont",
    label: "Grow font size",
    group: "Text formatting",
    combos: [modShift(">"), also(modShift("."))],
    detail: "Requires selected text of one size.",
  },
  {
    command: "shrinkFont",
    label: "Shrink font size",
    group: "Text formatting",
    combos: [modShift("<"), also(modShift(","))],
    detail: "Requires selected text of one size.",
  },
  {
    command: "clearFormatting",
    label: "Clear formatting",
    group: "Text formatting",
    combos: [modKey("\\")],
    detail: "Requires selected text.",
  },
  {
    command: "copyFormatting",
    label: "Copy formatting",
    group: "Text formatting",
    combos: [modAlt("c")],
    detail: "The format painter. ⇧⌘C is not used because ⇧⌘V is Paste and Match Style.",
  },
  {
    command: "pasteFormatting",
    label: "Paste formatting",
    group: "Text formatting",
    combos: [modAlt("v")],
    detail: "Applies the formatting copied with the shortcut above.",
  },

  // ---------- Paragraphs and lists ----------
  ...HEADING_LABELS.map((label, i): EditorShortcut => ({
    command: `heading${i + 1}` as EditorCommand,
    label,
    group: "Paragraphs and lists",
    combos: [modAlt(String(i + 1))],
  })),
  {
    command: "bodyText",
    label: "Normal paragraph",
    group: "Paragraphs and lists",
    combos: [modAlt("0")],
  },
  {
    command: "bullet",
    label: "Bulleted list",
    group: "Paragraphs and lists",
    combos: [modShift("l"), modShift("8"), also(modShift("*"))],
  },
  {
    command: "number",
    label: "Numbered list",
    group: "Paragraphs and lists",
    combos: [modShift("7"), also(modShift("&"))],
  },
  {
    command: "indentIn",
    label: "Increase indent",
    group: "Paragraphs and lists",
    combos: [modKey("]")],
  },
  {
    command: "indentOut",
    label: "Decrease indent",
    group: "Paragraphs and lists",
    combos: [modKey("[")],
  },
  {
    command: "alignLeft",
    label: "Align left",
    group: "Paragraphs and lists",
    combos: [modKey("l")],
  },
  {
    command: "alignCenter",
    label: "Center",
    group: "Paragraphs and lists",
    combos: [modKey("e")],
  },
  {
    command: "alignRight",
    label: "Align right",
    group: "Paragraphs and lists",
    combos: [modKey("r")],
  },
  {
    command: "justify",
    label: "Justify",
    group: "Paragraphs and lists",
    combos: [modKey("j")],
  },
  {
    command: "lineSpacingSingle",
    label: "Single line spacing",
    group: "Paragraphs and lists",
    combos: [modKey("1")],
  },
  {
    command: "lineSpacingOneAndHalf",
    label: "1.5 line spacing",
    group: "Paragraphs and lists",
    combos: [modKey("5")],
  },
  {
    command: "lineSpacingDouble",
    label: "Double line spacing",
    group: "Paragraphs and lists",
    combos: [modKey("2")],
  },

  // ---------- Insert and structure ----------
  {
    command: "link",
    label: "Insert or edit link",
    group: "Insert and structure",
    combos: [modKey("k")],
    detail: "Requires selected text.",
  },
  {
    command: "pageBreak",
    label: "Page break",
    group: "Insert and structure",
    combos: [modKey("Enter")],
  },
  {
    command: "columnBreak",
    label: "Column break",
    group: "Insert and structure",
    combos: [modShift("Enter")],
  },

  // ---------- Review ----------
  {
    command: "comment",
    label: "New comment",
    group: "Review",
    combos: [modAlt("m"), { key: "a", mod: true, alt: true, platform: "apple" }],
    detail: "Requires selected text.",
  },
  {
    command: "nextComment",
    label: "Next comment",
    group: "Review",
    combos: [modAlt("n")],
  },
  {
    command: "previousComment",
    label: "Previous comment",
    group: "Review",
    combos: [modAlt("p")],
  },
  {
    command: "trackChanges",
    label: "Track changes on or off",
    group: "Review",
    combos: [modShift("e")],
    detail: "Word's shortcut. ⇧⌘A is the desktop app's AI assistant.",
  },

  // ---------- Navigation ----------
  {
    command: "lineStart",
    label: "Move to the line start",
    group: "Navigation",
    combos: [modKey("ArrowLeft"), also(modShift("ArrowLeft"))],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "lineEnd",
    label: "Move to the line end",
    group: "Navigation",
    combos: [modKey("ArrowRight"), also(modShift("ArrowRight"))],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "paragraphUp",
    label: "Move to the previous paragraph",
    group: "Navigation",
    combos: [modKey("ArrowUp"), also(modShift("ArrowUp"))],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "paragraphDown",
    label: "Move to the next paragraph",
    group: "Navigation",
    combos: [modKey("ArrowDown"), also(modShift("ArrowDown"))],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "documentStart",
    label: "Move to the start of the story",
    group: "Navigation",
    combos: [modKey("Home"), also(modShift("Home"))],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "documentEnd",
    label: "Move to the end of the story",
    group: "Navigation",
    combos: [modKey("End"), also(modShift("End"))],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "pageUp",
    label: "Move to the previous page",
    group: "Navigation",
    combos: [{ key: "PageUp" }, also({ key: "PageUp", shift: true })],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "pageDown",
    label: "Move to the next page",
    group: "Navigation",
    combos: [{ key: "PageDown" }, also({ key: "PageDown", shift: true })],
    detail: "Add Shift to extend the selection.",
  },
  {
    command: "goToPage",
    label: "Go to page…",
    group: "Navigation",
    combos: [modAlt("g")],
  },
];

/**
 * Keys whose meaning depends on where the insertion point is, so they carry
 * no single command and are documented rather than dispatched from the table.
 * The help sheet lists them beside the shortcuts.
 */
export interface EditorKeyNote {
  label: string;
  keys: (apple: boolean) => string;
  group: ShortcutGroup;
  detail?: string;
}

export const EDITOR_KEY_NOTES: EditorKeyNote[] = [
  {
    label: "Cut, copy, paste",
    keys: (apple) => ["x", "c", "v"].map((k) => formatCombo(modKey(k), apple)).join(" / "),
    group: "History and selection",
    detail: "Handled by the browser or the desktop app.",
  },
  { label: "New paragraph", keys: () => "Enter", group: "Insert and structure" },
  { label: "Line break inside the paragraph", keys: () => "Shift+Enter", group: "Insert and structure" },
  {
    label: "Next or previous table cell",
    keys: () => "Tab / Shift+Tab",
    group: "Insert and structure",
    detail: "Tab in the last cell adds a row.",
  },
  {
    label: "Change list level",
    keys: () => "Tab / Shift+Tab",
    group: "Paragraphs and lists",
    detail: "Shift+Tab on a top-level item turns it back into a body paragraph.",
  },
  { label: "Insert a tab character", keys: () => "Tab", group: "Paragraphs and lists", detail: "In a body paragraph." },
  { label: "Move by character or visual line", keys: () => "Arrow keys", group: "Navigation" },
  { label: "Select by character or visual line", keys: () => "Shift+Arrow keys", group: "Navigation" },
  { label: "Move to the visual line edge", keys: () => "Home / End", group: "Navigation" },
  { label: "Select to the visual line edge", keys: () => "Shift+Home / Shift+End", group: "Navigation" },
  { label: "Nudge the selected object or ink", keys: () => "Arrow keys", group: "Navigation", detail: "One pixel; add Shift for ten." },
  { label: "Delete the selected object", keys: () => "Delete / Backspace", group: "Navigation" },
  {
    label: "Leave a drawing, text box, header/footer, or object selection",
    keys: () => "Escape",
    group: "Navigation",
  },
];

/**
 * Shortcuts owned by the embedding application rather than the editor — a
 * desktop app's menu accelerators, say. A menu accelerator is consumed by the
 * menu before the editor's keydown handler ever sees it, so the editor cannot
 * discover these; the host injects them (HelpGuide's `hostShortcuts` prop) and
 * the sheet becomes a reference for the whole keyboard instead of half of it.
 */
export interface HostShortcutSection {
  title: string;
  items: { label: string; keys: string; detail?: string }[];
}

/** True on macOS/iOS, where the modifier symbols and Cmd bindings apply. */
export function isApplePlatform(): boolean {
  return typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);
}

function comboApplies(combo: KeyCombo, apple: boolean): boolean {
  return combo.platform === undefined || (combo.platform === "apple") === apple;
}

/** Whether this event is exactly this combo (every modifier must agree). */
export function matchCombo(event: KeyboardEvent, combo: KeyCombo, apple: boolean): boolean {
  if (!comboApplies(combo, apple)) return false;
  const key = event.key.length === 1 ? event.key.toLowerCase() : event.key;
  return (
    combo.key === key &&
    !!combo.mod === (event.metaKey || event.ctrlKey) &&
    !!combo.shift === event.shiftKey &&
    !!combo.alt === event.altKey
  );
}

/** The command bound to this event, or null when nothing in the table matches. */
export function matchShortcut(event: KeyboardEvent, apple: boolean): EditorCommand | null {
  for (const shortcut of EDITOR_SHORTCUTS) {
    if (shortcut.combos.some((combo) => matchCombo(event, combo, apple))) return shortcut.command;
  }
  return null;
}

const KEY_SYMBOLS: Record<string, string> = {
  ArrowLeft: "←",
  ArrowRight: "→",
  ArrowUp: "↑",
  ArrowDown: "↓",
  Enter: "Enter",
  Home: "Home",
  End: "End",
  PageUp: "Page Up",
  PageDown: "Page Down",
};

/** One combo as the platform writes it: "⇧⌘X" on Apple, "Ctrl+Shift+X" else. */
export function formatCombo(combo: KeyCombo, apple: boolean): string {
  const key = KEY_SYMBOLS[combo.key] ?? (combo.key.length === 1 ? combo.key.toUpperCase() : combo.key);
  if (apple) {
    return `${combo.alt ? "⌥" : ""}${combo.shift ? "⇧" : ""}${combo.mod ? "⌘" : ""}${key}`;
  }
  const parts: string[] = [];
  if (combo.mod) parts.push("Ctrl");
  if (combo.alt) parts.push("Alt");
  if (combo.shift) parts.push("Shift");
  parts.push(key);
  return parts.join("+");
}

/** A shortcut's printable combos as one display string ("⇧⌘L or ⇧⌘8"). */
export function formatShortcutKeys(shortcut: EditorShortcut, apple: boolean): string {
  const shown: string[] = [];
  for (const combo of shortcut.combos) {
    if (combo.alias || !comboApplies(combo, apple)) continue;
    const text = formatCombo(combo, apple);
    if (!shown.includes(text)) shown.push(text);
  }
  return shown.join(" or ");
}
