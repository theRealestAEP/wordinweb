import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";

interface HelpGuideProps {
  open: boolean;
  onClose: () => void;
  returnFocus?: { current: HTMLElement | null };
}

interface GuideItem {
  title: string;
  description: string;
  steps?: string[];
}

interface GuideGroup {
  title: string;
  items: GuideItem[];
}

const GUIDE_GROUPS: GuideGroup[] = [
  {
    title: "Start editing",
    items: [
      {
        title: "Type and select text",
        description: "Click text or available white space to place the caret. Drag across text to select it, then use Home or the right-click menu to format it.",
      },
      {
        title: "Right-click text",
        description: "The text menu includes cut, copy, paste, formatting, links, comments, lists, alignment, and Select all. Selection-dependent actions become available after selecting text.",
      },
      {
        title: "Move through the document",
        description: "Use the arrow keys for character and line movement. Home and End move to visual line edges. Modified arrow keys jump to line edges or adjacent paragraphs.",
      },
      {
        title: "Save your work",
        description: "Use Download to save the edited DOCX. Hosts can also persist drafts from the editor's saved DOCX bytes.",
      },
    ],
  },
  {
    title: "Demo workspace",
    items: [
      {
        title: "Open or upload a document",
        description: "Choose File → Open .docx to edit a local Word file, or choose a starting document from the template menu.",
      },
      {
        title: "Zoom, review, and author",
        description: "Change Zoom without changing the DOCX. Switch between Editing and Suggesting, set the author stamped on comments and revisions, and show or hide comment balloons.",
      },
      {
        title: "Find and replace",
        description: "Press Cmd/Ctrl+F in the demo to open Find. Step through matches, replace the current match, or replace every match.",
      },
      {
        title: "Save, download, print",
        description: "Save keeps the demo workspace in this browser, Download writes a DOCX file, and Print opens the browser's print or PDF workflow.",
      },
    ],
  },
  {
    title: "Home tools",
    items: [
      {
        title: "Undo and redo",
        description: "Reverse or restore text, formatting, layout, and object edits with the arrow buttons or keyboard shortcuts.",
      },
      {
        title: "Styles, font, and size",
        description: "Apply document paragraph styles, choose an available typeface, and set the selected text size.",
      },
      {
        title: "Character formatting",
        description: "Bold, italicize, underline, strike, superscript, subscript, change case, clear formatting, or change text and highlight colors.",
      },
      {
        title: "Paragraph formatting",
        description: "Align text, adjust indentation, and set line or paragraph spacing. These controls apply to the selected paragraphs or the paragraph containing the caret.",
      },
      {
        title: "Lists",
        description: "Start or remove bulleted and numbered lists. Tab changes the level of the current list item; Shift+Tab promotes it.",
      },
    ],
  },
  {
    title: "Insert tools",
    items: [
      {
        title: "Pages, breaks, and sections",
        description: "Insert a cover page, full blank page, page break, column break, or next-page/continuous section break at the caret.",
      },
      {
        title: "Tables",
        description: "Choose a row and column count from the Table grid. Tab moves to the next cell, Shift+Tab moves back, and Tab in the last cell adds a row.",
      },
      {
        title: "Images, icons, and screenshots",
        description: "Insert an image or SVG file, or capture a screen, window, or browser tab. Select the result to resize, position, wrap, layer, replace, or add alternative text.",
      },
      {
        title: "3D models",
        description: "Insert a GLB file. Drag the model itself to rotate it, use the move grip to reposition it, resize from the handles, and use Reset 3D to restore its original view.",
      },
      {
        title: "Shapes and lines",
        description: "Insert a horizontal or vertical line, rectangle, rounded rectangle, ellipse, arrow, or callout. Select it to move, resize, rotate, change wrapping and layering, or edit its color, thickness, and solid, dashed, or dotted outline.",
      },
      {
        title: "Text boxes and WordArt",
        description: "Insert positioned text. Double-click a text box or text-bearing shape to edit its story, then use Home for font and text color. Escape returns to object selection for fill, outline, position, rotation, and wrapping.",
      },
      {
        title: "Charts and SmartArt",
        description: "Insert native editable charts or diagrams. Select an existing chart or SmartArt object before opening its Insert control to update its data or layout.",
      },
      {
        title: "Links, comments, and notes",
        description: "Select text before adding a hyperlink or comment. Footnotes insert at the caret and remain editable in the note area.",
      },
      {
        title: "Bookmarks and cross-references",
        description: "Create a named bookmark at a selection, then insert a live text or page reference to that bookmark elsewhere.",
      },
      {
        title: "Headers, footers, and page numbers",
        description: "Open the repeating header or footer layer, type or insert page fields and objects, then choose Close or press Escape to return to the body.",
      },
      {
        title: "Dates, fields, symbols, and drop caps",
        description: "Insert updating date/time and page fields, Unicode or custom symbols, and dropped or margin capitals at the caret.",
      },
      {
        title: "Equations",
        description: "Choose Insert → Equation and enter linear math. Use / for fractions, ^ for powers, √ for roots, [a&b;c&d] for matrices, and braces to group expressions.",
      },
      {
        title: "Media and embedded files",
        description: "Insert online-video metadata with a browser-safe poster, or embed another file in the DOCX package.",
      },
    ],
  },
  {
    title: "Draw tools",
    items: [
      {
        title: "Pen and highlighter",
        description: "Choose Pen or Highlighter, then drag on a page. Set color and width before drawing. Escape or Select returns to normal editing.",
      },
      {
        title: "Erase and select ink",
        description: "Stroke eraser removes a whole touched stroke. Lasso surrounds ink to select it as a group; drag or nudge the selected group to reposition it.",
      },
    ],
  },
  {
    title: "Layout tools",
    items: [
      {
        title: "Document or section scope",
        description: "Choose Whole document for a global change or This section when only the caret's current section should change.",
      },
      {
        title: "Margins, orientation, and page size",
        description: "Choose presets or exact custom page dimensions and margins. Orientation swaps the page between portrait and landscape.",
      },
      {
        title: "Columns",
        description: "Choose one, two, two with divider, or three columns. Insert a column break when text should begin in the next column immediately.",
      },
      {
        title: "Page borders and line numbers",
        description: "Apply a repeating page border with a preset or custom color and line weight, and enable continuous, per-page, per-section, or interval-based line numbering.",
      },
      {
        title: "Arrange selected objects",
        description: "With a floating object selected, align it to the page, rotate it by 90°, or change its stacking order with Bring to front and Send to back.",
      },
    ],
  },
  {
    title: "Object controls",
    items: [
      {
        title: "Move, resize, rotate, and delete",
        description: "Click an object to select it. Drag its body or move grip, drag a square handle to resize, choose Rotate for an exact angle, and press Delete or Backspace to remove it.",
      },
      {
        title: "Text wrapping and layering",
        description: "Inline moves with text. Wrap lets text flow around the object. Top+Bottom clears text from its sides. In front and Behind control its relationship to text.",
      },
      {
        title: "Exact object properties",
        description: "Set a line's color, weight, and solid, dashed, or dotted style directly in Insert → Shapes. After insertion, select it and choose Line style to revise those settings; Size and Position provide exact geometry.",
      },
    ],
  },
  {
    title: "Special-case recipes",
    items: [
      {
        title: "California pleading paper",
        description: "Create the fixed 1–28 pleading grid in the repeating header layer. Place the numbers in a text box as separate paragraphs, set paragraph spacing to 0 pt, and set the exact line height to 24 pt.",
        steps: [
          "Open Insert → Header, then insert a text box in the repeating header layer.",
          "Enter 1 through 28 as separate paragraphs and select all of those numbered paragraphs.",
          "Open Home → Line & paragraph spacing → Exactly 24 pt, then choose Remove space before and Remove space after.",
          "Open Insert → Shapes, set the line color, weight, and style, then insert a Vertical line for each rule and set its exact Size and Position in the repeating header layer.",
          "Choose Close to return to body text.",
        ],
      },
      {
        title: "Magazine columns with a middle rule",
        description: "Use the document's native column divider so the rule tracks layout automatically.",
        steps: [
          "Open Layout → Columns & divider → Two + divider line.",
          "Type normally to flow from the first column into the second.",
          "Use Insert → Break → Column break to move to the second column early.",
        ],
      },
      {
        title: "Resume horizontal rule",
        description: "Use a native paragraph divider so the rule stays attached to the résumé heading or contact line.",
        steps: [
          "Place the caret in the paragraph that should carry the rule and choose Insert → Divider.",
          "Choose a single, double, dashed, dotted, or thin + thick style, then set its color, width, and gap.",
          "Choose Apply divider. Return to the same paragraph and reopen Divider whenever you want to edit or remove it.",
        ],
      },
      {
        title: "Repeating page art",
        description: "Put lines, logos, and other repeating objects in the header or footer layer. A body object belongs to one page location; a header/footer object repeats with that section.",
      },
      {
        title: "Border around every page",
        description: "Choose Layout → Page border → Custom border, set the color and line weight, then apply it to the whole document or current section.",
      },
    ],
  },
];

interface ShortcutItem {
  action: string;
  keys: (apple: boolean) => string;
  detail?: string;
}

interface ShortcutGroup {
  title: string;
  items: ShortcutItem[];
}

const mod = (apple: boolean, key: string): string => apple ? `⌘${key}` : `Ctrl+${key}`;
const modShift = (apple: boolean, key: string): string => apple ? `⇧⌘${key}` : `Ctrl+Shift+${key}`;

const SHORTCUT_GROUPS: ShortcutGroup[] = [
  {
    title: "History and clipboard",
    items: [
      { action: "Undo", keys: (apple) => mod(apple, "Z") },
      { action: "Redo", keys: (apple) => apple ? "⇧⌘Z" : "Ctrl+Y or Ctrl+Shift+Z" },
      { action: "Cut", keys: (apple) => mod(apple, "X") },
      { action: "Copy", keys: (apple) => mod(apple, "C") },
      { action: "Paste", keys: (apple) => mod(apple, "V") },
      { action: "Select all in the active story", keys: (apple) => mod(apple, "A") },
      { action: "Find and replace in the demo", keys: (apple) => mod(apple, "F") },
    ],
  },
  {
    title: "Text and paragraphs",
    items: [
      { action: "Bold", keys: (apple) => mod(apple, "B"), detail: "Requires selected text." },
      { action: "Italic", keys: (apple) => mod(apple, "I"), detail: "Requires selected text." },
      { action: "Underline", keys: (apple) => mod(apple, "U"), detail: "Requires selected text." },
      { action: "Insert or edit link", keys: (apple) => mod(apple, "K"), detail: "Requires selected text." },
      { action: "New comment", keys: (apple) => apple ? "⌘⌥A" : "Ctrl+Alt+M", detail: "Requires selected text." },
      { action: "Bulleted list", keys: (apple) => modShift(apple, "L") },
      { action: "Numbered list", keys: (apple) => modShift(apple, "7") },
      { action: "Align left", keys: (apple) => mod(apple, "L") },
      { action: "Center", keys: (apple) => mod(apple, "E") },
      { action: "Align right", keys: (apple) => mod(apple, "R") },
      { action: "Justify", keys: (apple) => mod(apple, "J") },
      { action: "Heading 1–6", keys: (apple) => apple ? "⌘⌥1–6" : "Ctrl+Alt+1–6" },
      { action: "Normal paragraph", keys: (apple) => apple ? "⌘⌥0" : "Ctrl+Alt+0" },
    ],
  },
  {
    title: "Structure and navigation",
    items: [
      { action: "Page break", keys: (apple) => mod(apple, "Enter") },
      { action: "Column break", keys: (apple) => modShift(apple, "Enter") },
      { action: "Move by character or visual line", keys: () => "Arrow keys" },
      { action: "Select by character or visual line", keys: () => "Shift+Arrow keys" },
      { action: "Move to visual line edge", keys: () => "Home / End" },
      { action: "Select to visual line edge", keys: () => "Shift+Home / Shift+End" },
      { action: "Move to line edge or adjacent paragraph", keys: (apple) => apple ? "⌘+Arrow keys" : "Ctrl+Arrow keys" },
      { action: "Extend to line edge or adjacent paragraph", keys: (apple) => apple ? "⇧⌘+Arrow keys" : "Ctrl+Shift+Arrow keys" },
      { action: "Next or previous table cell", keys: () => "Tab / Shift+Tab" },
      { action: "Change list level", keys: () => "Tab / Shift+Tab" },
    ],
  },
  {
    title: "Objects and editing modes",
    items: [
      { action: "Nudge selected floating object or ink", keys: () => "Arrow keys", detail: "Moves one pixel." },
      { action: "Nudge selected object farther", keys: () => "Shift+Arrow keys", detail: "Moves ten pixels." },
      { action: "Delete selected object", keys: () => "Delete / Backspace" },
      { action: "Leave drawing, text-box, header/footer, or object selection", keys: () => "Escape" },
      { action: "Open this help guide", keys: (apple) => `${mod(apple, "/")} or F1` },
    ],
  },
];

const panel: React.CSSProperties = {
  background: "var(--dxw-popover-bg, #fff)",
  color: "var(--dxw-toolbar-fg, #3c4043)",
  border: "1px solid var(--dxw-toolbar-border, #dadce0)",
  borderRadius: 12,
};

export function HelpGuide({ open, onClose, returnFocus }: HelpGuideProps) {
  const [tab, setTab] = useState<"guides" | "shortcuts">("guides");
  const [query, setQuery] = useState("");
  const searchRef = useRef<HTMLInputElement | null>(null);
  const previousFocus = useRef<HTMLElement | null>(null);
  const apple = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);

  useEffect(() => {
    if (!open) return;
    previousFocus.current = document.activeElement as HTMLElement | null;
    const frame = requestAnimationFrame(() => searchRef.current?.focus());
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
      }
    };
    document.addEventListener("keydown", keydown);
    return () => {
      cancelAnimationFrame(frame);
      document.removeEventListener("keydown", keydown);
      requestAnimationFrame(() => (returnFocus?.current ?? previousFocus.current)?.focus({ preventScroll: true }));
    };
  }, [open, onClose, returnFocus]);

  const needle = query.trim().toLowerCase();
  const guideGroups = useMemo(() => GUIDE_GROUPS.map((group) => ({
    ...group,
    items: group.items.filter((item) =>
      !needle || `${group.title} ${item.title} ${item.description} ${(item.steps ?? []).join(" ")}`.toLowerCase().includes(needle),
    ),
  })).filter((group) => group.items.length > 0), [needle]);
  const shortcutGroups = useMemo(() => SHORTCUT_GROUPS.map((group) => ({
    ...group,
    items: group.items.filter((item) =>
      !needle || `${group.title} ${item.action} ${item.keys(apple)} ${item.detail ?? ""}`.toLowerCase().includes(needle),
    ),
  })).filter((group) => group.items.length > 0), [apple, needle]);

  if (!open) return null;
  return createPortal(
    <div
      data-dxw-help-backdrop=""
      onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 10000,
        display: "grid",
        placeItems: "center",
        padding: 18,
        background: "rgba(32,33,36,.48)",
        fontFamily: "system-ui, sans-serif",
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="dxw-help-title"
        data-dxw-help-dialog=""
        style={{
          ...panel,
          width: "min(980px, calc(100vw - 36px))",
          maxHeight: "min(820px, calc(100vh - 36px))",
          display: "flex",
          flexDirection: "column",
          overflow: "hidden",
          boxShadow: "0 18px 60px rgba(0,0,0,.28)",
        }}
      >
        <header style={{ padding: "18px 20px 14px", borderBottom: "1px solid var(--dxw-toolbar-border, #dadce0)" }}>
          <div style={{ display: "flex", alignItems: "flex-start", gap: 16 }}>
            <div style={{ flex: 1 }}>
              <h2 id="dxw-help-title" style={{ margin: 0, fontSize: 22, color: "var(--dxw-toolbar-fg, #202124)" }}>WordInWeb help</h2>
              <p style={{ margin: "4px 0 0", color: "var(--dxw-toolbar-muted, #5f6368)", fontSize: 13 }}>
                Tool guides, editing recipes, and keyboard shortcuts.
              </p>
            </div>
            <button
              type="button"
              aria-label="Close help"
              onClick={onClose}
              style={{ border: 0, background: "transparent", color: "inherit", cursor: "pointer", fontSize: 24, lineHeight: 1, padding: 4 }}
            >
              ×
            </button>
          </div>
          <input
            ref={searchRef}
            type="search"
            aria-label="Search help"
            placeholder="Search tools, lines, columns, shortcuts…"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            style={{
              width: "100%",
              boxSizing: "border-box",
              marginTop: 14,
              padding: "10px 12px",
              border: "1px solid var(--dxw-toolbar-border, #c7cbd1)",
              borderRadius: 8,
              background: "var(--dxw-popover-bg, #fff)",
              color: "inherit",
              font: "14px system-ui, sans-serif",
              outlineColor: "var(--dxw-accent, #1a73e8)",
            }}
          />
          <div role="tablist" aria-label="Help sections" style={{ display: "flex", gap: 6, marginTop: 12 }}>
            {(["guides", "shortcuts"] as const).map((name) => (
              <button
                key={name}
                type="button"
                role="tab"
                aria-selected={tab === name}
                onClick={() => setTab(name)}
                style={{
                  border: 0,
                  borderRadius: 7,
                  padding: "7px 12px",
                  background: tab === name ? "var(--dxw-tab-active-bg, #e8f0fe)" : "transparent",
                  color: tab === name ? "var(--dxw-accent, #1a73e8)" : "inherit",
                  font: "600 13px system-ui, sans-serif",
                  cursor: "pointer",
                  textTransform: "capitalize",
                }}
              >
                {name === "guides" ? "Guides" : "Shortcuts"}
              </button>
            ))}
          </div>
        </header>
        <div style={{ overflow: "auto", padding: "18px 20px 24px" }}>
          {tab === "guides" && guideGroups.map((group) => (
            <section key={group.title} style={{ marginBottom: 24 }}>
              <h3 style={{ margin: "0 0 10px", fontSize: 15 }}>{group.title}</h3>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 260px), 1fr))", gap: 10 }}>
                {group.items.map((item) => (
                  <article key={item.title} style={{ ...panel, padding: 14, background: "var(--dxw-help-card-bg, #f8fafd)" }}>
                    <h4 style={{ margin: "0 0 6px", fontSize: 14 }}>{item.title}</h4>
                    <p style={{ margin: 0, color: "var(--dxw-toolbar-muted, #5f6368)", fontSize: 12.5, lineHeight: 1.45 }}>{item.description}</p>
                    {item.steps && (
                      <ol style={{ margin: "9px 0 0", paddingLeft: 19, color: "var(--dxw-toolbar-muted, #5f6368)", fontSize: 12.5, lineHeight: 1.5 }}>
                        {item.steps.map((step) => <li key={step}>{step}</li>)}
                      </ol>
                    )}
                  </article>
                ))}
              </div>
            </section>
          ))}
          {tab === "shortcuts" && shortcutGroups.map((group) => (
            <section key={group.title} style={{ marginBottom: 24 }}>
              <h3 style={{ margin: "0 0 8px", fontSize: 15 }}>{group.title}</h3>
              <div style={{ ...panel, overflow: "hidden" }}>
                {group.items.map((item, index) => (
                  <div key={item.action} style={{ display: "grid", gridTemplateColumns: "minmax(180px, 1fr) minmax(150px, auto)", gap: 16, alignItems: "center", padding: "10px 12px", borderTop: index ? "1px solid var(--dxw-toolbar-border, #e6e8eb)" : 0 }}>
                    <div>
                      <div style={{ fontSize: 13 }}>{item.action}</div>
                      {item.detail && <div style={{ color: "var(--dxw-toolbar-muted, #5f6368)", fontSize: 11.5, marginTop: 2 }}>{item.detail}</div>}
                    </div>
                    <kbd style={{ justifySelf: "end", padding: "4px 7px", border: "1px solid var(--dxw-toolbar-border, #c7cbd1)", borderBottomWidth: 2, borderRadius: 5, background: "var(--dxw-help-key-bg, #fff)", font: "12px ui-monospace, SFMono-Regular, Menlo, monospace", whiteSpace: "nowrap" }}>
                      {item.keys(apple)}
                    </kbd>
                  </div>
                ))}
              </div>
            </section>
          ))}
          {(tab === "guides" ? guideGroups.length : shortcutGroups.length) === 0 && (
            <p role="status" style={{ color: "var(--dxw-toolbar-muted, #5f6368)", textAlign: "center", padding: 36 }}>
              No help topics match “{query}”.
            </p>
          )}
        </div>
      </div>
    </div>,
    document.body,
  );
}
