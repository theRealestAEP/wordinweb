import { useCallback, useEffect, useRef, useState } from "react";
import type { DocxViewApi } from "./index.js";

/**
 * Chrome theme tokens. Every color the toolbar paints routes through a CSS
 * custom property, so an embedder can retheme the bar (e.g. a dark toolbar)
 * just by setting these variables on any ancestor element — no fork needed.
 * The fallback in each `var(...)` is the stock Google-Docs-style value, so the
 * default look is byte-for-byte unchanged. See the README "Theming" section
 * for the full variable list and a dark-toolbar example. Icons paint with
 * `currentColor`, so they follow `--dxw-toolbar-fg` automatically.
 */
const T = {
  bg: "var(--dxw-toolbar-bg, #f9fbfd)",
  fg: "var(--dxw-toolbar-fg, #3c4043)",
  border: "var(--dxw-toolbar-border, #dadce0)",
  muted: "var(--dxw-toolbar-muted, #5f6368)",
  accent: "var(--dxw-accent, #1a73e8)",
  accentFg: "var(--dxw-accent-fg, #fff)",
  activeBg: "var(--dxw-btn-active-bg, #dfe7f5)",
  hoverBg: "var(--dxw-btn-hover-bg, #f1f3f4)",
  tabActiveBg: "var(--dxw-tab-active-bg, #e8f0fe)",
  popoverBg: "var(--dxw-popover-bg, #fff)",
  popoverShadow: "var(--dxw-popover-shadow, 0 4px 16px rgba(0,0,0,.15))",
} as const;

/** Candidate families: always-usable ones (bundled substitutes or web-safe)
 * plus common document fonts. The dropdown filters to fonts the browser can
 * actually render (canvas width probe against the generic fallback). */
const FONT_CANDIDATES = [
  "Arial", "Arial Black", "Arial Narrow", "Avenir", "Avenir Next", "Baskerville", "Bookman Old Style",
  "Brush Script MT", "Calibri", "Cambria", "Century Gothic", "Chalkboard", "Charter", "Comic Sans MS",
  "Copperplate", "Courier New", "Didot", "Futura", "Garamond", "Georgia", "Gill Sans", "Helvetica",
  "Helvetica Neue", "Hoefler Text", "Impact", "Lucida Grande", "Menlo", "Monaco", "Optima",
  "Palatino", "Rockwell", "Seravek", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana",
];

let availableFonts: string[] | null = null;
function detectFonts(): string[] {
  if (availableFonts) return availableFonts;
  try {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d")!;
    const probe = "mmmwwwlliiWQ@";
    const widthIn = (family: string): number => {
      ctx.font = `16px ${family.includes(" ") ? `"${family}"` : family}, monospace`;
      return ctx.measureText(probe).width;
    };
    ctx.font = "16px monospace";
    const base = ctx.measureText(probe).width;
    // Bundled substitutes make Calibri/Cambria always renderable.
    const always = new Set(["Calibri", "Cambria"]);
    availableFonts = FONT_CANDIDATES.filter((f) => always.has(f) || Math.abs(widthIn(f) - base) > 0.5);
  } catch {
    availableFonts = FONT_CANDIDATES;
  }
  return availableFonts;
}


const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];

const HIGHLIGHTS: { name: string; css: string }[] = [
  { name: "yellow", css: "#ffff00" },
  { name: "green", css: "#00ff00" },
  { name: "cyan", css: "#00ffff" },
  { name: "magenta", css: "#ff00ff" },
  { name: "lightGray", css: "#d3d3d3" },
];

const btnStyle = (active: boolean): React.CSSProperties => ({
  minWidth: 26,
  height: 26,
  border: "1px solid transparent",
  background: active ? T.activeBg : "transparent",
  borderRadius: 4,
  cursor: "pointer",
  fontSize: 13,
  padding: "0 5px",
  color: T.fg,
});

const selectStyle: React.CSSProperties = {
  height: 26,
  border: "1px solid transparent",
  background: "transparent",
  borderRadius: 4,
  fontSize: 13,
  color: T.fg,
  cursor: "pointer",
};

function Btn({ label, title, active, onClick }: { label: React.ReactNode; title: string; active?: boolean; onClick: () => void }) {
  return (
    <button
      title={title}
      style={btnStyle(!!active)}
      onMouseDown={(e) => e.preventDefault()}
      onMouseEnter={(e) => ((e.target as HTMLElement).style.background = active ? T.activeBg : T.hoverBg)}
      onMouseLeave={(e) => ((e.target as HTMLElement).style.background = active ? T.activeBg : "transparent")}
      onClick={onClick}
    >
      {label}
    </button>
  );
}

function Sep() {
  return <span style={{ width: 1, height: 18, background: T.border, margin: "0 4px", flexShrink: 0 }} />;
}

const icon = { width: 16, height: 16, display: "block" } as const;

function ImageIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <circle cx="5.2" cy="6" r="1.2" fill="currentColor" stroke="none" />
      <path d="M2.5 12l3.5-4 2.8 3 2-2.4 2.7 3.4" />
    </svg>
  );
}

function HighlightIcon({ color }: { color: string }) {
  return (
    <svg style={icon} viewBox="0 0 16 16">
      <path d="M3 9.5L9.5 3l3.5 3.5L6.5 13H4.5L3 11.5v-2z" fill="none" stroke="currentColor" strokeWidth="1.3" />
      <rect x="2" y="13.5" width="12" height="2.5" rx="0.5" fill={color} />
    </svg>
  );
}

/** Menu select that runs an action and resets (never shows a value). */
function ActionMenu({
  label,
  title,
  groups,
  onPick,
  width,
}: {
  label: string;
  title: string;
  groups: { label?: string; items: [value: string, text: string][] }[];
  onPick: (value: string) => void;
  width?: number;
}) {
  return (
    <select
      title={title}
      value=""
      onMouseDown={(e) => e.stopPropagation()}
      onChange={(e) => {
        if (e.target.value) onPick(e.target.value);
      }}
      style={{ ...selectStyle, width }}
    >
      <option value="" disabled>
        {label}
      </option>
      {groups.map((g, i) =>
        g.label ? (
          <optgroup key={i} label={g.label}>
            {g.items.map(([v, t]) => (
              <option key={v} value={v}>{t}</option>
            ))}
          </optgroup>
        ) : (
          g.items.map(([v, t]) => (
            <option key={v} value={v}>{t}</option>
          ))
        ),
      )}
    </select>
  );
}

/** Highlight swatch popover (marker icon + colors + none). */
function HighlightMenu({ current, onPick }: { current?: string; onPick: (v: string | null) => void }) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button
        title="Highlight color"
        style={btnStyle(open)}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        <HighlightIcon color={current ?? "#ffff00"} />
      </button>
      {open && (
        <div
          onMouseDown={(e) => e.preventDefault()}
          style={{
            position: "absolute", top: 28, left: 0, zIndex: 100, background: T.popoverBg,
            border: `1px solid ${T.border}`, borderRadius: 6, boxShadow: T.popoverShadow,
            padding: 8, display: "flex", gap: 4, alignItems: "center",
          }}
        >
          {HIGHLIGHTS.map((h) => (
            <div
              key={h.name}
              title={h.name}
              onClick={() => { onPick(h.name); setOpen(false); }}
              style={{ width: 20, height: 20, background: h.css, border: `1px solid ${T.border}`, borderRadius: 3, cursor: "pointer" }}
            />
          ))}
          <div
            title="No highlight"
            onClick={() => { onPick(null); setOpen(false); }}
            style={{
              width: 20, height: 20, border: `1px solid ${T.border}`, borderRadius: 3, cursor: "pointer",
              background: "linear-gradient(to top left, #fff 46%, #d93025 49%, #d93025 51%, #fff 54%)",
            }}
          />
        </div>
      )}
    </span>
  );
}

function BulletListIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <circle cx="3" cy="4" r="1.1" fill="currentColor" stroke="none" />
      <circle cx="3" cy="8" r="1.1" fill="currentColor" stroke="none" />
      <circle cx="3" cy="12" r="1.1" fill="currentColor" stroke="none" />
      <path d="M6.5 4h8M6.5 8h8M6.5 12h8" strokeLinecap="round" />
    </svg>
  );
}

function NumberListIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <text x="1" y="5.6" fontSize="5.4" fill="currentColor" stroke="none" fontFamily="system-ui">1</text>
      <text x="1" y="9.9" fontSize="5.4" fill="currentColor" stroke="none" fontFamily="system-ui">2</text>
      <text x="1" y="14.2" fontSize="5.4" fill="currentColor" stroke="none" fontFamily="system-ui">3</text>
      <path d="M6.5 4h8M6.5 8h8M6.5 12h8" strokeLinecap="round" />
    </svg>
  );
}

function LinkIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <path d="M6.5 9.5l3-3" strokeLinecap="round" />
      <path d="M7.5 4.5l1.2-1.2a2.6 2.6 0 013.7 3.7L11.2 8.2" strokeLinecap="round" />
      <path d="M8.5 11.5l-1.2 1.2a2.6 2.6 0 01-3.7-3.7l1.2-1.2" strokeLinecap="round" />
    </svg>
  );
}

function ClearFormatIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <path d="M5 3h8M9 3l-2.5 9" strokeLinecap="round" />
      <path d="M3 13.5l3.5-3.5M3 10l3.5 3.5" strokeLinecap="round" />
    </svg>
  );
}

function IndentIcon({ dir }: { dir: 1 | -1 }) {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <path d="M2 3h12M8 6.5h6M8 9.5h6M2 13h12" strokeLinecap="round" />
      {dir === 1 ? <path d="M2.5 6l3 2-3 2z" fill="currentColor" stroke="none" /> : <path d="M5.5 6l-3 2 3 2z" fill="currentColor" stroke="none" />}
    </svg>
  );
}

/** Insert/edit/remove hyperlink on the current selection. */
function LinkMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [url, setUrl] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    if (!open) return;
    setUrl(api?.getLinkAt() ?? "");
    inputRef.current?.focus();
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open, api]);
  const submit = () => {
    const v = url.trim();
    if (v) api?.setLink(/^[a-z][a-z0-9+.-]*:/i.test(v) ? v : `https://${v}`);
    setOpen(false);
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert link" style={btnStyle(open)} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>
        <LinkIcon />
      </button>
      {open && (
        <div
          style={{
            position: "absolute", top: 28, left: 0, zIndex: 100, background: T.popoverBg,
            border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow,
            padding: 10, width: 260, display: "flex", gap: 6, alignItems: "center",
          }}
        >
          <input
            ref={inputRef}
            value={url}
            placeholder="Paste or type a link"
            onChange={(e) => setUrl(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && submit()}
            style={{ flex: 1, border: `1px solid ${T.border}`, borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          <button style={pillBtn} disabled={!url.trim()} onClick={submit}>Apply</button>
          {api?.getLinkAt() && (
            <button
              title="Remove link"
              style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}
              onClick={() => { api?.setLink(null); setOpen(false); }}
            >
              ✕
            </button>
          )}
        </div>
      )}
    </span>
  );
}

function FootnoteIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <path d="M2.5 13V3.5M2.5 3.5h7" strokeLinecap="round" />
      <text x="10.2" y="6.6" fontSize="6.4" fill="currentColor" stroke="none" fontFamily="system-ui">1</text>
      <path d="M2.5 13h11" strokeLinecap="round" strokeDasharray="1.5 1.6" />
    </svg>
  );
}

function CommentIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <path d="M1.5 3.5h13v8h-7l-3 3v-3h-3z" strokeLinejoin="round" />
      <path d="M8 5.5v4M6 7.5h4" strokeLinecap="round" />
    </svg>
  );
}

/** Insert-footnote popover: a text box; the note lands at the caret. */
function FootnoteMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const [hint, setHint] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const inputRef = useRef<HTMLTextAreaElement | null>(null);
  useEffect(() => {
    if (!open) return;
    inputRef.current?.focus();
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const submit = () => {
    if (api?.addFootnote(text)) {
      setText("");
      setHint("");
      setOpen(false);
    } else {
      setHint("Click into the text first so the reference has a place to go.");
    }
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert footnote (at the caret)" style={btnStyle(open)} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>
        <FootnoteIcon />
      </button>
      {open && (
        <div
          style={{
            position: "absolute", top: 28, right: 0, zIndex: 100, background: T.popoverBg,
            border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow,
            padding: 10, width: 240,
          }}
        >
          <textarea
            ref={inputRef}
            value={text}
            placeholder="Footnote text…"
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                submit();
              }
            }}
            style={{
              width: "100%", minHeight: 54, resize: "vertical", boxSizing: "border-box",
              border: `1px solid ${T.border}`, borderRadius: 6, padding: 6,
              font: "13px system-ui, sans-serif", outline: "none",
            }}
          />
          {hint && <div style={{ color: "#c5221f", fontSize: 12, marginTop: 4 }}>{hint}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, marginTop: 6 }}>
            <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} onClick={() => setOpen(false)}>Cancel</button>
            <button style={pillBtn} disabled={!text.trim()} onClick={submit}>Insert</button>
          </div>
        </div>
      )}
    </span>
  );
}

/** Google-Docs-style "add comment": popover with a text box, anchored to the
 * current selection (the editor keeps its owned selection while typing). */
function CommentMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const inputRef = useRef<HTMLTextAreaElement | null>(null);
  useEffect(() => {
    if (!open) return;
    inputRef.current?.focus();
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const submit = () => {
    if (api?.addComment(text)) {
      setText("");
      setOpen(false);
    }
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button
        title="Add comment (select text first)"
        style={btnStyle(open)}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        <CommentIcon />
      </button>
      {open && (
        <div
          style={{
            position: "absolute", top: 28, right: 0, zIndex: 100, background: T.popoverBg,
            border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow,
            padding: 10, width: 240,
          }}
        >
          <textarea
            ref={inputRef}
            value={text}
            placeholder="Comment on the selection…"
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                submit();
              }
            }}
            style={{
              width: "100%", minHeight: 54, resize: "vertical", boxSizing: "border-box",
              border: `1px solid ${T.border}`, borderRadius: 6, padding: 6,
              font: "13px system-ui, sans-serif", outline: "none",
            }}
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, marginTop: 6 }}>
            <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} onClick={() => setOpen(false)}>
              Cancel
            </button>
            <button style={pillBtn} disabled={!text.trim()} onClick={submit}>
              Comment
            </button>
          </div>
        </div>
      )}
    </span>
  );
}

const pillBtn: React.CSSProperties = {
  border: `1px solid ${T.border}`,
  borderRadius: 14,
  padding: "3px 12px",
  fontSize: 12.5,
  cursor: "pointer",
  background: T.accent,
  color: T.accentFg,
};

/** Google-Docs-style table menu: hover grid picker + row/column operations. */
function TableMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [hover, setHover] = useState<{ r: number; c: number }>({ r: 0, c: 0 });
  const rootRef = useRef<HTMLSpanElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);

  const ROWS = 8, COLS = 10;
  const ops: [string, string][] = [
    ["rowAbove", "Insert row above"],
    ["rowBelow", "Insert row below"],
    ["deleteRow", "Delete row"],
    ["colLeft", "Insert column left"],
    ["colRight", "Insert column right"],
    ["deleteCol", "Delete column"],
    ["mergeRight", "Merge cell right"],
    ["mergeDown", "Merge cell down"],
    ["splitCell", "Split cell"],
    ["valign:top", "Cell align top"],
    ["valign:center", "Cell align middle"],
    ["valign:bottom", "Cell align bottom"],
    ["deleteTable", "Delete table"],
  ];
  const CELL_FILLS = ["FFF2CC", "D9E2F3", "E2EFDA", "FCE4EC", "F1F3F4"];

  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button
        title="Table"
        style={btnStyle(open)}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        Table ▾
      </button>
      {open && (
        <div
          style={{
            position: "absolute",
            top: 28,
            left: 0,
            zIndex: 100,
            background: T.popoverBg,
            border: `1px solid ${T.border}`,
            borderRadius: 6,
            boxShadow: T.popoverShadow,
            padding: 8,
            width: COLS * 18 + 16,
          }}
          onMouseDown={(e) => e.preventDefault()}
        >
          <div style={{ fontSize: 12, color: T.muted, marginBottom: 4 }}>
            {hover.r > 0 ? `${hover.r} × ${hover.c}` : "Insert table"}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: `repeat(${COLS}, 16px)`, gap: 2 }}>
            {Array.from({ length: ROWS * COLS }, (_, i) => {
              const r = Math.floor(i / COLS) + 1;
              const c = (i % COLS) + 1;
              const lit = r <= hover.r && c <= hover.c;
              return (
                <div
                  key={i}
                  onMouseEnter={() => setHover({ r, c })}
                  onClick={() => {
                    api?.insertTable(r, c);
                    setOpen(false);
                    setHover({ r: 0, c: 0 });
                  }}
                  style={{
                    width: 16,
                    height: 16,
                    border: `1px solid ${lit ? T.accent : T.border}`,
                    background: lit ? T.activeBg : T.popoverBg,
                    borderRadius: 2,
                    cursor: "pointer",
                  }}
                />
              );
            })}
          </div>
          <div style={{ borderTop: "1px solid #eee", marginTop: 8, paddingTop: 4 }}>
            {ops.map(([op, label]) => (
              <div
                key={op}
                onClick={() => {
                  if (op.startsWith("valign:")) {
                    api?.tableOp({ kind: "cellVAlign", v: op.slice(7) as "top" | "center" | "bottom" });
                  } else {
                    api?.tableOp(op as Parameters<NonNullable<typeof api>["tableOp"]>[0]);
                  }
                  setOpen(false);
                }}
                style={{ padding: "4px 6px", fontSize: 13, cursor: "pointer", borderRadius: 4, color: T.fg }}
                onMouseEnter={(e) => ((e.target as HTMLElement).style.background = T.hoverBg)}
                onMouseLeave={(e) => ((e.target as HTMLElement).style.background = "transparent")}
              >
                {label}
              </div>
            ))}
            <div style={{ display: "flex", gap: 4, alignItems: "center", padding: "4px 6px" }}>
              <span style={{ fontSize: 12, color: T.muted, marginRight: 2 }}>Cell fill</span>
              {CELL_FILLS.map((f) => (
                <div
                  key={f}
                  title={`#${f}`}
                  onClick={() => { api?.tableOp({ kind: "cellShading", fill: f }); setOpen(false); }}
                  style={{ width: 16, height: 16, background: `#${f}`, border: `1px solid ${T.border}`, borderRadius: 3, cursor: "pointer" }}
                />
              ))}
              <div
                title="No fill"
                onClick={() => { api?.tableOp({ kind: "cellShading", fill: null }); setOpen(false); }}
                style={{
                  width: 16, height: 16, border: `1px solid ${T.border}`, borderRadius: 3, cursor: "pointer",
                  background: "linear-gradient(to top left, #fff 46%, #d93025 49%, #d93025 51%, #fff 54%)",
                }}
              />
            </div>
          </div>
        </div>
      )}
    </span>
  );
}

/**
 * Default formatting toolbar for an editable DocxView. Compact, grouped like
 * a word processor; every control preserves the selection/caret.
 */
/** Word's Layout ribbon, scoped to the whole document or the caret's
 * section (per-page layout = section breaks + section scope). */
function LayoutTab({ api }: { api: DocxViewApi | null }) {
  const [scope, setScope] = useState<"document" | "section">("document");
  const set = (patch: Parameters<NonNullable<typeof api>["setPageLayout"]>[0]) => api?.setPageLayout(patch, scope);
  const setLn = (patch: Parameters<NonNullable<typeof api>["setLineNumbering"]>[0]) => api?.setLineNumbering(patch, scope);
  const sel = (title: string, entries: [string, string][], onPick: (v: string) => void, width = 96) => (
    <select
      title={title}
      value=""
      onMouseDown={(e) => e.stopPropagation()}
      onChange={(e) => { if (e.target.value) onPick(e.target.value); e.target.value = ""; }}
      style={{ ...selectStyle, width }}
    >
      <option value="" disabled>{title}</option>
      {entries.map(([v, l]) => (
        <option key={v} value={v}>{l}</option>
      ))}
    </select>
  );
  return (
    <>
      <select
        title="Apply layout changes to"
        value={scope}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => setScope(e.target.value as "document" | "section")}
        style={{ ...selectStyle, width: 118 }}
      >
        <option value="document">Whole document</option>
        <option value="section">This section</option>
      </select>
      <Sep />
      {sel("Margins", [["m:normal", 'Normal (1")'], ["m:narrow", 'Narrow (0.5")'], ["m:wide", 'Wide (1.5")']], (v) => {
        if (v === "m:normal") set({ margins: { top: 1, right: 1, bottom: 1, left: 1 } });
        else if (v === "m:narrow") set({ margins: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 } });
        else set({ margins: { top: 1, right: 1.5, bottom: 1, left: 1.5 } });
      }, 88)}
      {sel("Orientation", [["portrait", "Portrait"], ["landscape", "Landscape"]], (v) => set({ orientation: v as "portrait" | "landscape" }), 96)}
      {sel("Size", [["letter", "Letter"], ["legal", "Legal"], ["a4", "A4"]], (v) => {
        if (v === "letter") set({ size: { width: 8.5, height: 11 } });
        else if (v === "legal") set({ size: { width: 8.5, height: 14 } });
        else set({ size: { width: 8.27, height: 11.69 } });
      }, 64)}
      {sel("Columns", [["1", "One column"], ["2", "Two columns"], ["3", "Three columns"]], (v) => set({ columns: parseInt(v, 10) }), 84)}
      {sel("Page border", [["none", "No border"], ["thin", "Thin box (\u00bdpt)"], ["thick", "Thick box (1\u00bdpt)"], ["accent", "Blue box"]], (v) => {
        if (v === "none") set({ pageBorders: null });
        else if (v === "thin") set({ pageBorders: { sz: 4 } });
        else if (v === "thick") set({ pageBorders: { sz: 12 } });
        else set({ pageBorders: { sz: 8, color: "4472C4" } });
      }, 96)}
      {sel("Line numbers", [
        ["off", "None"],
        ["continuous", "Continuous"],
        ["eachPage", "Restart each page"],
        ["eachSection", "Restart each section"],
        ["by5", "Count by 5"],
        ["by10", "Count by 10"],
      ], (v) => {
        if (v === "off") setLn({ enabled: false });
        else if (v === "continuous") setLn({ enabled: true, countBy: 1, restart: "continuous" });
        else if (v === "eachPage") setLn({ enabled: true, countBy: 1, restart: "newPage" });
        else if (v === "eachSection") setLn({ enabled: true, countBy: 1, restart: "newSection" });
        else if (v === "by5") setLn({ enabled: true, countBy: 5 });
        else setLn({ enabled: true, countBy: 10 });
      }, 118)}
    </>
  );
}

/** Toolbar control groups a host can disable via the `features` prop. */
export type ToolbarFeature =
  | "history"
  | "styles"
  | "font"
  | "size"
  | "format"
  | "color"
  | "highlight"
  | "alignment"
  | "indent"
  | "spacing"
  | "link"
  | "lists"
  | "table"
  | "image"
  | "comment"
  | "footnote"
  | "layout"
  | "download";

export function DocxToolbar({
  api,
  onSave,
  features,
  className,
  style,
}: {
  api: DocxViewApi | null;
  onSave?: (bytes: Uint8Array) => void;
  /** Per-group overrides; every group defaults to enabled. */
  features?: Partial<Record<ToolbarFeature, boolean>>;
  /** Extra class on the toolbar root (e.g. a scope for CSS-variable overrides). */
  className?: string;
  /** Inline overrides merged onto the toolbar root; wins over the defaults. */
  style?: React.CSSProperties;
}) {
  const on = (k: ToolbarFeature) => features?.[k] !== false;
  // Ribbon-style tabs: complex tool groups get their own surface instead of
  // one overloaded row (Layout especially).
  const [tab, setTab] = useState<"home" | "insert" | "layout">("home");
  // Subtle delayed tooltips: controls declare `title`; on first hover the
  // title moves to data-tip (suppressing the OS tooltip) and a quiet custom
  // one fades in under the control after a beat.
  const [tip, setTip] = useState<{ text: string; x: number; y: number } | null>(null);
  const tipTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const onTipOver = useCallback((e: React.MouseEvent) => {
    const el = (e.target as HTMLElement).closest("[title], [data-tip]") as HTMLElement | null;
    if (!el) return;
    const title = el.getAttribute("title");
    if (title) {
      el.setAttribute("data-tip", title);
      el.removeAttribute("title");
    }
    const text = el.getAttribute("data-tip");
    if (!text) return;
    if (tipTimer.current) clearTimeout(tipTimer.current);
    tipTimer.current = setTimeout(() => {
      const r = el.getBoundingClientRect();
      setTip({ text, x: r.left + r.width / 2, y: r.bottom + 6 });
    }, 550);
  }, []);
  const onTipOut = useCallback(() => {
    if (tipTimer.current) clearTimeout(tipTimer.current);
    setTip(null);
  }, []);
  const [fmt, setFmt] = useState<ReturnType<NonNullable<DocxViewApi["getSelectionFormat"]>> | null>(null);
  const [curStyle, setCurStyle] = useState<string | null>(null);
  const [listKind, setListKind] = useState<"bullet" | "number" | null>(null);
  // Native <select>/<input type=color> steal focus and collapse the document
  // selection; remember the last real range and restore it before applying.
  const savedRange = useRef<Range | null>(null);
  const imageInput = useRef<HTMLInputElement | null>(null);

  const refresh = useCallback(() => {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && sel.rangeCount > 0) {
      savedRange.current = sel.getRangeAt(0).cloneRange();
    }
    setFmt(api?.getSelectionFormat() ?? null);
    setCurStyle(api?.getParagraphStyleId?.() ?? null);
    setListKind(api?.getListType?.() ?? null);
  }, [api]);

  useEffect(() => {
    document.addEventListener("selectionchange", refresh);
    document.addEventListener("dxw-selection", refresh);
    return () => {
      document.removeEventListener("selectionchange", refresh);
      document.removeEventListener("dxw-selection", refresh);
    };
  }, [refresh]);

  const restoreSelection = () => {
    const sel = window.getSelection();
    if (sel && sel.isCollapsed && savedRange.current) {
      try {
        sel.removeAllRanges();
        sel.addRange(savedRange.current);
      } catch {
        /* range may be stale after re-render */
      }
    }
  };

  const apply = (patch: Parameters<DocxViewApi["applyFormat"]>[0]) => {
    restoreSelection();
    api?.applyFormat(patch);
    setFmt(api?.getSelectionFormat() ?? null);
  };

  return (
    <div
      className={className}
      onMouseOver={onTipOver}
      onMouseOut={onTipOut}
      onMouseDownCapture={onTipOut}
      style={{
        display: "flex",
        gap: 2,
        alignItems: "center",
        padding: "4px 10px",
        borderBottom: `1px solid ${T.border}`,
        background: T.bg,
        flexWrap: "wrap",
        fontFamily: "system-ui, sans-serif",
        ...style,
      }}
    >
      {tip && (
        <div
          style={{
            position: "fixed",
            left: tip.x,
            top: tip.y,
            transform: "translateX(-50%)",
            background: "rgba(32,33,36,.92)",
            color: T.accentFg,
            font: "11.5px system-ui, sans-serif",
            padding: "4px 8px",
            borderRadius: 4,
            pointerEvents: "none",
            whiteSpace: "nowrap",
            zIndex: 1000,
            boxShadow: "0 2px 6px rgba(0,0,0,.2)",
          }}
        >
          {tip.text}
        </div>
      )}
      <div style={{ display: "flex", gap: 2, marginRight: 8 }}>
        {(["home", "insert", "layout"] as const).map((t) => (
          <button
            key={t}
            data-tab={t}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => setTab(t)}
            style={{
              border: "none",
              background: tab === t ? T.tabActiveBg : "transparent",
              color: tab === t ? T.accent : T.fg,
              font: "600 12.5px system-ui, sans-serif",
              padding: "5px 10px",
              borderRadius: 6,
              cursor: "pointer",
              textTransform: "capitalize",
            }}
          >
            {t}
          </button>
        ))}
      </div>
      <Sep />
      {tab === "home" && on("history") && (
        <>
          <Btn label={"↶"} title="Undo (⌘Z)" onClick={() => { api?.undo(); refresh(); }} />
          <Btn label={"↷"} title="Redo (⇧⌘Z)" onClick={() => { api?.redo(); refresh(); }} />
          <Sep />
        </>
      )}
      {tab === "home" && on("styles") && (
      <select
        title="Paragraph style"
        value={curStyle ?? "__normal"}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => {
          if (e.target.value) {
            api?.setParagraphStyle(e.target.value === "__normal" ? null : e.target.value);
            setCurStyle(api?.getParagraphStyleId?.() ?? null);
          }
        }}
        style={{ ...selectStyle, width: 92 }}
      >
        <option value="__normal">Normal</option>
        {(api?.listParagraphStyles() ?? [])
          .filter((s) => !/^normal$/i.test(s.name))
          .map((s) => (
            <option key={s.id} value={s.id}>{s.name}</option>
          ))}
        {curStyle !== null &&
          !(api?.listParagraphStyles() ?? []).some((s) => s.id === curStyle) && (
            <option value={curStyle}>
              {api?.document.styles.byId.get(curStyle)?.name ?? curStyle}
            </option>
          )}
      </select>
      )}
      {tab === "home" && on("font") && (
      <select
        title="Font"
        value={fmt?.fontFamily ?? ""}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => e.target.value && apply({ fontFamily: e.target.value })}
        style={{ ...selectStyle, width: 130 }}
      >
        <option value="" disabled>Font</option>
        {(fmt?.fontFamily && !detectFonts().includes(fmt.fontFamily) ? [fmt.fontFamily, ...detectFonts()] : detectFonts()).map((f) => (
          <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>
        ))}
      </select>
      )}
      {tab === "home" && on("size") && (
      <select
        title="Font size"
        value={fmt?.fontSizePt ?? ""}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => e.target.value && apply({ fontSizePt: parseFloat(e.target.value) })}
        style={{ ...selectStyle, width: 58 }}
      >
        <option value="" disabled>Size</option>
        {SIZES.map((s) => (
          <option key={s} value={s}>{s}</option>
        ))}
      </select>
      )}
      <Sep />
      {tab === "home" && on("format") && (
      <>
      <Btn label={<b>B</b>} title="Bold (⌘B)" active={!!fmt?.bold} onClick={() => apply({ bold: !fmt?.bold })} />
      <Btn label={<i>I</i>} title="Italic" active={!!fmt?.italic} onClick={() => apply({ italic: !fmt?.italic })} />
      <Btn label={<u>U</u>} title="Underline" active={!!fmt?.underline} onClick={() => apply({ underline: !fmt?.underline })} />
      <Btn label={<s>S</s>} title="Strikethrough" active={!!fmt?.strike} onClick={() => apply({ strike: !fmt?.strike })} />
      <Btn
        label={<span style={{ fontSize: 12 }}>x<sup style={{ fontSize: 9 }}>2</sup></span>}
        title="Superscript"
        active={fmt?.verticalAlign === "superscript"}
        onClick={() => apply({ verticalAlign: fmt?.verticalAlign === "superscript" ? null : "superscript" })}
      />
      <Btn
        label={<span style={{ fontSize: 12 }}>x<sub style={{ fontSize: 9 }}>2</sub></span>}
        title="Subscript"
        active={fmt?.verticalAlign === "subscript"}
        onClick={() => apply({ verticalAlign: fmt?.verticalAlign === "subscript" ? null : "subscript" })}
      />
      <Btn label={<ClearFormatIcon />} title="Clear formatting" onClick={() => apply({ clear: true })} />
      <ActionMenu
        label="Aa"
        title="Change case"
        width={52}
        groups={[{ items: [["upper", "UPPERCASE"], ["lower", "lowercase"], ["title", "Title Case"]] }]}
        onPick={(v) => { restoreSelection(); api?.changeCase(v as "upper" | "lower" | "title"); }}
      />
      </>
      )}
      {tab === "home" && on("color") && (
      <label title="Text color" style={{ display: "inline-flex", alignItems: "center", cursor: "pointer" }}>
        <span style={{ fontSize: 13, borderBottom: `3px solid ${fmt?.color && fmt.color !== "auto" ? fmt.color : "#000"}`, padding: "0 3px", color: T.fg }}>A</span>
        <input
          type="color"
          value={fmt?.color && fmt.color !== "auto" ? fmt.color : "#000000"}
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => apply({ color: e.target.value })}
          style={{ width: 0, height: 0, opacity: 0, border: "none", padding: 0 }}
        />
      </label>
      )}
      {tab === "home" && on("highlight") && <HighlightMenu current={fmt?.highlight} onPick={(v) => apply({ highlight: v })} />}
      <Sep />
      {tab === "home" && on("alignment") && (
        <>
          <Btn label={"≡"} title="Align left" onClick={() => api?.setAlignment("left")} />
          <Btn label={"≣"} title="Center" onClick={() => api?.setAlignment("center")} />
          <Btn label={"≢"} title="Align right" onClick={() => api?.setAlignment("right")} />
          <Btn label={"☰"} title="Justify" onClick={() => api?.setAlignment("justify")} />
          <Sep />
        </>
      )}
      {tab === "home" && on("indent") && (
        <>
          <Btn label={<IndentIcon dir={-1} />} title="Decrease indent" onClick={() => api?.adjustIndent(-1)} />
          <Btn label={<IndentIcon dir={1} />} title="Increase indent" onClick={() => api?.adjustIndent(1)} />
        </>
      )}
      {tab === "home" && on("spacing") && (
        <ActionMenu
          label="↕"
          title="Line & paragraph spacing"
          width={44}
          groups={[
            { label: "Line spacing", items: [["l:1", "Single"], ["l:1.15", "1.15"], ["l:1.5", "1.5"], ["l:2", "Double"]] },
            { label: "Paragraph", items: [["b:add", "Add space before"], ["b:none", "Remove space before"], ["a:add", "Add space after"], ["a:none", "Remove space after"]] },
          ]}
          onPick={(v) => {
            if (v.startsWith("l:")) api?.setParagraphSpacing({ lineMultiple: parseFloat(v.slice(2)) });
            else if (v === "b:add") api?.setParagraphSpacing({ beforePt: 10 });
            else if (v === "b:none") api?.setParagraphSpacing({ beforePt: null });
            else if (v === "a:add") api?.setParagraphSpacing({ afterPt: 10 });
            else if (v === "a:none") api?.setParagraphSpacing({ afterPt: null });
          }}
        />
      )}
      {tab === "home" && on("lists") && (
        <>
          <Btn
            label={<BulletListIcon />}
            title="Bulleted list"
            active={listKind === "bullet"}
            onClick={() => { api?.toggleList("bullet"); refresh(); }}
          />
          <Btn
            label={<NumberListIcon />}
            title="Numbered list"
            active={listKind === "number"}
            onClick={() => { api?.toggleList("number"); refresh(); }}
          />
          <Sep />
        </>
      )}
      {tab === "insert" && (
        <>
          {on("table") && <TableMenu api={api} />}
          {on("image") && <Btn label={<ImageIcon />} title="Insert image" onClick={() => imageInput.current?.click()} />}
          {on("link") && <LinkMenu api={api} />}
          {on("comment") && <CommentMenu api={api} />}
          {on("footnote") && <FootnoteMenu api={api} />}
          <Sep />
          <ActionMenu
            label="Page number"
            title="Insert a dynamic page number at the caret"
            width={104}
            groups={[{ items: [["pn:page", "Page number"], ["pn:pageof", "Page X of Y"]] }]}
            onPick={(v) => {
              if (v === "pn:page") api?.insertPageNumber("page");
              else if (v === "pn:pageof") api?.insertPageNumber("pageOfTotal");
            }}
          />
          <ActionMenu
            label="Break"
            title="Insert a page, column or section break at the caret"
            width={64}
            groups={[
              { label: "Breaks", items: [["br:page", "Page break"], ["br:column", "Column break"]] },
              { label: "Section breaks", items: [["br:next", "Section break (next page)"], ["br:cont", "Section break (continuous)"]] },
            ]}
            onPick={(v) => {
              if (v === "br:page") api?.insertBreak("page");
              else if (v === "br:column") api?.insertBreak("column");
              else if (v === "br:next") api?.insertBreak("sectionNextPage");
              else if (v === "br:cont") api?.insertBreak("sectionContinuous");
            }}
          />
        </>
      )}
      <input
        ref={imageInput}
        type="file"
        accept="image/png,image/jpeg,image/gif,image/webp"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void api?.insertImage(f);
          e.target.value = "";
        }}
      />
      {tab === "layout" && on("layout") && <LayoutTab api={api} />}
      {on("download") && onSave && (
        <>
          <span style={{ flex: 1 }} />
          <Btn label="Download" title="Save edited .docx" onClick={() => api && onSave(api.save())} />
        </>
      )}
    </div>
  );
}
