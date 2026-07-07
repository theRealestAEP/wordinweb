import { useCallback, useEffect, useRef, useState } from "react";
import type { DocxViewApi } from "./index.js";

const FONTS = [
  "Arial", "Calibri", "Cambria", "Courier New", "Garamond",
  "Georgia", "Helvetica", "Times New Roman", "Trebuchet MS", "Verdana",
];

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
  background: active ? "#dfe7f5" : "transparent",
  borderRadius: 4,
  cursor: "pointer",
  fontSize: 13,
  padding: "0 5px",
  color: "#3c4043",
});

const selectStyle: React.CSSProperties = {
  height: 26,
  border: "1px solid transparent",
  background: "transparent",
  borderRadius: 4,
  fontSize: 13,
  color: "#3c4043",
  cursor: "pointer",
};

function Btn({ label, title, active, onClick }: { label: React.ReactNode; title: string; active?: boolean; onClick: () => void }) {
  return (
    <button
      title={title}
      style={btnStyle(!!active)}
      onMouseDown={(e) => e.preventDefault()}
      onMouseEnter={(e) => ((e.target as HTMLElement).style.background = active ? "#dfe7f5" : "#f1f3f4")}
      onMouseLeave={(e) => ((e.target as HTMLElement).style.background = active ? "#dfe7f5" : "transparent")}
      onClick={onClick}
    >
      {label}
    </button>
  );
}

function Sep() {
  return <span style={{ width: 1, height: 18, background: "#dadce0", margin: "0 4px", flexShrink: 0 }} />;
}

const icon = { width: 16, height: 16, display: "block" } as const;

function ImageIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <circle cx="5.2" cy="6" r="1.2" fill="#3c4043" stroke="none" />
      <path d="M2.5 12l3.5-4 2.8 3 2-2.4 2.7 3.4" />
    </svg>
  );
}

function HighlightIcon({ color }: { color: string }) {
  return (
    <svg style={icon} viewBox="0 0 16 16">
      <path d="M3 9.5L9.5 3l3.5 3.5L6.5 13H4.5L3 11.5v-2z" fill="none" stroke="#3c4043" strokeWidth="1.3" />
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
            position: "absolute", top: 28, left: 0, zIndex: 100, background: "#fff",
            border: "1px solid #dadce0", borderRadius: 6, boxShadow: "0 4px 16px rgba(0,0,0,.15)",
            padding: 8, display: "flex", gap: 4, alignItems: "center",
          }}
        >
          {HIGHLIGHTS.map((h) => (
            <div
              key={h.name}
              title={h.name}
              onClick={() => { onPick(h.name); setOpen(false); }}
              style={{ width: 20, height: 20, background: h.css, border: "1px solid #dadce0", borderRadius: 3, cursor: "pointer" }}
            />
          ))}
          <div
            title="No highlight"
            onClick={() => { onPick(null); setOpen(false); }}
            style={{
              width: 20, height: 20, border: "1px solid #dadce0", borderRadius: 3, cursor: "pointer",
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
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <circle cx="3" cy="4" r="1.1" fill="#3c4043" stroke="none" />
      <circle cx="3" cy="8" r="1.1" fill="#3c4043" stroke="none" />
      <circle cx="3" cy="12" r="1.1" fill="#3c4043" stroke="none" />
      <path d="M6.5 4h8M6.5 8h8M6.5 12h8" strokeLinecap="round" />
    </svg>
  );
}

function NumberListIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <text x="1" y="5.6" fontSize="5.4" fill="#3c4043" stroke="none" fontFamily="system-ui">1</text>
      <text x="1" y="9.9" fontSize="5.4" fill="#3c4043" stroke="none" fontFamily="system-ui">2</text>
      <text x="1" y="14.2" fontSize="5.4" fill="#3c4043" stroke="none" fontFamily="system-ui">3</text>
      <path d="M6.5 4h8M6.5 8h8M6.5 12h8" strokeLinecap="round" />
    </svg>
  );
}

function LinkIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <path d="M6.5 9.5l3-3" strokeLinecap="round" />
      <path d="M7.5 4.5l1.2-1.2a2.6 2.6 0 013.7 3.7L11.2 8.2" strokeLinecap="round" />
      <path d="M8.5 11.5l-1.2 1.2a2.6 2.6 0 01-3.7-3.7l1.2-1.2" strokeLinecap="round" />
    </svg>
  );
}

function ClearFormatIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <path d="M5 3h8M9 3l-2.5 9" strokeLinecap="round" />
      <path d="M3 13.5l3.5-3.5M3 10l3.5 3.5" strokeLinecap="round" />
    </svg>
  );
}

function IndentIcon({ dir }: { dir: 1 | -1 }) {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <path d="M2 3h12M8 6.5h6M8 9.5h6M2 13h12" strokeLinecap="round" />
      {dir === 1 ? <path d="M2.5 6l3 2-3 2z" fill="#3c4043" stroke="none" /> : <path d="M5.5 6l-3 2 3 2z" fill="#3c4043" stroke="none" />}
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
            position: "absolute", top: 28, left: 0, zIndex: 100, background: "#fff",
            border: "1px solid #dadce0", borderRadius: 8, boxShadow: "0 4px 16px rgba(0,0,0,.15)",
            padding: 10, width: 260, display: "flex", gap: 6, alignItems: "center",
          }}
        >
          <input
            ref={inputRef}
            value={url}
            placeholder="Paste or type a link"
            onChange={(e) => setUrl(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && submit()}
            style={{ flex: 1, border: "1px solid #dadce0", borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          <button style={pillBtn} disabled={!url.trim()} onClick={submit}>Apply</button>
          {api?.getLinkAt() && (
            <button
              title="Remove link"
              style={{ ...pillBtn, background: "#fff", color: "#3c4043" }}
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

function CommentIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="#3c4043" strokeWidth="1.4">
      <path d="M1.5 3.5h13v8h-7l-3 3v-3h-3z" strokeLinejoin="round" />
      <path d="M8 5.5v4M6 7.5h4" strokeLinecap="round" />
    </svg>
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
            position: "absolute", top: 28, right: 0, zIndex: 100, background: "#fff",
            border: "1px solid #dadce0", borderRadius: 8, boxShadow: "0 4px 16px rgba(0,0,0,.15)",
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
              border: "1px solid #dadce0", borderRadius: 6, padding: 6,
              font: "13px system-ui, sans-serif", outline: "none",
            }}
          />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, marginTop: 6 }}>
            <button style={{ ...pillBtn, background: "#fff", color: "#3c4043" }} onClick={() => setOpen(false)}>
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
  border: "1px solid #dadce0",
  borderRadius: 14,
  padding: "3px 12px",
  fontSize: 12.5,
  cursor: "pointer",
  background: "#1a73e8",
  color: "#fff",
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
    ["deleteTable", "Delete table"],
  ];

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
            background: "#fff",
            border: "1px solid #dadce0",
            borderRadius: 6,
            boxShadow: "0 4px 16px rgba(0,0,0,.15)",
            padding: 8,
            width: COLS * 18 + 16,
          }}
          onMouseDown={(e) => e.preventDefault()}
        >
          <div style={{ fontSize: 12, color: "#5f6368", marginBottom: 4 }}>
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
                    border: `1px solid ${lit ? "#1a73e8" : "#dadce0"}`,
                    background: lit ? "#dfe7f5" : "#fff",
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
                  api?.tableOp(op as Parameters<NonNullable<typeof api>["tableOp"]>[0]);
                  setOpen(false);
                }}
                style={{ padding: "4px 6px", fontSize: 13, cursor: "pointer", borderRadius: 4, color: "#3c4043" }}
                onMouseEnter={(e) => ((e.target as HTMLElement).style.background = "#f1f3f4")}
                onMouseLeave={(e) => ((e.target as HTMLElement).style.background = "transparent")}
              >
                {label}
              </div>
            ))}
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
  | "layout"
  | "download";

export function DocxToolbar({
  api,
  onSave,
  features,
}: {
  api: DocxViewApi | null;
  onSave?: (bytes: Uint8Array) => void;
  /** Per-group overrides; every group defaults to enabled. */
  features?: Partial<Record<ToolbarFeature, boolean>>;
}) {
  const on = (k: ToolbarFeature) => features?.[k] !== false;
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
      onMouseOver={onTipOver}
      onMouseOut={onTipOut}
      onMouseDownCapture={onTipOut}
      style={{
        display: "flex",
        gap: 2,
        alignItems: "center",
        padding: "4px 10px",
        borderBottom: "1px solid #dadce0",
        background: "#f9fbfd",
        flexWrap: "wrap",
        fontFamily: "system-ui, sans-serif",
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
            color: "#fff",
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
      {on("history") && (
        <>
          <Btn label={"↶"} title="Undo (⌘Z)" onClick={() => { api?.undo(); refresh(); }} />
          <Btn label={"↷"} title="Redo (⇧⌘Z)" onClick={() => { api?.redo(); refresh(); }} />
          <Sep />
        </>
      )}
      {on("styles") && (
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
      {on("font") && (
      <select
        title="Font"
        value={fmt?.fontFamily ?? ""}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => e.target.value && apply({ fontFamily: e.target.value })}
        style={{ ...selectStyle, width: 130 }}
      >
        <option value="" disabled>Font</option>
        {FONTS.map((f) => (
          <option key={f} value={f} style={{ fontFamily: f }}>{f}</option>
        ))}
      </select>
      )}
      {on("size") && (
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
      {on("format") && (
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
      {on("color") && (
      <label title="Text color" style={{ display: "inline-flex", alignItems: "center", cursor: "pointer" }}>
        <span style={{ fontSize: 13, borderBottom: `3px solid ${fmt?.color && fmt.color !== "auto" ? fmt.color : "#000"}`, padding: "0 3px", color: "#3c4043" }}>A</span>
        <input
          type="color"
          value={fmt?.color && fmt.color !== "auto" ? fmt.color : "#000000"}
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => apply({ color: e.target.value })}
          style={{ width: 0, height: 0, opacity: 0, border: "none", padding: 0 }}
        />
      </label>
      )}
      {on("highlight") && <HighlightMenu current={fmt?.highlight} onPick={(v) => apply({ highlight: v })} />}
      <Sep />
      {on("alignment") && (
        <>
          <Btn label={"≡"} title="Align left" onClick={() => api?.setAlignment("left")} />
          <Btn label={"≣"} title="Center" onClick={() => api?.setAlignment("center")} />
          <Btn label={"≢"} title="Align right" onClick={() => api?.setAlignment("right")} />
          <Btn label={"☰"} title="Justify" onClick={() => api?.setAlignment("justify")} />
          <Sep />
        </>
      )}
      {on("indent") && (
        <>
          <Btn label={<IndentIcon dir={-1} />} title="Decrease indent" onClick={() => api?.adjustIndent(-1)} />
          <Btn label={<IndentIcon dir={1} />} title="Increase indent" onClick={() => api?.adjustIndent(1)} />
        </>
      )}
      {on("spacing") && (
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
      {on("lists") && (
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
      {on("table") && <TableMenu api={api} />}
      {on("image") && <Btn label={<ImageIcon />} title="Insert image" onClick={() => imageInput.current?.click()} />}
      {on("link") && <LinkMenu api={api} />}
      {on("comment") && <CommentMenu api={api} />}
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
      {on("layout") && (
      <ActionMenu
        label="Layout"
        title="Page layout"
        width={70}
        groups={[
          { label: "Margins", items: [["m:normal", 'Normal (1")'], ["m:narrow", 'Narrow (0.5")'], ["m:wide", 'Wide (1.5")']] },
          { label: "Orientation", items: [["o:portrait", "Portrait"], ["o:landscape", "Landscape"]] },
          { label: "Size", items: [["s:letter", "Letter"], ["s:legal", "Legal"], ["s:a4", "A4"]] },
        ]}
        onPick={(v) => {
          if (v === "m:normal") api?.setPageLayout({ margins: { top: 1, right: 1, bottom: 1, left: 1 } });
          else if (v === "m:narrow") api?.setPageLayout({ margins: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 } });
          else if (v === "m:wide") api?.setPageLayout({ margins: { top: 1, right: 1.5, bottom: 1, left: 1.5 } });
          else if (v === "o:portrait" || v === "o:landscape") api?.setPageLayout({ orientation: v.slice(2) as "portrait" | "landscape" });
          else if (v === "s:letter") api?.setPageLayout({ size: { width: 8.5, height: 11 } });
          else if (v === "s:legal") api?.setPageLayout({ size: { width: 8.5, height: 14 } });
          else if (v === "s:a4") api?.setPageLayout({ size: { width: 8.27, height: 11.69 } });
        }}
      />
      )}
      {on("download") && onSave && (
        <>
          <span style={{ flex: 1 }} />
          <Btn label="Download" title="Save edited .docx" onClick={() => api && onSave(api.save())} />
        </>
      )}
    </div>
  );
}
