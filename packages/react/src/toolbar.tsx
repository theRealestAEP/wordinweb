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
export function DocxToolbar({ api, onSave }: { api: DocxViewApi | null; onSave?: (bytes: Uint8Array) => void }) {
  const [fmt, setFmt] = useState<ReturnType<NonNullable<DocxViewApi["getSelectionFormat"]>> | null>(null);
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
      <Btn label={"↶"} title="Undo (⌘Z)" onClick={() => { api?.undo(); refresh(); }} />
      <Btn label={"↷"} title="Redo (⇧⌘Z)" onClick={() => { api?.redo(); refresh(); }} />
      <Sep />
      <select
        title="Paragraph style"
        value=""
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => {
          if (e.target.value) api?.setParagraphStyle(e.target.value === "__normal" ? null : e.target.value);
        }}
        style={{ ...selectStyle, width: 92 }}
      >
        <option value="" disabled>
          Styles
        </option>
        <option value="__normal">Normal</option>
        {(api?.listParagraphStyles() ?? [])
          .filter((s) => !/^normal$/i.test(s.name))
          .map((s) => (
            <option key={s.id} value={s.id}>{s.name}</option>
          ))}
      </select>
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
      <Sep />
      <Btn label={<b>B</b>} title="Bold (⌘B)" active={!!fmt?.bold} onClick={() => apply({ bold: !fmt?.bold })} />
      <Btn label={<i>I</i>} title="Italic" active={!!fmt?.italic} onClick={() => apply({ italic: !fmt?.italic })} />
      <Btn label={<u>U</u>} title="Underline" active={!!fmt?.underline} onClick={() => apply({ underline: !fmt?.underline })} />
      <Btn label={<s>S</s>} title="Strikethrough" active={!!fmt?.strike} onClick={() => apply({ strike: !fmt?.strike })} />
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
      <HighlightMenu current={fmt?.highlight} onPick={(v) => apply({ highlight: v })} />
      <Sep />
      <Btn label={"≡"} title="Align left" onClick={() => api?.setAlignment("left")} />
      <Btn label={"≣"} title="Center" onClick={() => api?.setAlignment("center")} />
      <Btn label={"≢"} title="Align right" onClick={() => api?.setAlignment("right")} />
      <Btn label={"☰"} title="Justify" onClick={() => api?.setAlignment("justify")} />
      <Sep />
      <TableMenu api={api} />
      <Btn label={<ImageIcon />} title="Insert image" onClick={() => imageInput.current?.click()} />
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
      {onSave && (
        <>
          <span style={{ flex: 1 }} />
          <Btn label="Download" title="Save edited .docx" onClick={() => api && onSave(api.save())} />
        </>
      )}
    </div>
  );
}
