import { useCallback, useEffect, useState } from "react";
import type { DocxViewApi } from "./index.js";

const FONTS = [
  "Arial",
  "Calibri",
  "Cambria",
  "Courier New",
  "Garamond",
  "Georgia",
  "Helvetica",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];

const HIGHLIGHTS: { name: string; css: string }[] = [
  { name: "yellow", css: "#ffff00" },
  { name: "green", css: "#00ff00" },
  { name: "cyan", css: "#00ffff" },
  { name: "magenta", css: "#ff00ff" },
  { name: "lightGray", css: "#d3d3d3" },
];

const btn = (active: boolean): React.CSSProperties => ({
  minWidth: 28,
  height: 26,
  border: `1px solid ${active ? "#1a73e8" : "#dadce0"}`,
  background: active ? "#e8f0fe" : "#fff",
  borderRadius: 4,
  cursor: "pointer",
  fontSize: 13,
  padding: "0 6px",
});

/**
 * Default formatting toolbar for an editable DocxView. Reads current
 * selection state from the api; every control preserves the text selection
 * by preventing default on mousedown.
 */
export function DocxToolbar({ api, onSave }: { api: DocxViewApi | null; onSave?: (bytes: Uint8Array) => void }) {
  const [fmt, setFmt] = useState<ReturnType<NonNullable<DocxViewApi["getSelectionFormat"]>> | null>(null);

  const refresh = useCallback(() => {
    setFmt(api?.getSelectionFormat() ?? null);
  }, [api]);

  useEffect(() => {
    document.addEventListener("selectionchange", refresh);
    return () => document.removeEventListener("selectionchange", refresh);
  }, [refresh]);

  const apply = (patch: Parameters<DocxViewApi["applyFormat"]>[0]) => {
    api?.applyFormat(patch);
    refresh();
  };
  const keepSelection = (e: React.SyntheticEvent) => e.preventDefault();

  return (
    <div
      style={{
        display: "flex",
        gap: 6,
        alignItems: "center",
        padding: "6px 10px",
        borderBottom: "1px solid #dadce0",
        background: "#fff",
        flexWrap: "wrap",
      }}
    >
      <button style={btn(false)} title="Undo (Cmd+Z)" onMouseDown={keepSelection} onClick={() => { api?.undo(); refresh(); }}>
        {"\u21b6"}
      </button>
      <button style={btn(false)} title="Redo (Shift+Cmd+Z)" onMouseDown={keepSelection} onClick={() => { api?.redo(); refresh(); }}>
        {"\u21b7"}
      </button>
      <select
        title="Font"
        value={fmt?.fontFamily ?? ""}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => e.target.value && apply({ fontFamily: e.target.value })}
        style={{ height: 26, width: 150 }}
      >
        <option value="" disabled>
          Font
        </option>
        {FONTS.map((f) => (
          <option key={f} value={f} style={{ fontFamily: f }}>
            {f}
          </option>
        ))}
      </select>
      <select
        title="Font size"
        value={fmt?.fontSizePt ?? ""}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => e.target.value && apply({ fontSizePt: parseFloat(e.target.value) })}
        style={{ height: 26 }}
      >
        <option value="" disabled>
          Size
        </option>
        {SIZES.map((s) => (
          <option key={s} value={s}>
            {s}
          </option>
        ))}
      </select>
      <button style={btn(!!fmt?.bold)} title="Bold" onMouseDown={keepSelection} onClick={() => apply({ bold: !fmt?.bold })}>
        <b>B</b>
      </button>
      <button style={btn(!!fmt?.italic)} title="Italic" onMouseDown={keepSelection} onClick={() => apply({ italic: !fmt?.italic })}>
        <i>I</i>
      </button>
      <button
        style={btn(!!fmt?.underline)}
        title="Underline"
        onMouseDown={keepSelection}
        onClick={() => apply({ underline: !fmt?.underline })}
      >
        <u>U</u>
      </button>
      <button style={btn(!!fmt?.strike)} title="Strikethrough" onMouseDown={keepSelection} onClick={() => apply({ strike: !fmt?.strike })}>
        <s>S</s>
      </button>
      <label title="Text color" style={{ display: "inline-flex", alignItems: "center", gap: 2 }}>
        <span style={{ fontSize: 13 }}>A</span>
        <input
          type="color"
          value={fmt?.color && fmt.color !== "auto" ? fmt.color : "#000000"}
          onMouseDown={(e) => e.stopPropagation()}
          onChange={(e) => apply({ color: e.target.value })}
          style={{ width: 26, height: 22, padding: 0, border: "1px solid #dadce0" }}
        />
      </label>
      <span style={{ display: "inline-flex", gap: 2 }} title="Alignment">
        {([
          ["left", "\u2b62L", "Align left"],
          ["center", "\u2b0d", "Center"],
          ["right", "R\u2b60", "Align right"],
          ["justify", "\u2630", "Justify"],
        ] as const).map(([a, glyph, title]) => (
          <button
            key={a}
            style={btn(false)}
            title={title}
            onMouseDown={keepSelection}
            onClick={() => api?.setAlignment(a)}
          >
            {glyph}
          </button>
        ))}
      </span>
      <select
        title="Insert table"
        defaultValue=""
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => {
          if (e.target.value) {
            const [r, c] = e.target.value.split("x").map(Number);
            api?.insertTable(r, c);
          }
          e.target.value = "";
        }}
        style={{ height: 26 }}
      >
        <option value="" disabled>
          Table
        </option>
        {["2x2", "3x3", "4x4", "5x2", "2x5"].map((v) => (
          <option key={v} value={v}>
            {v.replace("x", " \u00d7 ")}
          </option>
        ))}
      </select>
      <select
        title="Page layout"
        defaultValue=""
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => {
          const v = e.target.value;
          if (v === "portrait" || v === "landscape") api?.setPageLayout({ orientation: v });
          else if (v === "normal") api?.setPageLayout({ margins: { top: 1, right: 1, bottom: 1, left: 1 } });
          else if (v === "narrow") api?.setPageLayout({ margins: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 } });
          else if (v === "wide") api?.setPageLayout({ margins: { top: 1, right: 1.5, bottom: 1, left: 1.5 } });
          else if (v === "letter") api?.setPageLayout({ size: { width: 8.5, height: 11 } });
          else if (v === "legal") api?.setPageLayout({ size: { width: 8.5, height: 14 } });
          else if (v === "a4") api?.setPageLayout({ size: { width: 8.27, height: 11.69 } });
          e.target.value = "";
        }}
        style={{ height: 26 }}
      >
        <option value="" disabled>
          Layout
        </option>
        <optgroup label="Margins">
          <option value="normal">Normal (1&quot;)</option>
          <option value="narrow">Narrow (0.5&quot;)</option>
          <option value="wide">Wide (1.5&quot;)</option>
        </optgroup>
        <optgroup label="Orientation">
          <option value="portrait">Portrait</option>
          <option value="landscape">Landscape</option>
        </optgroup>
        <optgroup label="Size">
          <option value="letter">Letter</option>
          <option value="legal">Legal</option>
          <option value="a4">A4</option>
        </optgroup>
      </select>
      <span style={{ display: "inline-flex", gap: 2 }} title="Highlight">
        {HIGHLIGHTS.map((h) => (
          <button
            key={h.name}
            onMouseDown={keepSelection}
            onClick={() => apply({ highlight: h.name })}
            style={{ ...btn(false), minWidth: 18, width: 18, background: h.css, padding: 0 }}
          />
        ))}
        <button onMouseDown={keepSelection} onClick={() => apply({ highlight: null })} style={{ ...btn(false), minWidth: 22 }} title="Remove highlight">
          ✕
        </button>
      </span>
      {onSave && (
        <button style={{ ...btn(false), marginLeft: "auto" }} onMouseDown={keepSelection} onClick={() => api && onSave(api.save())}>
          Download
        </button>
      )}
    </div>
  );
}
