import { useCallback, useEffect, useState } from "react";
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

/**
 * Default formatting toolbar for an editable DocxView. Compact, grouped like
 * a word processor; every control preserves the selection/caret.
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
      <ActionMenu
        label="🖊"
        title="Highlight"
        width={44}
        groups={[
          { items: HIGHLIGHTS.map((h) => [h.name, h.name] as [string, string]) },
          { items: [["__none", "remove"]] },
        ]}
        onPick={(v) => apply({ highlight: v === "__none" ? null : v })}
      />
      <Sep />
      <Btn label={"≡"} title="Align left" onClick={() => api?.setAlignment("left")} />
      <Btn label={"≣"} title="Center" onClick={() => api?.setAlignment("center")} />
      <Btn label={"≢"} title="Align right" onClick={() => api?.setAlignment("right")} />
      <Btn label={"☰"} title="Justify" onClick={() => api?.setAlignment("justify")} />
      <Sep />
      <ActionMenu
        label="Table"
        title="Insert or edit table"
        width={64}
        groups={[
          { label: "Insert", items: [["i:2x2", "2 × 2"], ["i:3x3", "3 × 3"], ["i:4x4", "4 × 4"], ["i:2x5", "2 × 5"]] },
          {
            label: "Rows / columns (at caret)",
            items: [
              ["op:rowAbove", "Insert row above"],
              ["op:rowBelow", "Insert row below"],
              ["op:deleteRow", "Delete row"],
              ["op:colLeft", "Insert column left"],
              ["op:colRight", "Insert column right"],
              ["op:deleteCol", "Delete column"],
            ],
          },
          { label: "Table", items: [["op:deleteTable", "Delete table"]] },
        ]}
        onPick={(v) => {
          if (v.startsWith("i:")) {
            const [r, c] = v.slice(2).split("x").map(Number);
            api?.insertTable(r, c);
          } else {
            api?.tableOp(v.slice(3) as Parameters<NonNullable<typeof api>["tableOp"]>[0]);
          }
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
