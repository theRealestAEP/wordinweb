import { Fragment, useCallback, useEffect, useId, useLayoutEffect, useReducer, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  availableObjectCommands,
  requestTextInputDialog,
  styleIdFromName,
  uniqueStyleId,
  CELL_SCOPE_EDGES,
  TABLE_BORDER_STYLES,
  TABLE_SCOPE_EDGES,
  NUMBERING_PRESETS,
  SHAPE_GALLERY,
  presetShapeGeometry,
  type NumberingPresetId,
  type SelectionFormat,
  type TabStopSpec,
  type TableBorderEdge,
  type TableBorderStyle,
} from "@wordinweb/core";
import type { DocxViewApi } from "./index.js";
import { HelpGuide } from "./help.js";

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
    // Keep the standard document families selectable even when a browser's
    // synchronous width probe runs before their @font-face files finish loading.
    const always = new Set(["Arial", "Calibri", "Cambria", "Courier New", "Times New Roman"]);
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

const PAGE_SIZES = [
  { value: "letter", label: "Letter", description: '8.5" × 11"', width: 8.5, height: 11 },
  { value: "legal", label: "Legal", description: '8.5" × 14"', width: 8.5, height: 14 },
  { value: "3.5x5", label: "3.5 × 5", description: '3.5" × 5"', width: 3.5, height: 5 },
  { value: "3.5x5-borderless", label: "3.5 × 5 Borderless", description: '3.5" × 5"', width: 3.5, height: 5 },
  { value: "4x6", label: "4 × 6", description: '4" × 6"', width: 4, height: 6 },
  { value: "4x6-borderless", label: "4 × 6 Borderless", description: '4" × 6"', width: 4, height: 6 },
  { value: "5x7", label: "5 × 7", description: '5" × 7"', width: 5, height: 7 },
  { value: "5x7-borderless", label: "5 × 7 Borderless", description: '5" × 7"', width: 5, height: 7 },
  { value: "8x10", label: "8 × 10", description: '8" × 10"', width: 8, height: 10 },
  { value: "8x10-borderless", label: "8 × 10 Borderless", description: '8" × 10"', width: 8, height: 10 },
  { value: "a4", label: "A4", description: '8.27" × 11.69"', width: 8.27, height: 11.69 },
  { value: "a4-borderless", label: "A4 Borderless", description: '8.27" × 11.69"', width: 8.27, height: 11.69 },
  { value: "a6", label: "A6", description: '4.13" × 5.83"', width: 4.13, height: 5.83 },
  { value: "envelope10", label: "Envelope #10", description: '4.13" × 9.5"', width: 4.13, height: 9.5 },
] as const;

function Btn({ label, title, active, onClick, buttonRef }: { label: React.ReactNode; title: string; active?: boolean; onClick: () => void; buttonRef?: React.Ref<HTMLButtonElement> }) {
  return (
    <button
      ref={buttonRef}
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

function OverflowIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="currentColor">
      <circle cx="8" cy="3" r="1.4" />
      <circle cx="8" cy="8" r="1.4" />
      <circle cx="8" cy="13" r="1.4" />
    </svg>
  );
}

/**
 * "More" (⋮) menu holding the toolbar groups that don't fit the current width.
 * On a phone/tablet the low-frequency groups collapse in here (Google-Docs
 * pattern) so the primary row stays a single clean strip; every control stays
 * reachable. The grouped controls render stacked, wrapping as needed.
 */
function OverflowMenu({ children }: { children: React.ReactNode }) {
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
    <span ref={rootRef} style={{ position: "relative", display: "inline-flex", marginLeft: "auto" }}>
      <button
        title="More tools"
        data-dxw-overflow=""
        style={btnStyle(open)}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        <OverflowIcon />
      </button>
      {open && (
        <div
          data-dxw-overflow-menu=""
          style={{
            position: "absolute", top: 30, right: 0, zIndex: 100, background: T.popoverBg,
            border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow,
            padding: 8, display: "flex", flexWrap: "wrap", alignItems: "center", gap: 4,
            width: "min(280px, calc(100vw - 16px))",
            boxSizing: "border-box",
          }}
        >
          {children}
        </div>
      )}
    </span>
  );
}

const icon = { width: 16, height: 16, display: "block" } as const;

function ExpandChevronIcon({ up }: { up: boolean }) {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
      {up ? <path d="M4 9.5L8 5.5l4 4" /> : <path d="M4 6.5l4 4 4-4" />}
    </svg>
  );
}

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

export interface ToolbarMenuSelectOption {
  value: string;
  label: React.ReactNode;
  group?: string;
  disabled?: boolean;
  fontFamily?: string;
}

export interface ToolbarMenuSelectProps {
  value: string;
  options: ToolbarMenuSelectOption[];
  onChange: (value: string) => void;
  placeholder?: React.ReactNode;
  title?: string;
  ariaLabel?: string;
  triggerAriaLabel?: string;
  width?: number | string;
  menuWidth?: number;
  className?: string;
  style?: React.CSSProperties;
}

/** CSS-overridable replacement for visible native selects. An inert,
 * transparent select is kept as an event bridge for existing integrations;
 * every user-facing part is a button/listbox rendered by us. */
export function ToolbarMenuSelect({
  value,
  options,
  onChange,
  placeholder = "Choose…",
  title,
  ariaLabel,
  triggerAriaLabel,
  width,
  menuWidth,
  className,
  style,
}: ToolbarMenuSelectProps) {
  const [open, setOpen] = useState(false);
  const [position, setPosition] = useState({ left: 8, top: 8, width: menuWidth ?? 180, maxHeight: 320 });
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const keyboardOpen = useRef<"first" | "last" | null>(null);
  const id = useId();
  const selected = options.find((option) => option.value === value);

  useLayoutEffect(() => {
    if (!open) return;
    const update = () => {
      const trigger = triggerRef.current;
      if (!trigger) return;
      const rect = trigger.getBoundingClientRect();
      const nextWidth = Math.min(
        window.innerWidth - 16,
        menuWidth ?? Math.max(rect.width, menuRef.current?.scrollWidth ?? 180),
      );
      const below = window.innerHeight - rect.bottom - 8;
      const above = rect.top - 8;
      const placeAbove = below < 140 && above > below;
      const maxHeight = Math.max(96, Math.min(320, placeAbove ? above : below));
      const shownHeight = Math.min(menuRef.current?.scrollHeight ?? maxHeight, maxHeight);
      setPosition({
        left: Math.max(8, Math.min(rect.left, window.innerWidth - nextWidth - 8)),
        top: placeAbove ? Math.max(8, rect.top - shownHeight - 4) : rect.bottom + 4,
        width: nextWidth,
        maxHeight,
      });
    };
    update();
    const frame = requestAnimationFrame(() => {
      update();
      if (!keyboardOpen.current) return;
      const items = menuRef.current?.querySelectorAll<HTMLButtonElement>('[role="option"]');
      const item = keyboardOpen.current === "last" ? items?.[items.length - 1] : items?.[0];
      item?.focus({ preventScroll: true });
      keyboardOpen.current = null;
    });
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setOpen(false);
        requestAnimationFrame(() => triggerRef.current?.focus({ preventScroll: true }));
      }
    };
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", update);
    window.addEventListener("scroll", update, true);
    return () => {
      cancelAnimationFrame(frame);
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", update);
      window.removeEventListener("scroll", update, true);
    };
  }, [open, menuWidth]);

  const pick = (next: string) => {
    const option = options.find((item) => item.value === next);
    if (!option || option.disabled) return;
    onChange(next);
    setOpen(false);
  };
  const onMenuKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const items = menuRef.current
      ? Array.from(menuRef.current.querySelectorAll<HTMLButtonElement>('[role="option"]:not(:disabled)'))
      : [];
    const current = items.indexOf(document.activeElement as HTMLButtonElement);
    let next = current;
    if (event.key === "ArrowDown") next = current < 0 ? 0 : (current + 1) % items.length;
    else if (event.key === "ArrowUp") next = current < 0 ? items.length - 1 : (current - 1 + items.length) % items.length;
    else if (event.key === "Home") next = 0;
    else if (event.key === "End") next = items.length - 1;
    else if ((event.key === "Enter" || event.key === " ") && current >= 0) {
      event.preventDefault();
      items[current].click();
      return;
    } else if (event.key === "Escape" || event.key === "Tab") {
      event.preventDefault();
      setOpen(false);
      requestAnimationFrame(() => triggerRef.current?.focus({ preventScroll: true }));
      return;
    } else return;
    event.preventDefault();
    items[next]?.focus({ preventScroll: true });
  };

  let previousGroup: string | undefined;
  return (
    <span
      ref={rootRef}
      className={`dxw-menu-select${className ? ` ${className}` : ""}`}
      data-dxw-menu-select=""
      style={{ position: "relative", display: "inline-flex", width }}
    >
      <select
        tabIndex={-1}
        title={title}
        aria-label={ariaLabel}
        aria-hidden="true"
        value={value}
        onChange={(event) => pick(event.target.value)}
        data-dxw-native-bridge=""
        style={{
          position: "absolute",
          width: 1,
          height: 1,
          inset: 0,
          opacity: 0,
          pointerEvents: "none",
        }}
      >
        {!options.some((option) => option.value === value) && <option value="">{String(placeholder)}</option>}
        {options.map((option) => <option key={option.value} value={option.value} disabled={option.disabled}>{String(option.label)}</option>)}
      </select>
      <button
        ref={triggerRef}
        type="button"
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        aria-label={triggerAriaLabel}
        data-tip={title}
        className="dxw-menu-select-trigger"
        data-dxw-menu-select-trigger=""
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => setOpen(!open)}
        onKeyDown={(event) => {
          if (!open && (event.key === "Enter" || event.key === " " || event.key === "ArrowDown" || event.key === "ArrowUp")) {
            event.preventDefault();
            keyboardOpen.current = event.key === "ArrowUp" ? "last" : "first";
            setOpen(true);
          }
        }}
        style={{
          width: "100%",
          minWidth: 0,
          height: "var(--dxw-select-height, 26px)",
          border: "1px solid var(--dxw-select-border, transparent)",
          borderRadius: "var(--dxw-select-radius, 4px)",
          background: "var(--dxw-select-bg, transparent)",
          color: "var(--dxw-select-fg, var(--dxw-toolbar-fg, #3c4043))",
          padding: "var(--dxw-select-padding, 0 6px)",
          display: "inline-flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 6,
          cursor: "pointer",
          font: "var(--dxw-select-font, 13px system-ui, sans-serif)",
          ...style,
        }}
      >
        <span style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontFamily: selected?.fontFamily }}>
          {selected?.label ?? placeholder}
        </span>
        <span aria-hidden="true" className="dxw-menu-select-chevron" style={{ flexShrink: 0, fontSize: 10 }}>⌄</span>
      </button>
      {open && (
        <div
          ref={menuRef}
          id={id}
          role="listbox"
          aria-label={ariaLabel ?? title}
          className="dxw-menu-select-menu"
          data-dxw-menu-select-menu=""
          onMouseDown={(event) => event.preventDefault()}
          onKeyDown={onMenuKeyDown}
          style={{
            position: "fixed",
            left: position.left,
            top: position.top,
            zIndex: "var(--dxw-toolbar-z-index, 100)",
            width: position.width,
            maxHeight: `min(var(--dxw-select-menu-max-height, 320px), ${position.maxHeight}px)`,
            overflowY: "auto",
            overscrollBehavior: "contain",
            boxSizing: "border-box",
            padding: "var(--dxw-select-menu-padding, 5px)",
            border: "1px solid var(--dxw-select-menu-border, var(--dxw-toolbar-border, #dadce0))",
            borderRadius: "var(--dxw-select-menu-radius, 8px)",
            background: "var(--dxw-select-menu-bg, var(--dxw-popover-bg, #fff))",
            boxShadow: "var(--dxw-select-menu-shadow, var(--dxw-popover-shadow, 0 4px 16px rgba(0,0,0,.15)))",
          }}
        >
          {options.map((option) => {
            const groupChanged = option.group !== previousGroup;
            previousGroup = option.group;
            const active = option.value === value;
            return (
              <Fragment key={option.value}>
                {groupChanged && option.group && (
                  <div className="dxw-menu-select-group" style={{ padding: "6px 8px 3px", color: T.muted, font: "600 10.5px system-ui, sans-serif" }}>
                    {option.group}
                  </div>
                )}
                <button
                  type="button"
                  role="option"
                  aria-selected={active}
                  disabled={option.disabled}
                  tabIndex={-1}
                  className="dxw-menu-select-option"
                  data-dxw-menu-select-option={option.value}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={() => pick(option.value)}
                  style={{
                    width: "100%",
                    minHeight: 30,
                    border: 0,
                    borderRadius: 6,
                    padding: "5px 8px",
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    background: active ? T.activeBg : "transparent",
                    color: option.disabled ? T.muted : T.fg,
                    textAlign: "left",
                    cursor: option.disabled ? "default" : "pointer",
                    font: "13px system-ui, sans-serif",
                    fontFamily: option.fontFamily,
                  }}
                  onMouseEnter={(event) => { if (!active && !option.disabled) event.currentTarget.style.background = T.hoverBg; }}
                  onMouseLeave={(event) => { if (!active) event.currentTarget.style.background = "transparent"; }}
                >
                  <span style={{ flex: 1 }}>{option.label}</span>
                  {active && <span aria-hidden="true" style={{ color: T.accent, fontWeight: 700 }}>✓</span>}
                </button>
              </Fragment>
            );
          })}
        </div>
      )}
    </span>
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
    <ToolbarMenuSelect
      title={title}
      triggerAriaLabel={title}
      value=""
      placeholder={label}
      width={width}
      menuWidth={Math.max(width ?? 0, 190)}
      options={groups.flatMap((group) => group.items.map(([value, text]) => ({
        value,
        label: text,
        group: group.label,
      })))}
      onChange={onPick}
    />
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

const COLOR_SWATCHES = [
  "#000000", "#434343", "#666666", "#999999", "#b7b7b7", "#cccccc", "#d9d9d9", "#ffffff",
  "#980000", "#ff0000", "#ff9900", "#ffff00", "#00ff00", "#00ffff", "#4a86e8", "#0000ff",
  "#9900ff", "#ff00ff", "#e6b8af", "#f4cccc", "#fce5cd", "#fff2cc", "#d9ead3", "#d0e0e3",
  "#c9daf8", "#cfe2f3", "#d9d2e9", "#ead1dc", "#a61c00", "#cc0000", "#e69138", "#6aa84f",
] as const;

function normalizedColor(value: string): string | null {
  const match = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(value.trim());
  if (!match) return null;
  const digits = match[1].length === 3
    ? match[1].split("").map((digit) => digit + digit).join("")
    : match[1];
  return `#${digits.toLowerCase()}`;
}

/** Custom palette + hex entry used anywhere a native color picker used to
 * appear. All surfaces expose stable classes and inherit the toolbar tokens. */
function ColorMenu({
  current,
  title,
  trigger,
  onPick,
}: {
  current: string;
  title: string;
  trigger: React.ReactNode;
  onPick: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [custom, setCustom] = useState(current);
  const [position, setPosition] = useState({ left: 8, top: 8 });
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const valid = normalizedColor(custom);
  useLayoutEffect(() => {
    if (!open) return;
    setCustom(current);
    const update = () => {
      const rect = triggerRef.current?.getBoundingClientRect();
      if (!rect) return;
      const width = Math.min(236, window.innerWidth - 16);
      const menuHeight = menuRef.current?.offsetHeight ?? 188;
      const top = window.innerHeight - rect.bottom >= menuHeight + 8
        ? rect.bottom + 4
        : Math.max(8, rect.top - menuHeight - 4);
      setPosition({ left: Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)), top });
    };
    update();
    const frame = requestAnimationFrame(update);
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        setOpen(false);
        requestAnimationFrame(() => triggerRef.current?.focus({ preventScroll: true }));
      }
    };
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", update);
    window.addEventListener("scroll", update, true);
    return () => {
      cancelAnimationFrame(frame);
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", update);
      window.removeEventListener("scroll", update, true);
    };
  }, [open, current]);
  const pick = (value: string) => {
    onPick(value);
    setOpen(false);
  };
  return (
    <span ref={rootRef} className="dxw-color-control" style={{ position: "relative", display: "inline-flex" }}>
      <button
        ref={triggerRef}
        type="button"
        title={title}
        aria-haspopup="dialog"
        aria-expanded={open}
        className="dxw-color-trigger"
        data-dxw-color-trigger=""
        style={{ ...btnStyle(open), display: "inline-flex", alignItems: "center", gap: 5 }}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => setOpen(!open)}
      >
        {trigger}
      </button>
      {open && (
        <div
          ref={menuRef}
          role="dialog"
          aria-label={title}
          className="dxw-color-menu"
          data-dxw-color-menu=""
          onMouseDown={(event) => event.stopPropagation()}
          style={{
            position: "fixed", left: position.left, top: position.top,
            zIndex: "var(--dxw-toolbar-z-index, 100)",
            width: "min(var(--dxw-color-menu-width, 236px), calc(100vw - 16px))",
            boxSizing: "border-box", padding: 8,
            border: `1px solid ${T.border}`, borderRadius: 8,
            background: T.popoverBg, boxShadow: T.popoverShadow,
          }}
        >
          <div className="dxw-color-swatches" style={{ display: "grid", gridTemplateColumns: "repeat(8, 1fr)", gap: 4 }}>
            {COLOR_SWATCHES.map((color) => (
              <button
                key={color}
                type="button"
                title={color}
                aria-label={`Choose ${color}`}
                className="dxw-color-swatch"
                data-dxw-color={color}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => pick(color)}
                style={{
                  width: 23, height: 23, padding: 0, borderRadius: 4, cursor: "pointer",
                  border: color.toLowerCase() === current.toLowerCase() ? `2px solid ${T.accent}` : `1px solid ${T.border}`,
                  background: color,
                }}
              />
            ))}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 8 }}>
            <span aria-hidden="true" style={{ width: 24, height: 24, borderRadius: 4, border: `1px solid ${T.border}`, background: valid ?? current, flexShrink: 0 }} />
            <input
              aria-label="Custom hex color"
              className="dxw-color-value"
              value={custom}
              onChange={(event) => setCustom(event.target.value)}
              onKeyDown={(event) => { if (event.key === "Enter" && valid) pick(valid); }}
              spellCheck={false}
              placeholder="#1a73e8"
              style={{ minWidth: 0, flex: 1, height: 28, boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 6px", color: T.fg, background: T.popoverBg }}
            />
            <button type="button" disabled={!valid} onClick={() => valid && pick(valid)} style={pillBtn}>Apply</button>
          </div>
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
            type="url"
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

/** The footnote mark over its separator, plus the closing rule that says the
 * note sits at the END of the document rather than the foot of the page. */
function EndnoteIcon() {
  return (
    <svg style={icon} viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.4">
      <path d="M2.5 10.5V3.5M2.5 3.5h7" strokeLinecap="round" />
      <text x="10.2" y="6.6" fontSize="6.4" fill="currentColor" stroke="none" fontFamily="system-ui">i</text>
      <path d="M2.5 10.5h11" strokeLinecap="round" strokeDasharray="1.5 1.6" />
      <path d="M2.5 13.5h11" strokeLinecap="round" />
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

/** Insert-note popover: a text box; the note lands at the caret. A footnote
 * goes to the foot of its page, an endnote to the end of the document. */
function NoteMenu({ api, kind }: { api: DocxViewApi | null; kind: "footnote" | "endnote" }) {
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
    const added = kind === "footnote" ? api?.addFootnote(text) : api?.addEndnote(text);
    if (added) {
      setText("");
      setHint("");
      setOpen(false);
    } else {
      setHint("Click into the text first so the reference has a place to go.");
    }
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title={`Insert ${kind} (at the caret)`} style={btnStyle(open)} onMouseDown={(e) => e.preventDefault()} onClick={() => setOpen(!open)}>
        {kind === "footnote" ? <FootnoteIcon /> : <EndnoteIcon />}
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
            placeholder={kind === "footnote" ? "Footnote text…" : "Endnote text…"}
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
function CommentMenu({ api, mentions = [] }: { api: DocxViewApi | null; mentions?: string[] }) {
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
  const addMention = (name: string) => {
    const field = inputRef.current;
    const start = field?.selectionStart ?? text.length;
    const end = field?.selectionEnd ?? start;
    const prefix = start > 0 && !/\s$/.test(text.slice(0, start)) ? " " : "";
    const value = `${prefix}@${name} `;
    setText(text.slice(0, start) + value + text.slice(end));
    requestAnimationFrame(() => {
      field?.focus();
      field?.setSelectionRange(start + value.length, start + value.length);
    });
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
          {mentions.length > 0 && (
            <div aria-label="Mention a collaborator" style={{ display: "flex", flexWrap: "wrap", gap: 4, alignItems: "center", marginTop: 6 }}>
              <span style={{ color: T.muted, fontSize: 11 }}>Mention:</span>
              {[...new Set(mentions)].filter((name) => name.trim()).map((name) => (
                <button
                  type="button"
                  key={name}
                  onClick={() => addMention(name)}
                  style={{ ...pillBtn, minHeight: 24, padding: "0 7px", background: T.popoverBg, color: T.fg, borderColor: T.border }}
                >@{name}</button>
              ))}
            </div>
          )}
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

function BookmarkMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState("");
  const [error, setError] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    if (!open) return;
    inputRef.current?.focus();
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const submit = () => {
    const value = name.trim();
    if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(value)) {
      setError("Start with a letter; use letters, numbers, or underscores (40 characters max).");
      return;
    }
    if (!api?.addBookmark(value)) {
      setError("Select text or place the caret, and use a bookmark name that is not already present.");
      return;
    }
    setName("");
    setError("");
    setOpen(false);
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert bookmark" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        Bookmark
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, left: 0, zIndex: 100, width: 280, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <div style={{ font: "600 12px system-ui, sans-serif", marginBottom: 6, color: T.fg }}>Bookmark name</div>
          <input
            ref={inputRef}
            value={name}
            placeholder="Quarterly_Revenue"
            onChange={(event) => { setName(event.target.value); setError(""); }}
            onKeyDown={(event) => event.key === "Enter" && submit()}
            style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          {error && <div style={{ color: "#c5221f", fontSize: 11.5, marginTop: 5 }}>{error}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, marginTop: 8 }}>
            <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} onClick={() => setOpen(false)}>Cancel</button>
            <button style={pillBtn} disabled={!name.trim()} onClick={submit}>Add</button>
          </div>
        </div>
      )}
    </span>
  );
}

function CrossReferenceMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [refType, setRefType] = useState<"bookmark" | "heading" | "caption" | "numberedItem">("bookmark");
  const [choice, setChoice] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const bookmarks = open && refType === "bookmark" ? api?.listBookmarks() ?? [] : [];
  const targets = open && refType !== "bookmark"
    ? (api?.listCrossRefTargets() ?? []).filter((target) => target.kind === refType)
    : [];
  const options = refType === "bookmark"
    ? bookmarks.map((name) => ({ value: `b:${name}`, label: name }))
    : targets.map((target) => ({ value: `t:${target.key}`, label: target.text }));
  const selected = options.some((option) => option.value === choice) ? choice : options[0]?.value ?? "";
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const insert = (kind: "text" | "page") => {
    if (!selected) return;
    const done = selected.startsWith("b:")
      ? api?.insertCrossReference(selected.slice(2), kind)
      : api?.insertCrossRefToTarget(selected.slice(2), kind);
    if (done) setOpen(false);
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert cross-reference" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        Cross-reference
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, left: 0, zIndex: 100, width: 280, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow, display: "grid", gap: 8 }}>
          <ToolbarMenuSelect
            value={refType}
            ariaLabel="Reference type"
            width="100%"
            menuWidth={240}
            options={[
              { value: "bookmark", label: "Bookmark" },
              { value: "heading", label: "Heading" },
              { value: "caption", label: "Caption" },
              { value: "numberedItem", label: "Numbered item" },
            ]}
            onChange={(value) => { setRefType(value as "bookmark" | "heading" | "caption" | "numberedItem"); setChoice(""); }}
            style={{ borderColor: T.border }}
          />
          {options.length === 0 ? (
            <div style={{ color: T.muted, fontSize: 12 }}>
              {refType === "bookmark"
                ? "Add a bookmark first, then reference its text or page."
                : "Nothing of this type to reference yet."}
            </div>
          ) : (
            <>
              <ToolbarMenuSelect
                value={selected}
                ariaLabel="Target to reference"
                width="100%"
                menuWidth={260}
                options={options}
                onChange={setChoice}
                style={{ borderColor: T.border }}
              />
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
                <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} onClick={() => insert("page")}>Page number</button>
                <button style={pillBtn} onClick={() => insert("text")}>Referenced text</button>
              </div>
            </>
          )}
        </div>
      )}
    </span>
  );
}

/** Word's Insert Caption, popover-sized: label, optional text, above/below. */
function CaptionMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [label, setLabel] = useState("Figure");
  const [text, setText] = useState("");
  const [position, setPosition] = useState<"below" | "above">("below");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert a caption for the selected object or table" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        Caption
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, left: 0, zIndex: 100, width: 250, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow, display: "grid", gap: 7 }}>
          <label style={dialogFieldRow}>
            <span>Label</span>
            <select aria-label="Caption label" value={label} onChange={(event) => setLabel(event.target.value)} style={dialogInput}>
              <option value="Figure">Figure</option>
              <option value="Table">Table</option>
              <option value="Equation">Equation</option>
            </select>
          </label>
          <label style={dialogFieldRow}>
            <span>Text</span>
            <input aria-label="Caption text" value={text} onChange={(event) => setText(event.target.value)} placeholder="optional" style={dialogInput} />
          </label>
          <label style={dialogFieldRow}>
            <span>Position</span>
            <select aria-label="Caption position" value={position} onChange={(event) => setPosition(event.target.value as "below" | "above")} style={dialogInput}>
              <option value="below">Below</option>
              <option value="above">Above</option>
            </select>
          </label>
          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <button
              style={pillBtn}
              onClick={() => {
                if (api?.insertCaption(label, text.trim(), position)) {
                  setText("");
                  setOpen(false);
                }
              }}
            >
              Insert
            </button>
          </div>
        </div>
      )}
    </span>
  );
}

function EquationMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [linear, setLinear] = useState("x={-b±√{b^2-4ac}}/{2a}");
  const [error, setError] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    if (!open) return;
    inputRef.current?.focus();
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const submit = () => {
    if (api?.insertEquation(linear)) {
      setError("");
      setOpen(false);
    } else {
      setError("Place the caret in editable text and enter a valid equation.");
    }
  };
  const anchor = open ? rootRef.current?.getBoundingClientRect() : null;
  const viewportWidth = typeof window === "undefined" ? 356 : window.innerWidth;
  const popoverWidth = Math.min(340, viewportWidth - 16);
  const popoverLeft = Math.max(8, Math.min(anchor?.left ?? 8, viewportWidth - popoverWidth - 8));
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert equation" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        <span style={{ fontFamily: "'Cambria Math', serif", fontSize: 18 }}>π</span>
        <span style={{ marginLeft: 5 }}>Equation</span>
      </button>
      {open && (
        <div data-dxw-equation-menu="" style={{ position: "fixed", top: anchor?.bottom ?? 28, left: popoverLeft, zIndex: 100, width: popoverWidth, boxSizing: "border-box", padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <div style={{ font: "600 12px system-ui, sans-serif", marginBottom: 5, color: T.fg }}>Linear equation</div>
          <input
            ref={inputRef}
            aria-label="Linear equation"
            value={linear}
            onChange={(event) => { setLinear(event.target.value); setError(""); }}
            onKeyDown={(event) => event.key === "Enter" && submit()}
            style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "15px 'Cambria Math', serif", outline: "none" }}
          />
          <div style={{ color: T.muted, fontSize: 11.5, marginTop: 5 }}>Use ^, _, /, √&#123;…&#125;, ∫, matrices [a&amp;b;c&amp;d], and grouped &#123;…&#125; expressions.</div>
          {error && <div style={{ color: "#c5221f", fontSize: 11.5, marginTop: 5 }}>{error}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
            <button style={pillBtn} disabled={!linear.trim()} onClick={submit}>Insert</button>
          </div>
        </div>
      )}
    </span>
  );
}

/**
 * Table of contents, and the two update passes that keep a document's field
 * results current.
 *
 * Word puts these on a References tab this ribbon does not have. A TOC IS a
 * field, so it lands beside the other field controls — the same placement
 * INSERT_COMMANDS already records for `insertToc`.
 *
 * The two updates are different jobs and both are offered: refreshTocs
 * REBUILDS the entry paragraphs from the document's current headings, while
 * updateFields only recomputes every field's result (page numbers, dates,
 * cross-references) against the current layout.
 */
function ContentsMenu({ api }: { api: DocxViewApi | null }) {
  return (
    <ActionMenu
      label="Contents"
      title="Insert or update a table of contents"
      width={92}
      groups={[
        {
          items: [
            ["insert", "Table of contents"],
            ["figures", "Table of figures"],
            ["rebuild", "Update table of contents"],
            ["fields", "Update fields"],
          ],
        },
      ]}
      onPick={(value) => {
        if (value === "insert") api?.insertToc();
        else if (value === "figures") api?.insertToc({ captionLabel: "Figure" });
        else if (value === "rebuild") api?.refreshTocs();
        else api?.updateFields();
      }}
    />
  );
}

/**
 * The References citations cluster: insert a citation (picker over the
 * document's sources), manage sources (new / edit / delete), pick the
 * citation style, and insert the generated bibliography. One popover in the
 * chart-data-editor idiom — the source form is the only multi-field surface
 * the cluster needs.
 */
function CitationsMenu({ api }: { api: DocxViewApi | null }) {
  type Source = ReturnType<NonNullable<typeof api>["listCitationSources"]>[number];
  const SOURCE_TYPES = [
    ["book", "Book"],
    ["article", "Journal article"],
    ["website", "Website"],
    ["report", "Report"],
  ] as const;
  type SourceType = (typeof SOURCE_TYPES)[number][0];
  const [open, setOpen] = useState(false);
  const [sources, setSources] = useState<Source[]>([]);
  const [style, setStyle] = useState("APA");
  /** null = list view; "" = new-source form; a tag = edit that source. */
  const [editing, setEditing] = useState<string | null>(null);
  const [type, setType] = useState<SourceType>("book");
  const [tag, setTag] = useState("");
  const [authors, setAuthors] = useState("");
  const [corporate, setCorporate] = useState("");
  const [title, setTitle] = useState("");
  const [year, setYear] = useState("");
  const [container, setContainer] = useState("");
  const [error, setError] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);

  const refresh = () => {
    setSources(api?.listCitationSources() ?? []);
    setStyle(api?.getCitationStyle() ?? "APA");
  };
  const toggle = () => {
    if (open) {
      setOpen(false);
      return;
    }
    refresh();
    setEditing(null);
    setError("");
    setOpen(true);
  };

  const containerLabel = type === "article" ? "Journal" : type === "website" ? "URL" : "Publisher";
  const parseAuthors = (text: string): { last: string; first?: string }[] =>
    text
      .split(";")
      .map((part) => part.trim())
      .filter(Boolean)
      .map((part) => {
        const [last, first] = part.split(",").map((half) => half.trim());
        return first ? { last, first } : { last };
      })
      .filter((person) => person.last);
  const defaultTag = (): string => {
    const stem = (parseAuthors(authors)[0]?.last ?? corporate ?? title ?? "Source")
      .replace(/[^A-Za-z0-9_-]/g, "")
      .slice(0, 20) || "Source";
    const base = `${stem}${year.slice(-2)}`;
    let candidate = base;
    for (let n = 2; sources.some((source) => source.tag === candidate); n++) candidate = `${base}_${n}`;
    return candidate;
  };
  const openForm = (source: Source | null) => {
    setEditing(source?.tag ?? "");
    setType(
      source?.sourceType === "JournalArticle" ? "article"
        : source?.sourceType === "InternetSite" ? "website"
          : source?.sourceType === "Report" ? "report" : "book",
    );
    setTag(source?.tag ?? "");
    setAuthors((source?.authors ?? []).map((person) => (person.first ? `${person.last}, ${person.first}` : person.last)).join("; "));
    setCorporate(source?.corporate ?? "");
    setTitle(source?.title ?? "");
    setYear(source?.year ?? "");
    setContainer(source?.journal ?? source?.publisher ?? source?.url ?? "");
    setError("");
  };
  const submitForm = () => {
    const fields = {
      type,
      authors: parseAuthors(authors),
      corporate: parseAuthors(authors).length > 0 ? "" : corporate.trim(),
      title: title.trim(),
      year: year.trim(),
      publisher: type === "book" || type === "report" ? container.trim() : "",
      journal: type === "article" ? container.trim() : "",
      url: type === "website" ? container.trim() : "",
    };
    const done = editing
      ? api?.editCitationSource(editing, fields)
      : api?.createCitationSource({ tag: tag.trim() || defaultTag(), ...fields });
    if (!done) {
      setError(editing ? "Could not update the source." : "Could not add the source — is the tag already in use?");
      return;
    }
    setEditing(null);
    refresh();
  };
  const remove = (sourceTag: string) => {
    if (api?.deleteCitationSource(sourceTag)) {
      setError("");
      refresh();
    } else {
      setError("The source is cited in the document. Remove its citations first.");
    }
  };

  const anchor = open ? rootRef.current?.getBoundingClientRect() : null;
  const viewportWidth = typeof window === "undefined" ? 396 : window.innerWidth;
  const popoverWidth = Math.min(380, viewportWidth - 16);
  const popoverLeft = Math.max(8, Math.min(anchor?.left ?? 8, viewportWidth - popoverWidth - 8));
  const fieldStyle: React.CSSProperties = { width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "13px system-ui, sans-serif", color: T.fg, background: T.popoverBg };
  const fieldLabelStyle: React.CSSProperties = { display: "grid", gap: 3, color: T.muted, font: "11px system-ui, sans-serif" };
  const rowBtn: React.CSSProperties = { border: `1px solid ${T.border}`, borderRadius: 5, background: T.popoverBg, color: T.fg, cursor: "pointer", font: "12px system-ui, sans-serif", padding: "3px 8px" };

  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Citations and bibliography" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={toggle}>
        Citations
      </button>
      {open && (
        <div data-dxw-citations-menu="" style={{ position: "fixed", top: anchor?.bottom ?? 28, left: popoverLeft, zIndex: 100, width: popoverWidth, maxHeight: "calc(100vh - 48px)", overflow: "auto", boxSizing: "border-box", padding: 10, display: "grid", gap: 8, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <strong style={{ color: T.fg, font: "600 13px system-ui, sans-serif" }}>Citations</strong>
          <label style={{ ...fieldLabelStyle, gridTemplateColumns: "auto 1fr", alignItems: "center", display: "grid", gap: 6 }}>
            <span>Style</span>
            <select
              aria-label="Citation style"
              value={style === "MLA" ? "MLA" : "APA"}
              onChange={(event) => {
                const next = event.target.value as "APA" | "MLA";
                api?.setCitationStyle(next);
                refresh();
              }}
              style={fieldStyle}
            >
              <option value="APA">APA</option>
              <option value="MLA">MLA</option>
            </select>
          </label>
          {editing === null ? (
            <>
              <div role="list" aria-label="Bibliography sources" style={{ display: "grid", gap: 5 }}>
                {sources.length === 0 && (
                  <div style={{ color: T.muted, font: "12px system-ui, sans-serif" }}>No sources yet. Add one to cite it.</div>
                )}
                {sources.map((source) => (
                  <div key={source.tag} role="listitem" style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: T.fg, font: "12px system-ui, sans-serif" }} title={source.title ?? source.tag}>
                      {source.title || source.tag}
                      {source.year ? ` (${source.year})` : ""}
                    </span>
                    <button type="button" title={`Cite ${source.tag}`} style={rowBtn} onClick={() => { if (api?.insertCitation(source.tag)) setOpen(false); }}>Cite</button>
                    <button type="button" title={`Edit ${source.tag}`} style={rowBtn} onClick={() => openForm(source)}>Edit</button>
                    <button type="button" title={`Delete ${source.tag}`} style={rowBtn} onClick={() => remove(source.tag)}>✕</button>
                  </div>
                ))}
              </div>
              {error && <div role="alert" style={{ color: "#c5221f", fontSize: 11.5 }}>{error}</div>}
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", borderTop: `1px solid ${T.border}`, paddingTop: 8 }}>
                <button type="button" style={rowBtn} onClick={() => openForm(null)}>New source</button>
                <button type="button" style={pillBtn} onClick={() => { if (api?.insertBibliography()) setOpen(false); }}>Insert bibliography</button>
              </div>
            </>
          ) : (
            <div role="group" aria-label={editing ? "Edit source" : "New source"} style={{ display: "grid", gap: 7 }}>
              <strong style={{ color: T.fg, font: "600 11.5px system-ui, sans-serif" }}>{editing ? `Edit source (${editing})` : "New source"}</strong>
              <label style={fieldLabelStyle}>
                <span>Type</span>
                <select aria-label="Source type" value={type} onChange={(event) => setType(event.target.value as SourceType)} style={fieldStyle}>
                  {SOURCE_TYPES.map(([value, label]) => (
                    <option key={value} value={value}>{label}</option>
                  ))}
                </select>
              </label>
              <label style={fieldLabelStyle}>
                <span>Authors — Last, First; Last, First</span>
                <input aria-label="Authors" value={authors} onChange={(event) => setAuthors(event.target.value)} placeholder="Doe, Jane; Smith, Ada" style={fieldStyle} />
              </label>
              <label style={fieldLabelStyle}>
                <span>Corporate author (when no person author)</span>
                <input aria-label="Corporate author" value={corporate} onChange={(event) => setCorporate(event.target.value)} style={fieldStyle} />
              </label>
              <label style={fieldLabelStyle}>
                <span>Title</span>
                <input aria-label="Title" value={title} onChange={(event) => setTitle(event.target.value)} style={fieldStyle} />
              </label>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 2fr", gap: 6 }}>
                <label style={fieldLabelStyle}>
                  <span>Year</span>
                  <input aria-label="Year" value={year} onChange={(event) => setYear(event.target.value)} style={fieldStyle} />
                </label>
                <label style={fieldLabelStyle}>
                  <span>{containerLabel}</span>
                  <input aria-label={containerLabel} value={container} onChange={(event) => setContainer(event.target.value)} style={fieldStyle} />
                </label>
              </div>
              {!editing && (
                <label style={fieldLabelStyle}>
                  <span>Tag (blank = automatic)</span>
                  <input aria-label="Source tag" value={tag} onChange={(event) => setTag(event.target.value)} placeholder={defaultTag()} style={fieldStyle} />
                </label>
              )}
              {error && <div role="alert" style={{ color: "#c5221f", fontSize: 11.5 }}>{error}</div>}
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                <button type="button" style={rowBtn} onClick={() => { setEditing(null); setError(""); }}>Cancel</button>
                <button type="button" style={pillBtn} onClick={submitForm}>{editing ? "Save" : "Add source"}</button>
              </div>
            </div>
          )}
        </div>
      )}
    </span>
  );
}

const SYMBOLS = ["Ω", "±", "×", "÷", "≤", "≥", "≠", "≈", "∞", "∑", "√", "∫", "→", "↔", "©", "®", "™", "€", "£", "¥", "✓", "•", "§", "¶"];

function SymbolMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [custom, setCustom] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const anchor = open ? rootRef.current?.getBoundingClientRect() : null;
  const popoverLeft = Math.max(8, Math.min(anchor?.left ?? 8, (typeof window === "undefined" ? 280 : window.innerWidth) - 272));
  const insertCustom = () => {
    if (api?.insertSymbol(custom)) {
      setCustom("");
      setOpen(false);
    }
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert advanced symbol" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        <span style={{ fontFamily: "serif", fontSize: 14 }}>Ω <span style={{ fontFamily: "system-ui, sans-serif", fontSize: 12 }}>Advanced Symbol</span></span>
      </button>
      {open && (
        <div style={{ position: "fixed", top: anchor?.bottom ?? 28, left: popoverLeft, zIndex: 100, width: 264, padding: 8, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 32px)", gap: 5 }}>
            {SYMBOLS.map((symbol) => (
              <button key={symbol} title={`Insert ${symbol}`} onMouseDown={(event) => event.preventDefault()} onClick={() => { if (api?.insertSymbol(symbol)) setOpen(false); }} style={{ width: 32, height: 30, border: `1px solid ${T.border}`, borderRadius: 5, background: T.popoverBg, color: T.fg, cursor: "pointer", font: "17px 'Cambria Math', serif" }}>
                {symbol}
              </button>
            ))}
          </div>
          <div style={{ borderTop: `1px solid ${T.border}`, marginTop: 8, paddingTop: 8 }}>
            <label style={{ display: "block", color: T.muted, font: "11.5px system-ui, sans-serif", marginBottom: 4 }} htmlFor="dxw-advanced-symbol">Any Unicode symbol</label>
            <div style={{ display: "flex", gap: 6 }}>
              <input
                id="dxw-advanced-symbol"
                aria-label="Advanced symbol characters"
                value={custom}
                onChange={(event) => setCustom(event.target.value)}
                onKeyDown={(event) => event.key === "Enter" && insertCustom()}
                placeholder="Paste or type a symbol"
                style={{ minWidth: 0, flex: 1, border: `1px solid ${T.border}`, borderRadius: 5, padding: "5px 7px", background: T.popoverBg, color: T.fg, font: "15px 'Cambria Math', serif" }}
              />
              <button type="button" disabled={!custom} onClick={insertCustom} style={pillBtn}>Insert</button>
            </div>
          </div>
        </div>
      )}
    </span>
  );
}

/** The picker's Lines/Text row: presets with no gallery geometry of their own. */
const SHAPE_TOOLS = [
  ["line", "Line", "―"],
  ["verticalLine", "Vertical line", "│"],
  ["textBox", "Text box", "T"],
] as const;

/** A gallery icon: the shape's own preset outline evaluated at icon size. */
function ShapeGlyph({ preset }: { preset: string }) {
  const geom = presetShapeGeometry(preset, 26, 20);
  if (!geom) return null;
  return (
    <svg viewBox="-1 -1 28 22" width={26} height={20} aria-hidden="true" style={{ display: "block" }}>
      {geom.paths.map((path, index) => (
        <path key={index} d={path.d} fill="none" stroke="currentColor" strokeWidth={1.2} strokeLinejoin="round" />
      ))}
    </svg>
  );
}

function DividerMenu({ api }: { api: DocxViewApi | null }) {
  type Divider = NonNullable<ReturnType<DocxViewApi["getParagraphDivider"]>>;
  const [open, setOpen] = useState(false);
  const [style, setStyle] = useState<Divider["style"]>("single");
  const [color, setColor] = useState("#000000");
  const [widthPt, setWidthPt] = useState(1);
  const [spacePt, setSpacePt] = useState(1);
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const toggle = () => {
    if (!open) {
      const current = api?.getParagraphDivider();
      if (current) {
        setStyle(current.style);
        setColor(current.color);
        setWidthPt(current.widthPt);
        setSpacePt(current.spacePt);
      }
    }
    setOpen(!open);
  };
  const apply = () => {
    if (api?.setParagraphDivider({ style, color, widthPt, spacePt })) setOpen(false);
  };
  const previewStyle = style === "double" || style === "thinThickSmallGap"
    ? "double"
    : style;
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert or edit divider" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={toggle}>Divider</button>
      {open && (
        <div style={{ position: "absolute", top: 28, right: 0, zIndex: 100, width: 270, padding: 10, display: "grid", gap: 8, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <div style={{ color: T.muted, font: "11.5px system-ui, sans-serif" }}>Horizontal rule below the current paragraph</div>
          <div aria-hidden="true" style={{ height: 12, borderBottom: `${Math.max(widthPt, 1)}px ${previewStyle} ${color}` }} />
          <label style={{ display: "grid", gap: 3, color: T.muted, font: "11.5px system-ui, sans-serif" }}>
            Style
            <select aria-label="Divider style" value={style} onChange={(event) => setStyle(event.target.value as Divider["style"])} style={{ border: `1px solid ${T.border}`, borderRadius: 5, padding: "5px 7px", background: T.popoverBg, color: T.fg }}>
              <option value="single">Single</option>
              <option value="double">Double</option>
              <option value="dashed">Dashed</option>
              <option value="dotted">Dotted</option>
              <option value="thinThickSmallGap">Thin + thick</option>
            </select>
          </label>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 7 }}>
            <label style={{ display: "grid", gap: 3, color: T.muted, font: "11.5px system-ui, sans-serif" }}>
              Color
              <input aria-label="Divider color" type="color" value={color} onChange={(event) => setColor(event.target.value)} style={{ width: "100%", height: 30, border: `1px solid ${T.border}`, borderRadius: 5, background: T.popoverBg }} />
            </label>
            <label style={{ display: "grid", gap: 3, color: T.muted, font: "11.5px system-ui, sans-serif" }}>
              Width (pt)
              <input aria-label="Divider width in points" type="number" min="0.25" step="0.25" value={widthPt} onChange={(event) => setWidthPt(Number(event.target.value))} style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "5px 6px", color: T.fg, background: T.popoverBg }} />
            </label>
            <label style={{ display: "grid", gap: 3, color: T.muted, font: "11.5px system-ui, sans-serif" }}>
              Gap (pt)
              <input aria-label="Divider gap in points" type="number" min="0" step="1" value={spacePt} onChange={(event) => setSpacePt(Number(event.target.value))} style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "5px 6px", color: T.fg, background: T.popoverBg }} />
            </label>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
            <button type="button" onClick={() => { if (api?.setParagraphDivider(null)) setOpen(false); }} style={{ ...pillBtn, background: T.popoverBg, color: T.fg, border: `1px solid ${T.border}` }}>Remove</button>
            <button type="button" onClick={apply} style={pillBtn}>Apply divider</button>
          </div>
        </div>
      )}
    </span>
  );
}

function ShapeMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const [lineColor, setLineColor] = useState("#404040");
  const [lineWidth, setLineWidth] = useState("1.33");
  const [lineDash, setLineDash] = useState<"solid" | "dashed" | "dotted">("solid");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const insert = (preset: Parameters<DocxViewApi["insertShape"]>[0]) => {
    const isLine = preset === "line" || preset === "verticalLine";
    const width = Number(lineWidth);
    if (isLine && (!Number.isFinite(width) || width <= 0)) return;
    if (api?.insertShape(preset, text, isLine ? { color: lineColor, width, dash: lineDash } : undefined)) {
      setText("");
      setOpen(false);
    }
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert shape" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        <span style={{ fontSize: 17 }}>◇</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, right: 0, zIndex: 100, width: 290, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <input
            aria-label="Shape text"
            value={text}
            placeholder="Shape text (optional)"
            onChange={(event) => setText(event.target.value)}
            style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          <div style={{ marginTop: 9, color: T.muted, font: "11.5px system-ui, sans-serif" }}>Line appearance</div>
          <div style={{ display: "grid", gridTemplateColumns: "auto 82px 88px", gap: 6, alignItems: "end", marginTop: 5 }}>
            <ColorMenu
              current={lineColor}
              title="Line color"
              trigger={(
                <>
                  <span style={{ fontSize: 12 }}>Color</span>
                  <span aria-hidden="true" style={{ width: 16, height: 16, borderRadius: 3, border: `1px solid ${T.border}`, background: lineColor }} />
                </>
              )}
              onPick={setLineColor}
            />
            <label style={{ display: "grid", gap: 3, color: T.muted, font: "10.5px system-ui, sans-serif" }}>
              Weight (px)
              <input
                aria-label="Line width in pixels"
                type="number"
                min="0.25"
                step="0.25"
                value={lineWidth}
                onChange={(event) => setLineWidth(event.target.value)}
                style={{ width: "100%", height: 28, boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 5px", color: T.fg, background: T.popoverBg }}
              />
            </label>
            <label style={{ display: "grid", gap: 3, color: T.muted, font: "10.5px system-ui, sans-serif" }}>
              Style
              <select
                aria-label="Line style"
                value={lineDash}
                onChange={(event) => setLineDash(event.target.value as typeof lineDash)}
                style={{ width: "100%", height: 28, border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 4px", color: T.fg, background: T.popoverBg }}
              >
                <option value="solid">Solid</option>
                <option value="dashed">Dashed</option>
                <option value="dotted">Dotted</option>
              </select>
            </label>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 6, marginTop: 8 }}>
            {SHAPE_TOOLS.map(([preset, label, glyph]) => (
              <button key={preset} title={`Insert ${label}`} onClick={() => insert(preset)} style={{ minHeight: 40, border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, color: T.fg, cursor: "pointer", font: "12px system-ui, sans-serif" }}>
                <span style={{ display: "block", fontSize: 17, lineHeight: 1 }}>{glyph}</span>
                <span>{label}</span>
              </button>
            ))}
          </div>
          <div style={{ maxHeight: 320, overflowY: "auto", marginTop: 8, paddingRight: 2 }}>
            {SHAPE_GALLERY.map((section) => (
              <div key={section.category}>
                <div style={{ margin: "7px 0 4px", color: T.muted, font: "600 11px system-ui, sans-serif" }}>{section.category}</div>
                <div role="group" aria-label={section.category} style={{ display: "grid", gridTemplateColumns: "repeat(8, 1fr)", gap: 2 }}>
                  {section.items.map((entry) => (
                    <button
                      key={entry.preset}
                      title={entry.label}
                      aria-label={`Insert ${entry.label}`}
                      onClick={() => insert(entry.preset)}
                      style={{ padding: "4px 2px", border: "1px solid transparent", borderRadius: 4, background: "none", color: T.fg, cursor: "pointer", display: "flex", justifyContent: "center" }}
                      onMouseEnter={(event) => { event.currentTarget.style.borderColor = T.border; event.currentTarget.style.background = T.hoverBg; }}
                      onMouseLeave={(event) => { event.currentTarget.style.borderColor = "transparent"; event.currentTarget.style.background = "none"; }}
                    >
                      <ShapeGlyph preset={entry.preset} />
                    </button>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </span>
  );
}

function TextBoxMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const insert = () => {
    if (api?.insertShape("textBox", text)) {
      setText("");
      setOpen(false);
    }
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert text box" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>Text Box</button>
      {open && (
        <div style={{ position: "absolute", top: 28, right: 0, zIndex: 100, width: 240, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <label htmlFor="dxw-text-box-text" style={{ display: "block", color: T.muted, font: "11.5px system-ui, sans-serif", marginBottom: 4 }}>Initial text</label>
          <input
            id="dxw-text-box-text"
            aria-label="Text box text"
            value={text}
            onChange={(event) => setText(event.target.value)}
            onKeyDown={(event) => event.key === "Enter" && insert()}
            placeholder="Text box"
            style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
            <button type="button" onClick={insert} style={pillBtn}>Insert</button>
          </div>
        </div>
      )}
    </span>
  );
}

/** Insert-tab Header & Footer menu: enter either band, plus Word's two
 * page-variant toggles (different first page / different odd & even). The
 * toggle state is read from the api at render time; picking one re-reads it
 * immediately (this menu re-renders on its own, ahead of the next
 * dxw-selection announcement). */
function HeaderFooterMenu({ api }: { api: DocxViewApi | null }) {
  const [, force] = useReducer((n: number) => n + 1, 0);
  const first = api?.getDifferentFirstPage() ?? false;
  const oddEven = api?.getOddEvenHeaders() ?? false;
  return (
    <ActionMenu
      label="Header & footer"
      title="Edit the repeating header or footer"
      width={118}
      groups={[
        { items: [["header", "Header"], ["footer", "Footer"]] },
        {
          label: "Page setup",
          items: [
            ["first", `${first ? "✓ " : ""}Different first page`],
            ["oddEven", `${oddEven ? "✓ " : ""}Different odd & even pages`],
          ],
        },
      ]}
      onPick={(value) => {
        if (value === "header" || value === "footer") api?.openHeaderFooter(value);
        else if (value === "first") api?.setDifferentFirstPage(!first);
        else if (value === "oddEven") api?.setOddEvenHeaders(!oddEven);
        force();
      }}
    />
  );
}

/** Insert-tab Page Number menu: the two field inserts plus the section's
 * number format (w:pgNumType fmt) and start-at value. */
function PageNumberMenu({ api }: { api: DocxViewApi | null }) {
  const [, force] = useReducer((n: number) => n + 1, 0);
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const current = api?.getPageNumberFormat() ?? { fmt: "decimal" as const, start: null };
  const mark = (fmt: string) => (current.fmt === fmt ? "✓ " : "");
  return (
    <span ref={rootRef} style={{ display: "inline-block" }}>
      <ActionMenu
        label="Page number"
        title="Insert a page number and choose its format"
        width={104}
        groups={[
          { items: [["pn:page", "Page number"], ["pn:pageof", "Page X of Y"]] },
          {
            label: "Number format",
            items: [
              ["fmt:decimal", `${mark("decimal")}1, 2, 3`],
              ["fmt:lowerRoman", `${mark("lowerRoman")}i, ii, iii`],
              ["fmt:upperRoman", `${mark("upperRoman")}I, II, III`],
              ["fmt:lowerLetter", `${mark("lowerLetter")}a, b, c`],
              ["fmt:upperLetter", `${mark("upperLetter")}A, B, C`],
            ],
          },
          {
            label: "Page numbering",
            items: [
              ["start:set", current.start !== null ? `Start at ${current.start}…` : "Start at…"],
              ["start:continue", `${current.start === null ? "✓ " : ""}Continue from previous`],
            ],
          },
        ]}
        onPick={(value) => {
          if (value === "pn:page") api?.insertPageNumber("page");
          else if (value === "pn:pageof") api?.insertPageNumber("pageOfTotal");
          else if (value.startsWith("fmt:")) {
            api?.setPageNumberFormat({ fmt: value.slice(4) as Parameters<NonNullable<typeof api>["setPageNumberFormat"]>[0]["fmt"] });
          } else if (value === "start:continue") api?.setPageNumberFormat({ start: null });
          else if (value === "start:set") {
            const anchor = rootRef.current;
            if (!anchor) return;
            void requestTextInputDialog(anchor, {
              title: "Page numbering",
              label: "Start at",
              value: String(current.start ?? 1),
              submitLabel: "Apply",
              inputType: "number",
              min: 0,
              step: 1,
            }).then((next) => {
              if (next === null) return;
              const start = parseInt(next.trim(), 10);
              if (Number.isInteger(start) && start >= 0) api?.setPageNumberFormat({ start });
              force();
            });
            return;
          }
          force();
        }}
      />
    </span>
  );
}

/** Word's own gallery presets, the four people actually reach for. */
const WATERMARKS = ["CONFIDENTIAL", "DO NOT COPY", "DRAFT", "SAMPLE"] as const;

function WatermarkMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("CONFIDENTIAL");
  const [diagonal, setDiagonal] = useState(true);
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const stamp = (value: string) => {
    if (api?.insertWatermark({ text: value, diagonal })) setOpen(false);
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Watermark" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        <span style={{ color: "#9aa0a6", fontSize: 13, fontWeight: 700, letterSpacing: 0.5 }}>WM</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, right: 0, zIndex: 100, width: 240, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <div style={{ display: "grid", gap: 6 }}>
            {WATERMARKS.map((preset) => (
              <button
                key={preset}
                title={`Stamp ${preset} on every page`}
                onClick={() => stamp(preset)}
                style={{ padding: "7px 8px", border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, color: "#9aa0a6", cursor: "pointer", font: "700 12px system-ui, sans-serif", letterSpacing: 0.5 }}
              >
                {preset}
              </button>
            ))}
          </div>
          <input
            aria-label="Watermark text"
            value={text}
            onChange={(event) => setText(event.target.value)}
            style={{ width: "100%", boxSizing: "border-box", marginTop: 8, border: `1px solid ${T.border}`, borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          <label style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 8, font: "12px system-ui, sans-serif" }}>
            <input type="checkbox" checked={diagonal} onChange={(event) => setDiagonal(event.target.checked)} />
            Diagonal
          </label>
          <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
            <button
              title="Stamp this text on every page"
              disabled={!text}
              onClick={() => stamp(text)}
              style={{ flex: 1, padding: "6px 8px", border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, cursor: text ? "pointer" : "default", font: "600 12px system-ui, sans-serif" }}
            >
              Custom
            </button>
            <button
              title="Remove the watermark from every page"
              onClick={() => {
                api?.removeWatermark();
                setOpen(false);
              }}
              style={{ flex: 1, padding: "6px 8px", border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, cursor: "pointer", font: "600 12px system-ui, sans-serif" }}
            >
              Remove
            </button>
          </div>
        </div>
      )}
    </span>
  );
}

const WORD_ART = [
  ["plain", "Plain", "WordArt"],
  ["archUp", "Arch up", "⌒"],
  ["archDown", "Arch down", "⌣"],
  ["wave", "Wave", "∿"],
  ["chevron", "Chevron", "⌃"],
] as const;

function WordArtMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [text, setText] = useState("Your text here");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert WordArt" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        <span style={{ color: "#2e74b5", fontSize: 17, fontStyle: "italic", fontWeight: 700 }}>A</span>
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, right: 0, zIndex: 100, width: 270, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <input
            aria-label="WordArt text"
            value={text}
            onChange={(event) => setText(event.target.value)}
            style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "5px 8px", font: "13px system-ui, sans-serif", outline: "none" }}
          />
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 6, marginTop: 8 }}>
            {WORD_ART.map(([preset, label, glyph]) => (
              <button
                key={preset}
                title={`Insert WordArt ${label}`}
                disabled={!text}
                onClick={() => {
                  if (api?.insertWordArt(text, preset)) setOpen(false);
                }}
                style={{ minHeight: 48, border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, color: "#2e74b5", cursor: text ? "pointer" : "default", font: "600 12px system-ui, sans-serif" }}
              >
                <span style={{ display: "block", fontSize: 19, lineHeight: 1.2 }}>{glyph}</span>
                <span>{label}</span>
              </button>
            ))}
          </div>
        </div>
      )}
    </span>
  );
}

function ChartMenu({ api, label = "Chart" }: { api: DocxViewApi | null; label?: string }) {
  type Chart = Parameters<DocxViewApi["insertChart"]>[0];
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState(false);
  const [type, setType] = useState<Chart["type"]>("column");
  const [title, setTitle] = useState("");
  const [categories, setCategories] = useState(["", ""]);
  const [series, setSeries] = useState([{ name: "", values: ["", ""] }]);
  const [error, setError] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const submit = () => {
    const categoryValues = categories.map((value) => value.trim());
    const rawSeries = (type === "pie" ? series.slice(0, 1) : series).map((entry) => ({
      name: entry.name.trim(),
      values: entry.values.map((value) => value.trim()),
    }));
    if (categoryValues.some((value) => !value)) {
      setError(type === "pie" ? "Enter a label for every slice." : "Enter a name for every category.");
      return;
    }
    if (type !== "pie" && rawSeries.some((entry) => !entry.name)) {
      setError("Enter a name for every series.");
      return;
    }
    if (rawSeries.some((entry) => entry.values.some((value) => value === "" || !Number.isFinite(Number(value))))) {
      setError("Enter a number in every value field.");
      return;
    }
    if (type === "pie" && rawSeries.some((entry) => entry.values.some((value) => Number(value) < 0))) {
      setError("Pie chart values must be zero or greater.");
      return;
    }
    if (type === "pie" && rawSeries.every((entry) => entry.values.every((value) => Number(value) === 0))) {
      setError("Enter at least one pie chart value greater than zero.");
      return;
    }
    const seriesValues = rawSeries.map((entry) => ({
      name: type === "pie" ? entry.name || "Values" : entry.name,
      values: entry.values.map(Number),
    }));
    const data: Chart = { type, title: title.trim(), categories: categoryValues, series: seriesValues };
    if (api?.updateSelectedChart(data) || api?.insertChart(data)) {
      setError("");
      setOpen(false);
    }
  };
  const addCategory = () => {
    setCategories([...categories, ""]);
    setSeries(series.map((entry) => ({ ...entry, values: [...entry.values, ""] })));
  };
  const removeCategory = (index: number) => {
    if (categories.length === 1) return;
    setCategories(categories.filter((_, itemIndex) => itemIndex !== index));
    setSeries(series.map((entry) => ({ ...entry, values: entry.values.filter((_, itemIndex) => itemIndex !== index) })));
  };
  const changeType = (next: Chart["type"]) => {
    setType(next);
    if (next === "pie") {
      setSeries((current) => [
        current[0] ?? { name: "Values", values: categories.map(() => "") },
      ]);
    }
    setError("");
  };
  const anchor = open ? rootRef.current?.getBoundingClientRect() : null;
  const viewportWidth = typeof window === "undefined" ? 456 : window.innerWidth;
  const popoverWidth = Math.min(440, viewportWidth - 16);
  const popoverLeft = Math.max(8, Math.min(anchor?.left ?? 8, viewportWidth - popoverWidth - 8));
  const fieldStyle: React.CSSProperties = { width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "13px system-ui, sans-serif", color: T.fg, background: T.popoverBg };
  const fieldLabelStyle: React.CSSProperties = { display: "grid", gap: 3, color: T.muted, font: "11px system-ui, sans-serif" };
  const tableHeaderStyle: React.CSSProperties = { padding: "0 3px 4px", textAlign: "left", verticalAlign: "bottom", color: T.muted, font: "600 11px system-ui, sans-serif" };
  const tableCellStyle: React.CSSProperties = { padding: 3, verticalAlign: "bottom" };
  const toggle = () => {
    if (open) {
      setOpen(false);
      return;
    }
    const selected = api?.getSelectedChart();
    const nextCategories = selected?.categories.length ? [...selected.categories] : ["", ""];
    const nextSeries = selected?.series.length
      ? selected.series.map((entry) => ({
        name: entry.name,
        values: nextCategories.map((_, index) => String(entry.values[index] ?? "")),
      }))
      : [{ name: "", values: nextCategories.map(() => "") }];
    setEditing(!!selected);
    setType(selected?.type ?? "column");
    setTitle(selected?.title ?? "");
    setCategories(nextCategories);
    setSeries(selected?.type === "pie" ? nextSeries.slice(0, 1) : nextSeries);
    setError("");
    setOpen(true);
  };
  const chartTypes: { value: Chart["type"]; label: string }[] = [
    { value: "column", label: "Column" },
    { value: "bar", label: "Bar" },
    { value: "line", label: "Line" },
    { value: "pie", label: "Pie" },
  ];
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert or edit chart" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={toggle}>{label}</button>
      {open && (
        <div style={{ position: "fixed", top: anchor?.bottom ?? 28, left: popoverLeft, zIndex: 100, width: popoverWidth, maxHeight: "calc(100vh - 48px)", overflow: "auto", boxSizing: "border-box", padding: 10, display: "grid", gap: 8, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <strong style={{ color: T.fg, font: "600 13px system-ui, sans-serif" }}>{editing ? "Edit chart" : "Insert chart"}</strong>
          <fieldset style={{ minWidth: 0, margin: 0, padding: 0, border: 0 }}>
            <legend style={{ marginBottom: 4, color: T.muted, font: "11px system-ui, sans-serif" }}>Chart type</legend>
            <div role="radiogroup" aria-label="Chart type" style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 5 }}>
              {chartTypes.map((chartType) => {
                const selected = chartType.value === type;
                return (
                  <button
                    key={chartType.value}
                    type="button"
                    role="radio"
                    aria-checked={selected}
                    onClick={() => changeType(chartType.value)}
                    style={{
                      minHeight: 34,
                      border: `1px solid ${selected ? T.accent : T.border}`,
                      borderRadius: 6,
                      background: selected ? T.tabActiveBg : T.popoverBg,
                      color: selected ? T.accent : T.fg,
                      cursor: "pointer",
                      font: "600 12px system-ui, sans-serif",
                    }}
                  >
                    {chartType.label}
                  </button>
                );
              })}
            </div>
          </fieldset>
          <label style={fieldLabelStyle}>
            <span>Chart title (optional)</span>
            <input aria-label="Chart title" value={title} onChange={(event) => setTitle(event.target.value)} style={fieldStyle} />
          </label>
          <div role="group" aria-label="Chart data" style={{ display: "grid", gap: 7 }}>
            <strong style={{ color: T.fg, font: "600 11.5px system-ui, sans-serif" }}>{type === "pie" ? "Slices" : "Chart data"}</strong>
            <div style={{ overflowX: "auto" }}>
              {type === "pie" ? (
                <table aria-label="Pie chart data" style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      <th scope="col" style={tableHeaderStyle}>Slice label</th>
                      <th scope="col" style={tableHeaderStyle}>Value</th>
                      <th scope="col" aria-label="Actions" />
                    </tr>
                  </thead>
                  <tbody>
                    {categories.map((category, index) => (
                      <tr key={index}>
                        <th scope="row" style={{ ...tableCellStyle, width: "52%", fontWeight: 400 }}>
                          <label style={fieldLabelStyle}>
                            <span>Slice {index + 1}</span>
                            <input
                              aria-label={`Chart slice ${index + 1} label`}
                              value={category}
                              onChange={(event) => setCategories(categories.map((value, itemIndex) => itemIndex === index ? event.target.value : value))}
                              style={fieldStyle}
                            />
                          </label>
                        </th>
                        <td style={tableCellStyle}>
                          <input
                            aria-label={`Chart slice ${index + 1} value`}
                            type="number"
                            min="0"
                            step="any"
                            value={series[0]?.values[index] ?? ""}
                            onChange={(event) => setSeries([{
                              ...(series[0] ?? { name: "Values", values: categories.map(() => "") }),
                              values: categories.map((_, itemIndex) => itemIndex === index ? event.target.value : (series[0]?.values[itemIndex] ?? "")),
                            }])}
                            style={fieldStyle}
                          />
                        </td>
                        <td style={tableCellStyle}>
                          {categories.length > 1 && <button type="button" aria-label={`Remove chart slice ${index + 1}`} onClick={() => removeCategory(index)} style={{ ...pillBtn, padding: "0 7px" }}>Remove</button>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <table aria-label={`${type} chart data`} style={{ minWidth: series.length > 2 ? 430 : undefined, width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      <th scope="col" style={tableHeaderStyle}>Category</th>
                      {series.map((entry, seriesIndex) => (
                        <th key={seriesIndex} scope="col" style={tableHeaderStyle}>
                          <label style={fieldLabelStyle}>
                            <span>Series {seriesIndex + 1} name</span>
                            <input
                              aria-label={`Chart series ${seriesIndex + 1} name`}
                              value={entry.name}
                              onChange={(event) => setSeries(series.map((value, itemIndex) => itemIndex === seriesIndex ? { ...value, name: event.target.value } : value))}
                              style={fieldStyle}
                            />
                          </label>
                          {series.length > 1 && <button type="button" aria-label={`Remove chart series ${seriesIndex + 1}`} onClick={() => setSeries(series.filter((_, itemIndex) => itemIndex !== seriesIndex))} style={{ ...pillBtn, marginTop: 3, padding: "0 7px" }}>Remove</button>}
                        </th>
                      ))}
                      <th scope="col" aria-label="Actions" />
                    </tr>
                  </thead>
                  <tbody>
                    {categories.map((category, categoryIndex) => (
                      <tr key={categoryIndex}>
                        <th scope="row" style={{ ...tableCellStyle, minWidth: 120, fontWeight: 400 }}>
                          <label style={fieldLabelStyle}>
                            <span>Category {categoryIndex + 1}</span>
                            <input
                              aria-label={`Chart category ${categoryIndex + 1}`}
                              value={category}
                              onChange={(event) => setCategories(categories.map((value, itemIndex) => itemIndex === categoryIndex ? event.target.value : value))}
                              style={fieldStyle}
                            />
                          </label>
                        </th>
                        {series.map((entry, seriesIndex) => (
                          <td key={seriesIndex} style={{ ...tableCellStyle, minWidth: 82 }}>
                            <input
                              aria-label={`Chart series ${seriesIndex + 1} value ${categoryIndex + 1}`}
                              type="number"
                              step="any"
                              value={entry.values[categoryIndex] ?? ""}
                              onChange={(event) => setSeries(series.map((seriesValue, itemIndex) => itemIndex === seriesIndex ? {
                                ...seriesValue,
                                values: categories.map((_, itemValueIndex) => itemValueIndex === categoryIndex ? event.target.value : (seriesValue.values[itemValueIndex] ?? "")),
                              } : seriesValue))}
                              style={fieldStyle}
                            />
                          </td>
                        ))}
                        <td style={tableCellStyle}>
                          {categories.length > 1 && <button type="button" aria-label={`Remove chart category ${categoryIndex + 1}`} onClick={() => removeCategory(categoryIndex)} style={{ ...pillBtn, padding: "0 7px" }}>Remove</button>}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              {type !== "pie" && <button type="button" onClick={() => setSeries([...series, { name: "", values: categories.map(() => "") }])} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Add series</button>}
              <button type="button" onClick={addCategory} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>{type === "pie" ? "Add slice" : "Add category"}</button>
            </div>
          </div>
          {error && <div role="alert" style={{ color: "#c5221f", font: "11.5px system-ui, sans-serif" }}>{error}</div>}
          <button onClick={submit} style={{ border: 0, borderRadius: 6, padding: "7px 10px", background: T.accent, color: T.accentFg, cursor: "pointer", font: "600 12px system-ui, sans-serif" }}>{editing ? "Update chart" : "Insert chart"}</button>
        </div>
      )}
    </span>
  );
}

function SmartArtMenu({ api, label = "SmartArt" }: { api: DocxViewApi | null; label?: string }) {
  type SmartArt = Parameters<DocxViewApi["insertSmartArt"]>[0];
  const [open, setOpen] = useState(false);
  const [step, setStep] = useState<"layout" | "items">("layout");
  const [editing, setEditing] = useState(false);
  const [layout, setLayout] = useState<SmartArt["layout"]>("process");
  const [items, setItems] = useState([""]);
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const submit = () => {
    const values = items.map((value) => value.trim()).filter(Boolean);
    if (!values.length) return;
    const data: SmartArt = { layout, items: values };
    if (api?.updateSelectedSmartArt(data) || api?.insertSmartArt(data)) setOpen(false);
  };
  const fieldStyle: React.CSSProperties = { width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "13px system-ui, sans-serif", color: T.fg, background: T.popoverBg };
  const toggle = () => {
    if (open) {
      setOpen(false);
      return;
    }
    const selected = api?.getSelectedSmartArt();
    setEditing(!!selected);
    setLayout(selected?.layout ?? "process");
    setItems(selected ? [...selected.items] : [""]);
    setStep(selected ? "items" : "layout");
    setOpen(true);
  };
  const layouts: Array<{ value: SmartArt["layout"]; label: string }> = [
    { value: "process", label: "Process" },
    { value: "cycle", label: "Cycle" },
    { value: "hierarchy", label: "Hierarchy" },
    { value: "list", label: "List" },
  ];
  const preview = (value: SmartArt["layout"]) => {
    const node = { width: 22, height: 13, borderRadius: 3, background: T.accent };
    if (value === "list") return <div style={{ display: "grid", gap: 4 }}>{[0, 1, 2].map((key) => <span key={key} style={{ ...node, width: 72 }} />)}</div>;
    if (value === "hierarchy") return <div style={{ display: "grid", justifyItems: "center", gap: 8 }}><span style={node} /><div style={{ display: "flex", gap: 12 }}><span style={node} /><span style={node} /></div></div>;
    if (value === "cycle") return <div style={{ position: "relative", width: 76, height: 48 }}>{[[27, 0], [53, 18], [27, 35], [1, 18]].map(([left, top], key) => <span key={key} style={{ ...node, position: "absolute", left, top }} />)}</div>;
    return <div style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={node} /><span>→</span><span style={node} /><span>→</span><span style={node} /></div>;
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert or edit SmartArt" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={toggle}>{label}</button>
      {open && createPortal(
        <div style={{ position: "fixed", inset: 0, zIndex: 1000, display: "grid", placeItems: "center", padding: 16, background: "rgba(0,0,0,.34)" }} onMouseDown={(event) => event.target === event.currentTarget && setOpen(false)}>
          <div role="dialog" aria-modal="true" aria-label={editing ? "Edit SmartArt" : "Insert SmartArt"} style={{ width: "min(560px,calc(100vw - 32px))", maxHeight: "calc(100vh - 32px)", overflow: "auto", boxSizing: "border-box", padding: 18, display: "grid", gap: 14, color: T.fg, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 12, boxShadow: T.popoverShadow }}>
            <div>
              <strong style={{ display: "block", font: "600 18px system-ui, sans-serif" }}>{step === "layout" ? "Choose a SmartArt layout" : editing ? "Edit SmartArt" : "Add SmartArt text"}</strong>
              <span style={{ color: T.muted, font: "12px system-ui, sans-serif" }}>{step === "layout" ? "Choose one of the supported layout families." : layouts.find((item) => item.value === layout)?.label}</span>
            </div>
            {step === "layout" ? (
              <div role="group" aria-label="SmartArt layouts" style={{ display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 10 }}>
                {layouts.map((option) => (
                  <button key={option.value} type="button" aria-label={`${option.label} SmartArt`} onClick={() => { setLayout(option.value); setStep("items"); }} style={{ minHeight: 118, display: "grid", placeItems: "center", gap: 9, padding: 12, border: `1px solid ${layout === option.value ? T.accent : T.border}`, borderRadius: 9, background: layout === option.value ? T.activeBg : T.popoverBg, color: T.fg, cursor: "pointer" }}>
                    {preview(option.value)}
                    <span style={{ font: "600 13px system-ui, sans-serif" }}>{option.label}</span>
                  </button>
                ))}
              </div>
            ) : (
              <div role="group" aria-label="SmartArt items" style={{ display: "grid", gap: 7 }}>
                {items.map((item, index) => (
                  <div key={index} style={{ display: "grid", gridTemplateColumns: "1fr auto auto auto", gap: 5 }}>
                    <input aria-label={`SmartArt item ${index + 1}`} value={item} placeholder="Item" onChange={(event) => setItems(items.map((value, itemIndex) => itemIndex === index ? event.target.value : value))} style={fieldStyle} />
                    <button type="button" aria-label={`Move SmartArt item ${index + 1} up`} disabled={index === 0} onClick={() => setItems(items.map((value, itemIndex) => itemIndex === index - 1 ? items[index] : itemIndex === index ? items[index - 1] : value))} style={{ ...pillBtn, padding: "0 8px" }}>↑</button>
                    <button type="button" aria-label={`Move SmartArt item ${index + 1} down`} disabled={index === items.length - 1} onClick={() => setItems(items.map((value, itemIndex) => itemIndex === index ? items[index + 1] : itemIndex === index + 1 ? items[index] : value))} style={{ ...pillBtn, padding: "0 8px" }}>↓</button>
                    <button type="button" aria-label={`Remove SmartArt item ${index + 1}`} disabled={items.length === 1} onClick={() => setItems(items.filter((_, itemIndex) => itemIndex !== index))} style={{ ...pillBtn, padding: "0 8px" }}>×</button>
                  </div>
                ))}
                <button type="button" onClick={() => setItems([...items, ""])} style={{ ...pillBtn, justifySelf: "start", background: T.popoverBg, color: T.fg }}>Add item</button>
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between", gap: 8 }}>
              <div>{step === "items" && <button type="button" onClick={() => setStep("layout")} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Back</button>}</div>
              <div style={{ display: "flex", gap: 7 }}>
                <button type="button" onClick={() => setOpen(false)} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Cancel</button>
                {step === "items" && <button type="button" disabled={!items.some((item) => item.trim())} onClick={submit} style={{ border: 0, borderRadius: 6, padding: "7px 12px", background: items.some((item) => item.trim()) ? T.accent : T.border, color: T.accentFg, cursor: items.some((item) => item.trim()) ? "pointer" : "default", font: "600 12px system-ui, sans-serif" }}>{editing ? "Update" : "Insert"}</button>}
              </div>
            </div>
          </div>
        </div>,
        document.body,
      )}
    </span>
  );
}

function MediaMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [url, setUrl] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert online video" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>Media</button>
      {open && (
        <div style={{ position: "absolute", top: 28, right: 0, zIndex: 100, width: 300, padding: 10, display: "grid", gap: 7, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <input aria-label="Online video URL" type="url" value={url} onChange={(event) => setUrl(event.target.value)} placeholder="https://www.youtube.com/watch?v=…" style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "13px system-ui, sans-serif", color: T.fg, background: T.popoverBg }} />
          <button
            disabled={!url.trim()}
            onClick={() => void api?.insertOnlineVideo(url).then((inserted) => inserted && setOpen(false))}
            style={{ border: 0, borderRadius: 6, padding: "7px 10px", background: url.trim() ? T.accent : T.border, color: T.accentFg, cursor: url.trim() ? "pointer" : "default", font: "600 12px system-ui, sans-serif" }}
          >
            Insert online video
          </button>
        </div>
      )}
    </span>
  );
}

function DrawTab({ api }: { api: DocxViewApi | null }) {
  const [pen, setPen] = useState({ color: "#202124", width: 2 });
  const [highlighter, setHighlighter] = useState({ color: "#F9D949", width: 12 });
  const kindOf = (tool: ReturnType<NonNullable<DocxViewApi["getDrawingTool"]>>) =>
    tool ? tool.kind === "eraser" ? "eraser" : tool.kind === "lasso" ? "lasso" : tool.kind === "highlighter" ? "highlighter" : "pen" : "select";
  const [active, setActive] = useState<"select" | "pen" | "highlighter" | "eraser" | "lasso">(kindOf(api?.getDrawingTool() ?? null));
  useEffect(() => {
    const update = (event: Event) => {
      const tool = (event as CustomEvent<ReturnType<NonNullable<DocxViewApi["getDrawingTool"]>>>).detail;
      setActive(kindOf(tool));
      if (tool?.kind === "highlighter") setHighlighter({ color: tool.color, width: tool.width });
      else if (tool?.kind === "pen") setPen({ color: tool.color, width: tool.width });
    };
    document.addEventListener("dxw-drawing-tool", update);
    return () => document.removeEventListener("dxw-drawing-tool", update);
  }, []);
  useEffect(() => () => api?.setDrawingTool(null), [api]);
  const activate = (kind: "pen" | "highlighter", patch?: Partial<{ color: string; width: number }>) => {
    const next = { ...(kind === "highlighter" ? highlighter : pen), ...patch };
    if (kind === "highlighter") setHighlighter(next);
    else setPen(next);
    api?.setDrawingTool({ kind, color: next.color, width: next.width });
  };
  const inkKind = active === "highlighter" ? "highlighter" : "pen";
  const ink = inkKind === "highlighter" ? highlighter : pen;
  return (
    <>
      <Btn label="Select" title="Select objects" active={active === "select"} onClick={() => api?.setDrawingTool(null)} />
      <Btn label="Pen" title="Draw with pen" active={active === "pen"} onClick={() => activate("pen")} />
      <Btn label="Highlighter" title="Draw with highlighter" active={active === "highlighter"} onClick={() => activate("highlighter")} />
      <Btn label="Eraser" title="Stroke eraser" active={active === "eraser"} onClick={() => api?.setDrawingTool({ kind: "eraser", size: 14 })} />
      <Btn label="Lasso" title="Lasso ink" active={active === "lasso"} onClick={() => api?.setDrawingTool({ kind: "lasso" })} />
      <Sep />
      <ColorMenu
        current={ink.color}
        title={inkKind === "highlighter" ? "Highlighter color" : "Pen color"}
        trigger={(
          <>
            <span style={{ font: "12px system-ui, sans-serif" }}>Color</span>
            <span aria-hidden="true" style={{ width: 18, height: 18, borderRadius: 4, border: `1px solid ${T.border}`, background: ink.color }} />
          </>
        )}
        onPick={(value) => activate(inkKind, { color: value })}
      />
      <ActionMenu
        label={`${ink.width} px`}
        title={inkKind === "highlighter" ? "Highlighter width" : "Pen width"}
        width={70}
        groups={[{ items: [["1", "1 px"], ["2", "2 px"], ["4", "4 px"], ["8", "8 px"], ["12", "12 px"]] }]}
        onPick={(value) => activate(inkKind, { width: Number(value) })}
      />
    </>
  );
}

/** Byte count as the size a person would say out loud ("5 MB", "512 KB"). */
function formatBytes(n: number): string {
  const mb = n / (1024 * 1024);
  if (mb >= 1) return `${Number.isInteger(mb) ? mb : mb.toFixed(1)} MB`;
  const kb = n / 1024;
  return `${Math.max(1, Math.round(kb))} KB`;
}

function ScreenshotButton({ api }: { api: DocxViewApi | null }) {
  const [status, setStatus] = useState("");
  const capture = async () => {
    setStatus("Capturing screenshot…");
    const result = await api?.insertScreenshot();
    setStatus(result === "inserted"
      ? "Screenshot inserted."
      : result === "unsupported"
        ? "Screen capture is not supported in this browser."
        : result === "cancelled"
          ? "Screen capture was cancelled or denied."
          : result === "no-caret"
            ? "Click in the document before inserting a screenshot."
            : "Screenshot failed. Please try again.");
  };
  return (
    <span style={{ position: "relative", display: "inline-flex" }}>
      <Btn label="Screenshot" title="Capture and insert a screen, window, or tab" onClick={() => void capture()} />
      {status && (
        <span
          role={status === "Screenshot inserted." ? "status" : "alert"}
          data-dxw-screenshot-status=""
          style={{ position: "absolute", top: 30, left: 0, zIndex: 120, width: 210, padding: "6px 8px", border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, boxShadow: T.popoverShadow, color: T.fg, font: "12px system-ui, sans-serif" }}
        >
          {status}
        </span>
      )}
    </span>
  );
}

function CoverPageMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [title, setTitle] = useState("");
  const [subtitle, setSubtitle] = useState("");
  const [author, setAuthor] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const insert = () => {
    if (!title.trim() || !api?.insertCoverPage({ title, subtitle, author })) return;
    setOpen(false);
    setTitle("");
    setSubtitle("");
    setAuthor("");
  };
  const input = (label: string, value: string, set: (value: string) => void) => (
    <input
      aria-label={label}
      placeholder={label}
      value={value}
      onChange={(event) => set(event.target.value)}
      style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "13px system-ui, sans-serif" }}
    />
  );
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert cover page" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>Cover page</button>
      {open && (
        <div style={{ position: "absolute", top: 30, left: 0, zIndex: 100, width: 260, padding: 10, display: "grid", gap: 7, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          {input("Cover title", title, setTitle)}
          {input("Cover subtitle", subtitle, setSubtitle)}
          {input("Cover author", author, setAuthor)}
          <button disabled={!title.trim()} onClick={insert} style={{ border: 0, borderRadius: 6, padding: "7px 10px", background: title.trim() ? T.accent : T.border, color: T.accentFg, cursor: title.trim() ? "pointer" : "default", font: "600 12px system-ui, sans-serif" }}>Insert cover</button>
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
    ["mergeRight", "Merge cell right"],
    ["mergeDown", "Merge cell down"],
    ["splitCell", "Split cell"],
    ["valign:top", "Cell align top"],
    ["valign:center", "Cell align middle"],
    ["valign:bottom", "Cell align bottom"],
    ["convert:text", "Convert text to table"],
    ["convert:table", "Convert table to text"],
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
                  } else if (op === "convert:text") {
                    api?.convertTextToTable("tab");
                  } else if (op === "convert:table") {
                    api?.convertTableToText("tab");
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
 * Keep a popover under the control that opened it, and close it on a click
 * outside or on Escape. Returns the fixed position to paint at.
 *
 * The layout tab's three custom dialogs each grew their own copy of this
 * before there was a third; new popovers share this one.
 */
function useAnchoredPopover(
  anchorRef: React.RefObject<HTMLElement | null>,
  rootRef: React.RefObject<HTMLElement | null>,
  width: number,
  onClose: () => void,
): { left: number; top: number } {
  const [position, setPosition] = useState({ left: 8, top: 8 });
  useEffect(() => {
    const place = () => {
      const rect = anchorRef.current?.getBoundingClientRect();
      if (!rect) return;
      const w = Math.min(width, window.innerWidth - 16);
      // Clamped against the popover's OWN height once it has one, so a tall
      // form near the bottom of a short window still fits on screen.
      const height = rootRef.current?.offsetHeight ?? 0;
      setPosition({
        left: Math.max(8, Math.min(rect.left, window.innerWidth - w - 8)),
        top: Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - height - 8)),
      });
    };
    const outside = (event: MouseEvent) => {
      if (rootRef.current?.contains(event.target as Node)) return;
      if (anchorRef.current?.contains(event.target as Node)) return;
      onClose();
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    place();
    document.addEventListener("mousedown", outside);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", place);
    return () => {
      document.removeEventListener("mousedown", outside);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", place);
    };
  }, [anchorRef, rootRef, onClose, width]);
  return position;
}

/**
 * A popover form anchored under the control that opened it, ending with the
 * Cancel/Apply pair every form in this bar ends with.
 */
function AnchoredDialog({
  anchorRef,
  title,
  label,
  width = 236,
  onClose,
  onApply,
  applyDisabled,
  children,
}: {
  anchorRef: React.RefObject<HTMLElement | null>;
  title: string;
  /** Accessible name; defaults to the visible title. */
  label?: string;
  width?: number;
  onClose: () => void;
  onApply: () => void;
  applyDisabled?: boolean;
  children: React.ReactNode;
}) {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const position = useAnchoredPopover(anchorRef, rootRef, width, onClose);
  return (
    <div
      ref={rootRef}
      role="dialog"
      aria-label={label ?? title}
      onMouseDown={(event) => event.stopPropagation()}
      style={{
        position: "fixed", top: position.top, left: position.left, zIndex: 201,
        width: `min(${width}px, calc(100vw - 16px))`, boxSizing: "border-box",
        background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8,
        boxShadow: T.popoverShadow, padding: 10, display: "grid", gap: 7, color: T.fg,
      }}
    >
      <strong style={{ fontSize: 13 }}>{title}</strong>
      {children}
      {/* Marked, because a field inside the form can be a popover with its own
          Cancel/Apply pair (the color menu is): "the last button called Apply"
          is not a safe way to find this one. */}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
        <button type="button" data-dxw-dialog-cancel="" onClick={onClose} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Cancel</button>
        <button type="button" data-dxw-dialog-apply="" onClick={onApply} disabled={applyDisabled} style={pillBtn}>Apply</button>
      </div>
    </div>
  );
}

/** Shared row shape for the labelled fields inside an AnchoredDialog. */
const dialogFieldRow: React.CSSProperties = {
  display: "grid", gridTemplateColumns: "78px 1fr", gap: 8, alignItems: "center", fontSize: 12,
};

const dialogInput: React.CSSProperties = {
  width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`,
  borderRadius: 5, padding: "4px 6px", color: T.fg, background: T.popoverBg,
};

/** A number the user typed, or null when the box is empty or not a number. */
function typedNumber(value: string): number | null {
  const trimmed = value.trim();
  if (trimmed === "") return null;
  const n = Number(trimmed);
  return Number.isFinite(n) ? n : null;
}

/** Points, rounded the way a dialog shows them rather than to full float. */
function showPt(value: number): string {
  return String(Math.round(value * 100) / 100);
}

const MARGIN_SIDES = ["top", "left", "bottom", "right"] as const;

/**
 * Word's Table Properties dialog, for the numbers its ribbon buttons cannot
 * express: an exact table width, one column's width, the default cell
 * margins, and the size of the repeating header band.
 *
 * It PREFILLS from the document (api.getTableProperties) and applies only the
 * values the user actually changed. Applying everything would write a w:tblW
 * onto a table that never had one — and, in suggesting mode, record a tracked
 * change for a property nobody touched.
 */
function TablePropertiesDialog({ api, onChanged }: { api: DocxViewApi | null; onChanged: () => void }) {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  type Form = {
    widthUnit: "auto" | "pt" | "pct";
    widthValue: string;
    columnWidth: string;
    margins: Record<(typeof MARGIN_SIDES)[number], string>;
    headerRows: string;
  };
  const [initial, setInitial] = useState<{ form: Form; columnIdx: number; columnCount: number } | null>(null);
  const [form, setForm] = useState<Form | null>(null);

  const openDialog = () => {
    const info = api?.getTableProperties();
    if (!info) return;
    const next: Form = {
      widthUnit: info.width.unit,
      widthValue: info.width.unit === "auto" ? "" : showPt(info.width.value),
      columnWidth: showPt(info.columnWidthsPt[info.columnIdx] ?? 0),
      margins: {
        top: info.cellMargins.top === undefined ? "" : showPt(info.cellMargins.top),
        left: info.cellMargins.left === undefined ? "" : showPt(info.cellMargins.left),
        bottom: info.cellMargins.bottom === undefined ? "" : showPt(info.cellMargins.bottom),
        right: info.cellMargins.right === undefined ? "" : showPt(info.cellMargins.right),
      },
      headerRows: String(info.headerRows),
    };
    setInitial({ form: next, columnIdx: info.columnIdx, columnCount: info.columnCount });
    setForm(next);
    setOpen(true);
  };

  const close = useCallback(() => setOpen(false), []);

  const widthOk =
    form === null ||
    form.widthUnit === "auto" ||
    (() => {
      const n = typedNumber(form.widthValue);
      return n !== null && n > 0 && (form.widthUnit === "pct" ? n <= 100 : n <= 1584);
    })();
  const columnOk = form === null || (() => {
    const n = typedNumber(form.columnWidth);
    return n !== null && n >= 1 && n <= 1584;
  })();
  const marginsOk =
    form === null ||
    MARGIN_SIDES.every((side) => {
      const raw = form.margins[side];
      if (raw.trim() === "") return true;
      const n = typedNumber(raw);
      return n !== null && n >= 0 && n <= 144;
    });
  const headerOk = form === null || (() => {
    const n = typedNumber(form.headerRows);
    return n !== null && n >= 0 && Number.isInteger(n);
  })();
  const valid = widthOk && columnOk && marginsOk && headerOk;

  const apply = () => {
    if (!form || !initial || !valid) return;
    const was = initial.form;
    if (form.widthUnit !== was.widthUnit || form.widthValue !== was.widthValue) {
      if (form.widthUnit === "auto") api?.setTableWidth("auto");
      else api?.setTableWidth(form.widthUnit, typedNumber(form.widthValue) ?? 0);
    }
    if (form.columnWidth !== was.columnWidth) {
      api?.setTableColumnWidth(initial.columnIdx, typedNumber(form.columnWidth) ?? 0);
    }
    if (MARGIN_SIDES.some((side) => form.margins[side] !== was.margins[side])) {
      const margins: Record<string, number> = {};
      for (const side of MARGIN_SIDES) {
        const n = typedNumber(form.margins[side]);
        if (n !== null) margins[side] = n;
      }
      api?.setTableCellMargins("table", margins);
    }
    if (form.headerRows !== was.headerRows) {
      api?.setTableHeaderRows(typedNumber(form.headerRows) ?? 0);
    }
    setOpen(false);
    onChanged();
  };

  const marginField = (side: (typeof MARGIN_SIDES)[number], label: string) => (
    <label key={side} style={{ display: "grid", gap: 3, fontSize: 11, color: T.muted }}>
      <span>{label}</span>
      <input
        aria-label={`${label} cell margin (points)`}
        type="number"
        min="0"
        step="0.5"
        placeholder="—"
        value={form?.margins[side] ?? ""}
        onChange={(event) =>
          setForm((f) => (f ? { ...f, margins: { ...f.margins, [side]: event.target.value } } : f))
        }
        style={dialogInput}
      />
    </label>
  );

  return (
    <span style={{ display: "contents" }}>
      <Btn
        label="Properties"
        title="Table properties: exact widths, cell margins and header rows"
        active={open}
        buttonRef={triggerRef}
        onClick={() => (open ? close() : openDialog())}
      />
      {open && form && initial && (
        <AnchoredDialog
          anchorRef={triggerRef}
          title="Table Properties"
          width={252}
          onClose={close}
          onApply={apply}
          applyDisabled={!valid}
        >
          <label style={dialogFieldRow}>
            <span>Table width</span>
            <span style={{ display: "flex", gap: 6 }}>
              <select
                aria-label="Table width unit"
                value={form.widthUnit}
                onChange={(event) =>
                  setForm({ ...form, widthUnit: event.target.value as Form["widthUnit"] })
                }
                style={{ ...dialogInput, width: 78 }}
              >
                <option value="auto">Auto</option>
                <option value="pt">Points</option>
                <option value="pct">Percent</option>
              </select>
              <input
                aria-label="Table width"
                type="number"
                min="0"
                step="1"
                disabled={form.widthUnit === "auto"}
                value={form.widthUnit === "auto" ? "" : form.widthValue}
                onChange={(event) => setForm({ ...form, widthValue: event.target.value })}
                style={dialogInput}
              />
            </span>
          </label>
          <label style={dialogFieldRow}>
            <span>{`Column ${initial.columnIdx + 1} of ${initial.columnCount}`}</span>
            <input
              aria-label="Column width (points)"
              type="number"
              min="1"
              step="1"
              value={form.columnWidth}
              onChange={(event) => setForm({ ...form, columnWidth: event.target.value })}
              style={dialogInput}
            />
          </label>
          <span style={{ fontSize: 12 }}>Cell margins (points)</span>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 6 }}>
            {marginField("top", "Top")}
            {marginField("left", "Left")}
            {marginField("bottom", "Bottom")}
            {marginField("right", "Right")}
          </div>
          <label style={dialogFieldRow}>
            <span>Header rows</span>
            <input
              aria-label="Repeating header rows"
              type="number"
              min="0"
              step="1"
              value={form.headerRows}
              onChange={(event) => setForm({ ...form, headerRows: event.target.value })}
              style={dialogInput}
            />
          </label>
          <span style={{ color: T.muted, fontSize: 11 }}>
            Only the boxes you change are written. A blank margin keeps the table's current one.
          </span>
        </AnchoredDialog>
      )}
    </span>
  );
}

/** Word's border weights, in points. w:sz counts eighths of a point. */
const BORDER_WIDTHS_PT = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];

/** Display names for the border styles the engine writes. */
const BORDER_STYLE_NAMES: Record<TableBorderStyle, string> = {
  single: "Single",
  thick: "Thick",
  double: "Double",
  dotted: "Dotted",
  dashed: "Dashed",
  dotDash: "Dot dash",
  dotDotDash: "Dot dot dash",
  thinThickSmallGap: "Thin then thick",
  triple: "Triple",
  wave: "Wave",
  none: "None (suppress)",
};

const EDGE_NAMES: Record<TableBorderEdge, string> = {
  top: "Top",
  bottom: "Bottom",
  left: "Left",
  right: "Right",
  insideH: "Inside horizontal",
  insideV: "Inside vertical",
  tl2br: "Diagonal ↘",
  tr2bl: "Diagonal ↗",
};

/**
 * The full border editor behind the Borders menu's presets: any of the
 * engine's line styles, any weight w:sz can carry, any color, on whichever
 * edges the chosen scope allows.
 *
 * The scope decides the edge list rather than the user filtering it, because
 * setTableBorders REFUSES an edge the scope does not own (a table has no
 * diagonals; a cell has no inside rules) and a checkbox that silently does
 * nothing is worse than an absent one.
 */
function CustomBorderDialog({
  api,
  anchorRef,
  onClose,
  onChanged,
}: {
  api: DocxViewApi | null;
  anchorRef: React.RefObject<HTMLElement | null>;
  onClose: () => void;
  onChanged: () => void;
}) {
  const [scope, setScope] = useState<"table" | "cell">("table");
  const [style, setStyle] = useState<TableBorderStyle>("single");
  const [widthPt, setWidthPt] = useState("0.5");
  const [color, setColor] = useState("#000000");
  const [edges, setEdges] = useState<TableBorderEdge[]>(["top", "bottom", "left", "right"]);

  const allowed = scope === "cell" ? CELL_SCOPE_EDGES : TABLE_SCOPE_EDGES;
  const chosen = edges.filter((edge) => allowed.includes(edge));
  const width = typedNumber(widthPt);
  const validColor = normalizedColor(color);
  const valid =
    chosen.length > 0 &&
    validColor !== null &&
    (style === "none" || (width !== null && width >= 0.125 && width <= 12));

  const apply = () => {
    if (!valid) return;
    api?.setTableBorders(
      scope,
      [...chosen],
      style === "none"
        ? { style: "none" }
        : {
            style,
            // w:sz is eighths of a point, and the engine takes it in that
            // unit; the box asks for points because that is what Word's
            // weight list shows.
            sz: Math.min(96, Math.max(1, Math.round((width ?? 0.5) * 8))),
            color: validColor!.slice(1).toUpperCase(),
          },
    );
    onClose();
    onChanged();
  };

  return (
    <AnchoredDialog
      anchorRef={anchorRef}
      title="Custom Border"
      width={244}
      onClose={onClose}
      onApply={apply}
      applyDisabled={!valid}
    >
      <label style={dialogFieldRow}>
        <span>Apply to</span>
        <select
          aria-label="Border scope"
          value={scope}
          onChange={(event) => setScope(event.target.value as "table" | "cell")}
          style={dialogInput}
        >
          <option value="table">Whole table</option>
          <option value="cell">This cell</option>
        </select>
      </label>
      <label style={dialogFieldRow}>
        <span>Style</span>
        <select
          aria-label="Border style"
          value={style}
          onChange={(event) => setStyle(event.target.value as TableBorderStyle)}
          style={dialogInput}
        >
          {TABLE_BORDER_STYLES.map((value) => (
            <option key={value} value={value}>{BORDER_STYLE_NAMES[value]}</option>
          ))}
        </select>
      </label>
      <label style={dialogFieldRow}>
        <span>Weight</span>
        <select
          aria-label="Border width (points)"
          value={widthPt}
          disabled={style === "none"}
          onChange={(event) => setWidthPt(event.target.value)}
          style={dialogInput}
        >
          {BORDER_WIDTHS_PT.map((pt) => (
            <option key={pt} value={String(pt)}>{`${pt} pt`}</option>
          ))}
        </select>
      </label>
      <span style={dialogFieldRow}>
        <span>Color</span>
        <ColorMenu
          current={validColor ?? "#000000"}
          title="Border color"
          trigger={(
            <>
              <span
                aria-hidden="true"
                style={{ width: 14, height: 14, borderRadius: 3, border: `1px solid ${T.border}`, background: validColor ?? "#000000" }}
              />
              {validColor ?? "#000000"}
            </>
          )}
          onPick={setColor}
        />
      </span>
      <span style={{ fontSize: 12 }}>Edges</span>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 2 }}>
        {allowed.map((edge) => (
          <label key={edge} style={{ display: "flex", gap: 5, alignItems: "center", fontSize: 11.5 }}>
            <input
              type="checkbox"
              aria-label={EDGE_NAMES[edge]}
              checked={chosen.includes(edge)}
              onChange={(event) =>
                setEdges((current) =>
                  event.target.checked
                    ? [...current, edge]
                    : current.filter((other) => other !== edge),
                )
              }
            />
            {EDGE_NAMES[edge]}
          </label>
        ))}
      </div>
    </AnchoredDialog>
  );
}

const TAB_ALIGN_NAMES: Record<TabStopSpec["align"], string> = {
  left: "Left",
  center: "Center",
  right: "Right",
  decimal: "Decimal",
  bar: "Bar",
};

const TAB_LEADER_NAMES: Record<TabStopSpec["leader"], string> = {
  none: "None",
  dot: "Dots ….",
  hyphen: "Hyphens ---",
  underscore: "Underline ___",
  middleDot: "Middle dots ···",
};

/**
 * Word's Tabs dialog, popover-sized: the paragraph's direct tab stops as
 * editable rows (position in points, alignment, leader), plus add and
 * remove. Apply replaces the whole list; applying an empty list clears
 * w:tabs so style stops and the default grid take over again.
 */
function TabStopsDialog({
  api,
  anchorRef,
  onClose,
}: {
  api: DocxViewApi | null;
  anchorRef: React.RefObject<HTMLElement | null>;
  onClose: () => void;
}) {
  const [rows, setRows] = useState<{ pos: string; align: TabStopSpec["align"]; leader: TabStopSpec["leader"] }[]>(
    () => (api?.getTabStops() ?? []).map((stop) => ({ pos: showPt(stop.posPt), align: stop.align, leader: stop.leader })),
  );
  const parsed = rows.map((row) => ({ ...row, posPt: typedNumber(row.pos) }));
  const valid = parsed.every((row) => row.posPt !== null && row.posPt >= 0 && row.posPt <= 1584);
  const patch = (index: number, part: Partial<(typeof rows)[number]>) =>
    setRows((current) => current.map((row, i) => (i === index ? { ...row, ...part } : row)));
  const apply = () => {
    if (!valid) return;
    api?.setTabStops(parsed.map((row) => ({ posPt: row.posPt!, align: row.align, leader: row.leader })));
    onClose();
  };
  return (
    <AnchoredDialog
      anchorRef={anchorRef}
      title="Tab stops"
      width={288}
      onClose={onClose}
      onApply={apply}
      applyDisabled={!valid}
    >
      {rows.length === 0 && (
        <span style={{ fontSize: 12, color: T.muted }}>
          No custom stops — tabs use the default half-inch grid.
        </span>
      )}
      {rows.map((row, index) => (
        <div key={index} style={{ display: "grid", gridTemplateColumns: "64px 1fr 1fr 22px", gap: 6, alignItems: "center" }}>
          <input
            aria-label={`Tab stop ${index + 1} position (points)`}
            type="number"
            min={0}
            max={1584}
            step={1}
            value={row.pos}
            onChange={(event) => patch(index, { pos: event.target.value })}
            style={dialogInput}
          />
          <select
            aria-label={`Tab stop ${index + 1} alignment`}
            value={row.align}
            onChange={(event) => patch(index, { align: event.target.value as TabStopSpec["align"] })}
            style={dialogInput}
          >
            {(Object.keys(TAB_ALIGN_NAMES) as TabStopSpec["align"][]).map((align) => (
              <option key={align} value={align}>{TAB_ALIGN_NAMES[align]}</option>
            ))}
          </select>
          <select
            aria-label={`Tab stop ${index + 1} leader`}
            value={row.leader}
            onChange={(event) => patch(index, { leader: event.target.value as TabStopSpec["leader"] })}
            style={dialogInput}
          >
            {(Object.keys(TAB_LEADER_NAMES) as TabStopSpec["leader"][]).map((leader) => (
              <option key={leader} value={leader}>{TAB_LEADER_NAMES[leader]}</option>
            ))}
          </select>
          <button
            type="button"
            aria-label={`Remove tab stop ${index + 1}`}
            title="Remove this stop"
            onClick={() => setRows((current) => current.filter((_, i) => i !== index))}
            style={{ ...pillBtn, background: T.popoverBg, color: T.fg, padding: "2px 6px" }}
          >
            ×
          </button>
        </div>
      ))}
      <div>
        <button
          type="button"
          onClick={() =>
            setRows((current) => {
              const last = typedNumber(current[current.length - 1]?.pos ?? "");
              return [...current, { pos: showPt((last ?? 0) + 36), align: "left", leader: "none" }];
            })
          }
          style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}
        >
          Add stop
        </button>
      </div>
    </AnchoredDialog>
  );
}

const PARA_EDGE_NAMES: Record<"top" | "left" | "bottom" | "right" | "between", string> = {
  top: "Top",
  left: "Left",
  bottom: "Bottom",
  right: "Right",
  between: "Between paragraphs",
};

/**
 * Word's Borders and Shading dialog for PARAGRAPHS, popover-sized. The
 * style/weight/color rows are the table border picker's vocabulary; edges
 * are the paragraph's own (top/left/bottom/right/between), and the shading
 * row drives w:shd.
 */
function ParagraphBorderDialog({
  api,
  anchorRef,
  onClose,
}: {
  api: DocxViewApi | null;
  anchorRef: React.RefObject<HTMLElement | null>;
  onClose: () => void;
}) {
  const [style, setStyle] = useState<TableBorderStyle>("single");
  const [widthPt, setWidthPt] = useState("0.5");
  const [color, setColor] = useState("#000000");
  const [edges, setEdges] = useState<(keyof typeof PARA_EDGE_NAMES)[]>(["top", "bottom", "left", "right"]);
  const [shading, setShading] = useState<string | null>(() => api?.getParagraphBorders().shading ?? null);

  const width = typedNumber(widthPt);
  const validColor = normalizedColor(color);
  const valid =
    validColor !== null &&
    (style === "none" || (width !== null && width >= 0.125 && width <= 12));

  const apply = () => {
    if (!valid) return;
    const spec =
      style === "none"
        ? { style: "none" as const }
        : {
            style,
            sz: Math.min(96, Math.max(1, Math.round((width ?? 0.5) * 8))),
            color: validColor!.slice(1).toUpperCase(),
          };
    const borders = Object.fromEntries(edges.map((edge) => [edge, spec]));
    api?.setParagraphBorders({
      ...(edges.length > 0 ? { borders } : {}),
      shading: shading === null ? null : shading.replace(/^#/, "").toUpperCase(),
    });
    onClose();
  };

  return (
    <AnchoredDialog
      anchorRef={anchorRef}
      title="Paragraph borders"
      width={252}
      onClose={onClose}
      onApply={apply}
      applyDisabled={!valid}
    >
      <label style={dialogFieldRow}>
        <span>Style</span>
        <select
          aria-label="Paragraph border style"
          value={style}
          onChange={(event) => setStyle(event.target.value as TableBorderStyle)}
          style={dialogInput}
        >
          {TABLE_BORDER_STYLES.map((value) => (
            <option key={value} value={value}>{BORDER_STYLE_NAMES[value]}</option>
          ))}
        </select>
      </label>
      <label style={dialogFieldRow}>
        <span>Weight</span>
        <select
          aria-label="Paragraph border width (points)"
          value={widthPt}
          disabled={style === "none"}
          onChange={(event) => setWidthPt(event.target.value)}
          style={dialogInput}
        >
          {BORDER_WIDTHS_PT.map((pt) => (
            <option key={pt} value={String(pt)}>{`${pt} pt`}</option>
          ))}
        </select>
      </label>
      <span style={dialogFieldRow}>
        <span>Color</span>
        <ColorMenu
          current={validColor ?? "#000000"}
          title="Paragraph border color"
          trigger={(
            <>
              <span
                aria-hidden="true"
                style={{ width: 14, height: 14, borderRadius: 3, border: `1px solid ${T.border}`, background: validColor ?? "#000000" }}
              />
              {validColor ?? "#000000"}
            </>
          )}
          onPick={setColor}
        />
      </span>
      <span style={{ fontSize: 12 }}>Edges</span>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 2 }}>
        {(Object.keys(PARA_EDGE_NAMES) as (keyof typeof PARA_EDGE_NAMES)[]).map((edge) => (
          <label key={edge} style={{ display: "flex", gap: 5, alignItems: "center", fontSize: 11.5 }}>
            <input
              type="checkbox"
              aria-label={PARA_EDGE_NAMES[edge]}
              checked={edges.includes(edge)}
              onChange={(event) =>
                setEdges((list) =>
                  event.target.checked ? [...list, edge] : list.filter((other) => other !== edge),
                )
              }
            />
            {PARA_EDGE_NAMES[edge]}
          </label>
        ))}
      </div>
      <span style={dialogFieldRow}>
        <span>Shading</span>
        <span style={{ display: "inline-flex", gap: 6, alignItems: "center" }}>
          <ColorMenu
            current={shading ?? "#FFFF99"}
            title="Paragraph shading fill"
            trigger={(
              <span
                aria-hidden="true"
                style={{
                  width: 14, height: 14, borderRadius: 3, border: `1px solid ${T.border}`,
                  background: shading ?? "linear-gradient(to top left, #fff 46%, #d93025 49%, #d93025 51%, #fff 54%)",
                }}
              />
            )}
            onPick={setShading}
          />
          {shading !== null && (
            <button
              type="button"
              onClick={() => setShading(null)}
              style={{ ...pillBtn, background: T.popoverBg, color: T.fg, padding: "2px 8px" }}
            >
              No fill
            </button>
          )}
        </span>
      </span>
    </AnchoredDialog>
  );
}

/** Menu value standing for "remove the style reference". */
const NO_TABLE_STYLE = "(no table style)";

function TableFormatTab({
  api,
  fill,
  onChanged,
}: {
  api: DocxViewApi | null;
  fill: string | null;
  onChanged: () => void;
}) {
  const run = (op: Parameters<DocxViewApi["tableOp"]>[0]) => {
    api?.tableOp(op);
    onChanged();
  };
  const after = (act: () => void) => {
    act();
    onChanged();
  };
  const [customBorder, setCustomBorder] = useState(false);
  const borderMenuRef = useRef<HTMLSpanElement | null>(null);
  const closeCustomBorder = useCallback(() => setCustomBorder(false), []);
  // Word's "No Borders" and its eraser are different edits, and the engine
  // keeps them apart: SUPPRESS writes w:val="nil" so no rule is drawn even
  // when the table style asks for one, while CLEAR removes the direct edges
  // so the style's rules come back. Both are offered rather than guessing.
  const ALL_TABLE_EDGES = ["top", "bottom", "left", "right", "insideH", "insideV"] as const;
  const borderActions: Record<string, () => void> = {
    tableAll: () => api?.setTableBorders("table", [...ALL_TABLE_EDGES], { style: "single", sz: 4 }),
    tableOutside: () => api?.setTableBorders("table", ["top", "bottom", "left", "right"], { style: "single", sz: 4 }),
    tableInside: () => api?.setTableBorders("table", ["insideH", "insideV"], { style: "single", sz: 4 }),
    tableNone: () => api?.setTableBorders("table", [...ALL_TABLE_EDGES], { style: "none" }),
    tableClear: () => api?.setTableBorders("table", [...ALL_TABLE_EDGES], null),
    cellAll: () => api?.setTableBorders("cell", ["top", "bottom", "left", "right"], { style: "single", sz: 4 }),
    cellNone: () => api?.setTableBorders("cell", ["top", "bottom", "left", "right"], { style: "none" }),
    cellClear: () => api?.setTableBorders("cell", ["top", "bottom", "left", "right"], null),
  };
  const look = api?.getTableLook();
  const styleId = api?.getTableStyleId() ?? null;
  const tick = (on: boolean | undefined, text: string) => `${on ? "\u2713 " : "\u2007\u2007"}${text}`;
  return (
    <span data-dxw-table-format="" style={{ display: "contents" }}>
      <ColorMenu
        current={fill ?? "#FFFFFF"}
        title="Cell fill color"
        trigger={(
          <>
            <span aria-hidden="true" style={{ width: 14, height: 14, borderRadius: 3, border: `1px solid ${T.border}`, background: fill ?? "#FFFFFF" }} />
            Cell fill
          </>
        )}
        onPick={(color) => run({ kind: "cellShading", fill: color })}
      />
      {fill !== null && <Btn label="No fill" title="Remove cell fill" onClick={() => run({ kind: "cellShading", fill: null })} />}
      <ActionMenu
        label="Cell alignment"
        title="Align text inside the current cell"
        width={132}
        groups={[{ items: [["top", "Top"], ["center", "Middle"], ["bottom", "Bottom"]] }]}
        onPick={(value) => run({ kind: "cellVAlign", v: value as "top" | "center" | "bottom" })}
      />
      <ActionMenu
        label="Rows & columns"
        title="Edit rows and columns around the current cell"
        width={144}
        groups={[
          { label: "Rows", items: [["rowAbove", "Insert row above"], ["rowBelow", "Insert row below"], ["deleteRow", "Delete row"]] },
          { label: "Columns", items: [["colLeft", "Insert column left"], ["colRight", "Insert column right"], ["deleteCol", "Delete column"]] },
        ]}
        onPick={(value) => run(value as Parameters<DocxViewApi["tableOp"]>[0])}
      />
      <ActionMenu
        label="Merge"
        title="Merge or split the current cell"
        width={78}
        groups={[{ items: [["mergeRight", "Merge right"], ["mergeDown", "Merge down"], ["splitCell", "Split cell"]] }]}
        onPick={(value) => run(value as Parameters<DocxViewApi["tableOp"]>[0])}
      />
      <ActionMenu
        label="Sort"
        title="Sort rows by the current column (repeating header rows stay in place)"
        width={64}
        groups={[
          { label: "Text", items: [["text:asc", "Sort A → Z"], ["text:desc", "Sort Z → A"]] },
          { label: "Numbers", items: [["number:asc", "Sort 0 → 9"], ["number:desc", "Sort 9 → 0"]] },
        ]}
        onPick={(value) => {
          const [compare, order] = value.split(":") as ["text" | "number", "asc" | "desc"];
          const colIdx = api?.getTableProperties()?.columnIdx ?? 0;
          after(() => api?.sortTableRows(colIdx, order, compare));
        }}
      />
      <ActionMenu
        label="Convert"
        title="Convert this table to text"
        width={82}
        groups={[{ items: [["tab", "To text (tabs)"], ["comma", "To text (commas)"]] }]}
        onPick={(value) => after(() => api?.convertTableToText(value as "tab" | "comma"))}
      />
      <span ref={borderMenuRef} style={{ display: "inline-flex" }}>
        <ActionMenu
          label="Borders"
          title="Set or clear the borders of the table or the current cell"
          width={92}
          groups={[
            {
              label: "Table",
              items: [
                ["tableAll", "All borders"],
                ["tableOutside", "Outside borders"],
                ["tableInside", "Inside borders"],
                ["tableNone", "No borders"],
                ["tableClear", "Clear direct borders"],
              ],
            },
            {
              label: "Cell",
              items: [
                ["cellAll", "All borders"],
                ["cellNone", "No borders"],
                ["cellClear", "Clear direct borders"],
              ],
            },
            { items: [["custom", "Custom border…"]] },
          ]}
          onPick={(value) => {
            if (value === "custom") setCustomBorder(true);
            else after(() => borderActions[value]?.());
          }}
        />
      </span>
      {customBorder && (
        <CustomBorderDialog
          api={api}
          anchorRef={borderMenuRef}
          onClose={closeCustomBorder}
          onChanged={onChanged}
        />
      )}
      <ActionMenu
        label={styleId ? "Style" : "Table style"}
        title="Apply a table style defined in this document"
        width={112}
        groups={[
          {
            items: [
              // Not "": ToolbarMenuSelect treats the empty value as the
              // current selection and appends its own tick, which would
              // contradict the ticks this menu draws itself. A style id is
              // an XML name, so it can never collide with this sentinel.
              [NO_TABLE_STYLE, tick(styleId === null, "No style")],
              ...(api?.listTableStyles() ?? []).map(
                ({ id, name }) => [id, tick(id === styleId, name)] as [string, string],
              ),
            ],
          },
        ]}
        onPick={(value) => after(() => api?.setTableStyle(value === NO_TABLE_STYLE ? null : value))}
      />
      {/* The six checkboxes Word calls Table Style Options: they choose WHICH
          of a style's conditional formats apply, so they are meaningless
          without one. */}
      {styleId !== null && (
        <ActionMenu
          label="Style options"
          title="Choose which parts of the table style apply"
          width={118}
          groups={[
            {
              items: [
                ["firstRow", tick(look?.firstRow, "Header row")],
                ["lastRow", tick(look?.lastRow, "Total row")],
                ["firstColumn", tick(look?.firstColumn, "First column")],
                ["lastColumn", tick(look?.lastColumn, "Last column")],
                ["bandedRows", tick(look?.bandedRows, "Banded rows")],
                ["bandedCols", tick(look?.bandedCols, "Banded columns")],
              ],
            },
          ]}
          onPick={(value) => {
            const key = value as keyof NonNullable<typeof look>;
            if (!look) return;
            after(() => api?.setTableLook({ [key]: !look[key] }));
          }}
        />
      )}
      <ActionMenu
        label="AutoFit"
        title="Choose how the table sizes its columns"
        width={92}
        groups={[
          {
            items: [
              ["contents", "AutoFit to contents"],
              ["window", "AutoFit to window"],
              ["fixed", "Fixed column width"],
            ],
          },
        ]}
        onPick={(value) =>
          after(() => {
            if (value === "fixed") {
              api?.setTableLayout("fixed");
              return;
            }
            api?.setTableLayout("autofit");
            // "To window" is autofit measured against a full-width target
            // rather than against the content's own width.
            if (value === "window") api?.setTableWidth("pct", 100);
          })
        }
      />
      <ActionMenu
        label="Header rows"
        title="Repeat leading rows at the top of every page"
        width={112}
        groups={[{ items: [["0", "None"], ["1", "First row"], ["2", "First two rows"]] }]}
        onPick={(value) => after(() => api?.setTableHeaderRows(Number(value)))}
      />
      <TablePropertiesDialog api={api} onChanged={onChanged} />
      <Btn label="Delete table" title="Delete the current table" onClick={() => run("deleteTable")} />
    </span>
  );
}

// ---------------------------------------------------------------- styles pane

type StyleEntry = ReturnType<NonNullable<DocxViewApi["listStyles"]>>[number];

/** Paint a gallery row the way the style itself would paint text. RunProps
 * carries a resolved size in PX and a CSS color, so both go straight through. */
function previewStyle(preview: StyleEntry["preview"]): React.CSSProperties {
  return {
    fontFamily: preview.font,
    fontWeight: preview.bold ? 600 : 400,
    fontStyle: preview.italic ? "italic" : "normal",
    textDecoration: preview.underline && preview.underline !== "none" ? "underline" : undefined,
    // Clamped: a 36pt Title would push every other row off the panel.
    fontSize: preview.size ? Math.min(18, Math.max(11, preview.size)) : 13,
    color: preview.color && preview.color !== "auto" ? preview.color : undefined,
  };
}

/** The form both "New style" and "Modify style" fill in. */
interface StyleForm {
  name: string;
  type: "paragraph" | "character";
  basedOn: string;
  quickStyle: boolean;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  sizePt: string;
  color: string;
  /** "" leaves the alignment alone, which for a new style means inherit. */
  alignment: "" | "left" | "center" | "right" | "both";
}

const PX_PER_PT = 4 / 3;

function formFor(entry: StyleEntry | null): StyleForm {
  const preview = entry?.preview ?? {};
  return {
    name: entry?.name ?? "",
    type: entry?.type === "character" ? "character" : "paragraph",
    basedOn: entry?.basedOn ?? "",
    quickStyle: entry?.quickStyle ?? true,
    bold: !!preview.bold,
    italic: !!preview.italic,
    underline: !!preview.underline && preview.underline !== "none",
    sizePt: preview.size ? showPt(preview.size / PX_PER_PT) : "",
    color: preview.color && preview.color !== "auto" ? preview.color : "",
    alignment: "",
  };
}

/**
 * Word's Styles pane, as a ribbon popover.
 *
 * listStyles/createStyle/modifyStyle/deleteStyle have all been on the api
 * since the styles work landed, reachable only from the two dropdowns that
 * APPLY a style. Nothing in the toolbar could define one, rename one, or say
 * how much of the document a style is holding up.
 *
 * Paragraph and character styles only: those are the two an "apply" click at
 * the caret can honour. A table style is applied through the table tab, where
 * there is a table to apply it to.
 *
 * MODIFY WRITES ONLY WHAT CHANGED. The form prefills from the RESOLVED
 * preview — what the user is looking at — so saving the whole form would
 * copy every inherited property onto the definition and quietly cut it out of
 * its own cascade.
 */
function StylesPane({ api, onChanged }: { api: DocxViewApi | null; onChanged: () => void }) {
  const [open, setOpen] = useState(false);
  const [editing, setEditing] = useState<{ entry: StyleEntry | null; form: StyleForm; initial: StyleForm } | null>(null);
  const [confirmDelete, setConfirmDelete] = useState<string | null>(null);
  const [error, setError] = useState("");
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);
  const close = useCallback(() => {
    setOpen(false);
    setEditing(null);
    setConfirmDelete(null);
    setError("");
  }, []);
  const position = useAnchoredPopover(triggerRef, rootRef, 300, close);

  const entries = (open ? api?.listStyles?.() ?? [] : []).filter(
    (entry) => entry.type === "paragraph" || entry.type === "character",
  );

  const applyStyle = (entry: StyleEntry) => {
    if (entry.type === "character") api?.applyFormat({ characterStyleId: entry.id });
    else api?.setParagraphStyle(entry.id);
    onChanged();
    close();
  };

  const remove = (entry: StyleEntry) => {
    if (!api?.deleteStyle(entry.id)) {
      setError(`${entry.name} cannot be deleted — the default paragraph style holds up every other one.`);
      return;
    }
    setConfirmDelete(null);
    onChanged();
  };

  const startEdit = (entry: StyleEntry | null) => {
    const form = formFor(entry);
    setEditing({ entry, form, initial: form });
    setError("");
  };

  /** The run properties the user actually moved, and nothing else. */
  const runPatch = (form: StyleForm, initial: StyleForm) => {
    const patch: Record<string, unknown> = {};
    if (form.bold !== initial.bold) patch.bold = form.bold;
    if (form.italic !== initial.italic) patch.italic = form.italic;
    if (form.underline !== initial.underline) patch.underline = form.underline;
    if (form.sizePt !== initial.sizePt) {
      const pt = typedNumber(form.sizePt);
      if (pt !== null) patch.fontSizePt = pt;
    }
    if (form.color !== initial.color) {
      const color = normalizedColor(form.color);
      if (color) patch.color = color.slice(1).toUpperCase();
    }
    return patch;
  };

  const save = () => {
    if (!editing || !api) return;
    const { entry, form, initial } = editing;
    const name = form.name.trim();
    if (!name) {
      setError("Give the style a name.");
      return;
    }
    const run = runPatch(form, initial);
    const paragraph = form.alignment === "" ? undefined : { alignment: form.alignment };
    if (entry) {
      const patch = {
        ...(name === entry.name ? {} : { name }),
        ...(form.basedOn === (entry.basedOn ?? "") ? {} : { basedOn: form.basedOn === "" ? null : form.basedOn }),
        ...(form.quickStyle === entry.quickStyle ? {} : { quickStyle: form.quickStyle }),
        ...(Object.keys(run).length > 0 ? { run } : {}),
        ...(paragraph && entry.type === "paragraph" ? { paragraph } : {}),
      };
      if (Object.keys(patch).length === 0) {
        setEditing(null);
        return;
      }
      if (!api.modifyStyle(entry.id, patch)) {
        setError("That change was refused — a style cannot be based on itself.");
        return;
      }
    } else {
      const styleId = uniqueStyleId(api.document, styleIdFromName(name));
      const created = api.createStyle({
        styleId,
        type: form.type,
        name,
        ...(form.basedOn === "" ? {} : { basedOn: form.basedOn }),
        quickStyle: form.quickStyle,
        ...(Object.keys(run).length > 0 ? { run } : {}),
        ...(paragraph && form.type === "paragraph" ? { paragraph } : {}),
      });
      if (!created) {
        setError("The style could not be created. A shared document defines styles through the room.");
        return;
      }
    }
    setEditing(null);
    onChanged();
  };

  const check = (
    label: string,
    value: boolean,
    onToggle: (next: boolean) => void,
  ) => (
    <label style={{ display: "flex", gap: 5, alignItems: "center", fontSize: 12 }}>
      <input type="checkbox" aria-label={label} checked={value} onChange={(event) => onToggle(event.target.checked)} />
      {label}
    </label>
  );

  const form = editing?.form;
  return (
    <span style={{ position: "relative", display: "inline-flex" }}>
      <Btn
        label="Styles"
        title="Styles pane: create, change and remove this document's styles"
        active={open}
        buttonRef={triggerRef}
        onClick={() => (open ? close() : (setOpen(true), setEditing(null)))}
      />
      {open && (
        <div
          ref={rootRef}
          role="dialog"
          aria-label="Styles"
          data-dxw-styles-pane=""
          onMouseDown={(event) => event.stopPropagation()}
          style={{
            position: "fixed", top: position.top, left: position.left, zIndex: 201,
            width: "min(300px, calc(100vw - 16px))", boxSizing: "border-box",
            background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8,
            boxShadow: T.popoverShadow, padding: 10, display: "grid", gap: 7, color: T.fg,
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <strong style={{ fontSize: 13 }}>{editing ? (editing.entry ? "Modify Style" : "New Style") : "Styles"}</strong>
            {!editing && (
              <button type="button" style={pillBtn} onClick={() => startEdit(null)}>New style</button>
            )}
          </div>

          {!editing && (
            <div style={{ maxHeight: 260, overflowY: "auto", display: "grid", gap: 1 }}>
              {entries.map((entry) => (
                <div key={entry.id} style={{ display: "flex", gap: 4, alignItems: "center" }}>
                  <button
                    type="button"
                    title={`Apply ${entry.name}`}
                    aria-label={`Apply ${entry.name}`}
                    onClick={() => applyStyle(entry)}
                    style={{
                      flex: 1, minWidth: 0, textAlign: "left", background: "none", border: "none",
                      borderRadius: 5, padding: "4px 6px", cursor: "pointer", color: T.fg,
                      display: "flex", justifyContent: "space-between", gap: 8, alignItems: "baseline",
                    }}
                  >
                    <span style={{ ...previewStyle(entry.preview), overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {entry.type === "character" ? "\u{1D400} " : "¶ "}
                      {entry.name}
                    </span>
                    <span style={{ color: T.muted, fontSize: 11, flex: "none" }}>
                      {entry.usageCount === 0 ? "unused" : `${entry.usageCount} use${entry.usageCount === 1 ? "" : "s"}`}
                    </span>
                  </button>
                  <button
                    type="button"
                    title={`Modify ${entry.name}`}
                    aria-label={`Modify ${entry.name}`}
                    onClick={() => startEdit(entry)}
                    style={{ ...pillBtn, background: T.popoverBg, color: T.fg, padding: "2px 7px" }}
                  >
                    Edit
                  </button>
                  <button
                    type="button"
                    title={`Delete ${entry.name}`}
                    aria-label={confirmDelete === entry.id ? `Confirm delete ${entry.name}` : `Delete ${entry.name}`}
                    onClick={() => (confirmDelete === entry.id ? remove(entry) : setConfirmDelete(entry.id))}
                    style={{ ...pillBtn, background: T.popoverBg, color: T.fg, padding: "2px 7px" }}
                  >
                    {confirmDelete === entry.id ? "Sure?" : "×"}
                  </button>
                </div>
              ))}
              {entries.length === 0 && (
                <span style={{ color: T.muted, fontSize: 12 }}>This document declares no styles yet.</span>
              )}
            </div>
          )}

          {editing && form && (
            <>
              <label style={dialogFieldRow}>
                <span>Name</span>
                <input
                  aria-label="Style name"
                  autoFocus
                  value={form.name}
                  onChange={(event) => setEditing({ ...editing, form: { ...form, name: event.target.value } })}
                  style={dialogInput}
                />
              </label>
              {!editing.entry && (
                <label style={dialogFieldRow}>
                  <span>Type</span>
                  <select
                    aria-label="Style type"
                    value={form.type}
                    onChange={(event) =>
                      setEditing({ ...editing, form: { ...form, type: event.target.value as StyleForm["type"] } })
                    }
                    style={dialogInput}
                  >
                    <option value="paragraph">Paragraph</option>
                    <option value="character">Character</option>
                  </select>
                </label>
              )}
              <label style={dialogFieldRow}>
                <span>Based on</span>
                <select
                  aria-label="Based on"
                  value={form.basedOn}
                  onChange={(event) => setEditing({ ...editing, form: { ...form, basedOn: event.target.value } })}
                  style={dialogInput}
                >
                  <option value="">(none)</option>
                  {entries
                    .filter((other) => other.type === form.type && other.id !== editing.entry?.id)
                    .map((other) => (
                      <option key={other.id} value={other.id}>{other.name}</option>
                    ))}
                </select>
              </label>
              <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                {check("Bold", form.bold, (bold) => setEditing({ ...editing, form: { ...form, bold } }))}
                {check("Italic", form.italic, (italic) => setEditing({ ...editing, form: { ...form, italic } }))}
                {check("Underline", form.underline, (underline) => setEditing({ ...editing, form: { ...form, underline } }))}
              </div>
              <label style={dialogFieldRow}>
                <span>Size (pt)</span>
                <input
                  aria-label="Style font size (points)"
                  type="number"
                  min="1"
                  step="1"
                  placeholder="—"
                  value={form.sizePt}
                  onChange={(event) => setEditing({ ...editing, form: { ...form, sizePt: event.target.value } })}
                  style={dialogInput}
                />
              </label>
              <span style={dialogFieldRow}>
                <span>Color</span>
                <ColorMenu
                  current={normalizedColor(form.color) ?? "#000000"}
                  title="Style text color"
                  trigger={(
                    <>
                      <span
                        aria-hidden="true"
                        style={{ width: 14, height: 14, borderRadius: 3, border: `1px solid ${T.border}`, background: normalizedColor(form.color) ?? "#000000" }}
                      />
                      {form.color === "" ? "Inherited" : form.color}
                    </>
                  )}
                  onPick={(color) => setEditing({ ...editing, form: { ...form, color } })}
                />
              </span>
              {form.type === "paragraph" && (
                <label style={dialogFieldRow}>
                  <span>Alignment</span>
                  <select
                    aria-label="Style alignment"
                    value={form.alignment}
                    onChange={(event) =>
                      setEditing({ ...editing, form: { ...form, alignment: event.target.value as StyleForm["alignment"] } })
                    }
                    style={dialogInput}
                  >
                    <option value="">(unchanged)</option>
                    <option value="left">Left</option>
                    <option value="center">Center</option>
                    <option value="right">Right</option>
                    <option value="both">Justified</option>
                  </select>
                </label>
              )}
              {check("Show in the quick-style gallery", form.quickStyle, (quickStyle) =>
                setEditing({ ...editing, form: { ...form, quickStyle } }))}
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
                <button type="button" data-dxw-dialog-cancel="" onClick={() => setEditing(null)} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Cancel</button>
                <button type="button" data-dxw-dialog-apply="" onClick={save} style={pillBtn}>Save</button>
              </div>
            </>
          )}
          {error && <span style={{ color: "#c5221f", fontSize: 11.5 }}>{error}</span>}
        </div>
      )}
    </span>
  );
}

/**
 * Default formatting toolbar for an editable DocxView. Compact, grouped like
 * a word processor; every control preserves the selection/caret.
 */
type LayoutPatch = Parameters<DocxViewApi["setPageLayout"]>[0];

type LayoutMenuOption = {
  value: string;
  label: string;
  description?: string;
  preview: React.ReactNode;
};

function PagePreview({
  kind,
  width = 8.5,
  height = 11,
  margins,
  mirrored,
  columns,
  columnSeparator,
  border,
  borderColor,
  lineNumbers,
}: {
  kind: string;
  width?: number;
  height?: number;
  margins?: [number, number, number, number];
  mirrored?: boolean;
  columns?: number;
  columnSeparator?: boolean;
  border?: "none" | "thin" | "thick" | "accent";
  borderColor?: string;
  lineNumbers?: boolean;
}) {
  const maxWidth = mirrored ? 21 : 34;
  const maxHeight = 42;
  const scale = Math.min(maxWidth / width, maxHeight / height);
  const paperWidth = Math.max(12, width * scale);
  const paperHeight = Math.max(18, height * scale);
  const papers = mirrored ? 2 : 1;
  const inset = margins ?? [0.8, 0.8, 0.8, 0.8];
  const borderWidth = border === "thick" ? 2 : border === "none" ? 0 : 1;
  const pageBorderColor = borderColor ?? (border === "accent" ? T.accent : T.muted);
  return (
    <span
      aria-hidden="true"
      className="dxw-layout-preview"
      data-dxw-layout-preview={kind}
      style={{ width: 52, height: 46, display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 3, flex: "0 0 52px" }}
    >
      {Array.from({ length: papers }, (_, paper) => (
        <span
          key={paper}
          className="dxw-layout-preview-page"
          style={{
            position: "relative", display: "block", boxSizing: "border-box",
            width: paperWidth, height: paperHeight, background: "var(--dxw-layout-preview-bg, #fff)",
            border: `${Math.max(1, borderWidth)}px solid ${borderWidth ? pageBorderColor : T.border}`,
          }}
        >
          {kind === "margins" && (
            <span style={{
              position: "absolute",
              top: `${Math.min(35, inset[0] * 18)}%`, right: `${Math.min(35, (mirrored && paper === 0 ? inset[3] : inset[1]) * 15)}%`,
              bottom: `${Math.min(35, inset[2] * 18)}%`, left: `${Math.min(35, (mirrored && paper === 0 ? inset[1] : inset[3]) * 15)}%`,
              border: `1px solid ${T.accent}`, boxSizing: "border-box",
            }} />
          )}
          {!!columns && (
            <span style={{ position: "absolute", inset: "5px 3px", display: "flex", gap: 2 }}>
              {Array.from({ length: columns }, (_, column) => (
                <span key={column} style={{ flex: 1, background: `repeating-linear-gradient(to bottom, ${T.muted} 0 1px, transparent 1px 4px)` }} />
              ))}
              {columnSeparator && (
                <span style={{ position: "absolute", left: "50%", top: 0, bottom: 0, width: 1, background: T.muted }} />
              )}
            </span>
          )}
          {border === "none" && <span style={{ position: "absolute", inset: 4, border: `1px dashed ${T.border}` }} />}
          {lineNumbers && (
            <span style={{ position: "absolute", inset: "4px 3px", display: "grid", gridTemplateColumns: "8px 1fr", gap: 2 }}>
              <span style={{ fontSize: 5, lineHeight: "6px", color: T.accent }}>1<br />2<br />3<br />4</span>
              <span style={{ background: `repeating-linear-gradient(to bottom, ${T.muted} 0 1px, transparent 1px 6px)` }} />
            </span>
          )}
        </span>
      ))}
    </span>
  );
}

function LayoutMenu({
  name,
  label,
  open,
  onOpenChange,
  options,
  onPick,
}: {
  name: string;
  label: string;
  open: boolean;
  onOpenChange: (open: boolean) => void;
  options: LayoutMenuOption[];
  onPick: (value: string) => void;
}) {
  const id = useId();
  const triggerRef = useRef<HTMLButtonElement | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const keyboardOpen = useRef<"first" | "last" | null>(null);
  const [position, setPosition] = useState({ left: 8, top: 8, maxHeight: 480 });
  const [portalTokens, setPortalTokens] = useState<React.CSSProperties>({});
  useLayoutEffect(() => {
    if (!open) return;
    const update = () => {
      const trigger = triggerRef.current;
      const menu = menuRef.current;
      if (!trigger) return;
      const rect = trigger.getBoundingClientRect();
      const menuWidth = menu?.offsetWidth ?? 304;
      const left = Math.max(8, Math.min(rect.left, window.innerWidth - menuWidth - 8));
      const below = window.innerHeight - rect.bottom - 12;
      const above = rect.top - 12;
      const placeAbove = below < 180 && above > below;
      const maxHeight = Math.max(120, Math.min(480, placeAbove ? above : below));
      const shownHeight = Math.min(menu?.scrollHeight ?? maxHeight, maxHeight);
      const top = placeAbove ? Math.max(8, rect.top - shownHeight - 4) : rect.bottom + 4;
      setPosition({ left, top, maxHeight });
      const computed = getComputedStyle(trigger);
      const tokens: Record<string, string> = {};
      for (const property of [
        "--dxw-toolbar-fg", "--dxw-toolbar-border", "--dxw-toolbar-muted",
        "--dxw-accent", "--dxw-btn-hover-bg", "--dxw-popover-bg",
        "--dxw-popover-shadow", "--dxw-layout-menu-width",
        "--dxw-layout-menu-max-height", "--dxw-layout-preview-bg",
        "--dxw-toolbar-z-index",
      ]) {
        const value = computed.getPropertyValue(property);
        if (value) tokens[property] = value;
      }
      setPortalTokens(tokens as React.CSSProperties);
    };
    update();
    if (keyboardOpen.current) {
      const items = menuRef.current?.querySelectorAll<HTMLButtonElement>('[role="menuitem"]');
      const item = keyboardOpen.current === "last" ? items?.[items.length - 1] : items?.[0];
      item?.focus({ preventScroll: true });
      keyboardOpen.current = null;
    }
    const frame = requestAnimationFrame(update);
    const close = (event: MouseEvent) => {
      const target = event.target as Node;
      if (!triggerRef.current?.contains(target) && !menuRef.current?.contains(target)) onOpenChange(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        const restore = menuRef.current?.contains(document.activeElement);
        onOpenChange(false);
        if (restore) requestAnimationFrame(() => triggerRef.current?.focus({ preventScroll: true }));
      }
    };
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", update);
    window.addEventListener("scroll", update, true);
    return () => {
      cancelAnimationFrame(frame);
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", update);
      window.removeEventListener("scroll", update, true);
    };
  }, [open]);

  const restoreTrigger = () => requestAnimationFrame(() => triggerRef.current?.focus({ preventScroll: true }));
  const onMenuKeyDown = (event: React.KeyboardEvent<HTMLDivElement>) => {
    const items = menuRef.current
      ? Array.from(menuRef.current.querySelectorAll<HTMLButtonElement>('[role="menuitem"]'))
      : [];
    const current = items.indexOf(document.activeElement as HTMLButtonElement);
    let next = current;
    if (event.key === "ArrowDown") next = current < 0 ? 0 : (current + 1) % items.length;
    else if (event.key === "ArrowUp") next = current < 0 ? items.length - 1 : (current - 1 + items.length) % items.length;
    else if (event.key === "Home") next = 0;
    else if (event.key === "End") next = items.length - 1;
    else if ((event.key === "Enter" || event.key === " ") && current >= 0) {
      event.preventDefault();
      event.stopPropagation();
      items[current].click();
      return;
    } else if (event.key === "Escape" || event.key === "Tab") {
      event.preventDefault();
      event.stopPropagation();
      onOpenChange(false);
      restoreTrigger();
      return;
    } else return;
    event.preventDefault();
    event.stopPropagation();
    items[next]?.focus({ preventScroll: true });
  };

  return (
    <span className="dxw-layout-control" style={{ display: "inline-flex", minWidth: 0 }}>
      <button
        ref={triggerRef}
        type="button"
        aria-haspopup="menu"
        aria-expanded={open}
        aria-controls={open ? id : undefined}
        title={label}
        className="dxw-layout-menu-trigger"
        data-dxw-layout-menu-trigger={name}
        onMouseDown={(event) => event.preventDefault()}
        onClick={() => onOpenChange(!open)}
        onKeyDown={(event) => {
          if (!open && (event.key === "Enter" || event.key === " " || event.key === "ArrowDown" || event.key === "ArrowUp")) {
            event.preventDefault();
            keyboardOpen.current = event.key === "ArrowUp" ? "last" : "first";
            onOpenChange(true);
          }
        }}
        style={{
          ...btnStyle(open), minWidth: 76, maxWidth: "100%", height: 30, padding: "0 8px",
          display: "inline-flex", alignItems: "center", justifyContent: "space-between", gap: 6,
          whiteSpace: "nowrap", fontWeight: 500,
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis" }}>{label}</span>
        <span aria-hidden="true" style={{ fontSize: 10 }}>⌄</span>
      </button>
      {open && createPortal(
        <div
          ref={menuRef}
          id={id}
          role="menu"
          aria-label={label}
          className="dxw-layout-menu"
          data-dxw-layout-menu={name}
          onMouseDown={(event) => event.preventDefault()}
          onKeyDown={onMenuKeyDown}
          style={{
            ...portalTokens,
            position: "fixed", left: position.left, top: position.top,
            zIndex: "var(--dxw-toolbar-z-index, 100)",
            width: "min(var(--dxw-layout-menu-width, 304px), calc(100vw - 16px))",
            maxHeight: `min(var(--dxw-layout-menu-max-height, ${position.maxHeight}px), ${position.maxHeight}px)`,
            overflowY: "auto", overscrollBehavior: "contain", boxSizing: "border-box",
            background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8,
            boxShadow: T.popoverShadow, padding: 6,
          }}
        >
          {options.map((option) => (
            <button
              key={option.value}
              type="button"
              role="menuitem"
              className="dxw-layout-menu-item"
              data-dxw-layout-option={option.value}
              tabIndex={-1}
              onMouseDown={(event) => event.preventDefault()}
              onClick={(event) => {
                onPick(option.value);
                onOpenChange(false);
                if (event.detail === 0) restoreTrigger();
              }}
              style={{
                width: "100%", border: 0, borderRadius: 6, background: "transparent", color: T.fg,
                display: "flex", alignItems: "center", gap: 10, padding: "5px 8px",
                textAlign: "left", cursor: "pointer", fontFamily: "inherit",
              }}
              onMouseEnter={(event) => (event.currentTarget.style.background = T.hoverBg)}
              onMouseLeave={(event) => (event.currentTarget.style.background = "transparent")}
            >
              {option.preview}
              <span style={{ minWidth: 0 }}>
                <span style={{ display: "block", fontSize: 13, fontWeight: 550 }}>{option.label}</span>
                {option.description && <span style={{ display: "block", fontSize: 11, color: T.muted, marginTop: 1 }}>{option.description}</span>}
              </span>
            </button>
          ))}
        </div>
      , document.body)}
    </span>
  );
}

function MarginMenu({
  scope,
  onApply,
  open,
  onOpenChange,
}: {
  scope: "document" | "section";
  onApply: (patch: LayoutPatch) => void;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const [customOpen, setCustomOpen] = useState(false);
  const [values, setValues] = useState({ top: "1", bottom: "1", left: "1", right: "1" });
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const [dialogPosition, setDialogPosition] = useState({ left: 8, top: 8 });
  useEffect(() => {
    if (!customOpen) return;
    const positionDialog = () => {
      const trigger = rootRef.current?.querySelector<HTMLElement>("[data-dxw-layout-menu-trigger]");
      if (!trigger) return;
      const rect = trigger.getBoundingClientRect();
      const width = Math.min(244, window.innerWidth - 16);
      setDialogPosition({
        left: Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)),
        top: Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - 294)),
      });
    };
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setCustomOpen(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setCustomOpen(false);
    };
    positionDialog();
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", positionDialog);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", positionDialog);
    };
  }, [customOpen]);

  const pick = (value: string) => {
    if (value === "m:custom") {
      setCustomOpen(true);
      return;
    }
    setCustomOpen(false);
    if (value === "m:normal") onApply({ margins: { top: 1, right: 1, bottom: 1, left: 1 }, mirrorMargins: false });
    else if (value === "m:narrow") onApply({ margins: { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 }, mirrorMargins: false });
    else if (value === "m:moderate") onApply({ margins: { top: 1, right: 0.75, bottom: 1, left: 0.75 }, mirrorMargins: false });
    else if (value === "m:wide") onApply({ margins: { top: 1, right: 2, bottom: 1, left: 2 }, mirrorMargins: false });
    else if (value === "m:mirrored") onApply({ margins: { top: 1, right: 1, bottom: 1, left: 1.25 }, mirrorMargins: true });
  };
  const valid = Object.values(values).every((value) =>
    value.trim() !== "" && Number.isFinite(Number(value)) && Number(value) >= 0,
  );
  const applyCustom = () => {
    if (!valid) return;
    onApply({
      margins: {
        top: Number(values.top), bottom: Number(values.bottom),
        left: Number(values.left), right: Number(values.right),
      },
      mirrorMargins: false,
    });
    setCustomOpen(false);
  };
  const field = (side: keyof typeof values, label: string) => (
    <label style={{ display: "grid", gridTemplateColumns: "54px 1fr", gap: 8, alignItems: "center", fontSize: 12 }}>
      <span>{label}</span>
      <input
        aria-label={`${label} margin (inches)`}
        type="number"
        min="0"
        step="0.05"
        required
        autoFocus={side === "top"}
        value={values[side]}
        onChange={(event) => setValues({ ...values, [side]: event.target.value })}
        onKeyDown={(event) => {
          if (event.key === "Escape") setCustomOpen(false);
          else if (event.key === "Enter") applyCustom();
        }}
        style={{ width: 92, boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 6px" }}
      />
    </label>
  );

  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-flex" }}>
      <LayoutMenu
        name="margins"
        label="Margins"
        open={open}
        onOpenChange={onOpenChange}
        onPick={pick}
        options={[
          { value: "m:normal", label: "Normal", description: '1" on every side', preview: <PagePreview kind="margins" margins={[1, 1, 1, 1]} /> },
          { value: "m:narrow", label: "Narrow", description: '0.5" on every side', preview: <PagePreview kind="margins" margins={[0.5, 0.5, 0.5, 0.5]} /> },
          { value: "m:moderate", label: "Moderate", description: '1" top/bottom, 0.75" left/right', preview: <PagePreview kind="margins" margins={[1, 0.75, 1, 0.75]} /> },
          { value: "m:wide", label: "Wide", description: '1" top/bottom, 2" left/right', preview: <PagePreview kind="margins" margins={[1, 2, 1, 2]} /> },
          { value: "m:mirrored", label: "Mirrored", description: "Facing pages; inside margin 1.25\"", preview: <PagePreview kind="margins" margins={[1, 1, 1, 1.25]} mirrored /> },
          { value: "m:custom", label: "Custom Margins…", description: "Set each side in inches", preview: <PagePreview kind="margins" margins={[0.7, 1.2, 0.7, 1.2]} /> },
        ]}
      />
      {customOpen && (
        <div
          role="dialog"
          aria-label="Custom Margins"
          onMouseDown={(event) => event.stopPropagation()}
          style={{
            position: "fixed", top: dialogPosition.top, left: dialogPosition.left, zIndex: 201,
            width: "min(224px, calc(100vw - 16px))", boxSizing: "border-box",
            background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8,
            boxShadow: T.popoverShadow, padding: 10, display: "grid", gap: 7,
          }}
        >
          <strong style={{ fontSize: 13 }}>Custom Margins</strong>
          {field("top", "Top")}
          {field("bottom", "Bottom")}
          {field("left", "Left")}
          {field("right", "Right")}
          <span style={{ color: T.muted, fontSize: 11 }}>
            Applies to {scope === "section" ? "this section" : "the whole document"} and turns mirrored margins off for the whole document.
          </span>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
            <button type="button" onClick={() => setCustomOpen(false)} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Cancel</button>
            <button type="button" onClick={applyCustom} disabled={!valid} style={pillBtn}>Apply</button>
          </div>
        </div>
      )}
    </span>
  );
}

function PageSizeMenu({
  scope,
  onApply,
  open,
  onOpenChange,
}: {
  scope: "document" | "section";
  onApply: (patch: LayoutPatch) => void;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const [customOpen, setCustomOpen] = useState(false);
  const [values, setValues] = useState({ width: "8.5", height: "11" });
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const [dialogPosition, setDialogPosition] = useState({ left: 8, top: 8 });
  useEffect(() => {
    if (!customOpen) return;
    const positionDialog = () => {
      const trigger = rootRef.current?.querySelector<HTMLElement>("[data-dxw-layout-menu-trigger]");
      if (!trigger) return;
      const rect = trigger.getBoundingClientRect();
      const width = Math.min(244, window.innerWidth - 16);
      setDialogPosition({
        left: Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)),
        top: Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - 210)),
      });
    };
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setCustomOpen(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setCustomOpen(false);
    };
    positionDialog();
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", positionDialog);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", positionDialog);
    };
  }, [customOpen]);

  const valid = Object.values(values).every((value) =>
    value.trim() !== "" && Number.isFinite(Number(value)) && Number(value) > 0,
  );
  const applyCustom = () => {
    if (!valid) return;
    onApply({ size: { width: Number(values.width), height: Number(values.height) } });
    setCustomOpen(false);
  };
  const field = (side: keyof typeof values, label: string) => (
    <label style={{ display: "grid", gridTemplateColumns: "54px 1fr", gap: 8, alignItems: "center", fontSize: 12 }}>
      <span>{label}</span>
      <input
        aria-label={`Page ${side} (inches)`}
        type="number"
        min="0.1"
        step="0.05"
        required
        autoFocus={side === "width"}
        value={values[side]}
        onChange={(event) => setValues({ ...values, [side]: event.target.value })}
        onKeyDown={(event) => {
          if (event.key === "Escape") setCustomOpen(false);
          else if (event.key === "Enter") applyCustom();
        }}
        style={{ width: 92, boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 6px" }}
      />
    </label>
  );

  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-flex" }}>
      <LayoutMenu
        name="size"
        label="Size"
        open={open}
        onOpenChange={onOpenChange}
        options={[
          ...PAGE_SIZES.map((size) => ({
            value: size.value,
            label: size.label,
            description: size.description,
            preview: <PagePreview kind="size" width={size.width} height={size.height} />,
          })),
          { value: "custom", label: "Custom Paper Size…", description: "Set width and height in inches", preview: <PagePreview kind="size" width={7.5} height={10} /> },
        ]}
        onPick={(value) => {
          if (value === "custom") {
            setCustomOpen(true);
            return;
          }
          setCustomOpen(false);
          const size = PAGE_SIZES.find((entry) => entry.value === value);
          if (size) onApply({ size: { width: size.width, height: size.height } });
        }}
      />
      {customOpen && (
        <div
          role="dialog"
          aria-label="Custom Paper Size"
          onMouseDown={(event) => event.stopPropagation()}
          style={{
            position: "fixed", top: dialogPosition.top, left: dialogPosition.left, zIndex: 201,
            width: "min(224px, calc(100vw - 16px))", boxSizing: "border-box",
            background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8,
            boxShadow: T.popoverShadow, padding: 10, display: "grid", gap: 7,
          }}
        >
          <strong style={{ fontSize: 13 }}>Custom Paper Size</strong>
          {field("width", "Width")}
          {field("height", "Height")}
          <span style={{ color: T.muted, fontSize: 11 }}>
            Applies to {scope === "section" ? "this section" : "the whole document"}.
          </span>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
            <button type="button" onClick={() => setCustomOpen(false)} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Cancel</button>
            <button type="button" onClick={applyCustom} disabled={!valid} style={pillBtn}>Apply</button>
          </div>
        </div>
      )}
    </span>
  );
}

function PageBorderMenu({
  scope,
  onApply,
  open,
  onOpenChange,
}: {
  scope: "document" | "section";
  onApply: (patch: LayoutPatch) => void;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const [customOpen, setCustomOpen] = useState(false);
  const [color, setColor] = useState("#4472c4");
  const [widthPt, setWidthPt] = useState("1");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const [dialogPosition, setDialogPosition] = useState({ left: 8, top: 8 });
  useEffect(() => {
    if (!customOpen) return;
    const positionDialog = () => {
      const trigger = rootRef.current?.querySelector<HTMLElement>("[data-dxw-layout-menu-trigger]");
      if (!trigger) return;
      const rect = trigger.getBoundingClientRect();
      const width = Math.min(244, window.innerWidth - 16);
      setDialogPosition({
        left: Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)),
        top: Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - 224)),
      });
    };
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setCustomOpen(false);
    };
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setCustomOpen(false);
    };
    positionDialog();
    document.addEventListener("mousedown", close);
    document.addEventListener("keydown", keydown);
    window.addEventListener("resize", positionDialog);
    return () => {
      document.removeEventListener("mousedown", close);
      document.removeEventListener("keydown", keydown);
      window.removeEventListener("resize", positionDialog);
    };
  }, [customOpen]);

  const validColor = normalizedColor(color);
  const width = Number(widthPt);
  const valid = validColor !== null && Number.isFinite(width) && width >= 0.25 && width <= 12;
  const applyCustom = () => {
    if (!validColor || !valid) return;
    onApply({ pageBorders: { sz: Math.round(width * 8), color: validColor } });
    setCustomOpen(false);
  };
  const pick = (value: string) => {
    if (value === "custom") {
      setCustomOpen(true);
      return;
    }
    setCustomOpen(false);
    if (value === "none") onApply({ pageBorders: null });
    else if (value === "thin") onApply({ pageBorders: { sz: 4 } });
    else if (value === "thick") onApply({ pageBorders: { sz: 12 } });
    else onApply({ pageBorders: { sz: 8, color: "4472C4" } });
  };

  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-flex" }}>
      <LayoutMenu
        name="page-border"
        label="Page border"
        open={open}
        onOpenChange={onOpenChange}
        onPick={pick}
        options={[
          { value: "none", label: "None", description: "No page border", preview: <PagePreview kind="page-border" border="none" /> },
          { value: "thin", label: "Thin box", description: "½ pt solid line", preview: <PagePreview kind="page-border" border="thin" /> },
          { value: "thick", label: "Thick box", description: "1½ pt solid line", preview: <PagePreview kind="page-border" border="thick" /> },
          { value: "accent", label: "Accent box", description: "Blue 1 pt line", preview: <PagePreview kind="page-border" border="accent" /> },
          { value: "custom", label: "Custom border…", description: "Choose a color and line weight", preview: <PagePreview kind="page-border" border="accent" borderColor={validColor ?? "#4472c4"} /> },
        ]}
      />
      {customOpen && (
        <div
          role="dialog"
          aria-label="Custom Page Border"
          onMouseDown={(event) => event.stopPropagation()}
          style={{
            position: "fixed", top: dialogPosition.top, left: dialogPosition.left, zIndex: 201,
            width: "min(224px, calc(100vw - 16px))", boxSizing: "border-box",
            background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8,
            boxShadow: T.popoverShadow, padding: 10, display: "grid", gap: 8,
          }}
        >
          <strong style={{ fontSize: 13 }}>Custom Page Border</strong>
          <label style={{ display: "grid", gridTemplateColumns: "54px 1fr", gap: 8, alignItems: "center", fontSize: 12 }}>
            <span>Color</span>
            <span style={{ display: "flex", gap: 6 }}>
              <input aria-label="Page border color picker" type="color" value={validColor ?? "#4472c4"} onChange={(event) => setColor(event.target.value)} style={{ width: 34, height: 28, padding: 1, border: `1px solid ${T.border}`, borderRadius: 5, background: T.popoverBg }} />
              <input aria-label="Page border color" autoFocus value={color} onChange={(event) => setColor(event.target.value)} onKeyDown={(event) => event.key === "Enter" && applyCustom()} spellCheck={false} style={{ width: 92, boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 6px", color: T.fg, background: T.popoverBg }} />
            </span>
          </label>
          <label style={{ display: "grid", gridTemplateColumns: "54px 1fr", gap: 8, alignItems: "center", fontSize: 12 }}>
            <span>Weight</span>
            <select aria-label="Page border width" value={widthPt} onChange={(event) => setWidthPt(event.target.value)} style={{ width: 132, border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 6px", color: T.fg, background: T.popoverBg }}>
              <option value="0.5">½ pt</option>
              <option value="1">1 pt</option>
              <option value="1.5">1½ pt</option>
              <option value="2.25">2¼ pt</option>
              <option value="3">3 pt</option>
            </select>
          </label>
          <span style={{ color: T.muted, fontSize: 11 }}>
            Applies to {scope === "section" ? "this section" : "the whole document"}.
          </span>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
            <button type="button" onClick={() => setCustomOpen(false)} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Cancel</button>
            <button type="button" onClick={applyCustom} disabled={!valid} style={pillBtn}>Apply</button>
          </div>
        </div>
      )}
    </span>
  );
}

/** Whether the selected object accepts rotation, per the shared offered set. */
function rotatable(api: DocxViewApi | null): boolean {
  const context = api?.getSelectedObjectContext();
  return !!context && availableObjectCommands(context).includes("rotate");
}

/** Word's Layout ribbon, scoped to the whole document or the caret's
 * section (per-page layout = section breaks + section scope). */
function LayoutTab({ api, showArrange }: { api: DocxViewApi | null; showArrange: boolean }) {
  const [scope, setScope] = useState<"document" | "section">("document");
  const [section, setSection] = useState<{ index: number; count: number } | null>(null);
  const [objectSelected, setObjectSelected] = useState(false);
  const [openMenu, setOpenMenu] = useState<string | null>(null);
  const rootRef = useRef<HTMLSpanElement | null>(null);
  useEffect(() => {
    const refresh = () => {
      setSection(api?.getSectionContext() ?? null);
      setObjectSelected(api?.hasSelectedObject() ?? false);
    };
    refresh();
    document.addEventListener("dxw-selection", refresh);
    document.addEventListener("dxw-object-selection", refresh);
    return () => {
      document.removeEventListener("dxw-selection", refresh);
      document.removeEventListener("dxw-object-selection", refresh);
    };
  }, [api]);
  const set = (patch: Parameters<NonNullable<typeof api>["setPageLayout"]>[0]) => api?.setPageLayout(patch, scope);
  const setLn = (patch: Parameters<NonNullable<typeof api>["setLineNumbering"]>[0]) => api?.setLineNumbering(patch, scope);
  const menuState = (name: string) => ({
    open: openMenu === name,
    onOpenChange: (open: boolean) => setOpenMenu(open ? name : null),
  });
  return (
    <span
      ref={rootRef}
      className="dxw-layout-ribbon"
      data-dxw-layout-ribbon=""
      style={{
        display: "flex",
        flex: "0 0 100%",
        order: 2,
        minWidth: 0,
        boxSizing: "border-box",
        alignItems: "center",
        flexWrap: "wrap",
        gap: 2,
        paddingTop: 4,
        marginTop: 2,
        borderTop: `1px solid ${T.border}`,
      }}
    >
      <ToolbarMenuSelect
        title="Apply layout changes to"
        ariaLabel="Apply layout changes to"
        value={scope}
        width={section ? 160 : 142}
        menuWidth={205}
        options={[
          { value: "document", label: "Whole document" },
          {
            value: "section",
            label: section ? `This section · ${section.index} of ${section.count}` : "This section",
          },
        ]}
        onChange={(value) => setScope(value as "document" | "section")}
        style={{ maxWidth: "100%" }}
      />
      <Sep />
      <MarginMenu scope={scope} onApply={set} {...menuState("margins")} />
      <LayoutMenu
        name="orientation"
        label="Orientation"
        {...menuState("orientation")}
        options={[
          { value: "portrait", label: "Portrait", description: "Vertical page", preview: <PagePreview kind="orientation" /> },
          { value: "landscape", label: "Landscape", description: "Horizontal page", preview: <PagePreview kind="orientation" width={11} height={8.5} /> },
        ]}
        onPick={(value) => set({ orientation: value as "portrait" | "landscape" })}
      />
      <PageSizeMenu scope={scope} onApply={set} {...menuState("size")} />
      <LayoutMenu
        name="columns"
        label="Columns & divider"
        {...menuState("columns")}
        options={[
          { value: "1", label: "One", description: "Single text column", preview: <PagePreview kind="columns" columns={1} /> },
          { value: "2", label: "Two", description: "Two equal columns", preview: <PagePreview kind="columns" columns={2} /> },
          { value: "2-divider", label: "Two + divider line", description: "Two columns with Word's automatic vertical line between them", preview: <PagePreview kind="columns" columns={2} columnSeparator /> },
          { value: "3", label: "Three", description: "Three equal columns", preview: <PagePreview kind="columns" columns={3} /> },
        ]}
        onPick={(value) => set({ columns: parseInt(value, 10), columnSeparator: value === "2-divider" })}
      />
      <PageBorderMenu scope={scope} onApply={set} {...menuState("page-border")} />
      <LayoutMenu
        name="line-numbers"
        label="Line numbers"
        {...menuState("line-numbers")}
        options={[
          { value: "off", label: "None", description: "Hide line numbers", preview: <PagePreview kind="line-numbers" /> },
          { value: "continuous", label: "Continuous", description: "Number every line continuously", preview: <PagePreview kind="line-numbers" lineNumbers /> },
          { value: "eachPage", label: "Restart each page", description: "Start at 1 on every page", preview: <PagePreview kind="line-numbers" lineNumbers /> },
          { value: "eachSection", label: "Restart each section", description: "Start at 1 in each section", preview: <PagePreview kind="line-numbers" lineNumbers /> },
          { value: "by5", label: "Count by 5", description: "Show every fifth line", preview: <PagePreview kind="line-numbers" lineNumbers /> },
          { value: "by10", label: "Count by 10", description: "Show every tenth line", preview: <PagePreview kind="line-numbers" lineNumbers /> },
        ]}
        onPick={(value) => {
          if (value === "off") setLn({ enabled: false });
          else if (value === "continuous") setLn({ enabled: true, countBy: 1, restart: "continuous" });
          else if (value === "eachPage") setLn({ enabled: true, countBy: 1, restart: "newPage" });
          else if (value === "eachSection") setLn({ enabled: true, countBy: 1, restart: "newSection" });
          else if (value === "by5") setLn({ enabled: true, countBy: 5 });
          else setLn({ enabled: true, countBy: 10 });
        }}
      />
      {showArrange && objectSelected && (
        <>
          <Sep />
          <ActionMenu
            label="Align"
            title="Align selected object to page"
            width={76}
            groups={[
              { label: "Horizontal", items: [["alignLeft", "Align left"], ["alignCenter", "Align center"], ["alignRight", "Align right"]] },
              { label: "Vertical", items: [["alignTop", "Align top"], ["alignMiddle", "Align middle"], ["alignBottom", "Align bottom"]] },
            ]}
            onPick={(value) => api?.arrangeObject(value as Parameters<NonNullable<typeof api>["arrangeObject"]>[0])}
          />
          {/* Offered-set gate, shared with the contextual tab and the
              capability matrix: rotation only exists for shape-geometry
              objects; SmartArt/chart graphic frames have no rotatable xfrm,
              and the menu silently no-op'd for them. */}
          {rotatable(api) && (
            <ActionMenu
              label="Rotate"
              title="Rotate selected object"
              width={78}
              groups={[{ items: [["rotateRight", "Rotate right 90°"], ["rotateLeft", "Rotate left 90°"]] }]}
              onPick={(value) => api?.arrangeObject(value as Parameters<NonNullable<typeof api>["arrangeObject"]>[0])}
            />
          )}
          <ActionMenu
            label="Arrange"
            title="Change selected object stacking order"
            width={86}
            groups={[{ items: [["bringToFront", "Bring to front"], ["sendToBack", "Send to back"]] }]}
            onPick={(value) => api?.arrangeObject(value as Parameters<NonNullable<typeof api>["arrangeObject"]>[0])}
          />
        </>
      )}
    </span>
  );
}

/** Match-count status suffix: per-story breakdown when replacements landed
 * outside the body ("2 in headers, 1 in footnotes"). */
const FIND_STORY_LABELS: Record<string, string> = {
  body: "body", header: "headers", footer: "footers", footnote: "footnotes", endnote: "endnotes",
};

/** Find & Replace popover: find selects the first match and reports the
 * count; replace-all reports how many replacements were applied, per story.
 * The Go To row jumps to a page number or a named bookmark. */
function FindReplaceMenu({ api }: { api: DocxViewApi | null }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [replacement, setReplacement] = useState("");
  const [matchCase, setMatchCase] = useState(false);
  const [wholeWord, setWholeWord] = useState(false);
  const [gotoPage, setGotoPage] = useState("");
  const [status, setStatus] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    if (!open) return;
    inputRef.current?.focus();
    const close = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const opts = { matchCase, wholeWord };
  const runFind = () => {
    const n = api?.find(query, opts) ?? 0;
    setStatus(n === 1 ? "1 match" : `${n} matches`);
  };
  const runReplaceAll = () => {
    const result = api?.replaceAll(query, replacement, opts);
    if (!result) return;
    const head = result.total === 1 ? "Replaced 1 match" : `Replaced ${result.total} matches`;
    const stories = Object.entries(result.byStory);
    const breakdown = stories
      .map(([story, n]) => `${n} in ${FIND_STORY_LABELS[story] ?? story}`)
      .join(", ");
    setStatus(stories.length > 1 || (stories.length === 1 && stories[0][0] !== "body") ? `${head} (${breakdown})` : head);
  };
  const runGoToPage = () => {
    const n = Number.parseInt(gotoPage, 10);
    if (!Number.isInteger(n) || n < 1) return;
    setStatus(api?.goToPage(n) ? `Page ${n}` : `No page ${n}`);
  };
  const bookmarks = open ? api?.listBookmarks() ?? [] : [];
  const field: React.CSSProperties = {
    width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`,
    borderRadius: 6, padding: 6, font: "13px system-ui, sans-serif", outline: "none",
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button
        title="Find & replace"
        style={btnStyle(open)}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => { setStatus(""); setOpen(!open); }}
      >
        Find & replace
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, left: 0, zIndex: 100, width: 240, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow, display: "grid", gap: 6 }}>
          <input
            ref={inputRef}
            aria-label="Find text"
            placeholder="Find…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") {
                e.preventDefault();
                runFind();
              }
            }}
            style={field}
          />
          <input
            aria-label="Replace with"
            placeholder="Replace with…"
            value={replacement}
            onChange={(e) => setReplacement(e.target.value)}
            style={field}
          />
          <div style={{ display: "flex", gap: 10 }}>
            <label style={{ display: "flex", gap: 5, alignItems: "center", fontSize: 12 }}>
              <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
              Match case
            </label>
            <label style={{ display: "flex", gap: 5, alignItems: "center", fontSize: 12 }}>
              <input type="checkbox" checked={wholeWord} onChange={(e) => setWholeWord(e.target.checked)} />
              Whole word
            </label>
          </div>
          {status && <div data-dxw-find-status="" style={{ color: T.muted, fontSize: 12 }}>{status}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 6 }}>
            <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} disabled={!query} onClick={runFind}>Find</button>
            <button style={pillBtn} disabled={!query} onClick={runReplaceAll}>Replace all</button>
          </div>
          <div style={{ display: "flex", gap: 6, alignItems: "center", borderTop: `1px solid ${T.border}`, paddingTop: 6 }}>
            <input
              aria-label="Go to page"
              placeholder="Page…"
              inputMode="numeric"
              value={gotoPage}
              onChange={(e) => setGotoPage(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  runGoToPage();
                }
              }}
              style={{ ...field, width: 64 }}
            />
            <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} disabled={!gotoPage} onClick={runGoToPage}>Go</button>
            <select
              aria-label="Go to bookmark"
              value=""
              disabled={bookmarks.length === 0}
              onChange={(e) => {
                const name = e.target.value;
                if (name) setStatus(api?.goToBookmark(name) ? `Bookmark ${name}` : `Bookmark ${name} not found`);
              }}
              style={{ flex: 1, minWidth: 0, border: `1px solid ${T.border}`, borderRadius: 5, padding: "5px 4px", background: T.popoverBg, color: T.fg, font: "12px system-ui, sans-serif" }}
            >
              <option value="">Bookmark…</option>
              {bookmarks.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>
        </div>
      )}
    </span>
  );
}

/** Word's Review ribbon: track-changes toggle, accept/reject, the live
 * revision count, comments, and find & replace. All state is read from the
 * api at render time; the toolbar already re-renders on selectionchange /
 * dxw-selection (the `refresh` subscription), and each command here calls
 * `onChanged` (that same refresh) so the count and toggle update at once. */
function ReviewTab({
  api,
  onChanged,
  showComment,
  mentions,
}: {
  api: DocxViewApi | null;
  onChanged: () => void;
  showComment: boolean;
  mentions?: string[];
}) {
  // The parent's refresh bails out of re-rendering when the selection format
  // is unchanged (focus sits in the toolbar during accept/reject clicks), and
  // in a session-backed view the command applies asynchronously — so read the
  // count on every dxw-selection announcement with an unconditional re-render
  // of this tab alone.
  const [, force] = useReducer((n: number) => n + 1, 0);
  useEffect(() => {
    const bump = () => force();
    document.addEventListener("dxw-selection", bump);
    document.addEventListener("selectionchange", bump);
    return () => {
      document.removeEventListener("dxw-selection", bump);
      document.removeEventListener("selectionchange", bump);
    };
  }, []);
  const suggesting = api?.isSuggesting() ?? false;
  const count = api?.revisionCount() ?? 0;
  return (
    <>
      <Btn
        label="Track changes"
        title={suggesting ? "Stop tracking changes" : "Record edits as tracked changes"}
        active={suggesting}
        onClick={() => {
          api?.setSuggesting(!suggesting);
          onChanged();
        }}
      />
      <Sep />
      <ActionMenu
        label="Accept"
        title="Accept tracked changes"
        width={72}
        groups={[{ items: [["next", "Accept and move to next"], ["all", "Accept all changes"]] }]}
        onPick={(value) => {
          if (value === "next") api?.acceptRevisionAtCaret();
          else api?.acceptAllRevisions();
          onChanged();
        }}
      />
      <ActionMenu
        label="Reject"
        title="Reject tracked changes"
        width={72}
        groups={[{ items: [["next", "Reject and move to next"], ["all", "Reject all changes"]] }]}
        onPick={(value) => {
          if (value === "next") api?.rejectRevisionAtCaret();
          else api?.rejectAllRevisions();
          onChanged();
        }}
      />
      <span data-dxw-revision-count="" style={{ color: T.muted, font: "12px system-ui, sans-serif", padding: "0 4px", whiteSpace: "nowrap" }}>
        {count === 1 ? "1 change" : `${count} changes`}
      </span>
      <Sep />
      {showComment && <CommentMenu api={api} mentions={mentions} />}
      <Btn label="◀" title="Go to previous comment" onClick={() => api?.stepComment(-1)} />
      <Btn label="▶" title="Go to next comment" onClick={() => api?.stepComment(1)} />
      <FindReplaceMenu api={api} />
    </>
  );
}

type SelectedObjectContext = NonNullable<ReturnType<DocxViewApi["getSelectedObjectContext"]>>;

function SmartArtTextControls({ api, nodeIndex }: { api: DocxViewApi | null; nodeIndex?: number }) {
  const nodeSelected = nodeIndex !== undefined;
  const [format, setFormat] = useState(() => api?.getSelectedSmartArtTextFormat() ?? null);
  useEffect(() => setFormat(api?.getSelectedSmartArtTextFormat() ?? null), [api, nodeIndex]);
  if (!format) return null;
  const apply = (patch: Parameters<DocxViewApi["setSelectedSmartArtTextFormat"]>[0]) => {
    if (api?.setSelectedSmartArtTextFormat(patch)) setFormat(api.getSelectedSmartArtTextFormat());
  };
  const fonts = detectFonts();
  const sizes = SIZES.includes(format.fontSizePt) ? SIZES : [format.fontSizePt, ...SIZES];
  const scope = nodeSelected ? "selected node" : "all nodes";
  return (
    <>
      <Sep />
      <ToolbarMenuSelect
        title={`Font for ${scope}`}
        value={format.fontFamily}
        width={126}
        menuWidth={210}
        options={(fonts.includes(format.fontFamily) ? fonts : [format.fontFamily, ...fonts]).map((font) => ({
          value: font,
          label: font,
          fontFamily: font,
        }))}
        onChange={(fontFamily) => fontFamily && apply({ fontFamily })}
      />
      <ToolbarMenuSelect
        title={`Font size for ${scope}`}
        value={String(format.fontSizePt)}
        width={58}
        menuWidth={92}
        options={sizes.map((size) => ({ value: String(size), label: String(size) }))}
        onChange={(value) => value && apply({ fontSizePt: Number(value) })}
      />
      <Btn label={<b>B</b>} title={`Bold ${scope}`} active={format.bold} onClick={() => apply({ bold: !format.bold })} />
      <Btn label={<i>I</i>} title={`Italic ${scope}`} active={format.italic} onClick={() => apply({ italic: !format.italic })} />
      <ColorMenu
        current={format.color}
        title={`Text color for ${scope}`}
        trigger={<span style={{ fontSize: 13, borderBottom: `3px solid ${format.color}`, padding: "0 3px" }}>A</span>}
        onPick={(color) => apply({ color })}
      />
      <ActionMenu
        label={format.alignment === "center" ? "Center" : format.alignment === "right" ? "Right" : "Left"}
        title={`Text alignment for ${scope}`}
        width={74}
        groups={[{ items: [["left", "Align left"], ["center", "Align center"], ["right", "Align right"]] }]}
        onPick={(alignment) => apply({ alignment: alignment as "left" | "center" | "right" })}
      />
    </>
  );
}

function ObjectFormatTab({
  api,
  context,
  showArrange,
}: {
  api: DocxViewApi | null;
  context: SelectedObjectContext;
  showArrange: boolean;
}) {
  const run = (command: Parameters<DocxViewApi["runSelectedObjectCommand"]>[0]) => api?.runSelectedObjectCommand(command);
  // What this tab offers comes from core's shared offered set, so the buttons
  // rendered here and the capability matrix's `offered` column cannot drift.
  const offered = new Set(availableObjectCommands(context, { arrange: showArrange }));
  const wrapItems = ([
    ["wrapInline", "Inline with text"],
    ["wrapSquare", "Square"],
    ["wrapTopAndBottom", "Top and bottom"],
    ["wrapFront", "In front of text"],
    ["wrapBehind", "Behind text"],
  ] as [Parameters<DocxViewApi["runSelectedObjectCommand"]>[0], string][]).filter(([command]) => offered.has(command));
  const wrap = wrapItems.length > 0 && (
    <ActionMenu
      label="Wrap"
      title="Wrap"
      width={72}
      groups={[{ items: wrapItems }]}
      onPick={(value) => run(value as Parameters<DocxViewApi["runSelectedObjectCommand"]>[0])}
    />
  );
  return (
    <span data-dxw-object-format="" style={{ display: "contents" }}>
      {context.kind === "chart" && <ChartMenu api={api} label="Edit data" />}
      {context.kind === "smartArt" && <SmartArtMenu api={api} label="Edit SmartArt" />}
      {context.kind === "smartArt" && <SmartArtTextControls api={api} nodeIndex={context.smartArtNodeIndex} />}
      {offered.has("editText") && (
        <Btn
          label="Edit text"
          title={context.kind === "smartArt" ? "Edit selected SmartArt node text" : "Edit shape text"}
          onClick={() => run("editText")}
        />
      )}
      {offered.has("fill") && (
        <Btn
          label={context.kind === "smartArt" ? (context.smartArtNodeSelected ? "Node fill" : "Fill all") : "Fill"}
          title={context.kind === "smartArt" ? (context.smartArtNodeSelected ? "Selected SmartArt node fill" : "All SmartArt node fills") : "Fill color"}
          onClick={() => run("fill")}
        />
      )}
      {offered.has("outline") && <Btn label="Outline" title="Outline color, weight, and style" onClick={() => run("outline")} />}
      {offered.has("lineStyle") && <Btn label="Line style" title="Line color, weight, and style" onClick={() => run("lineStyle")} />}
      {offered.has("altText") && <Btn label="Alt text" title="Alternative text" onClick={() => run("altText")} />}
      {wrap}
      {offered.has("size") && <Btn label="Size" title="Exact size" onClick={() => run("size")} />}
      {offered.has("position") && <Btn label="Position" title="Exact page position" onClick={() => run("position")} />}
      {offered.has("crop") && <Btn label="Crop" title="Crop to part of the picture" onClick={() => run("crop")} />}
      {offered.has("rotate") && <Btn label="Rotate" title="Set rotation" onClick={() => run("rotate")} />}
      {offered.has("bringForward") && (
        <>
          <Btn label="Bring forward" title="Bring selected object forward" onClick={() => run("bringForward")} />
          <Btn label="Send backward" title="Send selected object backward" onClick={() => run("sendBackward")} />
        </>
      )}
      {offered.has("reset3d") && <Btn label="Reset 3D" title="Reset 3D rotation" onClick={() => run("reset3d")} />}
      {offered.has("delete") && <Btn label="Delete" title="Delete selected object" onClick={() => run("delete")} />}
    </span>
  );
}

/** Toolbar control groups a host can disable via the `features` prop. */
export type ToolbarFeature =
  | "history"
  | "styles"
  | "charStyles"
  | "formatPainter"
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
  | "borders"
  | "table"
  | "image"
  | "icon"
  | "screenshot"
  | "model3D"
  | "media"
  | "object"
  | "chart"
  | "smartArt"
  | "comment"
  | "footnote"
  | "bookmark"
  | "crossReference"
  | "dateTime"
  | "field"
  | "citations"
  | "equation"
  | "symbol"
  | "shape"
  | "divider"
  | "textBox"
  | "wordArt"
  | "drawing"
  | "arrange"
  | "dropCap"
  | "headerFooter"
  | "watermark"
  | "coverPage"
  | "pageNumber"
  | "break"
  | "layout"
  | "review"
  | "help"
  | "download";

/**
 * Every `DocxViewApi` command that INSERTS content at the caret, paired with
 * the toolbar feature group that gates its control.
 *
 * A single source of truth in the same spirit as core's
 * `SELECTED_OBJECT_COMMANDS`: the collab audit (react/test/insert-commands
 * INVARIANT C) reads this list and the real `COLLAB_TOOLBAR_DEFAULTS` gate
 * rather than restating either, so "what the toolbar offers in a room" and
 * "what the audit checks" cannot drift apart.
 *
 * THE RULE a new entry signs up to: in collab mode an insert command must
 * EMIT an intent, or its feature must be gated off so no control exists to
 * press. An absent button is honest; a present button that mutates the local
 * document without emitting forks the room silently — the exact shape that
 * shipped in `insertImage` and stayed green through the whole capability
 * matrix, because the matrix's collab mount withholds `submitOp` and so
 * never exercises an insert command's collab path.
 */
export interface InsertCommandSpec {
  /** The `DocxViewApi` method name. */
  command: string;
  /** Toolbar group whose gate decides whether a control is offered. */
  feature: ToolbarFeature;
}

export const INSERT_COMMANDS: readonly InsertCommandSpec[] = [
  { command: "insertTable", feature: "table" },
  { command: "insertImage", feature: "image" },
  { command: "insertScreenshot", feature: "screenshot" },
  { command: "insertModel3D", feature: "model3D" },
  { command: "insertOnlineVideo", feature: "media" },
  { command: "insertEmbeddedObject", feature: "object" },
  { command: "insertChart", feature: "chart" },
  { command: "insertSmartArt", feature: "smartArt" },
  { command: "insertShape", feature: "shape" },
  { command: "insertWordArt", feature: "wordArt" },
  { command: "insertEquation", feature: "equation" },
  { command: "insertSymbol", feature: "symbol" },
  { command: "insertPageNumber", feature: "pageNumber" },
  { command: "insertField", feature: "field" },
  // A table of contents IS a field, and it lands in the field group. It
  // emits like any other insert; the entry count rides in the payload so the
  // carried id allocation can be sized for a mutation whose size comes from
  // the document rather than from its arguments.
  { command: "insertToc", feature: "field" },
  // Citations are fields too, but they get their own group: the whole
  // References cluster (insert citation, bibliography, source manager,
  // style) hangs together and a host hides or shows it as one.
  { command: "insertCitation", feature: "citations" },
  { command: "insertBibliography", feature: "citations" },
  { command: "insertDateTime", feature: "dateTime" },
  { command: "insertCrossReference", feature: "crossReference" },
  { command: "insertCrossRefToTarget", feature: "crossReference" },
  { command: "insertCaption", feature: "crossReference" },
  { command: "insertBreak", feature: "break" },
  { command: "insertBlankPage", feature: "break" },
  { command: "insertCoverPage", feature: "coverPage" },
  // A watermark lands in the header parts rather than at the caret, so it
  // needs no addressable position — but it is an insert, and the rule the
  // list exists for applies to it unchanged.
  { command: "insertWatermark", feature: "watermark" },
  { command: "addComment", feature: "comment" },
  { command: "addFootnote", feature: "footnote" },
  { command: "addEndnote", feature: "footnote" },
  { command: "addBookmark", feature: "bookmark" },
];

export type ToolbarMode = "simple" | "advanced";

export interface DocxToolbarProps {
  api: DocxViewApi | null;
  onSave?: (bytes: Uint8Array) => void;
  /** Simple shows basic Home editing; advanced adds the Insert, Draw, Layout, and Review ribbons. */
  mode?: ToolbarMode;
  /** Per-group overrides; every group defaults to enabled. */
  features?: Partial<Record<ToolbarFeature, boolean>>;
  /** Extra class on the toolbar root (e.g. a scope for CSS-variable overrides). */
  className?: string;
  /** Inline overrides merged onto the toolbar root; wins over the defaults. */
  style?: React.CSSProperties;
  /** Connected collaborator names offered as @mention shortcuts in comments. */
  commentMentions?: string[];
  /** Start with the toolbar expanded (all groups inline, wrapping onto extra
   * rows). A choice the user makes via the chevron toggle is persisted in
   * localStorage and wins over this default. */
  defaultExpanded?: boolean;
}

/** localStorage key for the expand/collapse chevron choice. */
const EXPANDED_KEY = "dxw-toolbar-expanded";

export function DocxToolbar({
  api,
  onSave,
  mode = "advanced",
  features,
  className,
  style,
  commentMentions,
  defaultExpanded = false,
}: DocxToolbarProps) {
  const on = (k: ToolbarFeature) => features?.[k] !== false;
  // Ribbon-style tabs: complex tool groups get their own surface instead of
  // one overloaded row (Layout especially).
  type NormalTab = "home" | "insert" | "draw" | "layout" | "review";
  const [tab, setTab] = useState<NormalTab | "format" | "tableFormat">("home");
  const priorNormalTab = useRef<NormalTab>("home");
  const [objectContext, setObjectContext] = useState<SelectedObjectContext | null>(null);
  const [tableCellFill, setTableCellFill] = useState<string | null | undefined>(undefined);
  const tableContextRef = useRef<string | null | undefined>(undefined);
  const selectNormalTab = (next: NormalTab) => {
    priorNormalTab.current = next;
    setTab(next);
  };
  const [helpOpen, setHelpOpen] = useState(false);
  const [tabStopsOpen, setTabStopsOpen] = useState(false);
  const tabStopsAnchorRef = useRef<HTMLSpanElement | null>(null);
  const [paraBordersOpen, setParaBordersOpen] = useState(false);
  const paraBordersAnchorRef = useRef<HTMLSpanElement | null>(null);
  const apple = typeof navigator !== "undefined" && /Mac|iPhone|iPad/.test(navigator.platform);
  const shortcut = (key: string) => apple ? `⌘${key}` : `Ctrl+${key}`;
  const closeHelp = useCallback(() => setHelpOpen(false), []);
  const iconInput = useRef<HTMLInputElement | null>(null);
  const modelInput = useRef<HTMLInputElement | null>(null);
  const objectInput = useRef<HTMLInputElement | null>(null);
  const helpTrigger = useRef<HTMLButtonElement | null>(null);
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
  // Format painter: the copied formatting, held until it is painted or cleared.
  // Word keeps it on the toolbar rather than the document, so it lives here.
  const [painted, setPainted] = useState<SelectionFormat | null>(null);
  // Toolbar popovers can move focus away from the document selection; remember
  // the last real range and restore it before applying their choice.
  const savedRange = useRef<Range | null>(null);
  const imageInput = useRef<HTMLInputElement | null>(null);
  /**
   * Why the picked image did not land, shown beside the Pictures button.
   *
   * Every one of these was a silent `return` in insertImage, and a user hit
   * the format one: they chose an SVG through a picker that advertised SVG,
   * and the document did nothing — no skeleton, no error, no log. The picker
   * no longer offers what a shared document refuses (see `accept` below), so
   * these are now the paths reachable through the API or through a genuine
   * failure, and each of them says something.
   */
  const [imageStatus, setImageStatus] = useState("");
  useEffect(() => {
    if (!imageStatus) return;
    const t = setTimeout(() => setImageStatus(""), 8000);
    return () => clearTimeout(t);
  }, [imageStatus]);
  const insertPicture = async (file: File) => {
    setImageStatus("");
    const result = await api?.insertImage(file);
    if (result === undefined || result === "inserted") return;
    setImageStatus(
      result === "too-large"
        // The CONFIGURED number, never a constant: this deployment's cap is
        // whatever the server published (5MB in the public compose file, 50MB
        // in the dev stack), and a hardcoded figure would be wrong in one of
        // them. If no limit was published there is no number to name, and the
        // message says nothing about size rather than inventing one.
        ? (() => {
            const max = api?.imageMaxBytes() ?? null;
            return max === null
              ? "That image is too large for this document."
              : `Images must be under ${formatBytes(max)}.`;
          })()
        : result === "unsupported-format"
        ? "Images in a shared document must be PNG, JPEG, GIF, BMP or WebP."
        : result === "no-relay"
          ? "This shared document has no image relay, so images can’t be added to it."
          : result === "upload-failed"
            // DELIBERATELY DOES NOT NAME A SIZE. The relay's 413 carries the
            // configured maxBytes, but that number does not reach this layer
            // today (it would have to be threaded through the media client and
            // both connection classes), and the cap differs per deployment —
            // 5MB in the published compose file, 50MB in the dev stack. So the
            // copy points at the likely cause and the action that fixes it
            // without asserting a limit it cannot actually know.
            ? "Upload failed. The image may be too large, or the connection dropped. Try a smaller image."
            : result === "no-caret"
              ? "Click in the document before inserting an image."
              : "That file could not be read as an image.",
    );
  };
  // Responsive collapse: measure the toolbar width and pick a tier; the higher
  // the tier the more low-frequency Home groups fold into the ⋮ overflow menu,
  // so the strip stays single-row-clean on phones and tablets (Google Docs
  // does exactly this). Full width keeps everything inline, so desktop and the
  // e2e specs (1400px) are unchanged.
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [tier, setTier] = useState(0);
  // Expanded: the chevron at the right edge disables the overflow folding so
  // every group of the active tab stays inline and wraps onto extra rows (the
  // toolbar grows downward). Persisted so the choice survives reloads.
  const [expanded, setExpanded] = useState(() => {
    try {
      const stored = localStorage.getItem(EXPANDED_KEY);
      if (stored !== null) return stored === "1";
    } catch {
      /* storage may be unavailable (SSR, privacy mode) */
    }
    return defaultExpanded;
  });
  const toggleExpanded = () =>
    setExpanded((prev) => {
      const next = !prev;
      try {
        localStorage.setItem(EXPANDED_KEY, next ? "1" : "0");
      } catch {
        /* storage may be unavailable */
      }
      return next;
    });
  const effectiveTier = expanded ? 0 : tier;
  useEffect(() => {
    if (!on("help")) return;
    const keydown = (event: KeyboardEvent) => {
      if (event.key === "F1" || ((event.metaKey || event.ctrlKey) && event.key === "/")) {
        event.preventDefault();
        setHelpOpen(true);
      }
    };
    document.addEventListener("keydown", keydown);
    return () => document.removeEventListener("keydown", keydown);
  }, [features]);
  useEffect(() => {
    const refreshObject = () => {
      const next = api?.getSelectedObjectContext() ?? null;
      setObjectContext(next);
      setTab((current) => {
        if (next) {
          if (["home", "insert", "draw", "layout", "review"].includes(current)) priorNormalTab.current = current as NormalTab;
          return "format";
        }
        return current === "format"
          ? (tableContextRef.current !== undefined ? "tableFormat" : priorNormalTab.current)
          : current;
      });
    };
    refreshObject();
    document.addEventListener("dxw-object-selection", refreshObject);
    return () => document.removeEventListener("dxw-object-selection", refreshObject);
  }, [api]);
  useEffect(() => {
    const el = rootRef.current;
    if (!el || typeof ResizeObserver === "undefined") return;
    const measure = () => {
      const w = Math.min(el.clientWidth, window.innerWidth);
      setTier(w >= 1280 ? 0 : w >= 720 ? 1 : 2);
    };
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    window.addEventListener("resize", measure);
    measure();
    return () => {
      ro.disconnect();
      window.removeEventListener("resize", measure);
    };
  }, []);

  const refresh = useCallback(() => {
    const sel = window.getSelection();
    if (sel && !sel.isCollapsed && sel.rangeCount > 0) {
      savedRange.current = sel.getRangeAt(0).cloneRange();
    }
    setFmt(api?.getSelectionFormat() ?? null);
    setCurStyle(api?.getParagraphStyleId?.() ?? null);
    setListKind(api?.getListType?.() ?? null);
    const nextTableFill = api?.getTableCellFill();
    const wasInTable = tableContextRef.current !== undefined;
    tableContextRef.current = nextTableFill;
    setTableCellFill(nextTableFill);
    setTab((current) => {
      if (nextTableFill !== undefined && !wasInTable && current !== "format") {
        if (["home", "insert", "draw", "layout", "review"].includes(current)) priorNormalTab.current = current as NormalTab;
        return "tableFormat";
      }
      if (nextTableFill === undefined && current === "tableFormat") return priorNormalTab.current;
      return current;
    });
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

  // Home-tab controls as ordered groups so the low-frequency ones can fold into
  // the ⋮ overflow menu as the toolbar narrows. Keys drive the per-tier split.
  const renderHome = () => {
    const groups: { key: string; node: React.ReactNode }[] = [];
    if (on("history"))
      groups.push({
        key: "history",
        node: (
          <>
            <Btn label={"↶"} title="Undo (⌘Z)" onClick={() => { api?.undo(); refresh(); }} />
            <Btn label={"↷"} title="Redo (⇧⌘Z)" onClick={() => { api?.redo(); refresh(); }} />
            <Sep />
          </>
        ),
      });
    if (on("styles"))
      groups.push({
        key: "styles",
        node: (
          <>
          <ToolbarMenuSelect
            title="Paragraph style"
            value={curStyle ?? "__normal"}
            width={92}
            menuWidth={190}
            options={[
              { value: "__normal", label: "Normal" },
              ...(api?.listParagraphStyles() ?? [])
                .filter((style) => !/^normal$/i.test(style.name))
                .map((style) => ({ value: style.id, label: style.name })),
              ...(curStyle !== null && !(api?.listParagraphStyles() ?? []).some((style) => style.id === curStyle)
                ? [{ value: curStyle, label: api?.document.styles.byId.get(curStyle)?.name ?? curStyle }]
                : []),
            ]}
            onChange={(value) => {
              if (value) {
                api?.setParagraphStyle(value === "__normal" ? null : value);
                setCurStyle(api?.getParagraphStyleId?.() ?? null);
              }
            }}
          />
          <StylesPane
            api={api}
            onChanged={() => {
              refresh();
              setCurStyle(api?.getParagraphStyleId?.() ?? null);
            }}
          />
          </>
        ),
      });
    if (on("charStyles")) {
      const charStyles = api?.listStyles?.({ type: "character" }) ?? [];
      // A document with no character styles beyond Word's hidden defaults has
      // nothing to offer, so the control stays out of the way entirely.
      const offered = charStyles.filter((style) => style.quickStyle || style.usageCount > 0);
      if (offered.length > 0)
        groups.push({
          key: "charStyles",
          node: (
            <ToolbarMenuSelect
              title="Character style"
              value={fmt?.characterStyleId ?? "__none"}
              width={86}
              menuWidth={190}
              options={[
                { value: "__none", label: "None" },
                ...offered.map((style) => ({ value: style.id, label: style.name })),
              ]}
              onChange={(value) => {
                if (value) apply({ characterStyleId: value === "__none" ? null : value });
              }}
            />
          ),
        });
    }
    if (on("formatPainter"))
      groups.push({
        key: "formatPainter",
        node: (
          <Btn
            label={"\u{1F58C}"}
            title={painted ? "Paint the copied formatting" : "Copy formatting (format painter)"}
            active={!!painted}
            onClick={() => {
              // One button, two halves: the first click copies, the next
              // paints and hands the brush back — Word's single-use painter.
              if (painted) {
                api?.applyCopiedFormatting?.(painted);
                setPainted(null);
                setFmt(api?.getSelectionFormat() ?? null);
              } else {
                setPainted(api?.copyFormatting?.() ?? null);
              }
            }}
          />
        ),
      });
    if (on("font"))
      groups.push({
        key: "font",
        node: (
          <ToolbarMenuSelect
            title="Font"
            value={fmt?.fontFamily ?? ""}
            placeholder="Font"
            width={130}
            menuWidth={210}
            options={(fmt?.fontFamily && !detectFonts().includes(fmt.fontFamily) ? [fmt.fontFamily, ...detectFonts()] : detectFonts()).map((font) => ({
              value: font,
              label: font,
              fontFamily: font,
            }))}
            onChange={(value) => value && apply({ fontFamily: value })}
          />
        ),
      });
    if (on("size"))
      groups.push({
        key: "size",
        node: (
          <>
          <ToolbarMenuSelect
            title="Font size"
            value={fmt?.fontSizePt === undefined ? "" : String(fmt.fontSizePt)}
            placeholder="Size"
            width={58}
            menuWidth={92}
            options={SIZES.map((size) => ({ value: String(size), label: String(size) }))}
            onChange={(value) => value && apply({ fontSizePt: parseFloat(value) })}
          />
          <Sep />
          </>
        ),
      });
    if (on("format"))
      groups.push({
        key: "format",
        node: (
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
        ),
      });
    if (on("color"))
      groups.push({
        key: "color",
        node: (
          <ColorMenu
            current={fmt?.color && fmt.color !== "auto" ? fmt.color : "#000000"}
            title="Text color"
            trigger={<span style={{ fontSize: 13, borderBottom: `3px solid ${fmt?.color && fmt.color !== "auto" ? fmt.color : "#000"}`, padding: "0 3px", color: T.fg }}>A</span>}
            onPick={(value) => apply({ color: value })}
          />
        ),
      });
    if (on("highlight"))
      groups.push({
        key: "highlight",
        node: (
          <>
            <HighlightMenu current={fmt?.highlight} onPick={(v) => apply({ highlight: v })} />
            <Sep />
          </>
        ),
      });
    if (on("alignment"))
      groups.push({
        key: "alignment",
        node: (
          <>
            <Btn label={"≡"} title="Align left" onClick={() => api?.setAlignment("left")} />
            <Btn label={"≣"} title="Center" onClick={() => api?.setAlignment("center")} />
            <Btn label={"≢"} title="Align right" onClick={() => api?.setAlignment("right")} />
            <Btn label={"☰"} title="Justify" onClick={() => api?.setAlignment("justify")} />
            <Sep />
          </>
        ),
      });
    if (on("indent"))
      groups.push({
        key: "indent",
        node: (
          <>
            <Btn label={<IndentIcon dir={-1} />} title="Decrease indent" onClick={() => api?.adjustIndent(-1)} />
            <Btn label={<IndentIcon dir={1} />} title="Increase indent" onClick={() => api?.adjustIndent(1)} />
            <span ref={tabStopsAnchorRef} style={{ display: "inline-flex" }}>
              <Btn label={"⇥"} title="Tab stops" onClick={() => setTabStopsOpen(true)} />
            </span>
            {tabStopsOpen && (
              <TabStopsDialog api={api} anchorRef={tabStopsAnchorRef} onClose={() => setTabStopsOpen(false)} />
            )}
          </>
        ),
      });
    if (on("spacing"))
      groups.push({
        key: "spacing",
        node: (
          <ActionMenu
            label="↕"
            title="Line & paragraph spacing"
            width={44}
            groups={[
              { label: "Line spacing", items: [["l:1", "Single"], ["l:1.15", "1.15"], ["l:1.5", "1.5"], ["l:2", "Double"]] },
              { label: "Exact line height", items: [["e:12", "Exactly 12 pt"], ["e:18", "Exactly 18 pt"], ["e:24", "Exactly 24 pt"], ["e:custom", "Custom exact height…"]] },
              { label: "Paragraph", items: [["b:add", "Add space before"], ["b:none", "Remove space before"], ["a:add", "Add space after"], ["a:none", "Remove space after"]] },
            ]}
            onPick={(v) => {
              if (v.startsWith("l:")) api?.setParagraphSpacing({ lineMultiple: parseFloat(v.slice(2)) });
              else if (v === "e:custom") {
                const anchor = rootRef.current;
                if (!anchor) return;
                void requestTextInputDialog(anchor, {
                  title: "Exact line height",
                  label: "Line height (points)",
                  value: "24",
                  submitLabel: "Apply",
                  inputType: "number",
                  min: 1,
                  step: 0.5,
                }).then((next) => {
                  if (next === null) return;
                  const points = Number(next.trim());
                  if (Number.isFinite(points) && points > 0) api?.setParagraphSpacing({ exactLinePt: points });
                });
              } else if (v.startsWith("e:")) api?.setParagraphSpacing({ exactLinePt: parseFloat(v.slice(2)) });
              else if (v === "b:add") api?.setParagraphSpacing({ beforePt: 10 });
              else if (v === "b:none") api?.setParagraphSpacing({ beforePt: 0 });
              else if (v === "a:add") api?.setParagraphSpacing({ afterPt: 10 });
              else if (v === "a:none") api?.setParagraphSpacing({ afterPt: 0 });
            }}
          />
        ),
      });
    if (on("lists"))
      groups.push({
        key: "lists",
        node: (
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
            <ActionMenu
              label="⇶"
              title="Multilevel list gallery"
              width={44}
              groups={[{
                label: "Multilevel list",
                items: Object.entries(NUMBERING_PRESETS).map(([id, preset]) => [id, preset.name] as [string, string]),
              }]}
              onPick={(value) => { api?.applyNumberingPreset(value as NumberingPresetId); refresh(); }}
            />
            <Sep />
          </>
        ),
      });
    if (on("borders"))
      groups.push({
        key: "borders",
        node: (
          <>
            <span ref={paraBordersAnchorRef} style={{ display: "inline-flex" }}>
              <ActionMenu
                label="▦"
                title="Paragraph borders and shading"
                width={44}
                groups={[
                  {
                    label: "Borders",
                    items: [
                      ["outside", "Outside borders"],
                      ["all", "All borders"],
                      ["top", "Top border"],
                      ["bottom", "Bottom border"],
                      ["left", "Left border"],
                      ["right", "Right border"],
                      ["none", "No border"],
                    ],
                  },
                  { items: [["custom", "Borders and shading…"]] },
                ]}
                onPick={(v) => {
                  const rule = { style: "single" as const, sz: 4, color: "auto" };
                  if (v === "custom") setParaBordersOpen(true);
                  else if (v === "none")
                    api?.setParagraphBorders({ borders: { top: null, left: null, bottom: null, right: null, between: null, bar: null } });
                  else if (v === "outside")
                    api?.setParagraphBorders({ borders: { top: rule, bottom: rule, left: rule, right: rule } });
                  else if (v === "all")
                    api?.setParagraphBorders({ borders: { top: rule, bottom: rule, left: rule, right: rule, between: rule } });
                  else api?.setParagraphBorders({ borders: { [v]: rule } });
                }}
              />
            </span>
            {paraBordersOpen && (
              <ParagraphBorderDialog api={api} anchorRef={paraBordersAnchorRef} onClose={() => setParaBordersOpen(false)} />
            )}
            <Sep />
          </>
        ),
      });

    // Per-tier overflow: which group keys fold into ⋮. Tier 0 keeps all inline.
    const overflowKeys =
      effectiveTier === 0
        ? new Set<string>()
        : effectiveTier === 1
          ? new Set(["styles", "indent", "spacing"])
          : new Set(["styles", "font", "size", "color", "highlight", "alignment", "indent", "spacing"]);
    const inline = groups.filter((g) => !overflowKeys.has(g.key));
    const overflow = groups.filter((g) => overflowKeys.has(g.key));
    return (
      <>
        {inline.map((g) => (
          <Fragment key={g.key}>{g.node}</Fragment>
        ))}
        {overflow.length > 0 && (
          <OverflowMenu>
            {overflow.map((g) => (
              <Fragment key={g.key}>{g.node}</Fragment>
            ))}
          </OverflowMenu>
        )}
      </>
    );
  };

  return (
    <div
      ref={rootRef}
      className={className}
      data-dxw-toolbar-mode={mode}
      onMouseOver={onTipOver}
      onMouseOut={onTipOut}
      onMouseDownCapture={onTipOut}
      style={{
        position: "relative",
        zIndex: "var(--dxw-toolbar-z-index, 100)",
        display: "flex",
        width: "100%",
        maxWidth: "100%",
        minWidth: 0,
        boxSizing: "border-box",
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
      {mode === "advanced" && (
        <>
          <div style={{ display: "flex", gap: 2, marginRight: 8 }}>
            {(["home", "insert", "draw", "layout", "review"] as const)
              .filter((t) => (t !== "draw" || on("drawing")) && (t !== "layout" || on("layout")) && (t !== "review" || on("review")))
              .map((t) => (
              <button
                key={t}
                data-tab={t}
                aria-pressed={tab === t}
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => selectNormalTab(t)}
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
            {objectContext && (
              <button
                data-tab="format"
                aria-pressed={tab === "format"}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => setTab("format")}
                style={{
                  border: "none",
                  background: tab === "format" ? T.tabActiveBg : "transparent",
                  color: tab === "format" ? T.accent : T.fg,
                  font: "600 12.5px system-ui, sans-serif",
                  padding: "5px 10px",
                  borderRadius: 6,
                  cursor: "pointer",
                }}
              >
                {{
                  shape: "Shape Format",
                  line: "Line Format",
                  smartArt: "SmartArt Format",
                  chart: "Chart Format",
                  image: "Picture Format",
                  model3d: "3D Format",
                }[objectContext.kind]}
              </button>
            )}
            {!objectContext && tableCellFill !== undefined && (
              <button
                data-tab="tableFormat"
                aria-pressed={tab === "tableFormat"}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => setTab("tableFormat")}
                style={{
                  border: "none",
                  background: tab === "tableFormat" ? T.tabActiveBg : "transparent",
                  color: tab === "tableFormat" ? T.accent : T.fg,
                  font: "600 12.5px system-ui, sans-serif",
                  padding: "5px 10px",
                  borderRadius: 6,
                  cursor: "pointer",
                }}
              >
                Table Format
              </button>
            )}
            {on("help") && (
              <button
                ref={helpTrigger}
                type="button"
                title={`Help and keyboard shortcuts (${shortcut("/")})`}
                aria-haspopup="dialog"
                aria-expanded={helpOpen}
                data-dxw-help-trigger=""
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => setHelpOpen(true)}
                style={{
                  border: "none",
                  background: "transparent",
                  color: T.fg,
                  font: "600 12.5px system-ui, sans-serif",
                  padding: "5px 10px",
                  borderRadius: 6,
                  cursor: "pointer",
                }}
              >
                Help
              </button>
            )}
          </div>
          <Sep />
        </>
      )}
      {(mode === "simple" || tab === "home") && renderHome()}
      {mode === "advanced" && tab === "insert" && (
        <>
          {on("coverPage") && <CoverPageMenu api={api} />}
          {on("table") && <TableMenu api={api} />}
          {on("image") && (
            <span style={{ position: "relative", display: "inline-flex" }}>
              <Btn label={<ImageIcon />} title="Insert image" onClick={() => imageInput.current?.click()} />
              {imageStatus && (
                <span
                  role="alert"
                  data-dxw-image-status=""
                  style={{ position: "absolute", top: 30, left: 0, zIndex: 120, width: 230, padding: "6px 8px", border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, boxShadow: T.popoverShadow, color: T.fg, font: "12px system-ui, sans-serif" }}
                >
                  {imageStatus}
                </span>
              )}
            </span>
          )}
          {on("icon") && <Btn label="Icons" title="Insert SVG icon" onClick={() => iconInput.current?.click()} />}
          {on("screenshot") && <ScreenshotButton api={api} />}
          {effectiveTier === 0 ? (
            <>
          {on("model3D") && <Btn label="3D Models" title="Insert a GLB 3D model" onClick={() => modelInput.current?.click()} />}
          {on("smartArt") && <SmartArtMenu api={api} />}
          {on("chart") && <ChartMenu api={api} />}
          {on("media") && <MediaMenu api={api} />}
          {on("shape") && <ShapeMenu api={api} />}
          {on("divider") && <DividerMenu api={api} />}
          {on("textBox") && <TextBoxMenu api={api} />}
          {on("wordArt") && <WordArtMenu api={api} />}
          {on("link") && <LinkMenu api={api} />}
          {on("comment") && <CommentMenu api={api} mentions={commentMentions} />}
          {on("footnote") && <NoteMenu api={api} kind="footnote" />}
          {on("footnote") && <NoteMenu api={api} kind="endnote" />}
          {on("bookmark") && <BookmarkMenu api={api} />}
          {on("crossReference") && <CrossReferenceMenu api={api} />}
          {on("crossReference") && <CaptionMenu api={api} />}
          {on("headerFooter") && <HeaderFooterMenu api={api} />}
          {on("watermark") && <WatermarkMenu api={api} />}
          <Sep />
          {on("pageNumber") && <PageNumberMenu api={api} />}
          {on("break") && (
            <>
              <Btn label="Blank page" title="Insert blank page" onClick={() => api?.insertBlankPage()} />
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
          {on("dateTime") && (
            <ActionMenu
              label="Date & time"
              title="Insert an automatically updating date or time"
              width={100}
              groups={[
                { label: "Date", items: [["date:short", "Short date"], ["date:long", "Long date"], ["date:intl", "Day month year"]] },
                { label: "Time", items: [["time:12", "12-hour time"], ["time:24", "24-hour time"]] },
              ]}
              onPick={(value) => {
                if (value === "date:short") api?.insertDateTime("date", "M/d/yyyy");
                else if (value === "date:long") api?.insertDateTime("date", "MMMM d, yyyy");
                else if (value === "date:intl") api?.insertDateTime("date", "d MMMM yyyy");
                else if (value === "time:12") api?.insertDateTime("time", "h:mm am/pm");
                else if (value === "time:24") api?.insertDateTime("time", "HH:mm");
              }}
            />
          )}
          {on("field") && (
            <ActionMenu
              label="Field"
              title="Insert a Word field"
              width={68}
              groups={[{ items: [["PAGE", "Current page"], ["NUMPAGES", "Number of pages"], ["DATE", "Current date"], ["TIME", "Current time"]] }]}
              onPick={(value) => api?.insertField(`${value} \\* MERGEFORMAT`)}
            />
          )}
          {on("field") && <ContentsMenu api={api} />}
          {on("citations") && <CitationsMenu api={api} />}
          {on("equation") && <EquationMenu api={api} />}
          {on("symbol") && <SymbolMenu api={api} />}
          {on("dropCap") && (
            <ActionMenu
              label="Drop cap"
              title="Drop cap"
              width={84}
              groups={[{ items: [["drop", "Dropped"], ["margin", "In margin"], ["none", "None"]] }]}
              onPick={(value) => api?.setDropCap(value === "none" ? null : value as "drop" | "margin")}
            />
          )}
          {on("object") && <Btn label="Object" title="Embed a file in this document" onClick={() => objectInput.current?.click()} />}
            </>
          ) : (
            <OverflowMenu>
              {on("model3D") && <Btn label="3D Models" title="Insert a GLB 3D model" onClick={() => modelInput.current?.click()} />}
              {on("smartArt") && <SmartArtMenu api={api} />}
              {on("chart") && <ChartMenu api={api} />}
              {on("media") && <MediaMenu api={api} />}
              {on("shape") && <ShapeMenu api={api} />}
              {on("divider") && <DividerMenu api={api} />}
              {on("textBox") && <TextBoxMenu api={api} />}
              {on("wordArt") && <WordArtMenu api={api} />}
              {on("link") && <LinkMenu api={api} />}
              {on("comment") && <CommentMenu api={api} mentions={commentMentions} />}
              {on("footnote") && <NoteMenu api={api} kind="footnote" />}
              {on("footnote") && <NoteMenu api={api} kind="endnote" />}
              {on("bookmark") && <BookmarkMenu api={api} />}
              {on("crossReference") && <CrossReferenceMenu api={api} />}
              {on("crossReference") && <CaptionMenu api={api} />}
              {on("headerFooter") && <HeaderFooterMenu api={api} />}
              {on("watermark") && <WatermarkMenu api={api} />}
              {on("pageNumber") && <PageNumberMenu api={api} />}
              {on("break") && (
                <>
                  <Btn label="Blank page" title="Insert blank page" onClick={() => api?.insertBlankPage()} />
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
              {on("dateTime") && (
                <ActionMenu
                  label="Date & time"
                  title="Insert an automatically updating date or time"
                  width={100}
                  groups={[
                    { label: "Date", items: [["date:short", "Short date"], ["date:long", "Long date"], ["date:intl", "Day month year"]] },
                    { label: "Time", items: [["time:12", "12-hour time"], ["time:24", "24-hour time"]] },
                  ]}
                  onPick={(value) => {
                    if (value === "date:short") api?.insertDateTime("date", "M/d/yyyy");
                    else if (value === "date:long") api?.insertDateTime("date", "MMMM d, yyyy");
                    else if (value === "date:intl") api?.insertDateTime("date", "d MMMM yyyy");
                    else if (value === "time:12") api?.insertDateTime("time", "h:mm am/pm");
                    else if (value === "time:24") api?.insertDateTime("time", "HH:mm");
                  }}
                />
              )}
              {on("field") && (
                <ActionMenu label="Field" title="Insert a Word field" width={68} groups={[{ items: [["PAGE", "Current page"], ["NUMPAGES", "Number of pages"], ["DATE", "Current date"], ["TIME", "Current time"]] }]} onPick={(value) => api?.insertField(`${value} \\* MERGEFORMAT`)} />
              )}
              {on("field") && <ContentsMenu api={api} />}
              {on("citations") && <CitationsMenu api={api} />}
              {on("equation") && <EquationMenu api={api} />}
              {on("symbol") && <SymbolMenu api={api} />}
              {on("dropCap") && <ActionMenu label="Drop cap" title="Drop cap" width={84} groups={[{ items: [["drop", "Dropped"], ["margin", "In margin"], ["none", "None"]] }]} onPick={(value) => api?.setDropCap(value === "none" ? null : value as "drop" | "margin")} />}
              {on("object") && <Btn label="Object" title="Embed a file in this document" onClick={() => objectInput.current?.click()} />}
            </OverflowMenu>
          )}
        </>
      )}
      {mode === "advanced" && tab === "draw" && on("drawing") && <DrawTab api={api} />}
      {mode === "advanced" && tab === "format" && objectContext && (
        <ObjectFormatTab api={api} context={objectContext} showArrange={on("arrange")} />
      )}
      {mode === "advanced" && tab === "tableFormat" && tableCellFill !== undefined && (
        <TableFormatTab api={api} fill={tableCellFill} onChanged={refresh} />
      )}
      <input
        ref={imageInput}
        type="file"
        // ASKED, not asserted: a shared document can't carry every format a
        // local one can, and this picker used to advertise SVG that the
        // collab insert then silently refused. The document answers for
        // itself so the offer and the acceptance cannot drift apart again.
        accept={api?.imageAccept() ?? "image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml"}
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void insertPicture(f);
          e.target.value = "";
        }}
      />
      <input
        ref={iconInput}
        type="file"
        accept=".svg"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void insertPicture(f);
          e.target.value = "";
        }}
      />
      <input
        ref={modelInput}
        aria-label="3D model file"
        type="file"
        accept=".glb,model/gltf-binary"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void api?.insertModel3D(f);
          e.target.value = "";
        }}
      />
      <input
        ref={objectInput}
        aria-label="Embedded object file"
        type="file"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void api?.insertEmbeddedObject(f);
          e.target.value = "";
        }}
      />
      {mode === "advanced" && tab === "layout" && on("layout") && <LayoutTab api={api} showArrange={on("arrange")} />}
      {mode === "advanced" && tab === "review" && on("review") && (
        <ReviewTab api={api} onChanged={refresh} showComment={on("comment")} mentions={commentMentions} />
      )}
      {mode === "simple" && on("help") && (
        <span style={{ marginLeft: "auto" }}>
          <Btn buttonRef={helpTrigger} label="Help" title={`Help and keyboard shortcuts (${shortcut("/")})`} onClick={() => setHelpOpen(true)} />
        </span>
      )}
      {on("download") && onSave && (
        <>
          <span style={{ flex: 1 }} />
          <Btn label="Download" title="Save edited .docx" onClick={() => api && onSave(api.save())} />
        </>
      )}
      <button
        type="button"
        title={expanded ? "Collapse the toolbar" : "Expand the toolbar"}
        aria-expanded={expanded}
        data-dxw-toolbar-expand=""
        style={{ ...btnStyle(expanded), marginLeft: "auto", flexShrink: 0 }}
        onMouseDown={(e) => e.preventDefault()}
        onClick={toggleExpanded}
      >
        <ExpandChevronIcon up={expanded} />
      </button>
      <HelpGuide open={helpOpen} onClose={closeHelp} returnFocus={helpTrigger} />
    </div>
  );
}
