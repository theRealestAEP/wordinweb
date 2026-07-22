import { Fragment, useCallback, useEffect, useId, useLayoutEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { requestTextInputDialog } from "@wordinweb/core";
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
  const [bookmark, setBookmark] = useState("");
  const rootRef = useRef<HTMLSpanElement | null>(null);
  const bookmarks = open ? api?.listBookmarks() ?? [] : [];
  const selected = bookmarks.includes(bookmark) ? bookmark : bookmarks[0] ?? "";
  useEffect(() => {
    if (!open) return;
    const close = (event: MouseEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", close);
    return () => document.removeEventListener("mousedown", close);
  }, [open]);
  const insert = (kind: "text" | "page") => {
    if (selected && api?.insertCrossReference(selected, kind)) setOpen(false);
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert cross-reference" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={() => setOpen(!open)}>
        Cross-reference
      </button>
      {open && (
        <div style={{ position: "absolute", top: 28, left: 0, zIndex: 100, width: 260, padding: 10, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          {bookmarks.length === 0 ? (
            <div style={{ color: T.muted, fontSize: 12 }}>Add a bookmark first, then reference its text or page.</div>
          ) : (
            <>
              <ToolbarMenuSelect
                value={selected}
                ariaLabel="Bookmark to reference"
                width="100%"
                menuWidth={240}
                options={bookmarks.map((name) => ({ value: name, label: name }))}
                onChange={setBookmark}
                style={{ borderColor: T.border }}
              />
              <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, marginTop: 8 }}>
                <button style={{ ...pillBtn, background: T.popoverBg, color: T.fg }} onClick={() => insert("page")}>Page number</button>
                <button style={pillBtn} onClick={() => insert("text")}>Bookmark text</button>
              </div>
            </>
          )}
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

const SHAPES = [
  ["line", "Line", "―"],
  ["verticalLine", "Vertical line", "│"],
  ["rectangle", "Rectangle", "▭"],
  ["roundedRectangle", "Rounded rectangle", "▢"],
  ["ellipse", "Ellipse", "◯"],
  ["diamond", "Diamond", "◇"],
  ["textBox", "Text box", "T"],
] as const;

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
            {SHAPES.map(([preset, label, glyph]) => (
              <button key={preset} title={`Insert ${label}`} onClick={() => insert(preset)} style={{ minHeight: 48, border: `1px solid ${T.border}`, borderRadius: 6, background: T.popoverBg, color: T.fg, cursor: "pointer", font: "12px system-ui, sans-serif" }}>
                <span style={{ display: "block", fontSize: 20, lineHeight: 1 }}>{glyph}</span>
                <span>{label}</span>
              </button>
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
    const rawSeries = series.map((entry) => ({
      name: entry.name.trim(),
      values: entry.values.map((value) => value.trim()),
    }));
    if (
      categoryValues.some((value) => !value) ||
      rawSeries.some((entry) => !entry.name || entry.values.some((value) => value === "" || !Number.isFinite(Number(value))))
    ) {
      setError("Give every category and series a name, and enter a number in every data cell.");
      return;
    }
    const seriesValues = rawSeries.map((entry) => ({ name: entry.name, values: entry.values.map(Number) }));
    const data: Chart = { type, title, categories: categoryValues, series: seriesValues };
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
  const anchor = open ? rootRef.current?.getBoundingClientRect() : null;
  const viewportWidth = typeof window === "undefined" ? 456 : window.innerWidth;
  const popoverWidth = Math.min(440, viewportWidth - 16);
  const popoverLeft = Math.max(8, Math.min(anchor?.left ?? 8, viewportWidth - popoverWidth - 8));
  const fieldStyle: React.CSSProperties = { width: "100%", boxSizing: "border-box", border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 8px", font: "13px system-ui, sans-serif", color: T.fg, background: T.popoverBg };
  const toggle = () => {
    if (open) {
      setOpen(false);
      return;
    }
    const selected = api?.getSelectedChart();
    setType(selected?.type ?? "column");
    setTitle(selected?.title ?? "");
    setCategories(selected ? [...selected.categories] : ["", ""]);
    setSeries(selected
      ? selected.series.map((entry) => ({ name: entry.name, values: entry.values.map(String) }))
      : [{ name: "", values: ["", ""] }]);
    setError("");
    setOpen(true);
  };
  return (
    <span ref={rootRef} style={{ position: "relative", display: "inline-block" }}>
      <button title="Insert or edit chart" style={btnStyle(open)} onMouseDown={(event) => event.preventDefault()} onClick={toggle}>{label}</button>
      {open && (
        <div style={{ position: "fixed", top: anchor?.bottom ?? 28, left: popoverLeft, zIndex: 100, width: popoverWidth, maxHeight: "calc(100vh - 48px)", overflow: "auto", boxSizing: "border-box", padding: 10, display: "grid", gap: 8, background: T.popoverBg, border: `1px solid ${T.border}`, borderRadius: 8, boxShadow: T.popoverShadow }}>
          <ToolbarMenuSelect
            value={type}
            ariaLabel="Chart type"
            width="100%"
            menuWidth={300}
            options={[
              { value: "column", label: "Column" },
              { value: "bar", label: "Bar" },
              { value: "line", label: "Line" },
              { value: "pie", label: "Pie" },
            ]}
            onChange={(value) => setType(value as Chart["type"])}
            style={fieldStyle}
          />
          <input aria-label="Chart title" value={title} onChange={(event) => setTitle(event.target.value)} placeholder="Chart title" style={fieldStyle} />
          <div role="group" aria-label="Chart data" style={{ display: "grid", gap: 7 }}>
            <strong style={{ color: T.fg, font: "600 11.5px system-ui, sans-serif" }}>Data</strong>
            <div style={{ overflowX: "auto" }}>
              <div style={{ display: "grid", gridTemplateColumns: `minmax(110px,1.2fr) repeat(${categories.length},minmax(72px,1fr)) 62px`, gap: 5, minWidth: categories.length > 3 ? 430 : undefined }}>
                <span style={{ alignSelf: "center", color: T.muted, font: "11px system-ui, sans-serif" }}>Series</span>
                {categories.map((category, index) => (
                  <label key={index} style={{ display: "grid", gridTemplateColumns: "1fr auto", gap: 3 }}>
                    <input
                      aria-label={`Chart category ${index + 1}`}
                      value={category}
                      placeholder="Category"
                      onChange={(event) => setCategories(categories.map((value, itemIndex) => itemIndex === index ? event.target.value : value))}
                      style={fieldStyle}
                    />
                    {categories.length > 1 && <button type="button" aria-label={`Remove chart category ${index + 1}`} onClick={() => removeCategory(index)} style={{ ...pillBtn, padding: "0 6px" }}>×</button>}
                  </label>
                ))}
                <span />
                {series.map((entry, seriesIndex) => (
                  <Fragment key={seriesIndex}>
                    <input
                      aria-label={`Chart series ${seriesIndex + 1} name`}
                      value={entry.name}
                      placeholder="Series"
                      onChange={(event) => setSeries(series.map((value, itemIndex) => itemIndex === seriesIndex ? { ...value, name: event.target.value } : value))}
                      style={fieldStyle}
                    />
                    {entry.values.map((value, valueIndex) => (
                      <input
                        key={valueIndex}
                        aria-label={`Chart series ${seriesIndex + 1} value ${valueIndex + 1}`}
                        type="number"
                        step="any"
                        value={value}
                        placeholder="0"
                        onChange={(event) => setSeries(series.map((seriesValue, itemIndex) => itemIndex === seriesIndex ? { ...seriesValue, values: seriesValue.values.map((itemValue, itemValueIndex) => itemValueIndex === valueIndex ? event.target.value : itemValue) } : seriesValue))}
                        style={fieldStyle}
                      />
                    ))}
                    <button type="button" aria-label={`Remove chart series ${seriesIndex + 1}`} disabled={series.length === 1} onClick={() => setSeries(series.filter((_, itemIndex) => itemIndex !== seriesIndex))} style={{ ...pillBtn, padding: "0 7px" }}>Remove</button>
                  </Fragment>
                ))}
              </div>
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              <button type="button" onClick={() => setSeries([...series, { name: "", values: categories.map(() => "") }])} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Add series</button>
              <button type="button" onClick={addCategory} style={{ ...pillBtn, background: T.popoverBg, color: T.fg }}>Add category</button>
            </div>
          </div>
          {error && <div role="alert" style={{ color: "#c5221f", font: "11.5px system-ui, sans-serif" }}>{error}</div>}
          <button onClick={submit} style={{ border: 0, borderRadius: 6, padding: "7px 10px", background: T.accent, color: T.accentFg, cursor: "pointer", font: "600 12px system-ui, sans-serif" }}>Insert or update chart</button>
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
      <Btn label="Delete table" title="Delete the current table" onClick={() => run("deleteTable")} />
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
          <ActionMenu
            label="Rotate"
            title="Rotate selected object"
            width={78}
            groups={[{ items: [["rotateRight", "Rotate right 90°"], ["rotateLeft", "Rotate left 90°"]] }]}
            onPick={(value) => api?.arrangeObject(value as Parameters<NonNullable<typeof api>["arrangeObject"]>[0])}
          />
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
  const wrap = (
    <ActionMenu
      label="Wrap"
      title="Wrap"
      width={72}
      groups={[{ items: [
        ["wrapInline", "Inline with text"],
        ["wrapSquare", "Square"],
        ["wrapTopAndBottom", "Top and bottom"],
        ["wrapFront", "In front of text"],
        ["wrapBehind", "Behind text"],
      ] }]}
      onPick={(value) => run(value as Parameters<DocxViewApi["runSelectedObjectCommand"]>[0])}
    />
  );
  return (
    <span data-dxw-object-format="" style={{ display: "contents" }}>
      {context.kind === "chart" && <ChartMenu api={api} label="Edit data" />}
      {context.kind === "smartArt" && <SmartArtMenu api={api} label="Edit SmartArt" />}
      {context.kind === "smartArt" && <SmartArtTextControls api={api} nodeIndex={context.smartArtNodeIndex} />}
      {context.canEditText && (
        <Btn
          label="Edit text"
          title={context.kind === "smartArt" ? "Edit selected SmartArt node text" : "Edit shape text"}
          onClick={() => run("editText")}
        />
      )}
      {(context.kind === "shape" || context.kind === "smartArt") && (
        <Btn
          label={context.kind === "smartArt" ? (context.smartArtNodeSelected ? "Node fill" : "Fill all") : "Fill"}
          title={context.kind === "smartArt" ? (context.smartArtNodeSelected ? "Selected SmartArt node fill" : "All SmartArt node fills") : "Fill color"}
          onClick={() => run("fill")}
        />
      )}
      {context.kind === "shape" && <Btn label="Outline" title="Outline color, weight, and style" onClick={() => run("outline")} />}
      {context.kind === "line" && <Btn label="Line style" title="Line color, weight, and style" onClick={() => run("lineStyle")} />}
      {(context.kind === "image" || context.kind === "model3d") && <Btn label="Alt text" title="Alternative text" onClick={() => run("altText")} />}
      {wrap}
      <Btn label="Size" title="Exact size" onClick={() => run("size")} />
      <Btn label="Position" title="Exact page position" onClick={() => run("position")} />
      {(context.kind === "shape" || context.kind === "line" || context.kind === "image" || context.kind === "model3d") && (
        <Btn label="Rotate" title="Set rotation" onClick={() => run("rotate")} />
      )}
      {showArrange && (
        <>
          <Btn label="Bring forward" title="Bring selected object forward" onClick={() => run("bringForward")} />
          <Btn label="Send backward" title="Send selected object backward" onClick={() => run("sendBackward")} />
        </>
      )}
      {context.kind === "model3d" && <Btn label="Reset 3D" title="Reset 3D rotation" onClick={() => run("reset3d")} />}
      <Btn label="Delete" title="Delete selected object" onClick={() => run("delete")} />
    </span>
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
  | "coverPage"
  | "pageNumber"
  | "break"
  | "layout"
  | "help"
  | "download";

export type ToolbarMode = "simple" | "advanced";

export interface DocxToolbarProps {
  api: DocxViewApi | null;
  onSave?: (bytes: Uint8Array) => void;
  /** Simple shows basic Home editing; advanced adds the Insert, Draw, and Layout ribbons. */
  mode?: ToolbarMode;
  /** Per-group overrides; every group defaults to enabled. */
  features?: Partial<Record<ToolbarFeature, boolean>>;
  /** Extra class on the toolbar root (e.g. a scope for CSS-variable overrides). */
  className?: string;
  /** Inline overrides merged onto the toolbar root; wins over the defaults. */
  style?: React.CSSProperties;
}

export function DocxToolbar({
  api,
  onSave,
  mode = "advanced",
  features,
  className,
  style,
}: DocxToolbarProps) {
  const on = (k: ToolbarFeature) => features?.[k] !== false;
  // Ribbon-style tabs: complex tool groups get their own surface instead of
  // one overloaded row (Layout especially).
  type NormalTab = "home" | "insert" | "draw" | "layout";
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
  // Toolbar popovers can move focus away from the document selection; remember
  // the last real range and restore it before applying their choice.
  const savedRange = useRef<Range | null>(null);
  const imageInput = useRef<HTMLInputElement | null>(null);
  // Responsive collapse: measure the toolbar width and pick a tier; the higher
  // the tier the more low-frequency Home groups fold into the ⋮ overflow menu,
  // so the strip stays single-row-clean on phones and tablets (Google Docs
  // does exactly this). Full width keeps everything inline, so desktop and the
  // e2e specs (1400px) are unchanged.
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [tier, setTier] = useState(0);
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
          if (["home", "insert", "draw", "layout"].includes(current)) priorNormalTab.current = current as NormalTab;
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
        if (["home", "insert", "draw", "layout"].includes(current)) priorNormalTab.current = current as NormalTab;
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
            <Sep />
          </>
        ),
      });

    // Per-tier overflow: which group keys fold into ⋮. Tier 0 keeps all inline.
    const overflowKeys =
      tier === 0
        ? new Set<string>()
        : tier === 1
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
            {(["home", "insert", "draw", "layout"] as const)
              .filter((t) => (t !== "draw" || on("drawing")) && (t !== "layout" || on("layout")))
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
          {on("image") && <Btn label={<ImageIcon />} title="Insert image" onClick={() => imageInput.current?.click()} />}
          {on("icon") && <Btn label="Icons" title="Insert SVG icon" onClick={() => iconInput.current?.click()} />}
          {on("screenshot") && <ScreenshotButton api={api} />}
          {tier === 0 ? (
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
          {on("comment") && <CommentMenu api={api} />}
          {on("footnote") && <FootnoteMenu api={api} />}
          {on("bookmark") && <BookmarkMenu api={api} />}
          {on("crossReference") && <CrossReferenceMenu api={api} />}
          {on("headerFooter") && (
            <ActionMenu
              label="Header & footer"
              title="Edit the repeating header or footer"
              width={118}
              groups={[{ items: [["header", "Header"], ["footer", "Footer"]] }]}
              onPick={(value) => api?.openHeaderFooter(value as "header" | "footer")}
            />
          )}
          <Sep />
          {on("pageNumber") && (
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
          )}
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
              {on("comment") && <CommentMenu api={api} />}
              {on("footnote") && <FootnoteMenu api={api} />}
              {on("bookmark") && <BookmarkMenu api={api} />}
              {on("crossReference") && <CrossReferenceMenu api={api} />}
              {on("headerFooter") && (
                <ActionMenu
                  label="Header & footer"
                  title="Edit the repeating header or footer"
                  width={118}
                  groups={[{ items: [["header", "Header"], ["footer", "Footer"]] }]}
                  onPick={(value) => api?.openHeaderFooter(value as "header" | "footer")}
                />
              )}
              {on("pageNumber") && (
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
              )}
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
        accept="image/png,image/jpeg,image/gif,image/svg+xml,image/webp"
        style={{ display: "none" }}
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void api?.insertImage(f);
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
          if (f) void api?.insertImage(f);
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
      <HelpGuide open={helpOpen} onClose={closeHelp} returnFocus={helpTrigger} />
    </div>
  );
}
