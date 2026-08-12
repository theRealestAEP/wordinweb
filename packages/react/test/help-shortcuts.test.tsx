// @vitest-environment jsdom
/**
 * The help sheet is the shortcut table, rendered.
 *
 * These tests exist so the reference cannot drift from reality the way the
 * previous hand-kept copy did (it advertised ⌘F, which nothing bound). One
 * test pins the rendered rows to the table; the others pin the table to the
 * matcher, so a row can only appear once a key really produces its command.
 */
import { afterEach, describe, expect, it } from "vitest";
import { act, createElement } from "react";
import { createRoot } from "react-dom/client";
import {
  EDITOR_KEY_NOTES,
  EDITOR_SHORTCUTS,
  formatCombo,
  formatShortcutKeys,
  matchShortcut,
  type HostShortcutSection,
  type KeyCombo,
} from "@wordinweb/core";
import { HelpGuide } from "../src/help.js";
import { HELP_COMBOS } from "../src/toolbar.js";

afterEach(() => {
  document.body.replaceChildren();
});

async function openShortcutsTab(hostShortcuts?: HostShortcutSection[]) {
  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = createRoot(host);
  await act(async () => {
    root.render(createElement(HelpGuide, { open: true, onClose: () => {}, helpCombos: HELP_COMBOS, hostShortcuts }));
  });
  const tab = [...document.querySelectorAll<HTMLButtonElement>('[role="tab"]')]
    .find((button) => /shortcut/i.test(button.textContent ?? ""))!;
  await act(async () => { tab.click(); });
  const rows = [...document.querySelectorAll<HTMLElement>("[data-dxw-help-shortcut]")].map((row) => ({
    action: row.dataset.dxwHelpShortcut!,
    keys: row.querySelector("kbd")!.textContent!,
  }));
  return { rows, unmount: () => act(async () => { root.unmount(); }) };
}

/** A KeyboardEvent that presses exactly this combo (jsdom = non-Apple). */
function eventFor(combo: KeyCombo): KeyboardEvent {
  return new KeyboardEvent("keydown", {
    key: combo.key,
    ctrlKey: !!combo.mod,
    shiftKey: !!combo.shift,
    altKey: !!combo.alt,
  });
}

describe("the help sheet lists exactly the bound set", () => {
  it("has one row per table entry, and no rows of its own", async () => {
    const { rows, unmount } = await openShortcutsTab();
    const expected = [
      ...EDITOR_SHORTCUTS.map((s) => s.label),
      ...EDITOR_KEY_NOTES.map((n) => n.label),
      "Open this help guide",
    ];
    expect([...rows.map((r) => r.action)].sort()).toEqual([...expected].sort());
    await unmount();
  });

  it("prints each row's keys straight from the table", async () => {
    const { rows, unmount } = await openShortcutsTab();
    const printed = new Map(rows.map((r) => [r.action, r.keys]));
    for (const shortcut of EDITOR_SHORTCUTS) {
      expect(printed.get(shortcut.label), shortcut.label).toBe(formatShortcutKeys(shortcut, false));
    }
    for (const note of EDITOR_KEY_NOTES) {
      expect(printed.get(note.label), note.label).toBe(note.keys(false));
    }
    expect(printed.get("Open this help guide")).toBe("F1 or Ctrl+/");
    await unmount();
  });

  it("adds the host application's own accelerators when it declares them", async () => {
    // A desktop menu accelerator never reaches the editor, so the engine
    // cannot find it; the host injects it and the reference stays complete.
    const { rows, unmount } = await openShortcutsTab([
      { title: "LikeOffice menus", items: [
        { label: "Find…", keys: "⌘F" },
        { label: "Save", keys: "⌘S", detail: "Writes the open .docx." },
      ] },
    ]);
    const printed = new Map(rows.map((r) => [r.action, r.keys]));
    expect(printed.get("Find…")).toBe("⌘F");
    expect(printed.get("Save")).toBe("⌘S");
    // …and the editor's own rows are still all there.
    expect(printed.get("Strikethrough")).toBe(formatShortcutKeys(EDITOR_SHORTCUTS.find((s) => s.command === "strikethrough")!, false));
    await unmount();
  });

  it("writes the modifiers the way the platform does", () => {
    const combo: KeyCombo = { key: "x", mod: true, shift: true };
    expect(formatCombo(combo, true)).toBe("⇧⌘X");
    expect(formatCombo(combo, false)).toBe("Ctrl+Shift+X");
    expect(formatCombo({ key: "c", mod: true, alt: true }, true)).toBe("⌥⌘C");
    expect(formatCombo({ key: "ArrowLeft", mod: true }, true)).toBe("⌘←");
    expect(formatCombo({ key: "PageDown" }, false)).toBe("Page Down");
  });
});

describe("the table binds what it prints", () => {
  it("every combo produces its own command", () => {
    for (const shortcut of EDITOR_SHORTCUTS) {
      for (const combo of shortcut.combos) {
        if (combo.platform === "apple") continue; // jsdom is not Apple
        expect(matchShortcut(eventFor(combo), false), `${shortcut.label} / ${formatCombo(combo, false)}`)
          .toBe(shortcut.command);
      }
    }
  });

  it("no two shortcuts claim the same combo", () => {
    const seen = new Map<string, string>();
    for (const shortcut of EDITOR_SHORTCUTS) {
      for (const combo of shortcut.combos) {
        const id = `${combo.platform ?? "any"}|${combo.key}|${!!combo.mod}|${!!combo.shift}|${!!combo.alt}`;
        expect(seen.get(id), `${formatCombo(combo, false)} is claimed twice`).toBeUndefined();
        seen.set(id, shortcut.label);
      }
    }
  });

  it("never claims an accelerator the desktop app owns alone", () => {
    // The app's menus own these outright (LikeOffice menu.ts). A menu
    // accelerator is consumed before the editor's keydown handler runs, so a
    // binding here would be unreachable in the app and, worse, would mean two
    // declarations of one key. Keys both DO claim on purpose (⌘B, ⌘Z, ⌘A,
    // alignment, headings, lists, ⌘K, ⌘Enter, ⇧⌘X, ⇧⌘E) route to the same
    // behaviour on both sides and are deliberately absent from this list.
    const appOnly: KeyCombo[] = [
      { key: ",", mod: true },                  // Settings
      { key: "n", mod: true },                  // New Document
      { key: "o", mod: true },                  // Open
      { key: "s", mod: true },                  // Save
      { key: "s", mod: true, shift: true },     // Save As
      { key: "p", mod: true },                  // Print
      { key: "p", mod: true, shift: true },     // Page Setup
      { key: "w", mod: true },                  // Close Window
      { key: "v", mod: true, shift: true },     // Paste and Match Style
      { key: "a", mod: true, shift: true },     // AI Assistant
      { key: "f", mod: true },                  // Find
      { key: "h", mod: true, shift: true },     // Find and Replace
      { key: "g", mod: true },                  // Find Next
      { key: "g", mod: true, shift: true },     // Find Previous
      { key: "0", mod: true },                  // Actual Size
      { key: "-", mod: true },                  // Zoom Out
    ];
    for (const combo of appOnly) {
      expect(matchShortcut(eventFor(combo), false), formatCombo(combo, false)).toBeNull();
    }
  });
});
