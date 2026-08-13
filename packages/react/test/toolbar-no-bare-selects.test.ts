// @vitest-environment node
/**
 * #141's regression guard, over the whole toolbar source rather than over any
 * one popover.
 *
 * The defect was not a broken control, it was 30 controls drifting away from
 * the ones the toolbar already had, in four different directions. A test that
 * renders one popover cannot catch the 24th `<select>` being added to a
 * different one, and rendering all 31 popovers to look for native elements
 * would be a harness larger than the fix. So this reads the file.
 *
 * Both guards count only what a USER CAN SEE. Each replacement component
 * keeps a real native element inside itself — the select for its event
 * bridge, the checkbox for state, focus and the accessible name — so a naive
 * `<select>` or `type="checkbox"` count would either fail on day one or,
 * worse, pass by counting those. Each is identified by an explicit marker
 * attribute rather than by a guess about its styling.
 *
 * `@vitest-environment node`: under jsdom `import.meta.url` is an http URL and
 * cannot be turned back into a path. There is nothing to render here anyway.
 */
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const RAW = readFileSync(fileURLToPath(new URL("../src/toolbar.tsx", import.meta.url)), "utf8");

/**
 * Blank out comments, keeping every newline so reported line numbers stay
 * true. Without this the guards read prose: the components document
 * themselves by naming the `<input type="checkbox">` and `<select>` they
 * wrap, and a scanner that cannot tell code from a sentence about code
 * reports the documentation as the defect.
 */
const SOURCE = RAW
  .replace(/\/\*[\s\S]*?\*\//g, (block) => block.replace(/[^\n]/g, " "))
  .split("\n")
  .map((line) => (line.trimStart().startsWith("//") ? "" : line))
  .join("\n");

/** Every opening tag of `name` in the file, with its line number. Brace
 * counting, because a `>` inside a JSX expression does not end the tag. */
function openingTags(source: string, name: string): { line: number; tag: string }[] {
  const found: { line: number; tag: string }[] = [];
  for (const match of source.matchAll(new RegExp(`<${name}\\b`, "g"))) {
    const start = match.index!;
    let depth = 0;
    let i = start;
    while (i < source.length) {
      const c = source[i];
      if (c === ">" && depth === 0) break;
      if (c === "{") depth++;
      else if (c === "}") depth--;
      i++;
    }
    found.push({ line: source.slice(0, start).split("\n").length, tag: source.slice(start, i + 1) });
  }
  return found;
}

const selects = () => openingTags(SOURCE, "select");
const checkboxes = () => openingTags(SOURCE, "input").filter(({ tag }) => tag.includes('type="checkbox"'));

describe("#141: the toolbar draws its own form controls", () => {
  it("has no bare visible <select> left", () => {
    const bare = selects()
      .filter(({ tag }) => !tag.includes("data-dxw-native-bridge"))
      .map(({ line }) => line);
    expect(bare, `use ToolbarMenuSelect instead — bare <select> at toolbar.tsx line(s) ${bare.join(", ")}`)
      .toEqual([]);
  });

  it("has no bare visible checkbox left", () => {
    const bare = checkboxes()
      .filter(({ tag }) => !tag.includes("data-dxw-checkbox-input"))
      .map(({ line }) => line);
    expect(bare, `use ToolbarCheckbox instead — bare checkbox at toolbar.tsx line(s) ${bare.join(", ")}`)
      .toEqual([]);
  });

  it("gives every number box the attribute its spinner rule selects", () => {
    // The arrows are a pseudo-element and no inline style reaches them, so an
    // untagged number box silently keeps the browser's steppers — a miss that
    // looks like nothing at all until someone hovers it.
    const untagged = openingTags(SOURCE, "input")
      .filter(({ tag }) => tag.includes('type="number"') && !tag.includes("data-dxw-number"))
      .map(({ line }) => line);
    expect(untagged, `add data-dxw-number="" at toolbar.tsx line(s) ${untagged.join(", ")}`).toEqual([]);
  });

  it("keeps exactly one of each native element, so the guards are not vacuous", () => {
    // If a refactor deleted the markers, the two guards above would pass by
    // finding nothing at all. These are the counterweight.
    expect(selects().filter(({ tag }) => tag.includes("data-dxw-native-bridge")), "select bridge").toHaveLength(1);
    expect(checkboxes().filter(({ tag }) => tag.includes("data-dxw-checkbox-input")), "checkbox input").toHaveLength(1);
    expect(openingTags(SOURCE, "input").filter(({ tag }) => tag.includes('type="number"')).length, "number boxes")
      .toBeGreaterThan(0);
  });
});
