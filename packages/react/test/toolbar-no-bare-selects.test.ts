// @vitest-environment node
/**
 * #141's regression guard, over the whole toolbar source rather than over any
 * one popover.
 *
 * The defect was not a broken control, it was 23 controls drifting away from
 * the one the toolbar already had, in four different directions. A test that
 * renders one popover cannot catch the 24th `<select>` being added to a
 * different one, and rendering all 31 popovers to look for native elements
 * would be a harness larger than the fix. So this reads the file.
 *
 * `@vitest-environment node`: under jsdom `import.meta.url` is an http URL and
 * cannot be turned back into a path. There is nothing to render here anyway.
 */
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const SOURCE = readFileSync(fileURLToPath(new URL("../src/toolbar.tsx", import.meta.url)), "utf8");

/** Every `<select ...>` opening tag in the file, with its line number. Brace
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

describe("#141: the toolbar draws its own dropdowns", () => {
  it("has no bare visible <select> left", () => {
    const bare = openingTags(SOURCE, "select")
      // ToolbarMenuSelect's own bridge is the one legitimate native select:
      // inert, 1px, aria-hidden, kept so that integrations and the four older
      // suites that drive the DOM element by [aria-label] still work.
      .filter(({ tag }) => !tag.includes('aria-hidden="true"'))
      .map(({ line }) => line);
    expect(bare, `use ToolbarMenuSelect instead — bare <select> at toolbar.tsx line(s) ${bare.join(", ")}`)
      .toEqual([]);
  });

  it("still keeps exactly one bridge, so the guard above is not vacuous", () => {
    const bridges = openingTags(SOURCE, "select").filter(({ tag }) => tag.includes('aria-hidden="true"'));
    expect(bridges).toHaveLength(1);
  });
});
