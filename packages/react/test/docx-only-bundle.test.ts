// @vitest-environment node
import { build } from "esbuild";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { expect, it } from "vitest";

it("ships a DOCX-only browser entry without legacy Word converters", async () => {
  const root = fileURLToPath(new URL("../../../", import.meta.url));
  const result = await build({
    absWorkingDir: root,
    stdin: { contents: 'export * from "wordinweb/docx-only";', resolveDir: root, loader: "js" },
    bundle: true,
    platform: "browser",
    format: "esm",
    write: false,
    metafile: true,
  });
  const chunks = Object.keys(result.metafile!.inputs).filter((name) => name.startsWith("packages/react/dist/") && name.endsWith(".js"));
  expect(chunks.length).toBeGreaterThan(0);
  const sources: string[] = chunks.flatMap((name) => JSON.parse(readFileSync(`${root}${name}.map`, "utf8")).sources);
  expect(sources.filter((name) => /\/core\/src\/(doc|oldword)\//.test(name))).toEqual([]);
});
