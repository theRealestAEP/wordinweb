import { createHash } from "node:crypto";
import { existsSync, readFileSync, readdirSync, statSync } from "node:fs";
import { basename, relative, resolve } from "node:path";
import { describe, expect, it } from "vitest";
import type { Block, DocxDocument, Paragraph, Run, RunContent } from "@wordinweb/core";
import { AgentDocument, type AgentBlock, type AgentObjectResult } from "../src/index.js";

const enabled = process.env.WORDINWEB_FIXTURE_AUDIT === "1";

function fixtureRoot(): string | null {
  const candidates = [
    process.env.WORDINWEB_FIXTURES,
    resolve(process.cwd(), "../../wordinweb-parity"),
    resolve(process.cwd(), "../../../wordinweb-parity"),
    resolve(process.cwd(), "../wordinweb-parity"),
  ].filter((value): value is string => Boolean(value));
  return candidates.find(existsSync) ?? null;
}

function docxFiles(root: string): string[] {
  const files: string[] = [];
  const visit = (directory: string): void => {
    for (const name of readdirSync(directory)) {
      if (name === ".git" || name === "node_modules" || name === "dist") continue;
      const path = resolve(directory, name);
      const stat = statSync(path);
      if (stat.isDirectory()) visit(path);
      else if (stat.isFile() && name.toLowerCase().endsWith(".docx")) files.push(path);
    }
  };
  visit(root);
  return files.sort();
}

function runs(paragraph: Paragraph): Run[] {
  return paragraph.children.flatMap((child) => child.type === "run" ? [child] : child.runs);
}

function countModel(doc: DocxDocument): Record<string, number> {
  const counts: Record<string, number> = {};
  const count = (kind: string): void => { counts[kind] = (counts[kind] ?? 0) + 1; };
  const visitContent = (content: RunContent): void => {
    count(content.kind);
    if (content.kind === "anchor") {
      const shape = content.shape;
      if (shape.type === "textbox") visitBlocks(shape.blocks);
      else if (shape.type === "art") for (const text of shape.texts ?? []) visitBlocks(text.blocks);
    } else if (content.kind === "drawing") {
      if (content.textbox) visitBlocks(content.textbox.blocks);
      for (const text of content.texts ?? []) visitBlocks(text.blocks);
    }
  };
  const visitBlocks = (blocks: Block[]): void => {
    for (const block of blocks) {
      if (block.type === "table") {
        for (const row of block.rows) for (const cell of row.cells) visitBlocks(cell.blocks);
      } else {
        for (const run of runs(block)) for (const content of run.content) visitContent(content);
      }
    }
  };
  for (const section of doc.sections) visitBlocks(section.blocks);
  for (const story of doc.headers.values()) visitBlocks(story.blocks);
  for (const story of doc.footers.values()) visitBlocks(story.blocks);
  for (const blocks of doc.footnotes.values()) visitBlocks(blocks);
  for (const blocks of doc.endnotes.values()) visitBlocks(blocks);
  return counts;
}

function assertSerializable(value: unknown, label: string): void {
  const json = JSON.stringify(value);
  if (json === undefined) throw new Error(`${label}: result is not JSON serializable`);
  if (/"src(?:Drawing|T|Parent)?"\s*:/.test(json)) throw new Error(`${label}: leaked an XML source reference`);
  if (/"attrs"\s*:\s*\{/.test(json)) throw new Error(`${label}: leaked an XML element`);
}

function refsFromBlocks(blocks: AgentBlock[]): string[] {
  const refs: string[] = [];
  for (const block of blocks) {
    if (block.type !== "paragraph") continue;
    for (const run of block.runs) for (const component of run.components) refs.push(component.ref);
  }
  return refs;
}

function collectAssetRefs(value: unknown, out = new Set<string>()): Set<string> {
  if (typeof value === "string" && /^asset:\d+$/.test(value)) out.add(value);
  else if (Array.isArray(value)) for (const item of value) collectAssetRefs(item, out);
  else if (value && typeof value === "object") for (const item of Object.values(value)) collectAssetRefs(item, out);
  return out;
}

function expectedLoadFailure(path: string): boolean {
  return path.endsWith("fixtures-staging/robustness/damaged-truncated.docx");
}

async function audit(bytes: Uint8Array, label: string): Promise<void> {
  let agent: AgentDocument;
  try {
    agent = AgentDocument.load(bytes);
  } catch (error) {
    if (expectedLoadFailure(label)) return;
    throw error;
  }
  if (expectedLoadFailure(label)) throw new Error("expected the truncated fixture to fail loading");

  const overview = agent.inspect({ kind: "overview" });
  if (!("components" in overview)) throw new Error("overview returned the wrong result type");
  assertSerializable(overview, `${label} overview`);
  expect(overview.components, `${label}: model components must all reach the agent overview`).toEqual(countModel(agent.document));

  const componentRefs = new Set<string>();
  for (const story of overview.stories) {
    let cursor: { value: string } | undefined;
    do {
      const page = agent.inspect({ kind: "read", story: story.id, cursor, maxBlocks: 100, maxCharacters: 100_000 });
      if (!("blocks" in page)) throw new Error(`${label}: read returned the wrong result type`);
      assertSerializable(page, `${label} story ${story.id}`);
      for (const ref of refsFromBlocks(page.blocks)) componentRefs.add(ref);
      cursor = page.next;
    } while (cursor);
  }
  const expectedObjects = Object.entries(overview.components).reduce((sum, [kind, count]) => sum + (kind === "text" ? 0 : count), 0);
  expect(componentRefs.size, `${label}: every non-text component needs an inspectable reference`).toBe(expectedObjects);

  // Every story projects deterministically in every mode, and an identity
  // patch of a projected window is a fixed point.
  const projectable = new Set(["body", "header", "footer", "footnote", "endnote"]);
  for (const story of overview.stories) {
    if (!projectable.has(story.kind)) continue;
    for (const mode of ["text", "md", "outline"] as const) {
      let cursor: { value: string } | undefined;
      let windows = 0;
      do {
        const window = { story: story.id, mode, cursor, maxBlocks: 200, maxCharacters: 100_000 };
        const projection = agent.project(window);
        assertSerializable(projection, `${label} projection ${story.id} ${mode}`);
        expect(projection.text, `${label}: ${story.id} ${mode} must project byte-identically`).toBe(agent.project(window).text);
        expect(projection.revision, `${label}: every projection window carries its revision`).toBe(agent.revision);
        expect(projection.anchors.length, `${label}: ${story.id} ${mode} anchors must cover every line`).toBe(projection.lines);
        if (windows === 0 && mode !== "outline") {
          const edits = projection.text.split("\n")
            .map((newText, index) => ({ startLine: index + 1, endLine: index + 1, newText }))
            .filter((_, index) => projection.anchors[index].role === "paragraph" && projection.anchors[index].editable)
            .slice(0, 100);
          if (edits.length > 0) {
            const patched = await agent.patch({ revision: projection.revision, story: story.id, mode, edits });
            expect(patched.operations, `${label}: ${story.id} ${mode} identity patch must be a no-op`).toEqual([]);
            expect(patched.projection.text, `${label}: ${story.id} ${mode} must be a projection fixed point`).toBe(projection.text);
          }
        }
        cursor = projection.next;
        windows++;
      } while (cursor);
    }
  }

  const assets = new Set<string>();
  for (const ref of componentRefs) {
    const object = agent.inspect({ kind: "object", ref });
    if (!("detail" in object)) throw new Error(`${label}: object returned the wrong result type`);
    assertSerializable(object, `${label} object ${ref}`);
    collectAssetRefs((object as AgentObjectResult).detail, assets);
  }
  for (const ref of assets) {
    const asset = agent.asset(ref);
    expect(asset.bytes.length, `${label}: ${ref} must resolve to bytes`).toBeGreaterThan(0);
  }

  // The robustness corpus includes intentionally hostile documents such as a
  // single damaged 2 MB paragraph. Its contract is bounded parse/inspection,
  // while the parity and staging corpora own visual-layout coverage.
  if (label.includes("/robustness/")) return;

  let start = 1;
  let total = 1;
  do {
    const spatial = agent.inspect({ kind: "spatial", pages: { start, count: 100 }, includeOverlaps: true });
    if (!("totalPages" in spatial)) throw new Error(`${label}: spatial returned the wrong result type`);
    assertSerializable(spatial, `${label} spatial pages ${start}+`);
    total = spatial.totalPages;
    start += 100;
  } while (start <= total);
}

describe.runIf(enabled)("agent representation parity corpus", () => {
  it("represents every semantic component and every laid-out page in every DOCX fixture", async () => {
    const root = fixtureRoot();
    expect(root, "Set WORDINWEB_FIXTURES to the wordinweb-parity checkout").toBeTruthy();
    const files = docxFiles(root!);
    expect(files.length, "The fixture corpus must not be empty").toBeGreaterThan(100);

    const byHash = new Map<string, { bytes: Uint8Array; paths: string[] }>();
    for (const path of files) {
      const bytes = new Uint8Array(readFileSync(path));
      const hash = createHash("sha256").update(bytes).digest("hex");
      const group = byHash.get(hash);
      if (group) group.paths.push(relative(root!, path));
      else byHash.set(hash, { bytes, paths: [relative(root!, path)] });
    }

    const failures: string[] = [];
    for (const { bytes, paths } of byHash.values()) {
      const label = paths.find(expectedLoadFailure) ?? paths[0];
      try {
        await audit(bytes, label);
      } catch (error) {
        failures.push(`${label}${paths.length > 1 ? ` (+${paths.length - 1} identical copies)` : ""}: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
    expect(failures, failures.join("\n")).toEqual([]);
  }, 20 * 60 * 1000);
});
