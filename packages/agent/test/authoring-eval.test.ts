import { readFileSync, readdirSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { AgentDocument, type AgentBlock, type AgentObjectResult } from "../src/index.js";

type StoryKind = "body" | "header" | "footer" | "footnote" | "endnote" | "textbox";

interface ObjectRequirement {
  type: string | string[];
  storyKind?: StoryKind;
  minimum?: number;
  detail?: Record<string, unknown>;
  nonEmptyDetail?: string[];
}

interface EvalFixture {
  name: string;
  prompt: string;
  expected: {
    requiredText: string[];
    minimumParagraphs?: number;
    minimumOutlineEntries?: number;
    minimumTables?: number;
    table?: { minimumRows: number; minimumColumns: number };
    requiredObjectTypes?: string[];
    objectRequirements?: ObjectRequirement[];
    minimumSpatialObjects?: number;
    minimumSpatialLayers?: Partial<Record<"behind-text" | "body" | "in-front-of-text", number>>;
    minimumOverlaps?: number;
    requiredStories?: StoryKind[];
    minimumSections?: number;
    minimumPages?: number;
    pageWidth?: number;
    pageHeight?: number;
    margin?: number;
    minimumAssets?: number;
    requiredAssetMediaTypes?: string[];
    minimumComments?: number;
    requiredBookmarks?: string[];
    sectionLayout?: {
      minimumColumns?: number;
      columnSeparator?: boolean;
      pageBorders?: boolean;
      lineNumbering?: boolean;
    };
    requiredPackageParts?: string[];
    requiredPackageText?: string[];
  };
}

const evalDirectory = resolve(import.meta.dirname, "../evals");
const fixtures = readdirSync(evalDirectory)
  .filter((name) => name.endsWith(".fixture.json"))
  .sort()
  .map((name) => JSON.parse(readFileSync(resolve(evalDirectory, name), "utf8")) as EvalFixture);
const output = process.env.WORDINWEB_AGENT_EVAL_OUTPUT;
const selectedFixture = process.env.WORDINWEB_AGENT_EVAL_FIXTURE;

function componentRefs(blocks: AgentBlock[]): string[] {
  return blocks.flatMap((block) => block.type === "paragraph"
    ? block.runs.flatMap((run) => run.components.map((component) => component.ref))
    : []);
}

function collectAssetRefs(value: unknown, output = new Set<string>()): Set<string> {
  if (typeof value === "string" && value.startsWith("asset:")) output.add(value);
  else if (Array.isArray(value)) for (const item of value) collectAssetRefs(item, output);
  else if (value && typeof value === "object") for (const item of Object.values(value)) collectAssetRefs(item, output);
  return output;
}

describe("context-neutral agent authoring evaluation", () => {
  it.each(fixtures)("has a complete, natural fixture: $name", (fixture) => {
    expect(fixture.prompt.length).toBeGreaterThan(100);
    expect(fixture.prompt.toLocaleLowerCase()).not.toMatch(/\b(test|evaluation)\b/);
    expect(fixture.expected.requiredText.length).toBeGreaterThan(0);
  });

  it.runIf(Boolean(output))("scores a generated DOCX through the public agent representation", () => {
    if (!selectedFixture) throw new Error("WORDINWEB_AGENT_EVAL_FIXTURE is required");
    const fixture = fixtures.find((candidate) => candidate.name === selectedFixture);
    if (!fixture) throw new Error(`Unknown evaluation fixture: ${selectedFixture}`);
    const document = AgentDocument.load(readFileSync(output!));
    const overview = document.inspect({ kind: "overview" });
    if (!("stories" in overview)) throw new Error("overview result required");

    const blocks: AgentBlock[] = [];
    const refs = new Map<string, StoryKind>();
    const bookmarks = new Set<string>();
    let allText = "";
    for (const story of overview.stories) {
      let cursor: { value: string } | undefined;
      do {
        const page = document.inspect({ kind: "read", story: story.id, cursor, maxBlocks: 100, maxCharacters: 100_000 });
        if (!("blocks" in page)) throw new Error("read result required");
        blocks.push(...page.blocks);
        allText += `\n${page.blocks.flatMap((block) => block.type === "paragraph" ? [block.text] : []).join("\n")}`;
        for (const block of page.blocks) {
          if (block.type === "paragraph") for (const bookmark of block.bookmarks) bookmarks.add(bookmark);
        }
        for (const ref of componentRefs(page.blocks)) refs.set(ref, story.kind);
        cursor = page.next;
      } while (cursor);
    }

    const objects = [...refs].map(([ref, storyKind]) => {
      const object = document.inspect({ kind: "object", ref });
      if (!("type" in object)) throw new Error("object result required");
      return { object: object as AgentObjectResult, storyKind };
    });
    for (const { object } of objects) {
      if (typeof object.detail.text === "string") allText += `\n${object.detail.text}`;
    }
    const objectTypes = new Set(objects.map(({ object }) => object.type));
    const assets = new Map<string, string>();
    for (const { object } of objects) {
      for (const ref of collectAssetRefs(object.detail)) assets.set(ref, document.asset(ref).mediaType);
    }
    const tables = blocks.filter((block) => block.type === "table");
    const spatial = document.inspect({ kind: "spatial", pages: { start: 1, count: 100 }, includeOverlaps: true });
    if (!("objects" in spatial)) throw new Error("spatial result required");

    for (const text of fixture.expected.requiredText) expect(allText.toLocaleLowerCase()).toContain(text.toLocaleLowerCase());
    if (fixture.expected.minimumParagraphs !== undefined) expect(overview.blocks.paragraphs).toBeGreaterThanOrEqual(fixture.expected.minimumParagraphs);
    if (fixture.expected.minimumOutlineEntries !== undefined) expect(overview.outline.length).toBeGreaterThanOrEqual(fixture.expected.minimumOutlineEntries);
    if (fixture.expected.minimumTables !== undefined) expect(tables.length).toBeGreaterThanOrEqual(fixture.expected.minimumTables);
    if (fixture.expected.table) {
      expect(tables.some((table) => table.type === "table" && table.rows >= fixture.expected.table!.minimumRows && table.columns >= fixture.expected.table!.minimumColumns)).toBe(true);
    }
    for (const type of fixture.expected.requiredObjectTypes ?? []) expect(objectTypes, type).toContain(type);
    for (const requirement of fixture.expected.objectRequirements ?? []) {
      const types = Array.isArray(requirement.type) ? requirement.type : [requirement.type];
      const matches = objects.filter(({ object, storyKind }) =>
        types.includes(object.type) &&
        (requirement.storyKind === undefined || storyKind === requirement.storyKind) &&
        (requirement.nonEmptyDetail === undefined || requirement.nonEmptyDetail.every((field) => typeof object.detail[field] === "string" && object.detail[field].length > 0)) &&
        (requirement.detail === undefined || (() => {
          try {
            expect(object.detail).toMatchObject(requirement.detail!);
            return true;
          } catch {
            return false;
          }
        })()));
      expect(matches.length, `${requirement.storyKind ?? "any"} ${types.join(" or ")}`).toBeGreaterThanOrEqual(requirement.minimum ?? 1);
    }
    for (const kind of fixture.expected.requiredStories ?? []) expect(overview.stories.some((story) => story.kind === kind), kind).toBe(true);
    if (fixture.expected.minimumSections !== undefined) expect(overview.sections).toBeGreaterThanOrEqual(fixture.expected.minimumSections);
    if (fixture.expected.minimumPages !== undefined) expect(spatial.totalPages).toBeGreaterThanOrEqual(fixture.expected.minimumPages);
    if (fixture.expected.minimumSpatialObjects !== undefined) expect(spatial.objects.length).toBeGreaterThanOrEqual(fixture.expected.minimumSpatialObjects);
    if (fixture.expected.minimumOverlaps !== undefined) expect(spatial.overlaps.length).toBeGreaterThanOrEqual(fixture.expected.minimumOverlaps);
    for (const [layer, minimum] of Object.entries(fixture.expected.minimumSpatialLayers ?? {})) {
      expect(spatial.objects.filter((object) => object.layer === layer).length, layer).toBeGreaterThanOrEqual(minimum!);
    }
    if (fixture.expected.pageWidth !== undefined && fixture.expected.pageHeight !== undefined) {
      expect(spatial.pages[0]).toMatchObject({ width: fixture.expected.pageWidth, height: fixture.expected.pageHeight });
    }
    if (fixture.expected.margin !== undefined) {
      expect(overview.sectionLayouts[0].margins).toMatchObject({
        top: fixture.expected.margin,
        right: fixture.expected.margin,
        bottom: fixture.expected.margin,
        left: fixture.expected.margin,
      });
    }
    if (fixture.expected.minimumAssets !== undefined) expect(assets.size).toBeGreaterThanOrEqual(fixture.expected.minimumAssets);
    for (const mediaType of fixture.expected.requiredAssetMediaTypes ?? []) expect([...assets.values()], mediaType).toContain(mediaType);
    if (fixture.expected.minimumComments !== undefined) expect(overview.comments).toBeGreaterThanOrEqual(fixture.expected.minimumComments);
    for (const bookmark of fixture.expected.requiredBookmarks ?? []) expect(bookmarks, bookmark).toContain(bookmark);
    if (fixture.expected.sectionLayout) {
      expect(overview.sectionLayouts.some((layout) =>
        (fixture.expected.sectionLayout!.minimumColumns === undefined || layout.columns.count >= fixture.expected.sectionLayout!.minimumColumns) &&
        (fixture.expected.sectionLayout!.columnSeparator === undefined || Boolean(layout.columns.separator) === fixture.expected.sectionLayout!.columnSeparator) &&
        (fixture.expected.sectionLayout!.pageBorders === undefined || Boolean(layout.pageBorders) === fixture.expected.sectionLayout!.pageBorders) &&
        (fixture.expected.sectionLayout!.lineNumbering === undefined || Boolean(layout.lineNumbering) === fixture.expected.sectionLayout!.lineNumbering))).toBe(true);
    }
    const saved = AgentDocument.load(document.save());
    for (const part of fixture.expected.requiredPackageParts ?? []) expect(saved.document.pkg.names(), part).toContain(part);
    const packageText = saved.document.pkg.names()
      .filter((name) => name.endsWith(".xml") || name.endsWith(".rels"))
      .map((name) => saved.document.pkg.text(name) ?? "")
      .join("\n");
    for (const fragment of fixture.expected.requiredPackageText ?? []) expect(packageText, fragment).toContain(fragment);
    expect(document.save().length).toBeGreaterThan(0);
  });
});
