import { existsSync, readFileSync } from "node:fs";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { INTENT_KINDS } from "@wordinweb/collab/client";
import { AgentDocument, AGENT_EDIT_CAPABILITIES } from "../src/index.js";

const skillRoot = resolve(import.meta.dirname, "../../../skills/wordinweb-documents");
const skill = readFileSync(resolve(skillRoot, "SKILL.md"), "utf8");
const interfaceReference = readFileSync(resolve(skillRoot, "references/interface.md"), "utf8");
const createReference = readFileSync(resolve(skillRoot, "references/create.md"), "utf8");
const editReference = readFileSync(resolve(skillRoot, "references/inspect-edit.md"), "utf8");

function documentedOperations(): Map<string, Set<string>> {
  const catalog = interfaceReference
    .slice(interfaceReference.indexOf("## Complete operation catalog"), interfaceReference.indexOf("## Nested value shapes"));
  const operations = new Map<string, Set<string>>();
  for (const line of catalog.split("\n")) {
    const cells = line.match(/^\| `([A-Za-z0-9]+)` \|[^|]*\| (.*) \|$/);
    if (!cells) continue;
    operations.set(cells[1], new Set([...cells[2].matchAll(/`([A-Za-z0-9]+)`/g)].map((match) => match[1])));
  }
  return operations;
}

describe("wordinweb-documents skill", () => {
  it("is self-contained and routes agents to its complete references", () => {
    expect(skill).toMatch(/^---\nname: wordinweb-documents\ndescription: .+\n---/);
    expect(skill).toContain("references/interface.md");
    expect(skill).toContain("references/create.md");
    expect(skill).toContain("references/inspect-edit.md");
    expect(existsSync(resolve(skillRoot, "agents/openai.yaml"))).toBe(true);
    expect(interfaceReference).toContain("## Tool contracts");
    expect(interfaceReference).toContain("## Compose generic primitives");
    expect(createReference).toContain("AgentComposeBlock");
    expect(createReference).toContain('import { AgentDocument } from "@wordinweb/agent"');
    expect(createReference).toContain("Page sizes and page margins use inches");
    expect(createReference).toContain("A positioned image has type `anchored-image`");
    expect(createReference).toContain("Use either `text` or `runs`");
    expect(createReference).toContain("Floating drawing positions use page coordinates");
    expect(createReference).toContain("textStyle");
    expect(createReference).toContain("An unpositioned shape uses `topAndBottom` wrap");
    expect(editReference).toContain("Progressive inspection");
    expect(`${skill}\n${interfaceReference}`.toLocaleLowerCase()).not.toContain("pleading");
  });

  it("documents every runtime tool, inspection mode, and reference family", () => {
    for (const tool of AgentDocument.create().tools()) expect(interfaceReference).toContain(`\`${tool.name}\``);
    for (const kind of ["overview", "read", "search", "object", "spatial"]) expect(interfaceReference).toContain(`\"kind\": \"${kind}`);
    for (const prefix of ["block:*", "run:*", "object:*", "asset:*", "spatial:*", "view:*"]) expect(interfaceReference).toContain(`\`${prefix}\``);
  });

  it("keeps every operation and agent-facing field synchronized with runtime capabilities", () => {
    const documented = documentedOperations();
    expect([...documented.keys()].sort()).toEqual([...INTENT_KINDS].sort());
    for (const kind of INTENT_KINDS) {
      const capability = AGENT_EDIT_CAPABILITIES[kind];
      expect(documented.get(kind), `${kind} fields`).toEqual(new Set([...capability.required, ...(capability.optional ?? [])]));
    }
  });
});
