import { describe, expect, it } from "vitest";
import { INTENT_KINDS } from "@wordinweb/collab/client";
import { AgentDocument, agentOperationSchema, hoistRepeatedSubschemas } from "../src/index.js";

type JsonSchema = Record<string, unknown>;

/** Put every `$ref` back, so the result can be compared with the schema the
 * hoist started from. */
function expandRefs(node: unknown, defs: Record<string, unknown>): unknown {
  if (Array.isArray(node)) return node.map((entry) => expandRefs(entry, defs));
  if (!node || typeof node !== "object") return node;
  const ref = (node as JsonSchema).$ref;
  if (typeof ref === "string") return expandRefs(defs[ref.replace("#/$defs/", "")], defs);
  return Object.fromEntries(Object.entries(node as JsonSchema).map(([key, value]) => [key, expandRefs(value, defs)]));
}

function resolved(schema: JsonSchema): unknown {
  const { $defs, ...rest } = schema;
  return expandRefs(rest, ($defs ?? {}) as Record<string, unknown>);
}

describe("tool schema $defs hoisting", () => {
  it("leaves a schema with nothing repeated alone", () => {
    const schema: JsonSchema = {
      type: "object",
      properties: { a: { type: "string", pattern: "^a$" }, b: { type: "integer", minimum: 0 } },
      additionalProperties: false,
    };
    expect(hoistRepeatedSubschemas(schema)).toEqual(schema);
  });

  it("collapses a repeated shape and resolves back to the original", () => {
    const point: JsonSchema = {
      type: "object",
      properties: { blockRef: { type: "string", pattern: "^block:[0-9]+$" }, offset: { type: "integer", minimum: 0 } },
      required: ["blockRef", "offset"],
      additionalProperties: false,
    };
    const schema: JsonSchema = {
      type: "object",
      properties: { from: point, to: point, via: point },
      required: ["from"],
      additionalProperties: false,
    };
    const hoisted = hoistRepeatedSubschemas(schema);
    expect(Object.keys(hoisted.$defs as object)).toEqual(["from"]);
    expect((hoisted.properties as JsonSchema).to).toEqual({ $ref: "#/$defs/from" });
    expect(JSON.stringify(hoisted).length).toBeLessThan(JSON.stringify(schema).length);
    expect(resolved(hoisted)).toEqual(schema);
  });

  it("leaves a repeat too small to pay for a reference alone", () => {
    const schema: JsonSchema = {
      type: "object",
      properties: { a: { type: "boolean" }, b: { type: "boolean" }, c: { type: "boolean" } },
    };
    expect(hoistRepeatedSubschemas(schema)).toEqual(schema);
  });

  it("substitutes inside anyOf branches and array items", () => {
    const edge: JsonSchema = {
      type: "object",
      properties: { style: { enum: ["single", "double", "dashed", "dotted"] }, widthPt: { type: "number", minimum: 1 } },
      required: ["style"],
      additionalProperties: false,
    };
    const schema: JsonSchema = {
      type: "object",
      properties: {
        top: { anyOf: [edge, { type: "null" }] },
        bottom: { anyOf: [edge, { type: "null" }] },
        others: { type: "array", items: edge },
      },
    };
    const hoisted = hoistRepeatedSubschemas(schema);
    expect(resolved(hoisted)).toEqual(schema);
    expect(JSON.stringify(hoisted)).toContain('"$ref"');
  });

  it("writes an address out in full rather than behind a reference", () => {
    const shipped = AgentDocument.create().tools().find((t) => t.name === "word_document_edit")!.inputSchema;
    const wire = JSON.stringify(shipped);
    const defs = JSON.stringify(shipped.$defs);
    // The three reference patterns are the most repeated shapes in the union
    // and the most attractive to hoist. They ship spelled out, at every site,
    // and nothing in $defs stands in for one.
    for (const pattern of ["^block:[0-9]+$", "^run:[0-9]+$", "^object:[0-9]+:[0-9]+$"]) {
      expect(wire.split(pattern).length - 1).toBeGreaterThan(10);
    }
    expect(defs).not.toContain("block:[0-9]");
    expect(defs).not.toContain("run:[0-9]");
    expect(defs).not.toContain("object:[0-9]");
    for (const name of ["at", "blockRef", "runRef", "objectRef", "cellRef", "afterBlockRef", "offset"]) {
      expect(wire).not.toContain(`{"$ref":"#/$defs/${name}"}`);
    }
  });

  it("is a fixed point: hoisting the result changes nothing", () => {
    const schema = AgentDocument.create().tools().find((t) => t.name === "word_document_edit")!.inputSchema;
    expect(hoistRepeatedSubschemas(schema)).toEqual(schema);
  });

  it("keeps the shipped edit tool schema equal to the union it was built from", () => {
    const shipped = AgentDocument.create().tools().find((t) => t.name === "word_document_edit")!.inputSchema;
    expect(resolved(shipped)).toEqual({
      type: "object",
      properties: {
        revision: { type: "string", minLength: 1 },
        operations: { type: "array", minItems: 1, maxItems: 100, items: { anyOf: INTENT_KINDS.map(agentOperationSchema) } },
      },
      required: ["revision", "operations"],
      additionalProperties: false,
    });
    // The whole point: the wire form is materially smaller than the union.
    const union = JSON.stringify({ anyOf: INTENT_KINDS.map(agentOperationSchema) }).length;
    expect(JSON.stringify(shipped).length).toBeLessThan(union * 0.9);
  });
});
