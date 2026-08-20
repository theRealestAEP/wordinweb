import { describe, expect, it } from "vitest";
import type { AgentDocumentViewBinding } from "../src/index.js";

/**
 * The agent and view packages each BUNDLE their own copy of the engine core
 * (tsup `noExternal`, because @wordinweb/core is private and unpublished). Two
 * copies of a class with 79 private members are nominally incompatible, so the
 * desktop app bridges them with `as unknown as AgentDocumentViewBinding`.
 *
 * That cast is load-bearing and unchecked. The danger is not the cast — both
 * sides hold the SAME DocxDocument instance at runtime — it is SILENT DRIFT: if
 * the binding grows a member, or one is renamed, the app keeps compiling and
 * the feature quietly stops working. Word count already degrades to zeros this
 * way rather than failing.
 *
 * So this pins the binding's shape. It fails when the contract changes, which
 * is the moment someone has to go and check the other side of the cast.
 */
describe("AgentDocumentViewBinding shape", () => {
  it("still requires exactly the members the desktop app bridges", () => {
    // A value of the type, so the compiler checks the keys exist AND that none
    // were removed; the list below is then checked for additions.
    const required: Record<keyof AgentDocumentViewBinding, true> = {
      subscribe: true,
      getSnapshot: true,
      doc: true,
      submit: true,
      submitOp: true,
      allocIds: true,
      uploadMedia: true,
      noteHistory: true,
      takeRenderScope: true,
    };
    expect(Object.keys(required).sort()).toEqual(
      [
        "allocIds",
        "doc",
        "getSnapshot",
        "noteHistory",
        "submit",
        "submitOp",
        "subscribe",
        "takeRenderScope",
        "uploadMedia",
      ],
    );
  });
});
