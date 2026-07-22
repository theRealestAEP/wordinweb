import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument } from "./helpers.js";

describe("internal bookmark links", () => {
  it("connects an internal hyperlink to its laid-out bookmark target", () => {
    const body =
      `<w:p><w:hyperlink w:anchor="Target"><w:r><w:t>Jump to target</w:t></w:r></w:hyperlink></w:p>` +
      `<w:p><w:r><w:br w:type="page"/></w:r></w:p>` +
      `<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>Destination</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>`;
    const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
    const result = layoutDocument(doc, { measurer: new ApproxMeasurer() });

    const link = result.pages[0].items.find(
      (item) => item.kind === "text" && item.href === "#Target",
    );
    const destination = result.pages[1].items.find(
      (item) => item.kind === "text" && item.text === "Destination",
    );

    expect(link).toMatchObject({ kind: "text", href: "#Target" });
    expect(destination?.kind === "text" ? destination.bookmarks : undefined).toEqual(["Target"]);
  });
});
