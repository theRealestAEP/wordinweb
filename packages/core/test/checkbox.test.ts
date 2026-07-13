import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { checkboxStateElement, checkboxChecked, toggleCheckbox } from "../src/checkbox.js";
import { serializeXml, child, attr, localName } from "../src/xml.js";
import { Paragraph, Run } from "../src/model.js";
import { makeDocx, wrapDocument } from "./helpers.js";

const LEGACY = `<w:p><w:r>
  <w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Sub"/>
    <w:checkBox><w:default w:val="1"/><w:checked w:val="1"/></w:checkBox>
  </w:ffData></w:fldChar>
</w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;

const MODERN = `<w:p><w:sdt><w:sdtPr><w:id w:val="200"/>
  <w14:checkbox><w14:checked w14:val="0"/>
    <w14:checkedState w14:val="2612" w14:font="MS Gothic"/>
    <w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>
  </w14:checkbox></w:sdtPr>
  <w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic"/></w:rPr><w:t>&#9744;</w:t></w:r></w:sdtContent></w:sdt></w:p>`;

function loadDoc(body: string): DocxDocument {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

/** Every rendered checkbox glyph in the document, with its state element. */
function checkboxes(doc: DocxDocument): { run: Run; el: ReturnType<typeof checkboxStateElement> }[] {
  const out: { run: Run; el: ReturnType<typeof checkboxStateElement> }[] = [];
  for (const block of doc.sections[0].blocks) {
    if (block.type !== "paragraph") continue;
    for (const pc of (block as Paragraph).children) {
      const runs = pc.type === "run" ? [pc] : pc.runs;
      for (const r of runs) {
        // legacy: field content (t=null)
        let el = checkboxStateElement(r, null);
        if (el) out.push({ run: r, el });
        // modern: per text content's srcT
        for (const c of r.content) {
          if (c.kind === "text" && c.srcT) {
            el = checkboxStateElement(r, c.srcT);
            if (el) out.push({ run: r, el });
          }
        }
      }
    }
  }
  return out;
}

function docXml(doc: DocxDocument): string {
  return serializeXml(doc.editableRoots()[0]);
}

describe("checkbox parse markers", () => {
  it("legacy FORMCHECKBOX field content carries its w:checkBox element", () => {
    const doc = loadDoc(LEGACY);
    const boxes = checkboxes(doc);
    expect(boxes).toHaveLength(1);
    expect(localName(boxes[0].el!.name)).toBe("checkBox");
    expect(checkboxChecked(boxes[0].el!)).toBe(true);
  });

  it("modern content control glyph carries its w14:checkbox element", () => {
    const doc = loadDoc(MODERN);
    const boxes = checkboxes(doc);
    expect(boxes).toHaveLength(1);
    expect(localName(boxes[0].el!.name)).toBe("checkbox");
    expect(checkboxChecked(boxes[0].el!)).toBe(false);
  });

  it("ordinary text is not a checkbox target", () => {
    const doc = loadDoc(`<w:p><w:r><w:t>hello</w:t></w:r></w:p>`);
    expect(checkboxes(doc)).toHaveLength(0);
  });
});

describe("toggleCheckbox", () => {
  it("legacy: flips w:checked and re-derives the glyph on refresh", () => {
    const doc = loadDoc(LEGACY);
    const before = checkboxes(doc)[0];
    // Glyph starts checked (☒).
    const fieldOf = (d: DocxDocument) =>
      (d.sections[0].blocks[0] as Paragraph).children
        .flatMap((pc) => (pc.type === "run" ? [pc] : pc.runs))
        .flatMap((r) => r.content)
        .find((c) => c.kind === "field");
    expect((fieldOf(doc) as { cachedResult: string }).cachedResult).toBe("☒");

    const next = toggleCheckbox(doc, before.el!);
    expect(next).toBe(false);
    expect(attr(child(before.el!, "checked"), "val")).toBe("0");
    doc.refresh();
    expect((fieldOf(doc) as { cachedResult: string }).cachedResult).toBe("☐");
  });

  it("modern: flips w14:checked and swaps the glyph char to checkedState", () => {
    const doc = loadDoc(MODERN);
    const box = checkboxes(doc)[0];
    expect(box.run.content.find((c) => c.kind === "text")).toMatchObject({ text: "☐" });

    const next = toggleCheckbox(doc, box.el!);
    expect(next).toBe(true);
    expect(attr(child(box.el!, "checked"), "val")).toBe("1");
    // Glyph w:t rewritten to U+2612 (checkedState val="2612").
    expect(docXml(doc)).toContain("☒");
    doc.refresh();
    const box2 = checkboxes(doc)[0];
    expect(box2.run.content.find((c) => c.kind === "text")).toMatchObject({ text: "☒" });
    expect(checkboxChecked(box2.el!)).toBe(true);
  });

  it("modern: unchecking swaps back to uncheckedState char", () => {
    const doc = loadDoc(MODERN.replace('w14:val="0"', 'w14:val="1"').replace("&#9744;", "&#9746;"));
    const box = checkboxes(doc)[0];
    expect(checkboxChecked(box.el!)).toBe(true);
    toggleCheckbox(doc, box.el!);
    expect(attr(child(box.el!, "checked"), "val")).toBe("0");
    expect(docXml(doc)).toContain("☐");
  });

  it("round-trips a toggled state through save() + reload", () => {
    const doc = loadDoc(LEGACY + MODERN);
    const [legacy, modern] = checkboxes(doc);
    toggleCheckbox(doc, legacy.el!); // checked -> unchecked
    toggleCheckbox(doc, modern.el!); // unchecked -> checked
    const reloaded = DocxDocument.load(doc.save());
    const [legacy2, modern2] = checkboxes(reloaded);
    expect(checkboxChecked(legacy2.el!)).toBe(false);
    expect(checkboxChecked(modern2.el!)).toBe(true);
    // Modern glyph persisted as the checked char.
    expect(modern2.run.content.find((c) => c.kind === "text")).toMatchObject({ text: "☒" });
  });

  it("legacy: creates a w:checked element when absent (default unchecked)", () => {
    const noChecked = LEGACY.replace("<w:checked w:val=\"1\"/>", "");
    const doc = loadDoc(noChecked);
    const box = checkboxes(doc)[0];
    expect(checkboxChecked(box.el!)).toBe(false);
    const next = toggleCheckbox(doc, box.el!);
    expect(next).toBe(true);
    expect(attr(child(box.el!, "checked"), "val")).toBe("1");
  });
});
