import { describe, expect, it } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { DocxDocument, localName, type XmlElement } from "@wordinweb/core";

/**
 * DocxDocument memoizes two lookups that used to scan the whole document on
 * every edit (perf B9): the child→parent link behind findParentOf, and the
 * block-list location behind paragraphBySource / the targeted reparses.
 *
 * Neither memo is invalidated when the tree or the model changes — they are
 * re-derived on use instead. So what these pin is the only thing that can go
 * wrong: answering from a memo after the document has moved on.
 */

function docWith(bodyXml: string): DocxDocument {
  const documentXml =
    `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${bodyXml}</w:body></w:document>`;
  return DocxDocument.load(
    zipSync({
      "[Content_Types].xml": strToU8(
        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(documentXml),
    }),
  );
}

function find(doc: DocxDocument, name: string): XmlElement {
  const stack = [...doc.editableRoots()];
  while (stack.length) {
    const el = stack.pop()!;
    if (localName(el.name) === name) return el;
    stack.push(...el.children);
  }
  throw new Error(`no ${name}`);
}

describe("findParentOf under a stale parent memo", () => {
  it("re-derives the parent after an element is moved to a new one", () => {
    const doc = docWith(
      `<w:p><w:r><w:t xml:space="preserve">one</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>`,
    );
    const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const [p1, p2] = body.children.filter((c) => localName(c.name) === "p");
    const run = p1.children.find((c) => localName(c.name) === "r")!;
    expect(doc.findParentOf(run)).toBe(p1); // warms the memo

    p1.children = p1.children.filter((c) => c !== run);
    p2.children.push(run);
    expect(doc.findParentOf(run)).toBe(p2);
  });

  it("re-derives the parent when the memoized one was spliced out of the tree", () => {
    // Exactly the shape sub-range formatting produces: the original run is
    // replaced by new runs that take over its children, while the detached run
    // keeps listing those children. A memo that only checked containment would
    // hand back the detached run forever.
    const doc = docWith(`<w:p><w:r><w:t xml:space="preserve">hello</w:t></w:r></w:p>`);
    const para = find(doc, "p");
    const oldRun = para.children.find((c) => localName(c.name) === "r")!;
    const t = oldRun.children.find((c) => localName(c.name) === "t")!;
    expect(doc.findParentOf(t)).toBe(oldRun); // warms the memo

    const newRun: XmlElement = { name: oldRun.name, attrs: {}, text: "", children: [t] };
    para.children.splice(para.children.indexOf(oldRun), 1, newRun);
    expect(oldRun.children).toContain(t); // the detached run still lists it

    expect(doc.findParentOf(t)).toBe(newRun);
    expect(doc.findParentOf(oldRun)).toBeUndefined();
  });

  it("returns undefined for an element that left the tree entirely", () => {
    const doc = docWith(`<w:p><w:r><w:t xml:space="preserve">gone</w:t></w:r></w:p>`);
    const para = find(doc, "p");
    const run = para.children.find((c) => localName(c.name) === "r")!;
    expect(doc.findParentOf(run)).toBe(para);
    para.children = para.children.filter((c) => c !== run);
    expect(doc.findParentOf(run)).toBeUndefined();
    // …and its children go with it: the memo still links them to `run`, but
    // `run` no longer reaches a root.
    expect(doc.findParentOf(run.children[0])).toBeUndefined();
  });
});

describe("paragraphBySource under a stale location memo", () => {
  it("returns the freshly parsed paragraph after a full refresh", () => {
    const doc = docWith(
      `<w:p><w:r><w:t xml:space="preserve">one</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>`,
    );
    const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const source = body.children.filter((c) => localName(c.name) === "p")[1];
    const first = doc.paragraphBySource(source); // warms the memo
    expect(first).not.toBeNull();

    // refresh() replaces every block list AND every parsed paragraph. A memo
    // that survived it would hand back a paragraph in a detached list, and a
    // targeted reparse would then splice into a model nobody renders.
    doc.refresh();
    const after = doc.paragraphBySource(source);
    expect(after).not.toBeNull();
    expect(after).not.toBe(first);
    expect(doc.sections[0].blocks).toContain(after!);
  });

  it("follows a paragraph whose index shifted under an insert", () => {
    const doc = docWith(
      `<w:p><w:r><w:t xml:space="preserve">one</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t xml:space="preserve">two</w:t></w:r></w:p>`,
    );
    const body = doc.docRoot.children.find((c) => localName(c.name) === "body")!;
    const [p1, p2] = body.children.filter((c) => localName(c.name) === "p");
    expect(doc.paragraphBySource(p2)?.src).toBe(p2); // warms the memo at index 1

    // Enter at the start of the first paragraph inserts a sibling before it,
    // pushing every later paragraph one place along in the same block list.
    const blank: XmlElement = {
      name: p1.name,
      attrs: {},
      text: "",
      children: [{ name: "w:r", attrs: {}, text: "", children: [{ name: "w:t", attrs: {}, text: "", children: [] }] }],
    };
    body.children.splice(body.children.indexOf(p1), 0, blank);
    expect(doc.insertDirectBodyParagraphBefore(p1, blank)).not.toBeNull();

    expect(doc.paragraphBySource(p2)?.src).toBe(p2);
    expect(doc.sections[0].blocks.map((b) => b.src)).toEqual([blank, p1, p2]);
  });
});
