import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { addComment, deleteComment, editCommentText, replyToComment, setCommentResolved } from "../src/edit/comments.js";
import { SelectionSegment } from "../src/edit/commands.js";
import { Paragraph, Run, TextContent } from "../src/model.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";

function loadDoc(body: string) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

function segFor(doc: DocxDocument, start: number, end: number): SelectionSegment {
  const para = doc.sections[0].blocks[0] as Paragraph;
  const run = para.children[0] as Run;
  const t = (run.content.find((c) => c.kind === "text") as TextContent).srcT;
  return { run, t: t as SelectionSegment["t"], start, end, props: run.props };
}

function docWithThread(): DocxDocument {
  const doc = loadDoc(p("Please review this sentence."));
  expect(addComment(doc, [segFor(doc, 7, 13)], "Looks wrong", "Alice", "AL")).toBe(true);
  expect(replyToComment(doc, doc.comments[0].id, "Agreed", "Bob", "BO")).toBe(true);
  return doc;
}

describe("setCommentResolved", () => {
  it("marks the thread parent done=1 in commentsExtended and derives resolved", () => {
    const doc = docWithThread();
    const parent = doc.comments.find((c) => !c.parentId)!;
    expect(parent.resolved).toBeUndefined();
    expect(setCommentResolved(doc, parent.id, true)).toBe(true);
    expect(doc.comments.find((c) => !c.parentId)!.resolved).toBe(true);
    // The reply's own entry stays open — Word grays the thread from the parent.
    expect(doc.comments.find((c) => c.parentId)!.resolved).toBeUndefined();
    const files = unzipSync(doc.save());
    const ext = strFromU8(files["word/commentsExtended.xml"]);
    expect(ext).toContain('w15:done="1"');
  });

  it("resolving again is a no-op; reopen flips done back to 0", () => {
    const doc = docWithThread();
    const parent = doc.comments.find((c) => !c.parentId)!;
    expect(setCommentResolved(doc, parent.id, true)).toBe(true);
    expect(setCommentResolved(doc, parent.id, true)).toBe(false);
    expect(setCommentResolved(doc, parent.id, false)).toBe(true);
    expect(doc.comments.find((c) => !c.parentId)!.resolved).toBeUndefined();
    // Reopening an already-open thread with no commentsExtended entry: no-op.
    expect(setCommentResolved(doc, "does-not-exist", false)).toBe(false);
  });

  it("a reply id resolves its thread's parent", () => {
    const doc = docWithThread();
    const reply = doc.comments.find((c) => c.parentId)!;
    expect(setCommentResolved(doc, reply.id, true)).toBe(true);
    expect(doc.comments.find((c) => !c.parentId)!.resolved).toBe(true);
  });

  it("a comment with no thread entry yet gets one (single comment, no replies)", () => {
    const doc = loadDoc(p("Some commented text here."));
    expect(addComment(doc, [segFor(doc, 0, 4)], "Note", "Alice")).toBe(true);
    const id = doc.comments[0].id;
    expect(setCommentResolved(doc, id, true)).toBe(true);
    expect(doc.comments[0].resolved).toBe(true);
    // An open comment with no entry resolves to "already open".
    expect(setCommentResolved(doc, id, false)).toBe(true);
    expect(setCommentResolved(doc, id, false)).toBe(false);
  });
});

describe("editCommentText", () => {
  it("replaces the body text, keeping author, date, id, anchors and threading", () => {
    const doc = docWithThread();
    const parent = doc.comments.find((c) => !c.parentId)!;
    const { author, date, paraId } = parent;
    expect(editCommentText(doc, parent.id, "Actually it is fine")).toBe(true);
    const after = doc.comments.find((c) => c.id === parent.id)!;
    expect(after.text).toBe("Actually it is fine");
    expect(after.author).toBe(author);
    expect(after.date).toBe(date);
    // The threading key survives, so the reply still points at the parent.
    expect(after.paraId).toBe(paraId);
    expect(doc.comments.find((c) => c.parentId === parent.id)).toBeTruthy();
    // The document anchors are untouched.
    expect(doc.commentAnchors().get(parent.id)?.length).toBeGreaterThan(0);
  });

  it("newlines become paragraphs; empty/unchanged/unknown edits reject", () => {
    const doc = docWithThread();
    const parent = doc.comments.find((c) => !c.parentId)!;
    expect(editCommentText(doc, parent.id, "line one\nline two")).toBe(true);
    expect(doc.comments.find((c) => c.id === parent.id)!.text).toBe("line one\nline two");
    expect(editCommentText(doc, parent.id, "line one\nline two")).toBe(false);
    expect(editCommentText(doc, parent.id, "   ")).toBe(false);
    expect(editCommentText(doc, "nope", "text")).toBe(false);
  });

  it("editing a reply keeps its threading too", () => {
    const doc = docWithThread();
    const reply = doc.comments.find((c) => c.parentId)!;
    expect(editCommentText(doc, reply.id, "Disagreed after all")).toBe(true);
    const after = doc.comments.find((c) => c.id === reply.id)!;
    expect(after.text).toBe("Disagreed after all");
    expect(after.parentId).toBe(doc.comments.find((c) => !c.parentId)!.id);
  });
});

describe("delete thread vs delete reply", () => {
  it("deleting a reply keeps the parent; deleting the parent removes the thread", () => {
    const doc = docWithThread();
    const parent = doc.comments.find((c) => !c.parentId)!;
    const reply = doc.comments.find((c) => c.parentId)!;
    expect(deleteComment(doc, reply.id)).toBe(true);
    expect(doc.comments.map((c) => c.id)).toEqual([parent.id]);
    expect(deleteComment(doc, parent.id)).toBe(true);
    expect(doc.comments).toHaveLength(0);
  });
});
