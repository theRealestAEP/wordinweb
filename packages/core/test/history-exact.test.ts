import { describe, expect, it } from "vitest";
import { createHash } from "node:crypto";
import { unzipSync } from "fflate";
import { DocxDocument } from "../src/docx.js";
import { EditHistory } from "../src/edit/history.js";
import { applyRunFormat, SelectionSegment } from "../src/edit/commands.js";
import { applyTableOp } from "../src/edit/tables.js";
import { XmlElement, parseXml } from "../src/xml.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
import { Paragraph, Run, TextContent } from "../src/model.js";

/**
 * Undo/redo must restore EXACT document states — byte-identical saved files,
 * at every step of a mixed edit sequence, in both directions. This is the
 * guard for the delta-based history: a delta that under-captures (shares a
 * subtree it should have copied, or misclassifies a structural edit as
 * text-only) corrupts a state silently, and only a full-save byte comparison
 * catches it.
 */

function loadDoc(body: string) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(body) }));
}

/** Hash of every part of the saved package — the whole document, as bytes. */
function fingerprint(doc: DocxDocument): string {
  const parts = unzipSync(doc.save());
  const h = createHash("sha256");
  for (const name of Object.keys(parts).sort()) {
    h.update(name);
    h.update(parts[name]);
  }
  return h.digest("hex");
}

function findT(doc: DocxDocument, text: string): XmlElement {
  let found: XmlElement | null = null;
  const walk = (el: XmlElement) => {
    for (const c of el.children) {
      if (c.name.endsWith("t") && c.text === text) found = c;
      walk(c);
    }
  };
  walk(doc.editableRoots()[0]);
  if (!found) throw new Error(`no w:t with text ${JSON.stringify(text)}`);
  return found;
}

function segFor(run: Run, start: number, end: number): SelectionSegment {
  const t = (run.content.find((c) => c.kind === "text") as TextContent | undefined)?.srcT ?? null;
  return { run, t: t as SelectionSegment["t"], start, end, props: run.props };
}

const TABLE =
  `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
   <w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc></w:tr>
   </w:tbl>`;

describe("history restores byte-identical documents", () => {
  it("mixed edits undo and redo exactly at every step", () => {
    const doc = loadDoc(p("Hello world") + TABLE + p("tail"));
    const history = new EditHistory(doc);
    const body = doc.editableRoots()[0].children[0];

    // Each edit: checkpoint (no coalesce key -> one step per edit), mutate,
    // refresh, record the exact saved bytes.
    const states: string[] = [fingerprint(doc)];
    const step = (mutate: () => void) => {
      history.checkpoint();
      mutate();
      doc.refresh();
      states.push(fingerprint(doc));
    };

    // 1. Typing (text-only delta path).
    step(() => {
      findT(doc, "Hello world").text = "Hello brave world";
    });
    // 2. Deleting (text-only delta path).
    step(() => {
      findT(doc, "Hello brave world").text = "Hello brave";
    });
    // 3. Formatting — splits the run (structural fallback).
    step(() => {
      const para = doc.sections[0].blocks[0] as Paragraph;
      applyRunFormat(doc, [segFor(para.children[0] as Run, 6, 11)], { bold: true });
    });
    // 4. Paragraph split (structural fallback).
    step(() => {
      const tail = findT(doc, "tail");
      const paraIdx = body.children.findIndex((c) => {
        let hit = false;
        const scan = (el: XmlElement) => {
          if (el === tail) hit = true;
          for (const k of el.children) scan(k);
        };
        scan(c);
        return hit;
      });
      tail.text = "ta";
      body.children.splice(paraIdx + 1, 0, parseXml(`<w:p><w:r><w:t>il</w:t></w:r></w:p>`));
    });
    // 5. Table operation (structural fallback).
    step(() => {
      applyTableOp(doc, findT(doc, "A1"), "rowBelow");
    });
    // 6. Typing again after structural boundaries (delta path resumes).
    step(() => {
      findT(doc, "A1").text = "A1 edited";
    });

    // Undo the whole ladder; every intermediate state must be exact.
    for (let i = states.length - 1; i > 0; i--) {
      expect(history.undo()).toBe(true);
      expect(fingerprint(doc)).toBe(states[i - 1]);
    }
    expect(history.undo()).toBe(false);
    expect(fingerprint(doc)).toBe(states[0]);

    // Redo the whole ladder.
    for (let i = 1; i < states.length; i++) {
      expect(history.redo()).toBe(true);
      expect(fingerprint(doc)).toBe(states[i]);
    }
    expect(history.redo()).toBe(false);
    expect(fingerprint(doc)).toBe(states[states.length - 1]);

    // And back down again — redo must not have corrupted the undo chain.
    for (let i = states.length - 1; i > 0; i--) {
      expect(history.undo()).toBe(true);
      expect(fingerprint(doc)).toBe(states[i - 1]);
    }
  });

  it("click-type-click-type keeps one exact step per checkpoint", () => {
    // The caret-click pattern: every keystroke lands under a fresh coalesce
    // key, so every checkpoint is its own undo step (the case the delta
    // capture exists for).
    const doc = loadDoc(p("abc"));
    const history = new EditHistory(doc);
    const states: string[] = [fingerprint(doc)];
    for (let i = 0; i < 6; i++) {
      history.checkpoint(i % 2 ? "typing" : "click-typing");
      const t = doc.editableRoots()[0].children[0].children[0].children[0].children[0];
      t.text += String.fromCharCode(100 + i);
      doc.refresh();
      states.push(fingerprint(doc));
    }
    for (let i = states.length - 1; i > 0; i--) {
      expect(history.undo()).toBe(true);
      expect(fingerprint(doc)).toBe(states[i - 1]);
    }
    expect(history.undo()).toBe(false);
    for (let i = 1; i < states.length; i++) {
      expect(history.redo()).toBe(true);
      expect(fingerprint(doc)).toBe(states[i]);
    }
  });

  it("a structural edit after an undo takes the fallback and stays exact", () => {
    // Undo into the middle of history, edit structurally (clears redo),
    // then unwind — exercises the shadow rewind across a discarded branch.
    const doc = loadDoc(p("one") + p("two"));
    const history = new EditHistory(doc);
    const body = doc.editableRoots()[0].children[0];

    const s0 = fingerprint(doc);
    history.checkpoint();
    findT(doc, "one").text = "one!";
    doc.refresh();
    const s1 = fingerprint(doc);
    history.checkpoint();
    findT(doc, "two").text = "two!";
    doc.refresh();

    expect(history.undo()).toBe(true);
    expect(fingerprint(doc)).toBe(s1);

    history.checkpoint();
    body.children.splice(1, 0, parseXml(`<w:p><w:r><w:t>inserted</w:t></w:r></w:p>`));
    doc.refresh();
    const s2 = fingerprint(doc);
    expect(history.canRedo).toBe(false);

    expect(history.undo()).toBe(true);
    expect(fingerprint(doc)).toBe(s1);
    expect(history.undo()).toBe(true);
    expect(fingerprint(doc)).toBe(s0);
    expect(history.redo()).toBe(true);
    expect(fingerprint(doc)).toBe(s1);
    expect(history.redo()).toBe(true);
    expect(fingerprint(doc)).toBe(s2);
  });
});
