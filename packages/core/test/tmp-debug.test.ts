import { it } from "vitest";
import { readFileSync } from "fs";
import { DocxDocument } from "../src/docx.js";
import { insertPageField } from "../src/edit/fields.js";
import { XmlElement, localName } from "../src/xml.js";
it("debug pn insert in alexpickett", () => {
  const buf = readFileSync(new URL("../../../apps/demo/public/fixtures/alexpickett.docx", import.meta.url));
  const doc = DocxDocument.load(new Uint8Array(buf));
  // find the w:t containing EMAIL
  let target: XmlElement | null = null;
  const walk = (e: XmlElement): void => {
    if (localName(e.name) === "t" && e.text.includes("EMAIL")) target = target ?? e;
    for (const c of e.children) walk(c);
  };
  walk(doc.editableRoots()[0]);
  // maybe it is stored with caps formatting, search case-insensitively
  if (!target) {
    const walk2 = (e: XmlElement): void => {
      if (localName(e.name) === "t" && /email/i.test(e.text)) target = target ?? e;
      for (const c of e.children) walk2(c);
    };
    walk2(doc.editableRoots()[0]);
  }
  console.log("target text:", target ? (target as XmlElement).text : "NOT FOUND");
  if (!target) return;
  console.log("insert ok:", insertPageField(doc, target, (target as XmlElement).text.length, "pageOfTotal"));
  const parent = doc.findParentOf(target)!;
  const gp = doc.findParentOf(parent)!;
  const dump = (e: XmlElement): string => `<${e.name}>` + (e.text || "") + e.children.map(dump).join("");
  console.log("para after:", dump(gp).slice(0, 700));
  // model view: find the paragraph containing Email and print run contents
  let found = false;
  const scan = (blocks: import("../src/model.js").Block[]): void => {
    for (const b of blocks) {
      if (b.type === "table") { for (const r of b.rows) for (const c of r.cells) scan(c.blocks); continue; }
      const parts: string[] = [];
      for (const ch of b.children) {
        const runs = ch.type === "run" ? [ch] : ch.runs;
        for (const r of runs) for (const rc of r.content) {
          if (rc.kind === "text") parts.push("T:" + rc.text);
          else parts.push(rc.kind + (rc.kind === "field" ? ":" + rc.instruction.trim() : ""));
        }
      }
      if (parts.some((p) => p.includes("Email"))) { console.log("MODEL:", parts.join(" | ")); found = true; }
    }
  };
  for (const sec of doc.sections) scan(sec.blocks);
  if (!found) console.log("MODEL: paragraph not found");
});
