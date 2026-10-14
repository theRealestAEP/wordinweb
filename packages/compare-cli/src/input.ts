import { closeSync, openSync, readSync, fstatSync } from "node:fs";
import { join } from "node:path";
import { zipSync } from "fflate";
import { DocxDocument } from "../../core/src/docx.js";
import { XmlElement, child, cyrb53, localName, parseXml, serializeXml } from "../../core/src/xml.js";
import { acceptAllRevisions, collectRevisions } from "../../core/src/edit/suggest.js";
import { paragraphText } from "../../core/src/edit/compare/content.js";
import { blockFingerprint, sectionKey } from "../../core/src/edit/compare/index.js";
import { scanBody } from "./scan.js";
import { BlockStore } from "./store.js";
import { ZipReader } from "./zip.js";

export type Side = "original" | "revised";

/**
 * A binary part's stand-in inside a shell: the shell needs to know the part
 * exists, not what it holds. The output writer copies the real bytes from
 * the input zip the stub names.
 */
const STUB = "\u0000dxw-stub\u0000";

/** The input zip and part name a stub stands for, or null for real content. */
export function stubSource(bytes: Uint8Array): { side: Side; name: string } | null {
  if (bytes.length < STUB.length || bytes[0] !== 0 || Buffer.from(bytes.subarray(0, STUB.length)).toString("latin1") !== STUB) return null;
  const [side, name] = Buffer.from(bytes.subarray(STUB.length)).toString("utf8").split("\u0000");
  return { side: side as Side, name };
}

/** Tracked-change markup. */
const TRACKED_CHANGE =
  /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|trPrChange|tcPrChange|numberingChange|cellIns|cellDel|cellMerge)[\s>/]/;

/** Blocks with tracked changes are accepted in batches of about this much XML. */
const BATCH_BYTES = 4 << 20;

/** One input document: its body on disk, block by block, and everything else in a shell. */
export interface Input {
  side: Side;
  zip: ZipReader;
  docPart: string;
  /**
   * The document with every part but the body: the body holds only its own
   * section properties, and each binary part is a stub.
   */
  shell: DocxDocument;
  /** Each body block as XML, with tracked changes accepted. */
  blocks: BlockStore;
  /** Each block's paragraphText. */
  texts: BlockStore;
  /** Each block's blockFingerprint, hashed: the anchor pass's keys. */
  fingerprints: string[];
  /** A hash of each block's XML. */
  hashes: string[];
  /** Each block's local name. */
  kinds: string[];
  /** Whether each block is a paragraph that ends a section. */
  endsSection: boolean[];
  /** The sectionKey of each block that ends a section. */
  sectionKeys: Map<number, string>;
  /** The w:sectPr of each block that ends a section, in order. */
  sectPrs: XmlElement[];
  /** Note reference ids in document order. */
  noteRefs: Record<"footnote" | "endnote", string[]>;
  /** Whether the input held tracked changes, now accepted. */
  hadRevisions: boolean;
}

/** Read one input into `dir`. */
export async function readInput(zip: ZipReader, side: Side, dir: string): Promise<Input> {
  const docPart = documentPart(zip);
  const docEntry = zip.entries.get(docPart);
  if (!docEntry) throw new Error(`${zip.path}: missing ${docPart}`);
  const docFile = join(dir, `${side}-document.xml`);
  await zip.extract(docEntry, docFile);

  // First the body's outline: where it is, and where each child is.
  const starts: number[] = [];
  const ends: number[] = [];
  const names: string[] = [];
  let bounds;
  try {
    bounds = scanBody(docFile, (c) => {
      starts.push(c.start);
      ends.push(c.end);
      names.push(c.name);
    });
  } catch (error) {
    throw new Error(`${docPart}: ${error instanceof Error ? error.message : String(error)}`);
  }

  const fd = openSync(docFile, "r");
  try {
    // One buffer for every block read, so the reads leave no garbage outside the heap.
    let scratch = Buffer.allocUnsafe(64 * 1024);
    const range = (start: number, end: number): Buffer => {
      if (end - start > scratch.length) scratch = Buffer.allocUnsafe(Math.max(end - start, scratch.length * 2));
      const buf = scratch.subarray(0, end - start);
      let done = 0;
      while (done < buf.length) done += readSync(fd, buf, done, buf.length - done, start + done);
      return buf;
    };
    const isSectPr = (i: number): boolean => localName(names[i]) === "sectPr";

    // The shell: document.xml with only the body's own section properties.
    const files: Record<string, Uint8Array> = {};
    for (const entry of zip.entries.values()) {
      if (entry.name === docPart || entry.name.endsWith("/")) continue;
      files[entry.name] = /\.(xml|rels|vml)$/i.test(entry.name) ? zip.read(entry) : Buffer.from(`${STUB}${side}\u0000${entry.name}`, "utf8");
    }
    const skeleton: Buffer[] = [Buffer.from(range(0, bounds.contentStart))];
    for (let i = 0; i < names.length; i++) if (isSectPr(i)) skeleton.push(Buffer.from(range(starts[i], ends[i])));
    skeleton.push(Buffer.from(range(bounds.contentEnd, fstatSync(fd).size)));
    files[docPart] = Buffer.concat(skeleton);
    const shell = DocxDocument.load(zipSync(files, { level: 0 }));
    let hadRevisions = false;
    if (collectRevisions(shell).length > 0) {
      acceptAllRevisions(shell);
      hadRevisions = true;
    }

    const input: Input = {
      side,
      zip,
      docPart,
      shell,
      blocks: new BlockStore(join(dir, `${side}-blocks`)),
      texts: new BlockStore(join(dir, `${side}-texts`)),
      fingerprints: [],
      hashes: [],
      kinds: [],
      endsSection: [],
      sectionKeys: new Map(),
      sectPrs: [],
      noteRefs: { footnote: [], endnote: [] },
      hadRevisions,
    };

    const add = (el: XmlElement): void => {
      const i = input.blocks.count;
      const xml = serializeXml(el);
      input.blocks.add(xml);
      input.texts.add(paragraphText(el));
      input.fingerprints.push(cyrb53(blockFingerprint(el, shell)));
      input.hashes.push(cyrb53(xml));
      input.kinds.push(localName(el.name));
      const sectPr = localName(el.name) === "p" ? child(child(el, "pPr"), "sectPr") : undefined;
      input.endsSection.push(!!sectPr);
      if (sectPr) {
        input.sectionKeys.set(i, sectionKey(el, shell));
        input.sectPrs.push(sectPr);
      }
      const walk = (e: XmlElement): void => {
        const ln = localName(e.name);
        if (ln === "footnoteReference") input.noteRefs.footnote.push(e.attrs["w:id"]);
        else if (ln === "endnoteReference") input.noteRefs.endnote.push(e.attrs["w:id"]);
        for (const c of e.children) walk(c);
      };
      walk(el);
    };

    // Blocks with tracked changes go through the shell's body, a batch at a
    // time, to be accepted as acceptAllRevisions accepts a whole document. A
    // deleted paragraph mark joins its paragraph to the next block, so a
    // batch never ends on one.
    const body = child(shell.docRoot, "body");
    if (!body) throw new Error(`${zip.path}: document.xml has no w:body`);
    const tail = body.children;
    let batch: XmlElement[] = [];
    let batchBytes = 0;
    const flush = (): void => {
      if (batch.length === 0) return;
      body.children = [...batch, ...tail];
      acceptAllRevisions(shell);
      const accepted = body.children.filter((c) => !tail.includes(c));
      body.children = tail;
      batch = [];
      batchBytes = 0;
      for (const el of accepted) add(el);
    };
    for (let i = 0; i < names.length; i++) {
      if (isSectPr(i)) continue;
      const xml = range(starts[i], ends[i]).toString("utf8");
      const el = parseXml(xml);
      if (batch.length > 0 || TRACKED_CHANGE.test(xml)) {
        hadRevisions = true;
        batch.push(el);
        batchBytes += xml.length;
        if (batchBytes >= BATCH_BYTES && !joinsNext(el)) flush();
      } else {
        add(el);
      }
    }
    flush();
    input.hadRevisions = hadRevisions;
    return input;
  } finally {
    closeSync(fd);
  }
}

/** A paragraph whose mark is deleted or moved away: accepting it joins the next block. */
function joinsNext(el: XmlElement): boolean {
  const rPr = child(child(el, "pPr"), "rPr");
  return localName(el.name) === "p" && !!(child(rPr, "del") || child(rPr, "moveFrom"));
}

/** The main document part, as the package relationships name it. */
function documentPart(zip: ZipReader): string {
  const rels = zip.entries.get("_rels/.rels");
  if (rels) {
    const root = parseXml(zip.read(rels).toString("utf8"));
    for (const rel of root.children) {
      if ((rel.attrs.Type ?? "").endsWith("/officeDocument")) return (rel.attrs.Target ?? "").replace(/^\//, "");
    }
  }
  return "word/document.xml";
}
