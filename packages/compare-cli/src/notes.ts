import { closeSync, openSync, writeSync } from "node:fs";
import { XmlElement, cyrb53, localName, parseXml, serializeXml } from "../../core/src/xml.js";
import { paragraphText } from "../../core/src/edit/compare/content.js";
import { NotesPart } from "../../core/src/edit/compare/index.js";
import { normalizeText } from "../../core/src/edit/compare/text.js";
import { BlockStore } from "./store.js";

/**
 * A footnotes or endnotes part on disk: each child of the part as XML in a
 * store, and each note's id and pairing key in memory. compareNotes and
 * copyNotesForStruckReferences read and change it as a NotesPart; `write`
 * streams the part back out with the changes.
 */
export class DiskNotes implements NotesPart {
  /** The part's bytes before its first child, and from its end tag on. */
  head = "";
  tail = "";
  private readonly children: BlockStore;
  /** Each child's key: normalizeText of its paragraphText, hashed. */
  private readonly keys: number[] = [];
  /** id -> the last content note with it, and the first note of any type. */
  private readonly lastContent = new Map<string, number>();
  private readonly firstAny = new Map<string, Ref>();
  /** Changed and appended notes, kept apart from the part as read. */
  private readonly changes: BlockStore;
  private readonly changed = new Map<number, number>();
  private readonly appended: number[] = [];
  private next = 1;

  /** `path` is where the stores go: path-children and path-changes. */
  constructor(
    readonly kind: "footnote" | "endnote",
    /** The part's name in the package. */
    readonly part: string,
    path: string,
  ) {
    this.children = new BlockStore(`${path}-children`);
    this.changes = new BlockStore(`${path}-changes`);
  }

  content(id: string): XmlElement | undefined {
    const i = this.lastContent.get(id);
    return i === undefined ? undefined : this.load({ appended: false, at: i });
  }

  key(id: string): number {
    return this.keys[this.lastContent.get(id) as number];
  }

  first(id: string): XmlElement | undefined {
    const ref = this.firstAny.get(id);
    return ref && this.load(ref);
  }

  update(id: string, note: XmlElement): void {
    this.changed.set(this.lastContent.get(id) as number, this.changes.add(serializeXml(note)));
  }

  append(note: XmlElement): void {
    const at = this.changes.add(serializeXml(note));
    const id = note.attrs["w:id"];
    if (localName(note.name) === this.kind && !this.firstAny.has(id)) this.firstAny.set(id, { appended: true, at });
    this.appended.push(at);
    this.next = Math.max(this.next, 1 + (Number(id) || 0));
  }

  nextId(): number {
    return this.next;
  }

  /** Write the part, with its changes, to `path`. */
  write(path: string): void {
    const fd = openSync(path, "w");
    try {
      const out = (bytes: Buffer | string): void => {
        const buf = typeof bytes === "string" ? Buffer.from(bytes, "utf8") : bytes;
        let done = 0;
        while (done < buf.length) done += writeSync(fd, buf, done, buf.length - done);
      };
      out(this.head);
      for (let i = 0; i < this.children.count; i++) {
        const change = this.changed.get(i);
        if (change === undefined) this.children.withBytes(i, out);
        else this.changes.withBytes(change, out);
      }
      for (const at of this.appended) this.changes.withBytes(at, out);
      out(this.tail);
    } finally {
      closeSync(fd);
    }
  }

  close(): void {
    this.children.close();
    this.changes.close();
  }

  /** Add the part's next child, as read. */
  add(el: XmlElement): void {
    const i = this.children.add(serializeXml(el));
    const id = el.attrs["w:id"];
    const note = localName(el.name) === this.kind;
    this.keys.push(note ? parseInt(cyrb53(normalizeText(paragraphText(el))), 36) : 0);
    if (note && !el.attrs["w:type"]) this.lastContent.set(id, i);
    if (note && !this.firstAny.has(id)) this.firstAny.set(id, { appended: false, at: i });
    this.next = Math.max(this.next, 1 + (Number(id) || 0));
  }

  /** A note as it stands now: changed, appended, or as read. */
  private load(ref: Ref): XmlElement {
    if (ref.appended) return parseXml(this.changes.read(ref.at));
    const change = this.changed.get(ref.at);
    return parseXml(change === undefined ? this.children.read(ref.at) : this.changes.read(change));
  }
}

/** Where a note is: in the part as read, or appended. */
interface Ref {
  appended: boolean;
  at: number;
}

