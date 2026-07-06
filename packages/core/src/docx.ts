import { Package } from "./zip.js";
import { XmlElement, parseXml, serializeXml, child, intAttr, onOff, attr, localName } from "./xml.js";
import { strToU8, zipSync } from "fflate";
import { twipsToPx } from "./units.js";
import {
  Block,
  DocComment,
  HeaderFooter,
  Numbering,
  ParaProps,
  Paragraph,
  RunProps,
  Section,
  Styles,
  Theme,
} from "./model.js";
import { parseTheme } from "./parse/theme.js";
import { parseStyles, resolveCharacterStyleChain, resolveParagraphStyleChain } from "./parse/styles.js";
import { parseNumbering } from "./parse/numbering.js";
import { parseBody, parseBlocks, DocParseContext } from "./parse/document.js";
import { parseNotesPart } from "./parse/notes.js";
import { Relationships, parseRelationships, relsPathFor } from "./parse/rels.js";
import { mergeParaProps, mergeRunProps } from "./parse/properties.js";

const REL_TYPE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/**
 * A fully parsed .docx: sections of blocks, styles, numbering, theme, and
 * header/footer parts, with helpers to resolve effective formatting.
 */
export class DocxDocument {
  readonly pkg: Package;
  readonly theme: Theme;
  readonly styles: Styles;
  readonly numbering: Numbering;
  sections: Section[] = [];
  /** Header/footer parts keyed by relationship id from document.xml.rels. */
  readonly headers: Map<string, HeaderFooter> = new Map();
  readonly footers: Map<string, HeaderFooter> = new Map();
  /** Note content by note id (render-only; sources stripped). */
  readonly footnotes: Map<number, Block[]> = new Map();
  readonly endnotes: Map<number, Block[]> = new Map();
  readonly documentRels: Relationships;
  /** settings.xml w:evenAndOddHeaders — enables the "even" header/footer variants. */
  readonly evenAndOddHeaders: boolean = false;
  /** settings.xml w:defaultTabStop in px (Word default 0.5"). */
  readonly defaultTabStop: number = 48;
  /** Review comments from word/comments.xml (empty when the part is absent). */
  readonly comments: DocComment[] = [];

  /** Retained XML roots — source of truth for editing and save(). */
  private readonly docPart: string;
  private readonly docRoot: XmlElement;
  private readonly hfParts: { relId: string; target: string; root: XmlElement; isHeader: boolean; rels: Relationships }[] = [];
  private readonly ctxBase: { theme: Theme };
  private readonly relsPath: string;
  private relsRoot: XmlElement | null = null;
  private contentTypesRoot: XmlElement | null = null;
  private nextDocPrId = 1000;

  private constructor(pkg: Package) {
    this.pkg = pkg;

    const docPart = this.findDocumentPart();
    this.docPart = docPart;
    const docDir = docPart.slice(0, docPart.lastIndexOf("/") + 1);

    const themeXml = this.readXmlOptional(docDir + "theme/theme1.xml");
    this.theme = parseTheme(themeXml);
    this.ctxBase = { theme: this.theme };

    this.styles = parseStyles(this.readXmlOptional(docDir + "styles.xml"), this.ctxBase);
    this.numbering = parseNumbering(this.readXmlOptional(docDir + "numbering.xml"), this.ctxBase);

    this.relsPath = relsPathFor(docPart);
    this.relsRoot = this.readXmlOptional(this.relsPath) ?? null;
    this.contentTypesRoot = this.readXmlOptional("[Content_Types].xml") ?? null;
    this.documentRels = parseRelationships(this.relsRoot ?? undefined, docPart);

    const docRoot = this.readXmlOptional(docPart);
    if (!docRoot) throw new Error(`Missing ${docPart} in package`);
    this.docRoot = docRoot;

    const settings = this.readXmlOptional(docDir + "settings.xml");
    if (settings) {
      this.evenAndOddHeaders = onOff(child(settings, "evenAndOddHeaders")) ?? false;
      const tabStop = intAttr(child(settings, "defaultTabStop"), "val");
      if (tabStop !== undefined && tabStop > 0) this.defaultTabStop = twipsToPx(tabStop);
    }

    // Review comments (optional part).
    const commentsRoot = this.readXmlOptional(docDir + "comments.xml");
    if (commentsRoot) {
      for (const c of commentsRoot.children) {
        if (localName(c.name) !== "comment") continue;
        const paras: string[] = [];
        const collectPara = (el: XmlElement): void => {
          if (localName(el.name) === "p") {
            let text = "";
            const collectT = (e: XmlElement): void => {
              if (localName(e.name) === "t") text += e.text;
              for (const ch of e.children) collectT(ch);
            };
            collectT(el);
            paras.push(text);
            return;
          }
          for (const ch of el.children) collectPara(ch);
        };
        for (const ch of c.children) collectPara(ch);
        this.comments.push({
          id: attr(c, "id") ?? "",
          author: attr(c, "author") ?? "",
          initials: attr(c, "initials"),
          date: attr(c, "date"),
          text: paras.join("\n"),
        });
      }
    }

    // Collect header/footer parts referenced from the document rels.
    for (const rel of this.documentRels.values()) {
      const isHeader = rel.type.endsWith("/header");
      const isFooter = rel.type.endsWith("/footer");
      if (!isHeader && !isFooter) continue;
      const root = this.readXmlOptional(rel.target);
      if (!root) continue;
      const partRels = parseRelationships(this.readXmlOptional(relsPathFor(rel.target)), rel.target);
      this.hfParts.push({ relId: rel.id, target: rel.target, root, isHeader, rels: partRels });
    }

    // Footnote/endnote parts (static: notes aren't editable in v1).
    for (const rel of this.documentRels.values()) {
      const isFn = rel.type.endsWith("/footnotes");
      const isEn = rel.type.endsWith("/endnotes");
      if (!isFn && !isEn) continue;
      const root = this.readXmlOptional(rel.target);
      if (!root) continue;
      const partRels = parseRelationships(this.readXmlOptional(relsPathFor(rel.target)), rel.target);
      const notes = parseNotesPart(root, { ...this.ctxBase, rels: partRels });
      for (const [id, blocks] of notes) (isFn ? this.footnotes : this.endnotes).set(id, blocks);
    }

    this.refresh();
  }

  /**
   * Re-derive the document model from the retained XML trees. Called after
   * edit commands mutate the XML.
   */
  refresh(): void {
    const body = child(this.docRoot, "body");
    if (!body) throw new Error("document.xml has no w:body");
    const ctx: DocParseContext = { ...this.ctxBase, rels: this.documentRels };
    this.sections = parseBody(body, ctx);
    this.headers.clear();
    this.footers.clear();
    for (const part of this.hfParts) {
      const partCtx: DocParseContext = { ...this.ctxBase, rels: part.rels };
      const hf: HeaderFooter = { blocks: parseBlocks(part.root, partCtx) };
      (part.isHeader ? this.headers : this.footers).set(part.relId, hf);
    }
  }

  /**
   * The w:t elements covered by each comment's range, in document order.
   * Point comments (a bare commentReference with no range) anchor to the
   * nearest preceding w:t.
   */
  commentAnchors(): Map<string, XmlElement[]> {
    const map = new Map<string, XmlElement[]>();
    const active = new Set<string>();
    let lastT: XmlElement | null = null;
    const walk = (el: XmlElement): void => {
      const ln = localName(el.name);
      if (ln === "commentRangeStart") {
        const id = attr(el, "id");
        if (id !== undefined) active.add(id);
        return;
      }
      if (ln === "commentRangeEnd") {
        const id = attr(el, "id");
        if (id !== undefined) active.delete(id);
        return;
      }
      if (ln === "commentReference") {
        const id = attr(el, "id");
        if (id !== undefined && !map.has(id) && lastT) map.set(id, [lastT]);
        return;
      }
      if (ln === "t") {
        lastT = el;
        for (const id of active) {
          const list = map.get(id);
          if (list) list.push(el);
          else map.set(id, [el]);
        }
        return;
      }
      for (const c of el.children) walk(c);
    };
    walk(this.docRoot);
    return map;
  }

  /** The mutable XML roots (document body, then header/footer parts). */
  editableRoots(): XmlElement[] {
    return [this.docRoot, ...this.hfParts.map((p) => p.root)];
  }

  /**
   * Find the parent element of `target` in any modeled XML tree (document
   * body, headers, footers). Linear scan — documents are small and this only
   * runs on structural edits (Enter, paragraph merge).
   */
  findParentOf(target: XmlElement): XmlElement | undefined {
    const roots = [this.docRoot, ...this.hfParts.map((p) => p.root)];
    const walk = (el: XmlElement): XmlElement | undefined => {
      for (const c of el.children) {
        if (c === target) return el;
        const found = walk(c);
        if (found) return found;
      }
      return undefined;
    };
    for (const root of roots) {
      const found = walk(root);
      if (found) return found;
    }
    return undefined;
  }

  /**
   * Serialize the (possibly edited) document back to .docx bytes. Only the
   * XML parts we model are re-serialized; every other part round-trips
   * byte-for-byte.
   */
  save(): Uint8Array {
    const files: Record<string, Uint8Array> = { ...this.pkg.raw() };
    files[this.docPart] = strToU8(serializeXml(this.docRoot, true));
    for (const part of this.hfParts) {
      files[part.target] = strToU8(serializeXml(part.root, true));
    }
    if (this.relsRoot) files[this.relsPath] = strToU8(serializeXml(this.relsRoot, true));
    if (this.contentTypesRoot) files["[Content_Types].xml"] = strToU8(serializeXml(this.contentTypesRoot, true));
    return zipSync(files);
  }

  /** Fresh unique docPr id for inserted drawings. */
  nextDrawingId(): number {
    return this.nextDocPrId++;
  }

  /**
   * Add image bytes as a new media part + relationship (+ content-type
   * default). Returns the relationship id for use in a w:drawing.
   */
  addImageResource(bytes: Uint8Array, ext: string): string {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    // Unique media name
    let n = 1;
    while (this.pkg.has(`${docDir}media/image${n}.${ext}`)) n++;
    const part = `${docDir}media/image${n}.${ext}`;
    this.pkg.raw()[part] = bytes;

    // Relationship
    if (!this.relsRoot) {
      this.relsRoot = {
        name: "Relationships",
        attrs: { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
        children: [],
        text: "",
      };
    }
    let maxId = 0;
    for (const r of this.relsRoot.children) {
      const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
      if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
    }
    const relId = `rId${maxId + 1}`;
    this.relsRoot.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target: `media/image${n}.${ext}`,
      },
      children: [],
      text: "",
    });
    this.documentRels.set(relId, { id: relId, type: "image", target: part, external: false });

    // Content type default for the extension
    const MIME: Record<string, string> = {
      png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp",
    };
    if (this.contentTypesRoot && MIME[ext]) {
      const has = this.contentTypesRoot.children.some(
        (c) => c.name.endsWith("Default") && (c.attrs["Extension"] ?? "").toLowerCase() === ext,
      );
      if (!has) {
        this.contentTypesRoot.children.unshift({
          name: "Default",
          attrs: { Extension: ext, ContentType: MIME[ext] },
          children: [],
          text: "",
        });
      }
    }
    return relId;
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    return new DocxDocument(Package.from(data));
  }

  media(part: string): Uint8Array | undefined {
    return this.pkg.binary(part);
  }

  /** Effective paragraph properties: docDefaults → style chain → direct. */
  effectiveParaProps(para: Paragraph): ParaProps {
    const { pPr } = resolveParagraphStyleChain(this.styles, para.props.styleId);
    let merged = mergeParaProps(pPr, para.props);
    // Numbering level can contribute indentation when the paragraph doesn't set its own.
    const num = merged.numbering;
    if (num) {
      const lvl = this.numberingLevel(num.numId, num.ilvl);
      if (lvl?.pPr) {
        const withLvl = mergeParaProps(pPr, lvl.pPr);
        merged = mergeParaProps(withLvl, para.props);
      }
    }
    return merged;
  }

  /** Effective run properties for a run inside a paragraph. */
  effectiveRunProps(para: Paragraph, runProps: RunProps): RunProps {
    const { rPr: paraStyleRPr } = resolveParagraphStyleChain(this.styles, para.props.styleId);
    let props = paraStyleRPr;
    if (runProps.styleId) {
      props = mergeRunProps(props, resolveCharacterStyleChain(this.styles, runProps.styleId));
    }
    props = mergeRunProps(props, runProps);
    return props;
  }

  numberingLevel(numId: number, ilvl: number) {
    const inst = this.numbering.instances.get(numId);
    if (!inst) return undefined;
    const override = inst.overrides.get(ilvl);
    if (override?.level) return override.level;
    const abs = this.numbering.abstract.get(inst.abstractNumId);
    return abs?.levels.get(ilvl);
  }

  numberingInstance(numId: number) {
    return this.numbering.instances.get(numId);
  }

  private findDocumentPart(): string {
    const rootRels = parseRelationships(this.readXmlOptional("_rels/.rels"), "");
    for (const rel of rootRels.values()) {
      if (rel.type === REL_TYPE_DOCUMENT) return rel.target;
    }
    if (this.pkg.has("word/document.xml")) return "word/document.xml";
    throw new Error("Not a WordprocessingML package: no main document part");
  }

  private readXmlOptional(part: string) {
    const text = this.pkg.text(part);
    if (text === undefined) return undefined;
    return parseXml(text);
  }
}
