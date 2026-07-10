import { Package } from "./zip.js";
import { XmlElement, parseXml, serializeXml, child, children, intAttr, onOff, attr, localName } from "./xml.js";
import { strToU8, zipSync } from "fflate";
import { twipsToPx } from "./units.js";
import {
  Block,
  DocComment,
  HeaderFooter,
  Numbering,
  ParaProps,
  Paragraph,
  Run,
  RunProps,
  Section,
  Styles,
  Theme,
} from "./model.js";
import { parseTheme } from "./parse/theme.js";
import {
  DEFAULT_TBL_LOOK,
  parseStyles,
  resolveCharacterStyleChain,
  resolveParagraphStyleChain,
  resolveTableConditional,
  resolveTableStyleProps,
  tableCondOrder,
} from "./parse/styles.js";
import { parseNumbering } from "./parse/numbering.js";
import { parseBody, parseBlocks, DocParseContext } from "./parse/document.js";
import { parseNotesPart } from "./parse/notes.js";
import { Relationships, parseRelationships, relsPathFor } from "./parse/rels.js";
import { mergeParaProps, mergeRunProps } from "./parse/properties.js";

const REL_TYPE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/** Word's built-in heading/title looks (modern Office theme), injected when a
 * file uses one without declaring it. Sizes in half-points. */
const BUILTIN_PARA_STYLES: Record<string, string> = (() => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const heading = (n: number, sizeHalfPt: number, color: string, extraRpr = ""): string =>
    `<w:style ${W} w:type="paragraph" w:styleId="Heading${n}">
      <w:name w:val="Heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
      <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${n === 1 ? 240 : 40}" w:after="0"/><w:outlineLvl w:val="${n - 1}"/></w:pPr>
      <w:rPr><w:color w:val="${color}"/><w:sz w:val="${sizeHalfPt}"/><w:szCs w:val="${sizeHalfPt}"/>${extraRpr}</w:rPr>
    </w:style>`;
  return {
    Heading1: heading(1, 32, "2F5496"),
    Heading2: heading(2, 26, "2F5496"),
    Heading3: heading(3, 24, "1F3863"),
    Heading4: heading(4, 22, "2F5496", "<w:i/>"),
    Heading5: heading(5, 22, "2F5496"),
    Heading6: heading(6, 22, "1F3863"),
    Title: `<w:style ${W} w:type="paragraph" w:styleId="Title">
      <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
      <w:pPr><w:spacing w:after="80"/></w:pPr>
      <w:rPr><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>
    </w:style>`,
  };
})();

/**
 * A fully parsed .docx: sections of blocks, styles, numbering, theme, and
 * header/footer parts, with helpers to resolve effective formatting.
 */
export class DocxDocument {
  readonly pkg: Package;
  readonly theme: Theme;
  styles: Styles;
  numbering: Numbering;
  sections: Section[] = [];
  /** Header/footer parts keyed by relationship id from document.xml.rels. */
  readonly headers: Map<string, HeaderFooter> = new Map();
  readonly footers: Map<string, HeaderFooter> = new Map();
  /** Note content by note id (render-only; sources stripped). */
  readonly footnotes: Map<number, Block[]> = new Map();
  readonly endnotes: Map<number, Block[]> = new Map();
  /** `_Ref` cross-reference bookmark ranges (name → captured runs). REF
   * fields re-render the referenced text from these — Word recomputes REF on
   * open, so the cached field result in the file is stale. */
  refBookmarks: Map<string, Run[]> = new Map();
  readonly documentRels: Relationships;
  /** settings.xml w:evenAndOddHeaders — enables the "even" header/footer variants. */
  readonly evenAndOddHeaders: boolean = false;
  /** settings.xml w:defaultTabStop in px (Word default 0.5"). */
  readonly defaultTabStop: number = 48;
  /** settings.xml w:compat compatibilityMode (12=Word2007, 14=Word2010,
   * 15=Word2013+). Word 2013 (mode 15) introduced suppressing a paragraph's
   * space-before when it lands at the top of a page; mode 14 and earlier keep
   * it (nccih: a Heading1/2 after a page break sits at margin + its before).
   * Absent → treated as current (15). */
  readonly compatibilityMode: number = 15;
  /** settings.xml m:mathPr/m:defJc — default justification for display
   * equations whose m:oMathParaPr carries no explicit m:jc (Word default:
   * centerGroup — the rows of a broken equation left-align to each other and
   * the group is centered in the column). */
  readonly mathDefJc: "left" | "right" | "center" | "centerGroup" = "centerGroup";
  /** settings.xml m:mathPr/m:wrapIndent in px (Word default 1440tw = 1"):
   * indent of auto-wrapped display-equation continuation rows from the
   * equation group's left edge (dense p13: the "+Dc(...)" continuations sit
   * exactly 72pt right of the explicit rows). */
  readonly mathWrapIndent: number = 96;
  /** Review comments from word/comments.xml (empty when the part is absent).
   * Re-derived from the retained comments XML on every refresh(). */
  comments: DocComment[] = [];
  /** Retained comments.xml tree (editing + save round-trip), when present. */
  private commentsPart: string | null = null;
  private commentsRoot: XmlElement | null = null;
  /** Retained commentsExtended.xml tree (comment threading), when present. */
  private commentsExtPart: string | null = null;
  private commentsExtRoot: XmlElement | null = null;
  private commentsExtDirty = false;
  /** Conditional table formats per table style id, keyed by the Styles object
   * so re-parsing styles.xml (edits) naturally invalidates the cache. */
  private tableCondCache = new WeakMap<Styles, Map<string, ReturnType<typeof resolveTableConditional>>>();
  /** Retained styles.xml tree (built-in style injection + save). */
  private stylesPart: string | null = null;
  private stylesRoot: XmlElement | null = null;
  /** Retained numbering.xml tree (list creation + save round-trip). */
  private numberingPart: string | null = null;
  private numberingRoot: XmlElement | null = null;
  private numberingDirty = false;
  /** Retained footnotes.xml tree (footnote insertion + save round-trip). */
  private footnotesPart: string | null = null;
  private footnotesRoot: XmlElement | null = null;
  private footnotesDirty = false;
  private footnotesRels: Relationships = new Map();
  /** Serialize retained optional parts only once actually mutated, keeping
   * untouched parts byte-identical through save(). */
  private stylesDirty = false;
  private commentsDirty = false;

  /** Retained XML roots — source of truth for editing and save(). */
  private readonly docPart: string;
  private readonly docRoot: XmlElement;
  private readonly hfParts: { relId: string; target: string; root: XmlElement; isHeader: boolean; rels: Relationships }[] = [];
  private readonly ctxBase: { theme: Theme; revisionView?: "final" | "markup" };
  /** Tracked-changes display mode; refresh() re-derives after changes. */
  revisionView: "final" | "markup" = "final";
  private readonly relsPath: string;
  private relsRoot: XmlElement | null = null;
  private contentTypesRoot: XmlElement | null = null;
  private nextDocPrId = 1000;

  private constructor(pkg: Package) {
    this.pkg = pkg;

    const docPart = this.findDocumentPart();
    this.docPart = docPart;
    const docDir = docPart.slice(0, docPart.lastIndexOf("/") + 1);

    const settings = this.readXmlOptional(docDir + "settings.xml");
    const bidiThemeLanguage = attr(child(settings, "themeFontLang"), "bidi");
    const themeXml = this.readXmlOptional(docDir + "theme/theme1.xml");
    this.theme = parseTheme(themeXml, bidiThemeLanguage);
    this.ctxBase = { theme: this.theme };
    this.ctxBase.revisionView = this.revisionView;

    this.stylesPart = docDir + "styles.xml";
    this.stylesRoot = this.readXmlOptional(this.stylesPart) ?? null;
    this.styles = parseStyles(this.stylesRoot ?? undefined, this.ctxBase);
    const numberingRoot = this.readXmlOptional(docDir + "numbering.xml");
    if (numberingRoot) {
      this.numberingPart = docDir + "numbering.xml";
      this.numberingRoot = numberingRoot;
    }
    this.numbering = parseNumbering(this.numberingRoot ?? undefined, this.ctxBase);

    this.relsPath = relsPathFor(docPart);
    this.relsRoot = this.readXmlOptional(this.relsPath) ?? null;
    this.contentTypesRoot = this.readXmlOptional("[Content_Types].xml") ?? null;
    this.documentRels = parseRelationships(this.relsRoot ?? undefined, docPart);

    const docRoot = this.readXmlOptional(docPart);
    if (!docRoot) throw new Error(`Missing ${docPart} in package`);
    this.docRoot = docRoot;

    if (settings) {
      this.evenAndOddHeaders = onOff(child(settings, "evenAndOddHeaders")) ?? false;
      const tabStop = intAttr(child(settings, "defaultTabStop"), "val");
      if (tabStop !== undefined && tabStop > 0) this.defaultTabStop = twipsToPx(tabStop);
      const compat = child(settings, "compat");
      for (const cs of children(compat, "compatSetting")) {
        if (attr(cs, "name") === "compatibilityMode") {
          const v = Number(attr(cs, "val"));
          if (Number.isFinite(v)) (this as { compatibilityMode: number }).compatibilityMode = v;
        }
      }
      const mathPr = child(settings, "mathPr");
      const defJc = attr(child(mathPr, "defJc"), "val");
      if (defJc === "left" || defJc === "right" || defJc === "center" || defJc === "centerGroup") {
        (this as { mathDefJc: string }).mathDefJc = defJc;
      }
      const wrapIndent = intAttr(child(mathPr, "wrapIndent"), "val");
      if (wrapIndent !== undefined && wrapIndent >= 0) {
        (this as { mathWrapIndent: number }).mathWrapIndent = twipsToPx(wrapIndent);
      }
    }

    // Review comments (optional part). The XML tree is retained so comments
    // can be deleted (with undo) and round-trip through save().
    const commentsRoot = this.readXmlOptional(docDir + "comments.xml");
    if (commentsRoot) {
      this.commentsPart = docDir + "comments.xml";
      this.commentsRoot = commentsRoot;
    }
    const commentsExtRoot = this.readXmlOptional(docDir + "commentsExtended.xml");
    if (commentsExtRoot) {
      this.commentsExtPart = docDir + "commentsExtended.xml";
      this.commentsExtRoot = commentsExtRoot;
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

    // Footnote/endnote parts. Footnotes retain their tree so insertion can
    // mutate and serialize it; note bodies stay non-editable (source refs
    // stripped by the parser).
    for (const rel of this.documentRels.values()) {
      const isFn = rel.type.endsWith("/footnotes");
      const isEn = rel.type.endsWith("/endnotes");
      if (!isFn && !isEn) continue;
      const root = this.readXmlOptional(rel.target);
      if (!root) continue;
      const partRels = parseRelationships(this.readXmlOptional(relsPathFor(rel.target)), rel.target);
      if (isFn) {
        this.footnotesPart = rel.target;
        this.footnotesRoot = root;
        this.footnotesRels = partRels;
      }
      const notes = parseNotesPart(root, { ...this.ctxBase, rels: partRels });
      for (const [id, blocks] of notes) (isFn ? this.footnotes : this.endnotes).set(id, blocks);
    }

    this.refresh();
  }

  /**
   * Re-derive the document model from the retained XML trees. Called after
   * edit commands mutate the XML.
   */
  /** Switch tracked-changes display and re-derive the model. */
  setRevisionView(view: "final" | "markup"): void {
    this.revisionView = view;
    this.ctxBase.revisionView = view;
    this.refresh();
  }

  refresh(): void {
    const body = child(this.docRoot, "body");
    if (!body) throw new Error("document.xml has no w:body");
    // Some content (SmartArt cached drawings) lives in parts reachable only
    // through relationship indirection at parse time.
    const readPart = (part: string) => this.readXmlOptional(part);
    const refBookmarks = { open: new Map<string, Run[]>(), byName: new Map<string, Run[]>() };
    const ctx: DocParseContext = { ...this.ctxBase, rels: this.documentRels, readPart, refBookmarks };
    this.sections = parseBody(body, ctx);
    this.refBookmarks = refBookmarks.byName;
    this.headers.clear();
    this.footers.clear();
    for (const part of this.hfParts) {
      const partCtx: DocParseContext = { ...this.ctxBase, rels: part.rels, readPart };
      const hf: HeaderFooter = { blocks: parseBlocks(part.root, partCtx) };
      (part.isHeader ? this.headers : this.footers).set(part.relId, hf);
    }
    this.comments = this.deriveComments();
    this.styles = parseStyles(this.stylesRoot ?? undefined, this.ctxBase);
    this.numbering = parseNumbering(this.numberingRoot ?? undefined, this.ctxBase);
  }

  private deriveComments(): DocComment[] {
    const out: DocComment[] = [];
    if (!this.commentsRoot) return out;
    for (const c of this.commentsRoot.children) {
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
      // Threading key: the w14:paraId of the comment's last body paragraph.
      let paraId: string | undefined;
      const lastPara = (el: XmlElement): void => {
        if (localName(el.name) === "p") {
          paraId = attr(el, "paraId") ?? paraId;
          return;
        }
        for (const ch of el.children) lastPara(ch);
      };
      for (const ch of c.children) lastPara(ch);
      out.push({
        id: attr(c, "id") ?? "",
        author: attr(c, "author") ?? "",
        initials: attr(c, "initials"),
        date: attr(c, "date"),
        text: paras.join("\n"),
        paraId,
      });
    }
    // commentsExtended threading: paraIdParent links a reply to its parent.
    if (this.commentsExtRoot) {
      const parentOf = new Map<string, string>();
      for (const ex of this.commentsExtRoot.children) {
        if (localName(ex.name) !== "commentEx") continue;
        const pid = attr(ex, "paraId");
        const parent = attr(ex, "paraIdParent");
        if (pid && parent) parentOf.set(pid, parent);
      }
      const byParaId = new Map(out.filter((c) => c.paraId).map((c) => [c.paraId!, c]));
      for (const c of out) {
        const parentPara = c.paraId ? parentOf.get(c.paraId) : undefined;
        if (parentPara) c.parentId = byParaId.get(parentPara)?.id;
      }
    }
    return out;
  }

  /** Retained comments tree for edit commands (null when the doc has none). */
  /**
   * Retained comments tree. With create=true, a missing comments.xml part is
   * created and registered (content type + document relationship) so newly
   * added comments serialize and round-trip through Word.
   */
  commentsTree(create = false): XmlElement | null {
    if (this.commentsRoot || !create) return this.commentsRoot;
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    this.commentsPart = docDir + "comments.xml";
    this.commentsRoot = {
      name: "w:comments",
      attrs: {
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml",
      },
      children: [],
      text: "",
    };
    {
      const rels = this.ensureRelsRoot();
      let maxId = 0;
      for (const r of rels.children) {
        const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
        if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
      }
      rels.children.push({
        name: "Relationship",
        attrs: {
          Id: `rId${maxId + 1}`,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
          Target: "comments.xml",
        },
        children: [],
        text: "",
      });
    }
    if (this.contentTypesRoot) {
      const partName = "/" + this.commentsPart;
      if (!this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
          },
          children: [],
          text: "",
        });
      }
    }
    this.commentsDirty = true;
    return this.commentsRoot;
  }

  /**
   * Make sure a paragraph style is usable: Word ships built-in definitions
   * for Heading 1-6/Title even when a file doesn't declare them, so applying
   * one to such a file must inject a standard definition (otherwise the
   * paragraph would reference an undefined style and render as Normal).
   */
  ensureParagraphStyle(styleId: string): boolean {
    if (this.styles.byId.has(styleId)) return true;
    const def = BUILTIN_PARA_STYLES[styleId];
    if (!def || !this.stylesRoot) return false;
    this.stylesRoot.children.push(parseXml(def));
    this.styles = parseStyles(this.stylesRoot, this.ctxBase);
    this.stylesDirty = true;
    return true;
  }

  /**
   * Retained numbering tree. With create=true, a missing numbering.xml part
   * is created and registered (content type + document relationship) so list
   * definitions added by editing serialize and round-trip.
   */
  numberingTree(create = false): XmlElement | null {
    if (this.numberingRoot || !create) return this.numberingRoot;
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    this.numberingPart = docDir + "numbering.xml";
    this.numberingRoot = {
      name: "w:numbering",
      attrs: { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
      children: [],
      text: "",
    };
    {
      const rels = this.ensureRelsRoot();
      let maxId = 0;
      for (const r of rels.children) {
        const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
        if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
      }
      rels.children.push({
        name: "Relationship",
        attrs: {
          Id: `rId${maxId + 1}`,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
          Target: "numbering.xml",
        },
        children: [],
        text: "",
      });
    }
    if (this.contentTypesRoot) {
      const partName = "/" + this.numberingPart;
      if (!this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
          },
          children: [],
          text: "",
        });
      }
    }
    this.numberingDirty = true;
    return this.numberingRoot;
  }

  markNumberingChanged(): void {
    this.numberingDirty = true;
  }

  /**
   * Retained footnotes tree. With create=true, a missing footnotes.xml part
   * is created and registered (with Word's required separator footnotes) so
   * inserted footnotes serialize and round-trip.
   */
  footnotesTree(create = false): XmlElement | null {
    if (this.footnotesRoot || !create) return this.footnotesRoot;
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    this.footnotesPart = docDir + "footnotes.xml";
    const sep = (id: string, type: string, refEl: string): XmlElement => ({
      name: "w:footnote",
      attrs: { "w:type": type, "w:id": id },
      children: [
        {
          name: "w:p",
          attrs: {},
          children: [
            { name: "w:pPr", attrs: {}, children: [{ name: "w:spacing", attrs: { "w:after": "0", "w:line": "240", "w:lineRule": "auto" }, children: [], text: "" }], text: "" },
            { name: "w:r", attrs: {}, children: [{ name: refEl, attrs: {}, children: [], text: "" }], text: "" },
          ],
          text: "",
        },
      ],
      text: "",
    });
    this.footnotesRoot = {
      name: "w:footnotes",
      attrs: { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
      children: [sep("-1", "separator", "w:separator"), sep("0", "continuationSeparator", "w:continuationSeparator")],
      text: "",
    };
    {
      const rels = this.ensureRelsRoot();
      let maxId = 0;
      for (const r of rels.children) {
        const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
        if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
      }
      rels.children.push({
        name: "Relationship",
        attrs: {
          Id: `rId${maxId + 1}`,
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
          Target: "footnotes.xml",
        },
        children: [],
        text: "",
      });
    }
    if (this.contentTypesRoot) {
      const partName = "/" + this.footnotesPart;
      if (!this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
          },
          children: [],
          text: "",
        });
      }
    }
    this.footnotesDirty = true;
    return this.footnotesRoot;
  }

  /**
   * Create an empty header/footer part (with a default-type reference in
   * every sectPr) when the document has none - Word does this implicitly the
   * first time you edit the header area. Returns the part's root.
   */
  ensureHfPart(kind: "header" | "footer"): XmlElement {
    const isHeader = kind === "header";
    const existing = this.hfParts.find((p2) => p2.isHeader === isHeader);
    if (existing) return existing.root;
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    while (this.pkg.has(`${docDir}${kind}${n}.xml`)) n++;
    const target = `${docDir}${kind}${n}.xml`;
    const rootName = isHeader ? "w:hdr" : "w:ftr";
    const root: XmlElement = {
      name: rootName,
      attrs: { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
      children: [
        {
          name: "w:p",
          attrs: {},
          children: [
            {
              name: "w:r",
              attrs: {},
              children: [{ name: "w:t", attrs: { "xml:space": "preserve" }, children: [], text: "" }],
              text: "",
            },
          ],
          text: "",
        },
      ],
      text: "",
    };
    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const r of rels.children) {
      const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
      if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
    }
    const relId = `rId${maxId + 1}`;
    rels.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`,
        Target: `${kind}${n}.xml`,
      },
      children: [],
      text: "",
    });
    if (this.contentTypesRoot) {
      const partName = "/" + target;
      if (!this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType: `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`,
          },
          children: [],
          text: "",
        });
      }
    }
    this.hfParts.push({ relId, target, root, isHeader, rels: new Map() });
    // Reference from every sectPr (schema: hf references lead the sectPr).
    const refName = isHeader ? "w:headerReference" : "w:footerReference";
    const addRef = (e: XmlElement): void => {
      if (localName(e.name) === "sectPr") {
        e.children.unshift({
          name: refName,
          attrs: {
            "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "w:type": "default",
            "r:id": relId,
          },
          children: [],
          text: "",
        });
        return;
      }
      for (const c of e.children) addRef(c);
    };
    addRef(this.docRoot);
    this.refresh();
    return root;
  }

  markFootnotesChanged(): void {
    this.footnotesDirty = true;
    // Re-derive the id -> blocks map so layout sees the new note.
    this.footnotes.clear();
    if (this.footnotesRoot) {
      const notes = parseNotesPart(this.footnotesRoot, { ...this.ctxBase, rels: this.footnotesRels });
      for (const [id, blocks] of notes) this.footnotes.set(id, blocks);
    }
  }

  /** Called by comment edit commands after mutating the comments tree. */
  markCommentsChanged(): void {
    this.commentsDirty = true;
  }

  /**
   * Retained commentsExtended tree (threading). With create=true, a missing
   * part is created and registered (content type + document relationship) so
   * Word picks up reply threading.
   */
  commentsExtendedTree(create = false): XmlElement | null {
    if (this.commentsExtRoot || !create) return this.commentsExtRoot;
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    this.commentsExtPart = docDir + "commentsExtended.xml";
    this.commentsExtRoot = {
      name: "w15:commentsEx",
      attrs: {
        "xmlns:w15": "http://schemas.microsoft.com/office/word/2012/wordml",
        "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      },
      children: [],
      text: "",
    };
    {
      const rels = this.ensureRelsRoot();
      let maxId = 0;
      for (const r of rels.children) {
        const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
        if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
      }
      rels.children.push({
        name: "Relationship",
        attrs: {
          Id: `rId${maxId + 1}`,
          Type: "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
          Target: "commentsExtended.xml",
        },
        children: [],
        text: "",
      });
    }
    if (this.contentTypesRoot) {
      const partName = "/" + this.commentsExtPart;
      const has = this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName);
      if (!has) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
          },
          children: [],
          text: "",
        });
      }
    }
    this.commentsExtDirty = true;
    return this.commentsExtRoot;
  }

  markCommentsExtendedChanged(): void {
    this.commentsExtDirty = true;
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

  /** The mutable XML roots (document body, header/footer parts, comments).
   * The comments root is last so history snapshot indices stay stable. */
  editableRoots(): XmlElement[] {
    const roots = [this.docRoot, ...this.hfParts.map((p) => p.root)];
    if (this.commentsRoot) roots.push(this.commentsRoot);
    if (this.commentsExtRoot) roots.push(this.commentsExtRoot);
    return roots;
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
    if (this.commentsDirty && this.commentsRoot && this.commentsPart) {
      files[this.commentsPart] = strToU8(serializeXml(this.commentsRoot, true));
    }
    if (this.commentsExtDirty && this.commentsExtRoot && this.commentsExtPart) {
      files[this.commentsExtPart] = strToU8(serializeXml(this.commentsExtRoot, true));
    }
    if (this.stylesDirty && this.stylesRoot && this.stylesPart) {
      files[this.stylesPart] = strToU8(serializeXml(this.stylesRoot, true));
    }
    if (this.numberingDirty && this.numberingRoot && this.numberingPart) {
      files[this.numberingPart] = strToU8(serializeXml(this.numberingRoot, true));
    }
    if (this.footnotesDirty && this.footnotesRoot && this.footnotesPart) {
      files[this.footnotesPart] = strToU8(serializeXml(this.footnotesRoot, true));
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
  private ensureRelsRoot(): XmlElement {
    if (!this.relsRoot) {
      this.relsRoot = {
        name: "Relationships",
        attrs: { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
        children: [],
        text: "",
      };
    }
    return this.relsRoot;
  }

  /** Register an external hyperlink relationship and return its rId. */
  addHyperlinkRel(url: string): string {
    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const r of rels.children) {
      const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
      if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
    }
    const id = `rId${maxId + 1}`;
    rels.children.push({
      name: "Relationship",
      attrs: {
        Id: id,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        Target: url,
        TargetMode: "External",
      },
      children: [],
      text: "",
    });
    // documentRels is what refresh() resolves r:id through - keep it live.
    this.documentRels.set(id, { id, type: "hyperlink", target: url, external: true });
    return id;
  }

  /** Retarget an existing external relationship (hyperlink href edit). */
  setRelTarget(relId: string, url: string): boolean {
    const rel = this.documentRels.get(relId);
    if (!rel || !rel.external) return false;
    rel.target = url;
    const el = this.relsRoot?.children.find((r) => r.attrs["Id"] === relId);
    if (el) el.attrs["Target"] = url;
    return true;
  }

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

  /** Effective paragraph properties: docDefaults → table style → style chain → direct. */
  effectiveParaProps(para: Paragraph): ParaProps {
    let pPr: ParaProps;
    const tableStyleId = para.props.tableStyleId;
    if (tableStyleId) {
      // Precedence: docDefaults < table style < paragraph style < direct.
      // The table style's pPr sits just above docDefaults, so a paragraph
      // style that leaves spacing unset (e.g. ListParagraph) inherits the
      // table style's compact spacing rather than docDefaults'. The table
      // style resolves through its own basedOn chain (TableGrid basedOn
      // TableNormal).
      const tbl = resolveTableStyleProps(this.styles, tableStyleId);
      let base: ParaProps = { ...this.styles.defaultPPr };
      if (tbl.pPr) base = mergeParaProps(base, tbl.pPr);
      const contrib = resolveParagraphStyleChain(this.styles, para.props.styleId, false);
      pPr = mergeParaProps(base, contrib.pPr);
    } else {
      pPr = resolveParagraphStyleChain(this.styles, para.props.styleId).pPr;
    }
    let merged = mergeParaProps(pPr, para.props);
    // Numbering level can contribute indentation when the paragraph doesn't set its own.
    const num = merged.numbering;
    if (num) {
      const lvl = this.numberingLevel(num.numId, num.ilvl);
      if (lvl?.pPr) {
        if (para.props.numbering) {
          // Direct numPr: the level's pPr acts as direct-level formatting -
          // it beats the style chain's ind (classic ListParagraph left=720
          // replaced by the level's ind) but stays below the paragraph's own
          // direct pPr.
          const withLvl = mergeParaProps(pPr, lvl.pPr);
          merged = mergeParaProps(withLvl, para.props);
        } else {
          // Style-sourced numbering (pStyle -> numPr): the level's pPr slots
          // in BELOW the style chain, so a style's own w:ind beats the
          // level's, attribute by attribute. phase23's Heading3 carries
          // ind left=720 while its abs lvl says left=4410 hanging=720: Word
          // paints the number at the margin with text at 720 (style left
          // wins, level hanging survives because the style sets none).
          const contrib = resolveParagraphStyleChain(this.styles, para.props.styleId, false);
          const withLvl = mergeParaProps(pPr, mergeParaProps(lvl.pPr, contrib.pPr));
          merged = mergeParaProps(withLvl, para.props);
        }
      }
    }
    return merged;
  }

  /**
   * Run props contributed by the enclosing table style's conditional
   * w:tblStylePr blocks for this paragraph's cell (undefined when the
   * paragraph isn't in a styled table cell or nothing applies).
   */
  private tableCondRPr(para: Paragraph): RunProps | undefined {
    const cond = para.props.tableCellCond;
    const styleId = para.props.tableStyleId;
    if (!cond || !styleId) return undefined;
    let cache = this.tableCondCache.get(this.styles);
    if (!cache) {
      cache = new Map();
      this.tableCondCache.set(this.styles, cache);
    }
    let resolved = cache.get(styleId);
    if (!resolved) {
      resolved = resolveTableConditional(this.styles, styleId);
      cache.set(styleId, resolved);
    }
    if (resolved.formats.size === 0) return undefined;
    const order = tableCondOrder(
      cond.look ?? DEFAULT_TBL_LOOK,
      cond.rowIdx,
      cond.nRows,
      cond.colStart,
      cond.colSpan,
      cond.nCols,
      resolved.rowBandSize,
      resolved.colBandSize,
    );
    let out: RunProps | undefined;
    for (const type of order) {
      const rPr = resolved.formats.get(type)?.rPr;
      if (rPr) out = out ? mergeRunProps(out, rPr) : { ...rPr };
    }
    return out;
  }

  /** Effective run properties for a run inside a paragraph. */
  effectiveRunProps(para: Paragraph, runProps: RunProps): RunProps {
    let props: RunProps;
    const tableStyleId = para.props.tableStyleId;
    if (tableStyleId) {
      // Same layering as effectiveParaProps: the table style's rPr sits
      // between docDefaults and the paragraph style chain.
      const tbl = resolveTableStyleProps(this.styles, tableStyleId);
      let base: RunProps = { ...this.styles.defaultRPr };
      if (tbl.rPr) base = mergeRunProps(base, tbl.rPr);
      // Conditional w:tblStylePr run formats (firstRow bold/white, firstCol
      // bold, banding, …) layer above the table style's own rPr but below the
      // paragraph style chain and direct formatting.
      const condRPr = this.tableCondRPr(para);
      if (condRPr) base = mergeRunProps(base, condRPr);
      const contrib = resolveParagraphStyleChain(this.styles, para.props.styleId, false);
      props = mergeRunProps(base, contrib.rPr);
    } else {
      props = resolveParagraphStyleChain(this.styles, para.props.styleId).rPr;
    }
    const tocHyperlink = /^TOC[1-9]$/i.test(para.props.styleId ?? "")
      ? para.children.find(
          (child) =>
            child.type === "hyperlink" &&
            child.runs.some((run) => run.props === runProps) &&
            child.runs.some((run) =>
              run.content.some(
                (content) => content.kind === "field" && /^\s*PAGEREF\b/i.test(content.instruction),
              ),
            ),
        )
      : undefined;
    let generatedTocStyleColor: string | undefined;
    if (tocHyperlink?.type === "hyperlink") {
      // A styled run in a generated TOC hyperlink keeps its own character
      // style's font family while Word suppresses the style's other formatting.
      // Unstyled leader and PAGEREF runs keep the TOC paragraph's font; a style
      // on a sibling title run does not leak into them. A plain hyperlink in a
      // TOC-styled paragraph has no PAGEREF field and still uses the full style.
      if (runProps.styleId) {
        const linkProps = resolveCharacterStyleChain(this.styles, runProps.styleId);
        const tocLinkStyle = this.styles.byId.get(runProps.styleId);
        const keepTocLinkColor = /^Hyperlink-toc$/i.test(runProps.styleId) ||
          /^Hyperlink-toc$/i.test(tocLinkStyle?.name ?? "");
        const keptProps: RunProps = {
          font: linkProps.font,
          fontEastAsia: linkProps.fontEastAsia,
          fontComplex: linkProps.fontComplex,
        };
        if (keepTocLinkColor) {
          keptProps.color = linkProps.color;
          generatedTocStyleColor = linkProps.color;
        }
        props = mergeRunProps(props, keptProps);
      }
    } else if (runProps.styleId) {
      props = mergeRunProps(props, resolveCharacterStyleChain(this.styles, runProps.styleId));
    }
    // Generated TOC caches can put a direct size on the tab between the
    // heading number and text. Word renders that separator at the paragraph
    // mark size when one is present, otherwise at the TOC paragraph size. Keep
    // direct sizes on text/field runs so editing a TOC entry remains effective.
    const tocRun =
      tocHyperlink?.type === "hyperlink"
        ? tocHyperlink.runs.find((run) => run.props === runProps)
        : undefined;
    const cachedTocTab =
      tocRun !== undefined &&
      tocRun.content.length > 0 &&
      tocRun.content.every((content) => content.kind === "tab");
    let directProps =
      cachedTocTab && runProps.size !== undefined
        ? { ...runProps, size: para.props.markRunProps?.size }
        : runProps;
    // TOC parsing writes synthetic color=auto to suppress the standard
    // Hyperlink style. A custom Hyperlink-toc color is the exception Word
    // retains, so replace only that synthetic value; a real direct color wins.
    if (generatedTocStyleColor !== undefined && directProps.color === "auto") {
      directProps = { ...directProps, color: generatedTocStyleColor };
    }
    props = mergeRunProps(props, directProps);
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
