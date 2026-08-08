import { Package, FIXED_ZIP_MTIME } from "./zip.js";
import { XmlElement, parseXml, serializeXml, child, children, intAttr, onOff, attr, localName, cyrb53 } from "./xml.js";
import { strToU8, zip, zipSync } from "fflate";
import { pxToTwips, twipsToPx } from "./units.js";
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
  SmartArtData,
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
import { StableIds } from "./edit/ids.js";
import { parseBody, parseBlocks, parseParagraph, DocParseContext } from "./parse/document.js";
import { parseNotesPart } from "./parse/notes.js";
import { Relationships, parseRelationships, relsPathFor } from "./parse/rels.js";
import { mergeParaProps, mergeRunProps } from "./parse/properties.js";
import { extractOlePackage } from "./parse/ole.js";
import { isSourcesRoot } from "./citations.js";
import {
  buildSmartArtColorsXml,
  buildSmartArtDataXml,
  buildSmartArtDrawingXml,
  buildSmartArtLayoutXml,
  buildSmartArtStyleXml,
} from "./edit/smartart.js";

const REL_TYPE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/**
 * A fresh footnotes.xml / endnotes.xml root carrying the two stock notes Word
 * requires of the part: the separator rule drawn above the notes, and the
 * continuation separator drawn when a note spills onto the next page. Real
 * notes start at id 1 because these occupy -1 and 0.
 */
function notesPartRoot(kind: "footnote" | "endnote"): XmlElement {
  const stock = (id: string, type: string, refEl: string): XmlElement => ({
    name: `w:${kind}`,
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
  return {
    name: `w:${kind}s`,
    attrs: { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
    children: [stock("-1", "separator", "w:separator"), stock("0", "continuationSeparator", "w:continuationSeparator")],
    text: "",
  };
}

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
  // Word's "toc 1".."toc 9": Normal plus a per-level indent of 220 twips. A
  // generated table of contents references these by pStyle, so a document that
  // has never held one still needs them the moment a TOC is inserted.
  const toc = (n: number): string =>
    `<w:style ${W} w:type="paragraph" w:styleId="TOC${n}">
      <w:name w:val="toc ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
      <w:autoRedefine/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>
      <w:pPr><w:spacing w:after="100"/>${n > 1 ? `<w:ind w:left="${(n - 1) * 220}"/>` : ""}</w:pPr>
    </w:style>`;
  return {
    ...Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`TOC${i + 1}`, toc(i + 1)])),
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
    // Word's built-in "Bibliography" look: Normal with the utility flags Word
    // writes. A generated bibliography references it by pStyle, so a document
    // that has never held one still needs it the moment one is inserted.
    Bibliography: `<w:style ${W} w:type="paragraph" w:styleId="Bibliography">
      <w:name w:val="Bibliography"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
      <w:uiPriority w:val="37"/><w:semiHidden/><w:unhideWhenUsed/>
    </w:style>`,
    // Word's caption style: italic 9pt in the theme's dark accent, keep-next
    // so a caption never separates from its figure.
    Caption: `<w:style ${W} w:type="paragraph" w:styleId="Caption">
      <w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
      <w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/>
      <w:pPr><w:keepNext/><w:spacing w:after="200"/></w:pPr>
      <w:rPr><w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
    </w:style>`,
    // Word's "table of figures": Normal plus the TOC entries' after-spacing.
    TableofFigures: `<w:style ${W} w:type="paragraph" w:styleId="TableofFigures">
      <w:name w:val="table of figures"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
      <w:uiPriority w:val="99"/><w:unhideWhenUsed/>
      <w:pPr><w:spacing w:after="100"/></w:pPr>
    </w:style>`,
  };
})();

/** The two note parts a document can carry — see notePartsHolding. */
export type NotePart = "footnotes" | "endnotes";

/**
 * A fully parsed .docx: sections of blocks, styles, numbering, theme, and
 * header/footer parts, with helpers to resolve effective formatting.
 */
export class DocxDocument {
  /** Changes whenever refresh() rebuilds the parsed model. Plain in-place text
   * edits can keep this stable so incremental layout reuses model-only caches. */
  private _modelVersion = 0;
  private _packageResourceVersion = 0;
  get modelVersion(): number {
    return this._modelVersion;
  }

  /** Invalidate layout derived from related parts such as ChartML. */
  markPackageResourceChanged(): void {
    this._packageResourceVersion++;
    this._layoutGlobalSig = null;
  }
  readonly pkg: Package;
  readonly theme: Theme;
  styles: Styles;
  numbering: Numbering;
  sections: Section[] = [];
  /** Header/footer parts keyed by relationship id from document.xml.rels. */
  readonly headers: Map<string, HeaderFooter> = new Map();
  readonly footers: Map<string, HeaderFooter> = new Map();
  /** Note content by note id. */
  readonly footnotes: Map<number, Block[]> = new Map();
  readonly endnotes: Map<number, Block[]> = new Map();
  /** The separator paragraph controls the gap between its rule and the first
   * footnote. */
  readonly footnoteSeparator: Block[] = [];
  /** `_Ref` cross-reference bookmark ranges (name → captured runs). REF
   * fields re-render the referenced text from these — Word recomputes REF on
   * open, so the cached field result in the file is stale. */
  refBookmarks: Map<string, Run[]> = new Map();
  readonly documentRels: Relationships;
  /** settings.xml w:evenAndOddHeaders — enables the "even" header/footer variants. */
  readonly evenAndOddHeaders: boolean = false;
  /** settings.xml w:mirrorMargins — facing-page (book fold) margins: even
   * (verso) pages swap the left/right margins and place the gutter on the
   * inside (right) edge so the binding margin stays on the inner side of
   * each spread. */
  readonly mirrorMargins: boolean = false;
  /** settings.xml w:compat w:suppressTopSpacing: the first line of a page
   * takes its character height, not its authored (exact) line spacing. */
  readonly suppressTopSpacing: boolean = false;
  /** settings.xml w:defaultTabStop in px (Word default 0.5"). */
  readonly defaultTabStop: number = 48;
  /** settings.xml w:compat compatibilityMode (12=Word2007, 14=Word2010,
   * 15=Word2013+). Word 2013 (mode 15) introduced suppressing a paragraph's
   * space-before when it lands at the top of a page; mode 14 and earlier keep
   * it (nccih: a Heading1/2 after a page break sits at margin + its before).
   * Absent → treated as current (15). */
  readonly compatibilityMode: number = 15;
  /** The compatibilityMode settings.xml actually declares, or undefined when
   * the setting (or settings.xml itself) is missing. `compatibilityMode`
   * defaults an omitted value to 15, but a table's percentage width needs the
   * distinction: Word fits the horizontal cell margins inside the table box at
   * an EXPLICIT 15 and adds them around it when the setting is absent. */
  readonly declaredCompatibilityMode: number | undefined = undefined;
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
  /** Retained endnotes.xml tree (endnote insertion + save round-trip). */
  private endnotesPart: string | null = null;
  private endnotesRoot: XmlElement | null = null;
  private endnotesDirty = false;
  private endnotesRels: Relationships = new Map();
  /** Retained bibliography sources part (the customXml b:Sources data of
   * ECMA-376 §22.6), when present. Source-management editing mutates this
   * tree in place; a document with none gets the part on first create. */
  private sourcesPart: string | null = null;
  private sourcesRoot: XmlElement | null = null;
  private sourcesDirty = false;
  /** Companion parts written when THIS session created the sources part
   * (itemProps + its .rels); an arriving part keeps its own companions. */
  private sourcesAux: Record<string, string> | null = null;
  /** Serialize retained optional parts only once actually mutated, keeping
   * untouched parts byte-identical through save(). */
  private stylesDirty = false;
  private commentsDirty = false;

  /** Retained XML roots — source of truth for editing and save(). */
  private readonly docPart: string;
  /** Retained settings.xml tree. A synthetic empty root keeps history root
   * indices stable for documents that did not originally contain the part. */
  private readonly settingsPart: string;
  private readonly settingsRoot: XmlElement;
  private settingsDirty = false;
  /** Parsed document.xml root (read-only outside the class; the layout engine
   * scans it for incremental-reuse eligibility, tests walk it). */
  readonly docRoot: XmlElement;
  private readonly hfParts: {
    relId: string;
    target: string;
    root: XmlElement;
    isHeader: boolean;
    rels: Relationships;
    relsRoot: XmlElement | null;
  }[] = [];
  private readonly ctxBase: { theme: Theme; revisionView?: "final" | "markup" };
  /** Tracked-changes display mode; refresh() re-derives after changes. */
  revisionView: "final" | "markup" = "final";
  private readonly relsPath: string;
  private relsRoot: XmlElement | null = null;
  private contentTypesRoot: XmlElement | null = null;
  /** Canonical XML as first parsed from each always-modeled package part.
   * If the retained tree still matches on save, keep the part's original
   * bytes instead of replacing producer formatting such as CRLF line ends. */
  private readonly originalModeledXml = new Map<string, string>();
  private docPrIdCounter: number | null = null;

  /** Transient layout state: set by the engine while laying out a docGrid
   * type="charsAndLines" section so line measurement can give East-Asian
   * glyphs their true (uninflated) grid line height. Off outside such a
   * section. Safe as document-scoped mutable state because layout is
   * single-threaded and sequential per section. */
  charGridEa = false;

  private constructor(pkg: Package) {
    this.pkg = pkg;

    const docPart = this.findDocumentPart();
    this.docPart = docPart;
    const docDir = docPart.slice(0, docPart.lastIndexOf("/") + 1);

    this.settingsPart = docDir + "settings.xml";
    this.settingsRoot = this.readXmlOptional(this.settingsPart) ?? {
      name: "w:settings",
      attrs: { "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main" },
      children: [],
      text: "",
    };
    const settings = this.settingsRoot;
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
    if (this.relsRoot) this.rememberOriginalXml(this.relsPath, this.relsRoot);
    if (this.contentTypesRoot) this.rememberOriginalXml("[Content_Types].xml", this.contentTypesRoot);

    const docRoot = this.readXmlOptional(docPart);
    if (!docRoot) throw new Error(`Missing ${docPart} in package`);
    this.docRoot = docRoot;
    this.rememberOriginalXml(docPart, docRoot);
    const coreProperties = this.readXmlOptional("docProps/core.xml");
    this.hydrateCorePropertyControls(docRoot, coreProperties);
    this.repairLegacyWordInWebObjects();
    this.documentRels = parseRelationships(this.relsRoot ?? undefined, docPart);

    // Bibliography sources: probe every customXml relationship for a
    // b:Sources root (how Word's own packages are found without hard-coding
    // the item1.xml file name) and retain the first, so source-management
    // editing can mutate it and save() can serialize the mutated tree.
    for (const rel of this.documentRels.values()) {
      if (!rel.type.endsWith("/customXml") || rel.external) continue;
      const text = this.pkg.text(rel.target);
      if (!text) continue;
      let root: XmlElement;
      try {
        root = parseXml(text);
      } catch {
        continue;
      }
      if (!isSourcesRoot(root)) continue;
      this.sourcesPart = rel.target;
      this.sourcesRoot = root;
      break;
    }

    if (settings) {
      this.evenAndOddHeaders = onOff(child(settings, "evenAndOddHeaders")) ?? false;
      (this as { mirrorMargins: boolean }).mirrorMargins = onOff(child(settings, "mirrorMargins")) ?? false;
      const tabStop = intAttr(child(settings, "defaultTabStop"), "val");
      if (tabStop !== undefined && tabStop > 0) this.defaultTabStop = twipsToPx(tabStop);
      const compat = child(settings, "compat");
      // w:suppressTopSpacing (legacy compat option): line spacing beyond the
      // character height is suppressed for the first line of a page. See the
      // engine's page-top exact-line collapse for the measured behavior.
      (this as { suppressTopSpacing: boolean }).suppressTopSpacing =
        onOff(child(compat, "suppressTopSpacing")) ?? false;
      for (const cs of children(compat, "compatSetting")) {
        if (attr(cs, "name") === "compatibilityMode") {
          const v = Number(attr(cs, "val"));
          if (Number.isFinite(v)) {
            (this as { compatibilityMode: number }).compatibilityMode = v;
            (this as { declaredCompatibilityMode: number | undefined }).declaredCompatibilityMode = v;
          }
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

    // PENDING MEDIA ROUND-TRIP (plan doc 16 §6): an image relationship whose
    // target is absent from the package is a HOLE, not corruption — that is
    // exactly the shape a document carrying out-of-band media saves and
    // reloads as. Deriving it here is what lets a reloaded (or newly joined)
    // replica render a skeleton in the reserved box instead of silently
    // nothing. The declared sha is NOT recoverable from the package — it
    // lives in the sequenced intent — so the address is left empty and the
    // transfer layer treats such a part as unfetchable-but-known.
    for (const rel of this.documentRels.values()) {
      if (rel.type !== "image" && !rel.type.endsWith("/image")) continue;
      if (rel.external || this.pkg.binary(rel.target)) continue;
      this.pendingMedia.set(rel.target, { sha: "" });
      this.mediaMeta.set(rel.target, { sha: "" });
    }

    // Collect header/footer parts referenced from the document rels.
    for (const rel of this.documentRels.values()) {
      const isHeader = rel.type.endsWith("/header");
      const isFooter = rel.type.endsWith("/footer");
      if (!isHeader && !isFooter) continue;
      const root = this.readXmlOptional(rel.target);
      if (!root) continue;
      const partRelsRoot = this.readXmlOptional(relsPathFor(rel.target));
      const partRels = parseRelationships(partRelsRoot, rel.target);
      for (const imageRel of partRels.values()) {
        if (imageRel.type !== "image" && !imageRel.type.endsWith("/image")) continue;
        if (imageRel.external || this.pkg.binary(imageRel.target)) continue;
        this.pendingMedia.set(imageRel.target, { sha: "" });
        this.mediaMeta.set(imageRel.target, { sha: "" });
      }
      this.rememberOriginalXml(rel.target, root);
      this.hydrateCorePropertyControls(root, coreProperties);
      this.hfParts.push({ relId: rel.id, target: rel.target, root, isHeader, rels: partRels, relsRoot: partRelsRoot ?? null });
    }

    // Footnote/endnote parts. Both retain their tree so insertion can mutate
    // and serialize it, and both parse editable (source refs kept) so the
    // caret can bind to note text.
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
        const separator = root.children.find(
          (item) => localName(item.name) === "footnote" && attr(item, "type") === "separator",
        );
        if (separator) {
          this.footnoteSeparator.push(...parseBlocks(separator, { ...this.ctxBase, rels: partRels }));
        }
      } else {
        this.endnotesPart = rel.target;
        this.endnotesRoot = root;
        this.endnotesRels = partRels;
      }
      const notes = parseNotesPart(root, { ...this.ctxBase, rels: partRels });
      for (const [id, blocks] of notes) (isFn ? this.footnotes : this.endnotes).set(id, blocks);
    }

    this.refresh();
  }

  /** Resolve content controls mapped to standard package core properties.
   * Word refreshes these bindings on open, so the serialized sdtContent can
   * be stale even though the visible value comes from docProps/core.xml. */
  private hydrateCorePropertyControls(root: XmlElement, coreProperties: XmlElement | undefined): void {
    if (!coreProperties) return;
    const textNodes = (element: XmlElement): XmlElement[] => {
      const out = localName(element.name) === "t" ? [element] : [];
      for (const item of element.children) out.push(...textNodes(item));
      return out;
    };
    const walk = (element: XmlElement): void => {
      if (localName(element.name) === "sdt") {
        const binding = child(child(element, "sdtPr"), "dataBinding");
        const xpath = attr(binding, "xpath") ?? "";
        const propertyMatch = /\/(?:[^/:]+:)?([A-Za-z_][\w.-]*)(?:\[\d+\])?\s*$/.exec(xpath);
        if (xpath.includes("coreProperties") && propertyMatch) {
          const property = coreProperties.children.find(
            (item) => localName(item.name) === propertyMatch[1],
          );
          const content = child(element, "sdtContent");
          const targets = content ? textNodes(content) : [];
          // An empty bound property leaves the serialized placeholder visible
          // when w:showingPlcHdr is set; only a real value replaces it.
          if (property?.text && targets.length > 0) {
            targets[0].text = property.text;
            for (const target of targets.slice(1)) target.text = "";
          }
        }
      }
      for (const item of element.children) walk(item);
    };
    walk(root);
  }

  /** Repair only objects emitted by older WordInWeb builds that Word rejects. */
  private repairLegacyWordInWebObjects(): void {
    const raw = this.pkg.raw();
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    const all = (root: XmlElement, name: string): XmlElement[] => {
      const found = localName(root.name) === name ? [root] : [];
      for (const item of root.children) found.push(...all(item, name));
      return found;
    };
    const first = (root: XmlElement, name: string): XmlElement | undefined => all(root, name)[0];
    const setAttr = (element: XmlElement, name: string, value: string): void => {
      const key = Object.keys(element.attrs).find((item) => localName(item) === name) ?? name;
      element.attrs[key] = value;
    };
    const element = (name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement => ({
      name, attrs, children, text,
    });

    // Old WordInWeb SmartArt used non-schema connector ids (c1/c2/...) and
    // negative/zero cached extents. The urn signature is ours, so rebuilding
    // only these parts cannot touch third-party SmartArt.
    for (const part of Object.keys(raw)) {
      const match = new RegExp(`^${docDir.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}diagrams/data(\\d+)\\.xml$`).exec(part);
      if (!match) continue;
      const dataRoot = this.readXmlOptional(part);
      if (!dataRoot) continue;
      const typeId = attr(first(dataRoot, "prSet"), "loTypeId") ?? "";
      const layout = typeId.startsWith("urn:wordinweb:smartart:") ? typeId.slice("urn:wordinweb:smartart:".length) : "";
      if (layout !== "list" && layout !== "process" && layout !== "hierarchy" && layout !== "cycle") continue;
      const ptList = first(dataRoot, "ptLst");
      const items = ptList
        ? ptList.children
          .filter((item) => localName(item.name) === "pt" && attr(item, "type") !== "doc")
          .map((item) => all(item, "t").find((text) => text.text)?.text ?? "")
          .filter(Boolean)
        : [];
      if (!items.length) continue;
      const drawingRelId = attr(first(dataRoot, "dataModelExt"), "relId");
      if (!drawingRelId) continue;
      const n = match[1];
      const drawingPart = `${docDir}diagrams/drawing${n}.xml`;
      const drawingRoot = this.readXmlOptional(drawingPart);
      const modelIds = [
        ...all(dataRoot, "pt").map((item) => attr(item, "modelId") ?? ""),
        ...all(dataRoot, "cxn").map((item) => attr(item, "modelId") ?? ""),
      ];
      const legacyConnectionId = all(dataRoot, "cxn").some((item) => /^c\d+$/.test(attr(item, "modelId") ?? ""));
      const invalidExtent = !!drawingRoot && all(drawingRoot, "ext").some(
        (item) => Number(attr(item, "cx")) <= 0 || Number(attr(item, "cy")) <= 0,
      );
      const staleShapeId = !!drawingRoot && all(drawingRoot, "sp").some(
        (item) => !modelIds.includes(attr(item, "modelId") ?? ""),
      );
      if (!legacyConnectionId && !invalidExtent && !staleShapeId) continue;
      const data: SmartArtData = { layout, items };
      raw[part] = strToU8(buildSmartArtDataXml(data, drawingRelId));
      raw[`${docDir}diagrams/layout${n}.xml`] = strToU8(buildSmartArtLayoutXml(data));
      raw[`${docDir}diagrams/quickStyle${n}.xml`] = strToU8(buildSmartArtStyleXml());
      raw[`${docDir}diagrams/colors${n}.xml`] = strToU8(buildSmartArtColorsXml());
      raw[drawingPart] = strToU8(buildSmartArtDrawingXml(data));
    }

    if (!this.relsRoot || !this.contentTypesRoot) return;
    let migrated = 0;
    for (const object of all(this.docRoot, "object")) {
      const ole = first(object, "OLEObject");
      const imageData = first(object, "imagedata");
      const filename = attr(imageData, "title") ?? "";
      if (!ole || attr(ole, "ProgID") !== "Package" || !filename.toLowerCase().endsWith(".docx")) continue;
      const relId = attr(ole, "id");
      const rel = this.relsRoot.children.find((item) => item.attrs.Id === relId);
      if (!rel || !rel.attrs.Type?.endsWith("/oleObject") || !rel.attrs.Target?.startsWith("embeddings/")) continue;
      const oldPart = `${docDir}${rel.attrs.Target}`;
      const packaged = raw[oldPart] ? extractOlePackage(raw[oldPart]) : null;
      if (!packaged || !packaged.filename.toLowerCase().endsWith(".docx") ||
        packaged.data[0] !== 0x50 || packaged.data[1] !== 0x4b) continue;

      let index = 1;
      const packageName = () => `Microsoft_Word_Document${index === 1 ? "" : index}.docx`;
      while (raw[`${docDir}embeddings/${packageName()}`]) index++;
      const target = `embeddings/${packageName()}`;
      const newPart = `${docDir}${target}`;
      raw[newPart] = packaged.data;
      delete raw[oldPart];
      rel.attrs.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
      rel.attrs.Target = target;
      this.contentTypesRoot.children = this.contentTypesRoot.children.filter(
        (item) => item.attrs.PartName !== `/${oldPart}`,
      );
      if (!this.contentTypesRoot.children.some(
        (item) => localName(item.name) === "Default" && item.attrs.Extension?.toLowerCase() === "docx",
      )) {
        this.contentTypesRoot.children.unshift(element("Default", {
          Extension: "docx",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }));
      }

      const vmlId = 1025 + migrated++;
      const shapeTypeId = `_x0000_t${vmlId}`;
      const shapeId = `_x0000_i${vmlId}`;
      const shapeType = first(object, "shapetype");
      const shape = first(object, "shape");
      if (shapeType) {
        shapeType.attrs = {
          id: shapeTypeId,
          coordsize: "21600,21600",
          "o:spt": "75",
          "o:preferrelative": "t",
          path: "m@4@5l@4@11@9@11@9@5xe",
          filled: "f",
          stroked: "f",
        };
        shapeType.children = [
          element("v:stroke", { joinstyle: "miter" }),
          element("v:formulas", {}, [
            "if lineDrawn pixelLineWidth 0", "sum @0 1 0", "sum 0 0 @1", "prod @2 1 2",
            "prod @3 21600 pixelWidth", "prod @3 21600 pixelHeight", "sum @0 0 1", "prod @6 1 2",
            "prod @7 21600 pixelWidth", "sum @8 21600 0", "prod @7 21600 pixelHeight", "sum @10 21600 0",
          ].map((eqn) => element("v:f", { eqn }))),
          element("v:path", { "o:extrusionok": "f", gradientshapeok: "t", "o:connecttype": "rect" }),
          element("o:lock", { "v:ext": "edit", aspectratio: "t" }),
        ];
      }
      if (shape) {
        setAttr(shape, "id", shapeId);
        setAttr(shape, "type", `#${shapeTypeId}`);
      }
      setAttr(ole, "ProgID", "Word.Document.12");
      setAttr(ole, "ShapeID", shapeId);
      if (!first(ole, "FieldCodes")) ole.children.push(element("o:FieldCodes", {}, [], "\\s"));
    }
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

  /** Invalidated on refresh; see layoutGlobalSig. */
  private _layoutGlobalSig: string | null = null;

  /** Signature of everything OUTSIDE a paragraph's own XML that affects how it
   * breaks into lines: style + numbering definitions, doc-level layout scalars,
   * and the tracked-changes view mode. The line-break cache (layout/inline.ts)
   * combines this with a paragraph's own content signature so a style/numbering/
   * settings edit invalidates cached breaks even though the paragraph XML is
   * unchanged. Memoized until the next refresh() (styles/numbering trees are
   * stable across a plain text edit). */
  layoutGlobalSig(): string {
    if (this._layoutGlobalSig === null) {
      const parts = [
        String(this.defaultTabStop),
        String(this.compatibilityMode),
        String(this.charGridEa),
        String(this._packageResourceVersion),
        this.revisionView,
        this.styles.defaultRPr.font ?? "",
      ];
      if (this.stylesRoot) parts.push(serializeXml(this.stylesRoot));
      if (this.numberingRoot) parts.push(serializeXml(this.numberingRoot));
      // Hash to a short token: this is concatenated into every line-break cache
      // key, so it must stay small (the raw styles/numbering XML is tens of KB).
      // A collision would only mean a style edit fails to invalidate cached
      // breaks - astronomically unlikely, and the parity gate would catch it.
      this._layoutGlobalSig = cyrb53(parts.join(""));
    }
    return this._layoutGlobalSig;
  }

  refresh(): void {
    this._layoutGlobalSig = null;
    (this as { mirrorMargins: boolean }).mirrorMargins = onOff(child(this.settingsRoot, "mirrorMargins")) ?? false;
    (this as { evenAndOddHeaders: boolean }).evenAndOddHeaders = onOff(child(this.settingsRoot, "evenAndOddHeaders")) ?? false;
    const body = child(this.docRoot, "body");
    if (!body) throw new Error("document.xml has no w:body");
    // Some content (SmartArt cached drawings) lives in parts reachable only
    // through relationship indirection at parse time.
    const readPart = (part: string) => this.readXmlOptional(part);
    const refBookmarks = { open: new Map<string, Run[]>(), byName: new Map<string, Run[]>() };
    const ctx: DocParseContext = {
      ...this.ctxBase,
      rels: this.documentRels,
      readPart,
      refBookmarks,
      independentTextboxStories: true,
    };
    this.sections = parseBody(body, ctx);
    this.refBookmarks = refBookmarks.byName;
    this.headers.clear();
    this.footers.clear();
    for (const part of this.hfParts) {
      const partCtx: DocParseContext = {
        ...this.ctxBase,
        rels: part.rels,
        readPart,
        independentTextboxStories: true,
      };
      const hf: HeaderFooter = { blocks: parseBlocks(part.root, partCtx) };
      (part.isHeader ? this.headers : this.footers).set(part.relId, hf);
    }
    this.comments = this.deriveComments();
    this.styles = parseStyles(this.stylesRoot ?? undefined, this.ctxBase);
    this.numbering = parseNumbering(this.numberingRoot ?? undefined, this.ctxBase);
    // Re-derive note blocks from the retained trees (editable: keep source
    // refs) so an edit to a note's w:t re-measures.
    this.rederiveNotes();
    this._modelVersion++;
    if (this.stableIds) {
      // In-place XML mutation preserves element identity across refresh, so
      // survivors keep their ids; this fills ids for newly created nodes and
      // retires ids for deleted ones. Opt-in (collab only) — see stableIds.
      this.stableIds.assignFromRoots(this.editableRoots());
      this.stableIds.prune(this.editableRoots());
    }
  }

  /** Stable node-id side table for replicated editing. Null (and zero cost)
   * for local-only documents; call `enableStableIds()` to populate and
   * maintain it. Kept in memory only — never serialized into the XML. */
  stableIds: StableIds | null = null;

  /** Populate the stable-id table from current content and keep it updated
   * on every subsequent refresh(). Idempotent. */
  enableStableIds(): StableIds {
    if (!this.stableIds) {
      this.stableIds = new StableIds();
      this.stableIds.assignFromRoots(this.editableRoots());
    }
    return this.stableIds;
  }

  /** Reparse the two sibling body-story paragraphs created by Enter without
   * rebuilding the complete document model. Paragraphs nested in table cells
   * are included because legal documents spend most of their body inside
   * tables; revisions, bookmarks, fields, and section breaks use refresh(). */
  reparseDirectBodyParagraphSplit(
    beforeSource: XmlElement,
    afterSource: XmlElement,
  ): { before: Paragraph; after: Paragraph } | null {
    const parsed = this.reparseDirectBodyParagraphSplits(beforeSource, [afterSource]);
    return parsed ? { before: parsed[0], after: parsed[1] } : null;
  }

  /**
   * The parsed body-story block list holding the paragraph parsed from
   * `source`, and its index in it — searching sections and, recursively,
   * table cells. Null when the model doesn't hold that paragraph (it lives in
   * a header/footer/footnote part, or the model predates its creation), which
   * is every targeted-reparse helper's signal to fall back to a full
   * refresh(). Shared by those helpers and by paragraphBySource.
   */
  private locateParagraph(source: XmlElement): { blocks: Block[]; index: number } | null {
    // Memoized so a targeted reparse costs the paragraph rather than a scan of
    // every block in the document (perf B9 — this runs twice per edit). Two
    // guards make a stale entry impossible to use: the model VERSION, because
    // refresh() is the only thing that replaces these block-list arrays and a
    // remembered array would otherwise be a detached list a reparse could
    // splice into invisibly; and the paragraph at the remembered index, which
    // catches the index shifting under an insert or delete.
    const hit = this._paraLoc.get(source);
    if (hit && hit.version === this._modelVersion) {
      const block = hit.blocks[hit.index];
      if (block && block.type === "paragraph" && block.src === source) {
        return { blocks: hit.blocks, index: hit.index };
      }
    }
    const search = (blocks: Block[]): { blocks: Block[]; index: number } | null => {
      const index = blocks.findIndex((block) => block.type === "paragraph" && block.src === source);
      if (index >= 0) return { blocks, index };
      for (const block of blocks) {
        if (block.type !== "table") continue;
        for (const row of block.rows) {
          for (const cell of row.cells) {
            const found = search(cell.blocks);
            if (found) return found;
          }
        }
      }
      return null;
    };
    for (const section of this.sections) {
      const found = search(section.blocks);
      if (found) {
        this._paraLoc.set(source, { ...found, version: this._modelVersion });
        return found;
      }
    }
    return null;
  }

  private _paraLoc = new WeakMap<XmlElement, { blocks: Block[]; index: number; version: number }>();

  /** The parsed Paragraph for a retained w:p in the body story (table cells
   * included), or null when the model doesn't hold it. Lets a caller that
   * already knows which paragraph an edit addresses reach that paragraph's
   * runs without walking the whole model. */
  paragraphBySource(source: XmlElement): Paragraph | null {
    const at = this.locateParagraph(source);
    if (!at) return null;
    const block = at.blocks[at.index];
    return block.type === "paragraph" ? block : null;
  }

  /** Insert a new paragraph immediately before a retained body paragraph
   * without rebuilding the complete document model. Used by Enter at the
   * exact paragraph start, where the existing paragraph itself is unchanged. */
  insertDirectBodyParagraphBefore(
    referenceSource: XmlElement,
    insertedSource: XmlElement,
  ): Paragraph | null {
    const parent = this.findParentOf(referenceSource);
    if (!parent) return null;
    const referenceIndex = parent.children.indexOf(referenceSource);
    if (referenceIndex < 1 || parent.children[referenceIndex - 1] !== insertedSource) return null;
    if (localName(referenceSource.name) !== "p" || localName(insertedSource.name) !== "p") return null;

    const location = this.locateParagraph(referenceSource);
    if (!location || location.blocks.some((block) => block.src === insertedSource)) return null;

    const paragraph = parseParagraph(insertedSource, {
      ...this.ctxBase,
      rels: this.documentRels,
      readPart: (part: string) => this.readXmlOptional(part),
      independentTextboxStories: true,
    });
    if (paragraph.revisionHidden || paragraph.sectionBreak) return null;
    location.blocks.splice(location.index, 0, paragraph);
    return paragraph;
  }

  /** Reparse a body paragraph plus several new siblings created by
   * click-and-type without rebuilding the complete document model. */
  reparseDirectBodyParagraphSplits(
    beforeSource: XmlElement,
    afterSources: XmlElement[],
  ): Paragraph[] | null {
    if (afterSources.length === 0) return null;
    const parent = this.findParentOf(beforeSource);
    if (!parent) return null;
    const beforeIndex = parent.children.indexOf(beforeSource);
    if (beforeIndex < 0 || afterSources.some((source, i) => parent.children[beforeIndex + i + 1] !== source)) {
      return null;
    }
    if (localName(beforeSource.name) !== "p" || afterSources.some((source) => localName(source.name) !== "p")) {
      return null;
    }

    const location = this.locateParagraph(beforeSource);
    if (!location) return null;
    const { blocks, index: blockIndex } = location;
    if (afterSources.some((source) => blocks.some((block) => block.src === source))) return null;
    const old = blocks[blockIndex];
    if (old.type !== "paragraph" || old.sectionBreak) return null;

    const unsafe = (element: XmlElement): boolean => {
      const name = localName(element.name);
      if (
        name === "sectPr" ||
        name === "bookmarkStart" ||
        name === "bookmarkEnd" ||
        name === "fldChar" ||
        name === "instrText" ||
        name === "fldSimple" ||
        name === "sdt" ||
        name === "ins" ||
        name === "del" ||
        name.startsWith("move") ||
        name.endsWith("PrChange")
      ) return true;
      return element.children.some(unsafe);
    };
    if (unsafe(beforeSource) || afterSources.some(unsafe)) return null;

    const readPart = (part: string) => this.readXmlOptional(part);
    const ctx: DocParseContext = {
      ...this.ctxBase,
      rels: this.documentRels,
      readPart,
      independentTextboxStories: true,
    };
    const parsed = [beforeSource, ...afterSources].map((source) => parseParagraph(source, ctx));
    if (parsed.some((paragraph) => paragraph.revisionHidden || paragraph.sectionBreak)) return null;
    blocks.splice(blockIndex, 1, ...parsed);
    return parsed;
  }

  /** Reparse one retained body-story paragraph after a structural edit that
   * leaves its surrounding block list unchanged, such as inserting ink. */
  reparseBodyParagraph(source: XmlElement): Paragraph | null {
    if (localName(source.name) !== "p") return null;
    const location = this.locateParagraph(source);
    if (!location) return null;
    const old = location.blocks[location.index];
    if (old.type !== "paragraph" || old.sectionBreak) return null;

    // Bookmark ranges retain parsed Run identities in refBookmarks; a local
    // paragraph replacement must rebuild those captures or REF/PAGEREF fields
    // would read detached runs. Ranges fully INSIDE this paragraph (every
    // bookmarkStart's id also ends here, and vice versa) are re-captured from
    // the reparse below. A range crossing the paragraph boundary cannot be
    // rebuilt locally, so it falls back to the full refresh. This matters at
    // scale: rejecting bookmarks outright sent the first keystroke in any
    // heading (TOC targets are bookmarked) through doc.refresh() + a full
    // relayout — an inert multi-second stall per keystroke on long documents.
    const starts = new Map<string, string>(); // bookmark id -> name
    const ends = new Set<string>();
    let unsafe = false;
    const scan = (element: XmlElement): void => {
      const name = localName(element.name);
      if (name === "sectPr") {
        unsafe = true;
        return;
      }
      if (name === "bookmarkStart") {
        const id = attr(element, "id");
        if (id) starts.set(id, attr(element, "name") ?? "");
      } else if (name === "bookmarkEnd") {
        const id = attr(element, "id");
        if (id) ends.add(id);
      }
      for (const c of element.children) scan(c);
    };
    scan(source);
    if (unsafe) return null;
    if (starts.size !== ends.size) return null;
    for (const id of starts.keys()) if (!ends.has(id)) return null;

    // A bookmark opened in ANOTHER paragraph can span this one without any
    // marker inside it; its capture then holds this paragraph's old runs.
    if (this.refBookmarks.size > 0) {
      const paragraphNames = new Set(starts.values());
      const oldRuns = new Set<Run>();
      for (const c of old.children) {
        for (const run of c.type === "run" ? [c] : c.runs) oldRuns.add(run);
      }
      for (const [name, runs] of this.refBookmarks) {
        if (paragraphNames.has(name)) continue;
        if (runs.some((run) => oldRuns.has(run))) return null;
      }
    }

    const refBookmarks = { open: new Map<string, Run[]>(), byName: new Map<string, Run[]>() };
    const paragraph = parseParagraph(source, {
      ...this.ctxBase,
      rels: this.documentRels,
      readPart: (part: string) => this.readXmlOptional(part),
      refBookmarks,
      independentTextboxStories: true,
    });
    if (paragraph.revisionHidden || paragraph.sectionBreak) return null;
    for (const [name, runs] of refBookmarks.byName) this.refBookmarks.set(name, runs);
    location.blocks[location.index] = paragraph;
    return paragraph;
  }

  /** Reparse two sibling body-story paragraphs after Backspace/Delete merged
   * their XML into one. Keeps the parsed model generation stable so long
   * documents can use incremental layout instead of repaginating in full. */
  reparseDirectBodyParagraphMerge(
    beforeSource: XmlElement,
    afterSource: XmlElement,
    survivorSource: XmlElement,
  ): Paragraph | null {
    if (survivorSource !== beforeSource && survivorSource !== afterSource) return null;
    if (localName(beforeSource.name) !== "p" || localName(afterSource.name) !== "p") return null;
    const parent = this.findParentOf(survivorSource);
    if (!parent || parent.children.includes(survivorSource === beforeSource ? afterSource : beforeSource)) return null;

    const findBlockList = (blocks: Block[]): { blocks: Block[]; index: number } | null => {
      const index = blocks.findIndex((block, i) =>
        block.type === "paragraph" &&
        block.src === beforeSource &&
        blocks[i + 1]?.type === "paragraph" &&
        blocks[i + 1].src === afterSource,
      );
      if (index >= 0) return { blocks, index };
      for (const block of blocks) {
        if (block.type !== "table") continue;
        for (const row of block.rows) {
          for (const cell of row.cells) {
            const found = findBlockList(cell.blocks);
            if (found) return found;
          }
        }
      }
      return null;
    };
    let location: { blocks: Block[]; index: number } | null = null;
    for (const section of this.sections) {
      location = findBlockList(section.blocks);
      if (location) break;
    }
    if (!location) return null;
    const { blocks, index } = location;
    const before = blocks[index];
    const after = blocks[index + 1];
    if (before.type !== "paragraph" || after.type !== "paragraph" || before.sectionBreak || after.sectionBreak) return null;

    const unsafe = (element: XmlElement): boolean => {
      const name = localName(element.name);
      if (
        name === "sectPr" ||
        name === "bookmarkStart" ||
        name === "bookmarkEnd" ||
        name === "fldChar" ||
        name === "instrText" ||
        name === "fldSimple" ||
        name === "sdt" ||
        name === "ins" ||
        name === "del" ||
        name.startsWith("move") ||
        name.endsWith("PrChange")
      ) return true;
      return element.children.some(unsafe);
    };
    if (unsafe(beforeSource) || unsafe(afterSource)) return null;

    const readPart = (part: string) => this.readXmlOptional(part);
    const merged = parseParagraph(survivorSource, {
      ...this.ctxBase,
      rels: this.documentRels,
      readPart,
      independentTextboxStories: true,
    });
    if (merged.revisionHidden || merged.sectionBreak) return null;
    blocks.splice(index, 2, merged);
    return merged;
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
    // commentsExtended threading: paraIdParent links a reply to its parent,
    // and w15:done marks the thread resolved.
    if (this.commentsExtRoot) {
      const parentOf = new Map<string, string>();
      const doneOf = new Map<string, boolean>();
      for (const ex of this.commentsExtRoot.children) {
        if (localName(ex.name) !== "commentEx") continue;
        const pid = attr(ex, "paraId");
        const parent = attr(ex, "paraIdParent");
        if (pid && parent) parentOf.set(pid, parent);
        if (pid) doneOf.set(pid, attr(ex, "done") === "1");
      }
      const byParaId = new Map(out.filter((c) => c.paraId).map((c) => [c.paraId!, c]));
      for (const c of out) {
        const parentPara = c.paraId ? parentOf.get(c.paraId) : undefined;
        if (parentPara) c.parentId = byParaId.get(parentPara)?.id;
        if (c.paraId && doneOf.get(c.paraId)) c.resolved = true;
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
    this.numbering = parseNumbering(this.numberingRoot ?? undefined, this.ctxBase);
    this._layoutGlobalSig = null;
  }

  /** Retained styles.xml tree, or null when the package has no styles part.
   * Style-definition editing mutates this tree in place. */
  stylesTree(): XmlElement | null {
    return this.stylesRoot;
  }

  /**
   * Re-resolve the style cascade after a definition changed in place.
   *
   * A style definition is not addressed by any paragraph's XML, so mutating one
   * changes how EVERY paragraph resolves without changing a single w:p. Two
   * things therefore have to be reset by hand: the parsed Styles map every
   * effective-props call reads, and layoutGlobalSig, which is what makes the
   * line-break cache treat cached breaks for unchanged paragraph XML as stale.
   */
  markStylesChanged(): void {
    this.stylesDirty = true;
    this.styles = parseStyles(this.stylesRoot ?? undefined, this.ctxBase);
    this._layoutGlobalSig = null;
  }

  /**
   * Retained bibliography sources tree (b:Sources), or null when the package
   * has no sources part. With create=true, a missing part is created and
   * registered the way Word lays one out (ECMA-376 §22.6 + Part 2 custom XML
   * data storage): the data itself as customXml/itemN.xml, a datastore-item
   * properties part (itemPropsN.xml) declaring the Bibliography schema URI as
   * its schemaRef, a .rels from the item to its properties, a customXml
   * relationship from the main document part, and the content-type overrides.
   *
   * DETERMINISM. Everything written is a pure function of the package —
   * the item number is the first free one, and the datastore itemID GUID is a
   * fixed constant rather than Word's random draw — so two replicas creating
   * the part through the same sequenced operation write identical bytes.
   */
  sourcesTree(create = false): XmlElement | null {
    if (this.sourcesRoot || !create) return this.sourcesRoot;
    let n = 1;
    while (this.pkg.has(`customXml/item${n}.xml`) || this.pkg.has(`customXml/itemProps${n}.xml`)) n++;
    this.sourcesPart = `customXml/item${n}.xml`;
    this.sourcesRoot = {
      name: "b:Sources",
      attrs: {
        SelectedStyle: "\\APASixthEditionOfficeOnline.xsl",
        StyleName: "APA",
        Version: "6",
        "xmlns:b": "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
        xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/bibliography",
      },
      children: [],
      text: "",
    };
    // The datastore properties part and its relationship, as Word writes
    // them. The itemID is a GUID in Word; a FIXED one keeps replica creation
    // byte-identical, and uniqueness only matters within this package, which
    // holds at most one part this engine creates.
    this.sourcesAux = {
      [`customXml/itemProps${n}.xml`]:
        `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n` +
        `<ds:datastoreItem ds:itemID="{4AA3A060-2A83-4A31-8C77-13810E5B4FAB}" ` +
        `xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">` +
        `<ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"/>` +
        `</ds:schemaRefs></ds:datastoreItem>`,
      [`customXml/_rels/item${n}.xml.rels`]:
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" ` +
        `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" ` +
        `Target="itemProps${n}.xml"/></Relationships>`,
    };
    {
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
          Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
          Target: `../customXml/item${n}.xml`,
        },
        children: [],
        text: "",
      });
      this.documentRels.set(relId, {
        id: relId,
        type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
        target: this.sourcesPart,
        external: false,
      });
    }
    if (this.contentTypesRoot) {
      const overrides: [string, string][] = [
        // The item part is application/xml. Word covers it with the package's
        // Default rule for the xml extension; declare the override only when
        // that Default is missing, so the created package stays valid either
        // way without touching packages that already carry the Default.
        ...(this.contentTypesRoot.children.some(
          (c) => localName(c.name) === "Default" && c.attrs["Extension"] === "xml",
        )
          ? []
          : [["/" + this.sourcesPart, "application/xml"] as [string, string]]),
        [`/customXml/itemProps${n}.xml`, "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"],
      ];
      for (const [partName, contentType] of overrides) {
        if (this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) continue;
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: { PartName: partName, ContentType: contentType },
          children: [],
          text: "",
        });
      }
    }
    this.sourcesDirty = true;
    return this.sourcesRoot;
  }

  /**
   * Invalidate what a sources-part edit changes downstream. Like
   * markStylesChanged, the part is addressed by no paragraph's XML: a
   * CITATION field's painted text resolves from the part at layout, so the
   * layout signature has to drop for unchanged paragraph XML to repaint.
   */
  markSourcesChanged(): void {
    this.sourcesDirty = true;
    this._layoutGlobalSig = null;
  }

  /**
   * Retained footnotes tree. With create=true, a missing footnotes.xml part
   * is created and registered (with Word's required separator footnotes) so
   * inserted footnotes serialize and round-trip.
   */
  footnotesTree(create = false): XmlElement | null {
    if (this.footnotesRoot || !create) return this.footnotesRoot;
    this.footnotesPart = this.createNotesPart("footnotes");
    this.footnotesRoot = notesPartRoot("footnote");
    this.footnotesDirty = true;
    return this.footnotesRoot;
  }

  /** Retained endnotes tree; the endnote mirror of footnotesTree. */
  endnotesTree(create = false): XmlElement | null {
    if (this.endnotesRoot || !create) return this.endnotesRoot;
    this.endnotesPart = this.createNotesPart("endnotes");
    this.endnotesRoot = notesPartRoot("endnote");
    this.endnotesDirty = true;
    return this.endnotesRoot;
  }

  /** Register a new notes part in document.xml.rels and [Content_Types].xml,
   * and return the part path. The tree itself is the caller's. */
  private createNotesPart(kind: "footnotes" | "endnotes"): string {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    const part = `${docDir}${kind}.xml`;
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
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}`,
        Target: `${kind}.xml`,
      },
      children: [],
      text: "",
    });
    if (this.contentTypesRoot) {
      const partName = "/" + part;
      if (!this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType:
              `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`,
          },
          children: [],
          text: "",
        });
      }
    }
    return part;
  }

  /**
   * Create an empty header/footer part (with a default-type reference in
   * every sectPr) when the document has none - Word does this implicitly the
   * first time you edit the header area. Returns the part's root.
   */
  /** Whether the document declares any section properties at all. A blank
   * minimal document has none; see ensureHfPart. */
  private hasSectPr(): boolean {
    const walk = (e: XmlElement): boolean =>
      localName(e.name) === "sectPr" || e.children.some(walk);
    return walk(this.docRoot);
  }

  /** Whether a header (or footer) part already exists — the precondition
   * `ensureHfPart` tests internally, exposed so a caller can tell a real
   * creation from a no-op before mutating (the collab apply needs to know
   * whether an ensureHeaderFooter intent has anything to do). */
  hasHfPart(kind: "header" | "footer"): boolean {
    return this.hfParts.some((p) => p.isHeader === (kind === "header"));
  }

  ensureHfPart(kind: "header" | "footer"): XmlElement {
    const isHeader = kind === "header";
    const existing = this.hfParts.find((p2) => p2.isHeader === isHeader);
    if (existing) return existing.root;
    const { relId, root } = this.createHfPart(kind);
    // A minimal document (the blank the demo starts from) has NO sectPr at
    // all, so the reference walk below had nowhere to put the reference: the
    // part was created, never referenced, never laid out, and the caller found
    // no caret target — "the header won't open". Materialize the default body
    // section first, exactly as setPageLayout/setLineNumbering do for the same
    // shape of document.
    this.ensureBodySectPr();
    // Reference from every sectPr (schema: hf references lead the sectPr).
    const refName = isHeader ? "w:headerReference" : "w:footerReference";
    const addRef = (e: XmlElement): void => {
      if (localName(e.name) === "sectPr") {
        e.children.unshift(this.hfReference(refName, "default", relId));
        return;
      }
      for (const c of e.children) addRef(c);
    };
    addRef(this.docRoot);
    this.refresh();
    return root;
  }

  /** A w:headerReference / w:footerReference element. */
  private hfReference(refName: string, type: "default" | "first" | "even", relId: string): XmlElement {
    return {
      name: refName,
      attrs: {
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "w:type": type,
        "r:id": relId,
      },
      children: [],
      text: "",
    };
  }

  /** Materialize the default body-level sectPr when the document declares no
   * section properties at all (schema: the body's sectPr is its LAST child). */
  private ensureBodySectPr(): void {
    const bodyEl = child(this.docRoot, "body");
    if (bodyEl && !this.hasSectPr()) {
      const w = bodyEl.name.includes(":") ? bodyEl.name.slice(0, bodyEl.name.indexOf(":") + 1) : "";
      bodyEl.children.push({ name: `${w}sectPr`, attrs: {}, children: [], text: "" });
    }
  }

  /**
   * Give every section a header and footer reference of the given variant
   * type, creating one empty part per band the way Word does the first time
   * "different first page" (type "first") or "different odd & even pages"
   * (type "even") is enabled. Sections that already carry a reference of the
   * type keep it — and its part's content. True when anything was added.
   */
  ensureHfVariantParts(type: "first" | "even"): boolean {
    this.ensureBodySectPr();
    const sectPrs: XmlElement[] = [];
    const walk = (e: XmlElement): void => {
      if (localName(e.name) === "sectPr") sectPrs.push(e);
      else for (const c of e.children) walk(c);
    };
    walk(this.docRoot);
    if (sectPrs.length === 0) return false;
    let changed = false;
    for (const kind of ["header", "footer"] as const) {
      const refLocal = kind === "header" ? "headerReference" : "footerReference";
      const missing = sectPrs.filter(
        (sectPr) => !sectPr.children.some(
          (c) => localName(c.name) === refLocal && attr(c, "type") === type,
        ),
      );
      if (missing.length === 0) continue;
      const { relId } = this.createHfPart(kind);
      for (const sectPr of missing) {
        sectPr.children.unshift(this.hfReference(`w:${refLocal}`, type, relId));
      }
      changed = true;
    }
    if (changed) this.refresh();
    return changed;
  }

  /** Create and register a new empty header/footer part — package part,
   * document relationship, content-type override, hfParts entry — with no
   * section references of its own; the caller decides which sectPrs point at
   * it and with which w:type. */
  private createHfPart(kind: "header" | "footer"): { relId: string; root: XmlElement } {
    const isHeader = kind === "header";
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    // A part created earlier in this session is not in the package until
    // save(), so the numbering probe checks both stores.
    while (
      this.pkg.has(`${docDir}${kind}${n}.xml`) ||
      this.hfParts.some((p) => p.target === `${docDir}${kind}${n}.xml`)
    ) n++;
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
    this.hfParts.push({ relId, target, root, isHeader, rels: new Map(), relsRoot: null });
    return { relId, root };
  }

  markFootnotesChanged(): void {
    this.footnotesDirty = true;
    this.rederiveNotes();
  }

  markEndnotesChanged(): void {
    this.endnotesDirty = true;
    this.rederiveNotes();
  }

  /** Re-derive both id -> blocks maps from the retained trees, so layout sees
   * an inserted or edited note. */
  private rederiveNotes(): void {
    for (const [root, rels, into] of [
      [this.footnotesRoot, this.footnotesRels, this.footnotes],
      [this.endnotesRoot, this.endnotesRels, this.endnotes],
    ] as const) {
      if (!root) continue;
      into.clear();
      for (const [id, blocks] of parseNotesPart(root, { ...this.ctxBase, rels })) into.set(id, blocks);
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

  /** Flag a notes part dirty when `t` lives inside it, so save()
   * re-serializes footnotes.xml / endnotes.xml. Called by the editor after a
   * text edit; a no-op for body/header/footer targets. */
  markDirtyIfFootnote(t: XmlElement): void {
    if (this.footnotesDirty && this.endnotesDirty) return;
    this.markNotePartsChanged(this.notePartsHolding(t));
  }

  /**
   * Which note parts hold `el` right now. An operation that DETACHES `el` —
   * accepting or rejecting a revision, for one — has to ask BEFORE it mutates
   * and mark the answer afterwards: a detached element belongs to no part, so
   * a containment test run after the fact reports nothing and save() then
   * leaves the stale footnotes.xml / endnotes.xml in the package.
   */
  notePartsHolding(el: XmlElement): NotePart[] {
    const contains = (root: XmlElement): boolean => {
      if (root === el) return true;
      for (const c of root.children) if (contains(c)) return true;
      return false;
    };
    const parts: NotePart[] = [];
    if (this.footnotesRoot && contains(this.footnotesRoot)) parts.push("footnotes");
    if (this.endnotesRoot && contains(this.endnotesRoot)) parts.push("endnotes");
    return parts;
  }

  /** Flag the given note parts dirty, so save() re-serializes exactly those. */
  markNotePartsChanged(parts: readonly NotePart[]): void {
    for (const part of parts) {
      if (part === "footnotes") this.markFootnotesChanged();
      else this.markEndnotesChanged();
    }
  }

  /** The mutable XML roots (document body, related modeled parts, settings).
   * settingsRoot is always second and always present so its history snapshot
   * index stays stable even when optional related roots are created later. */
  editableRoots(): XmlElement[] {
    const roots = [this.docRoot, this.settingsRoot, ...this.hfParts.map((p) => p.root)];
    if (this.footnotesRoot) roots.push(this.footnotesRoot);
    if (this.endnotesRoot) roots.push(this.endnotesRoot);
    if (this.commentsRoot) roots.push(this.commentsRoot);
    if (this.commentsExtRoot) roots.push(this.commentsExtRoot);
    return roots;
  }

  /** CT_Settings children that precede w:mirrorMargins in the schema sequence
   * (§17.15.1.78). Insertions go before the first child that is not a
   * predecessor, so Word never has to repair settings.xml. */
  private static readonly SETTINGS_BEFORE_MIRROR = [
    "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
    "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
    "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
    "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
  ];

  /** The schema predecessors of w:evenAndOddHeaders (§17.10.1): everything up
   * to mirrorMargins, then the sequence through defaultTableStyle. */
  private static readonly SETTINGS_BEFORE_EVEN_AND_ODD = [
    ...DocxDocument.SETTINGS_BEFORE_MIRROR,
    "mirrorMargins", "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors",
    "activeWritingStyle", "proofState", "formsDesign", "attachedTemplate", "linkStyles",
    "stylePaneFormatFilter", "stylePaneSortMethod", "documentType", "mailMerge", "revisionView",
    "trackChanges", "doNotTrackMoves", "doNotTrackFormatting", "documentProtection",
    "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "defaultTabStop",
    "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
  ];

  /** Set or remove one on/off settings.xml toggle at its schema position. */
  private setSettingsToggle(local: string, predecessors: readonly string[], enabled: boolean): void {
    this.settingsRoot.children = this.settingsRoot.children.filter((c) => localName(c.name) !== local);
    if (enabled) {
      const before = new Set(predecessors);
      const index = this.settingsRoot.children.findIndex((c) => !before.has(localName(c.name)));
      const el = { name: `w:${local}`, attrs: {}, children: [], text: "" };
      this.settingsRoot.children.splice(index === -1 ? this.settingsRoot.children.length : index, 0, el);
    }
    this.registerSettingsPart();
    this.settingsDirty = true;
  }

  /** Toggle the document-global facing-page margin mode in settings.xml. */
  setMirrorMargins(enabled: boolean): void {
    this.setSettingsToggle("mirrorMargins", DocxDocument.SETTINGS_BEFORE_MIRROR, enabled);
    (this as { mirrorMargins: boolean }).mirrorMargins = enabled;
  }

  /** Toggle w:evenAndOddHeaders (§17.10.1) — the document-global switch that
   * makes even pages use the "even" header/footer variants. */
  setEvenAndOddHeaders(enabled: boolean): void {
    this.setSettingsToggle("evenAndOddHeaders", DocxDocument.SETTINGS_BEFORE_EVEN_AND_ODD, enabled);
    (this as { evenAndOddHeaders: boolean }).evenAndOddHeaders = enabled;
  }

  /** Register settings.xml in document.xml.rels and [Content_Types].xml when
   * the package was born without it (a minimal document). */
  private registerSettingsPart(): void {
    const rels = this.ensureRelsRoot();
    const settingsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
    if (!rels.children.some((r) => r.attrs["Type"]?.endsWith("/settings"))) {
      let maxId = 0;
      for (const r of rels.children) {
        const match = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
        if (match) maxId = Math.max(maxId, parseInt(match[1], 10));
      }
      rels.children.push({
        name: "Relationship",
        attrs: {
          Id: `rId${maxId + 1}`,
          Type: settingsRelType,
          Target: this.settingsPart.slice(this.docPart.lastIndexOf("/") + 1),
        },
        children: [],
        text: "",
      });
    }
    if (this.contentTypesRoot) {
      const partName = `/${this.settingsPart}`;
      if (!this.contentTypesRoot.children.some((c) => c.attrs["PartName"] === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: {
            PartName: partName,
            ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
          },
          children: [],
          text: "",
        });
      }
    }
  }

  /**
   * Find the parent element of `target` in any modeled XML tree (document
   * body, headers, footers). Linear scan — documents are small and this only
   * runs on structural edits (Enter, paragraph merge).
   */
  /** The header part roots, in package order. A watermark is authored into
   * every one of them: Word paints it on every page, and a document with a
   * first-page or even-page header has more than one header part. */
  headerRoots(): XmlElement[] {
    return this.hfParts.filter((p) => p.isHeader).map((p) => p.root);
  }

  /** XML roots that can carry tracked changes: body, headers/footers, notes. */
  revisionRoots(): XmlElement[] {
    return this.contentRoots();
  }

  /** Body, header/footer and note roots — every tree the caret can reach. */
  private contentRoots(): XmlElement[] {
    const roots = [this.docRoot, ...this.hfParts.map((p) => p.root)];
    if (this.footnotesRoot) roots.push(this.footnotesRoot);
    if (this.endnotesRoot) roots.push(this.endnotesRoot);
    return roots;
  }

  /**
   * Memoized child→parent links. Nothing invalidates this map — the tree is
   * spliced constantly — so every answer is RE-DERIVED before it is returned:
   * the memoized parent must still list the target among its children, and
   * must still hang off a root (see memoIsLive). A stale entry can then only
   * cost a cache miss, never a wrong answer.
   *
   * Both halves are needed. Containment alone is not proof of a live parent:
   * a run split for formatting is spliced OUT of its paragraph while its own
   * `children` array still lists the elements that moved into the replacement
   * runs, so a containment-only check would hand back the detached run where a
   * walk from the roots finds the new one.
   */
  private _parentMemo = new WeakMap<XmlElement, XmlElement>();

  private memoIsLive(parent: XmlElement, roots: XmlElement[]): boolean {
    // Walk memoized links up to a root. Bounded because each hop must be a
    // verified parent link, and a cycle or a gap ends the walk as "not live",
    // which just falls through to the authoritative tree walk.
    for (let cur = parent, hops = 0; hops < 128; hops++) {
      if (roots.includes(cur)) return true;
      const up = this._parentMemo.get(cur);
      if (!up || !up.children.includes(cur)) return false;
      cur = up;
    }
    return false;
  }

  findParentOf(target: XmlElement): XmlElement | undefined {
    const roots = this.contentRoots();
    const memo = this._parentMemo.get(target);
    if (memo && memo.children.includes(target) && this.memoIsLive(memo, roots)) return memo;
    // A miss walks the tree anyway, so record EVERY link the walk passes, not
    // just the target's: the ops that call this call it several times in a row
    // (split resolves run → paragraph → parent), and the neighbours are then
    // free. This is what keeps a structural edit from paying one full-document
    // walk per lookup (perf B9).
    const walk = (el: XmlElement): XmlElement | undefined => {
      for (const c of el.children) {
        this._parentMemo.set(c, el);
        if (c === target) return el;
        const hit = walk(c);
        if (hit) return hit;
      }
      return undefined;
    };
    for (const root of roots) {
      const hit = walk(root);
      if (hit) return hit;
    }
    return undefined;
  }

  /** Record a parent link the caller just created (a splice), so the next
   * lookup skips the full-document walk. Advisory only — findParentOf
   * verifies every memo against the live tree before trusting it. */
  noteParent(child: XmlElement, parent: XmlElement): void {
    this._parentMemo.set(child, parent);
  }

  /**
   * Serialize the (possibly edited) document back to .docx bytes. Only the
   * XML parts we model are re-serialized; every other part round-trips
   * byte-for-byte.
   */
  private rememberOriginalXml(part: string, root: XmlElement): void {
    this.originalModeledXml.set(part, serializeXml(root, true));
  }

  private writeModeledXml(files: Record<string, Uint8Array>, part: string, root: XmlElement): void {
    const xml = serializeXml(root, true);
    if (xml !== this.originalModeledXml.get(part)) files[part] = strToU8(xml);
  }

  /**
   * Canonicalize producer shorthand that Google Docs otherwise interprets as
   * a fixed, few-twip table. Word treats `tblW="100%"` plus a placeholder
   * grid as autofit; Google needs the standard pct value and a usable cached
   * grid. The cached widths follow the same content-dominant shape and do not
   * change Word's autofit result.
   */
  /** Undo closures recorded by save-time fixups while `saveJournal` is
   * active, so `save()` can revert every live-tree mutation and remain
   * side-effect-free. Collaborative checkpoints re-serialize the
   * authoritative document repeatedly; a save that mutated the tree would
   * change its hash outside the intent stream and desync the fleet. */
  private saveJournal: (() => void)[] | null = null;

  private journalSetAttr(element: XmlElement, key: string, value: string): void {
    if (this.saveJournal) {
      const had = Object.prototype.hasOwnProperty.call(element.attrs, key);
      const prev = element.attrs[key];
      this.saveJournal.push(() => {
        if (had) element.attrs[key] = prev;
        else delete element.attrs[key];
      });
    }
    element.attrs[key] = value;
  }

  private normalizePercentageTableGrids(): void {
    const setAttr = (element: XmlElement, name: string, value: string): void => {
      const key = Object.keys(element.attrs).find((item) => localName(item) === name);
      this.journalSetAttr(element, key ?? `${element.name.includes(":") ? element.name.split(":")[0] + ":" : ""}${name}`, value);
    };
    const textLength = (blocks: Block[]): number => {
      let length = 0;
      for (const block of blocks) {
        if (block.type === "table") {
          for (const row of block.rows) {
            for (const cell of row.cells) length += textLength(cell.blocks);
          }
          continue;
        }
        for (const item of block.children) {
          const runs = item.type === "run" ? [item] : item.runs;
          for (const run of runs) {
            for (const content of run.content) {
              if (content.kind === "text") length += content.text.length;
              else if (content.kind === "tab") length += 4;
            }
          }
        }
      }
      return length;
    };
    const normalize = (table: Extract<Block, { type: "table" }>, available: number): void => {
      const source = table.src;
      if (!source) return;
      const tableProps = child(source, "tblPr");
      const tableWidth = child(tableProps, "tblW");
      const rawWidth = attr(tableWidth, "w")?.trim();
      if (attr(tableWidth, "type") !== "pct" || !rawWidth?.endsWith("%")) return;

      const percent = Number.parseFloat(rawWidth);
      if (!Number.isFinite(percent) || percent <= 0) return;
      setAttr(tableWidth!, "w", String(Math.round(percent * 50)));

      const grid = child(source, "tblGrid");
      const columns = children(grid, "gridCol");
      if (columns.length === 0) return;
      const target = Math.round(pxToTwips(available) * percent / 100);
      const authoredTotal = columns.reduce((sum, column) => sum + (intAttr(column, "w") ?? 0), 0);
      if (target <= 0 || authoredTotal >= target * 0.1) return;

      const floor = 600;
      const widths = new Array<number>(columns.length).fill(floor);
      for (const row of table.rows) {
        let column = 0;
        for (const cell of row.cells) {
          const span = Math.max(1, Math.min(cell.props.gridSpan, columns.length - column));
          const demand = Math.max(floor * span, textLength(cell.blocks) * 100 + 300);
          for (let offset = 0; offset < span; offset++) {
            widths[column + offset] = Math.max(widths[column + offset], demand / span);
          }
          column += cell.props.gridSpan;
          if (column >= columns.length) break;
        }
      }

      const dominant = widths.indexOf(Math.max(...widths));
      const total = widths.reduce((sum, width) => sum + width, 0);
      if (total < target) {
        widths[dominant] += target - total;
      } else if (total > target) {
        const slack = widths.reduce((sum, width) => sum + Math.max(0, width - floor), 0);
        const scale = slack > 0 ? Math.min(1, (total - target) / slack) : 0;
        for (let index = 0; index < widths.length; index++) {
          widths[index] -= Math.max(0, widths[index] - floor) * scale;
        }
      }
      const rounded = widths.map(Math.round);
      rounded[dominant] += target - rounded.reduce((sum, width) => sum + width, 0);
      columns.forEach((column, index) => setAttr(column, "w", String(rounded[index])));
    };

    for (const section of this.sections) {
      const contentWidth = section.props.pageWidth - section.props.marginLeft -
        section.props.marginRight - section.props.gutter;
      const columnCount = Math.max(1, section.props.columns.count);
      const available = section.props.columns.widths?.[0] ??
        (contentWidth - section.props.columns.space * (columnCount - 1)) / columnCount;
      for (const block of section.blocks) {
        if (block.type === "table") normalize(block, available);
      }
    }
  }

  save(): Uint8Array {
    const journal: (() => void)[] = [];
    this.saveJournal = journal;
    try {
      return this.buildPackage();
    } finally {
      // Revert save-time fixups in reverse so the live tree is byte-identical
      // to before the save (the collab-checkpoint purity invariant).
      for (let i = journal.length - 1; i >= 0; i--) journal[i]();
      this.saveJournal = null;
    }
  }

  saveAsync(): Promise<Uint8Array> {
    const journal: (() => void)[] = [];
    this.saveJournal = journal;
    let files: Record<string, Uint8Array>;
    try {
      files = this.buildPackageFiles();
    } finally {
      for (let i = journal.length - 1; i >= 0; i--) journal[i]();
      this.saveJournal = null;
    }
    return new Promise((resolve, reject) => {
      zip(files, { mtime: FIXED_ZIP_MTIME }, (error, bytes) => {
        if (error) reject(error);
        else resolve(bytes);
      });
    });
  }

  private buildPackage(): Uint8Array {
    return zipSync(this.buildPackageFiles(), { mtime: FIXED_ZIP_MTIME });
  }

  private buildPackageFiles(): Record<string, Uint8Array> {
    this.normalizePercentageTableGrids();
    const files: Record<string, Uint8Array> = { ...this.pkg.raw() };
    if (files["docProps/custom.xml"] && this.contentTypesRoot && !this.contentTypesRoot.children.some(
      (item) => localName(item.name) === "Override" && item.attrs.PartName === "/docProps/custom.xml",
    )) {
      const prefixEnd = this.contentTypesRoot.name.indexOf(":") + 1;
      const prefix = prefixEnd > 0 ? this.contentTypesRoot.name.slice(0, prefixEnd) : "";
      const ctRoot = this.contentTypesRoot;
      ctRoot.children.push({
        name: `${prefix}Override`,
        attrs: {
          PartName: "/docProps/custom.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        },
        children: [],
        text: "",
      });
      if (this.saveJournal) this.saveJournal.push(() => ctRoot.children.pop());
    }
    this.writeModeledXml(files, this.docPart, this.docRoot);
    for (const part of this.hfParts) {
      this.writeModeledXml(files, part.target, part.root);
      if (part.relsRoot) this.writeModeledXml(files, relsPathFor(part.target), part.relsRoot);
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
    if (this.endnotesDirty && this.endnotesRoot && this.endnotesPart) {
      files[this.endnotesPart] = strToU8(serializeXml(this.endnotesRoot, true));
    }
    if (this.sourcesDirty && this.sourcesRoot && this.sourcesPart) {
      files[this.sourcesPart] = strToU8(serializeXml(this.sourcesRoot, true));
    }
    if (this.sourcesAux) {
      for (const [name, xml] of Object.entries(this.sourcesAux)) files[name] = strToU8(xml);
    }
    if (this.settingsDirty) files[this.settingsPart] = strToU8(serializeXml(this.settingsRoot, true));
    if (this.relsRoot) this.writeModeledXml(files, this.relsPath, this.relsRoot);
    if (this.contentTypesRoot) this.writeModeledXml(files, "[Content_Types].xml", this.contentTypesRoot);
    return files;
  }

  /** Fresh unique docPr id for inserted drawings. Seeded once past the
   * highest id already present in any editable root (floor 1000, matching
   * the historical seed for documents with no drawings) so a document that
   * already carries drawings never gets a colliding id. */
  nextDrawingId(): number {
    if (this.docPrIdCounter === null) {
      let max = 999;
      const scan = (el: XmlElement): void => {
        if (localName(el.name) === "docPr") {
          const idKey = Object.keys(el.attrs).find((k) => localName(k) === "id");
          const v = idKey ? parseInt(el.attrs[idKey], 10) : NaN;
          if (Number.isFinite(v) && v > max) max = v;
        }
        for (const c of el.children) scan(c);
      };
      for (const root of this.editableRoots()) scan(root);
      this.docPrIdCounter = max + 1;
    }
    return this.docPrIdCounter++;
  }

  /** Next unused revision id (w:id on w:ins/w:del). Seeded once past the
   * highest id already present in any editable root so a document that
   * already has tracked changes never collides. */
  private revIdCounter: number | null = null;
  nextRevisionId(): number {
    if (this.revIdCounter === null) {
      let max = 0;
      const scan = (el: XmlElement): void => {
        const ln = localName(el.name);
        if (ln === "ins" || ln === "del" || ln === "moveTo" || ln === "moveFrom") {
          const idKey = Object.keys(el.attrs).find((k) => localName(k) === "id");
          const v = idKey ? parseInt(el.attrs[idKey], 10) : NaN;
          if (Number.isFinite(v) && v > max) max = v;
        }
        for (const c of el.children) scan(c);
      };
      for (const root of this.editableRoots()) scan(root);
      this.revIdCounter = max + 1;
    }
    return this.revIdCounter++;
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

  private addImageResourceAt(bytes: Uint8Array, ext: string, source?: XmlElement): { relId: string; part: string } {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    while (this.pkg.has(`${docDir}media/image${n}.${ext}`)) n++;
    const part = `${docDir}media/image${n}.${ext}`;
    this.pkg.raw()[part] = bytes;

    const contains = (root: XmlElement, target: XmlElement): boolean =>
      root === target || root.children.some((item) => contains(item, target));
    const owner = source ? this.hfParts.find((candidate) => contains(candidate.root, source)) : undefined;
    let relsRoot: XmlElement;
    let relationships: Relationships;
    if (owner) {
      owner.relsRoot ??= {
        name: "Relationships",
        attrs: { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" },
        children: [],
        text: "",
      };
      relsRoot = owner.relsRoot;
      relationships = owner.rels;
    } else {
      relsRoot = this.ensureRelsRoot();
      relationships = this.documentRels;
    }
    let maxId = 0;
    for (const r of relsRoot.children) {
      const m = /^rId(\d+)$/.exec(r.attrs["Id"] ?? "");
      if (m) maxId = Math.max(maxId, parseInt(m[1], 10));
    }
    const relId = `rId${maxId + 1}`;
    relsRoot.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target: `media/image${n}.${ext}`,
      },
      children: [],
      text: "",
    });
    relationships.set(relId, { id: relId, type: "image", target: part, external: false });

    // Content type default for the extension
    const MIME: Record<string, string> = {
      png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", svg: "image/svg+xml", webp: "image/webp",
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
    return { relId, part };
  }

  addImageResource(bytes: Uint8Array, ext: string): string {
    return this.addImageResourceAt(bytes, ext).relId;
  }

  /** Add a GLB model part and its Office 2019 model3d relationship. */
  addModel3DResource(bytes: Uint8Array): { relId: string; part: string } {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    while (this.pkg.has(`${docDir}media/model3d${n}.glb`)) n++;
    const part = `${docDir}media/model3d${n}.glb`;
    this.pkg.raw()[part] = bytes;
    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const rel of rels.children) {
      const match = /^rId(\d+)$/.exec(rel.attrs.Id ?? "");
      if (match) maxId = Math.max(maxId, Number(match[1]));
    }
    const relId = `rId${maxId + 1}`;
    rels.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: "http://schemas.microsoft.com/office/2017/06/relationships/model3d",
        Target: `media/model3d${n}.glb`,
      },
      children: [],
      text: "",
    });
    this.documentRels.set(relId, { id: relId, type: "model3d", target: part, external: false });
    if (this.contentTypesRoot && !this.contentTypesRoot.children.some(
      (item) => localName(item.name) === "Default" && item.attrs.Extension?.toLowerCase() === "glb",
    )) {
      this.contentTypesRoot.children.unshift({
        name: "Default",
        attrs: { Extension: "glb", ContentType: "model/gltf-binary" },
        children: [],
        text: "",
      });
    }
    return { relId, part };
  }

  /** Add an OLE package part used by a Word w:object. */
  addEmbeddedObjectResource(bytes: Uint8Array): { relId: string; part: string } {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    while (this.pkg.has(`${docDir}embeddings/oleObject${n}.bin`)) n++;
    const part = `${docDir}embeddings/oleObject${n}.bin`;
    this.pkg.raw()[part] = bytes;
    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const rel of rels.children) {
      const match = /^rId(\d+)$/.exec(rel.attrs.Id ?? "");
      if (match) maxId = Math.max(maxId, Number(match[1]));
    }
    const relId = `rId${maxId + 1}`;
    rels.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject",
        Target: `embeddings/oleObject${n}.bin`,
      },
      children: [],
      text: "",
    });
    this.documentRels.set(relId, { id: relId, type: "oleObject", target: part, external: false });
    const partName = `/${part}`;
    if (this.contentTypesRoot && !this.contentTypesRoot.children.some((item) => item.attrs.PartName === partName)) {
      this.contentTypesRoot.children.push({
        name: "Override",
        attrs: { PartName: partName, ContentType: "application/vnd.openxmlformats-officedocument.oleObject" },
        children: [],
        text: "",
      });
    }
    return { relId, part };
  }

  /** Add a DOCX package embedded as an activatable Word.Document.12 object. */
  addEmbeddedWordDocumentResource(bytes: Uint8Array): { relId: string; part: string } {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    const name = () => `Microsoft_Word_Document${n === 1 ? "" : n}.docx`;
    while (this.pkg.has(`${docDir}embeddings/${name()}`)) n++;
    const part = `${docDir}embeddings/${name()}`;
    this.pkg.raw()[part] = bytes;
    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const rel of rels.children) {
      const match = /^rId(\d+)$/.exec(rel.attrs.Id ?? "");
      if (match) maxId = Math.max(maxId, Number(match[1]));
    }
    const relId = `rId${maxId + 1}`;
    rels.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
        Target: `embeddings/${name()}`,
      },
      children: [],
      text: "",
    });
    this.documentRels.set(relId, { id: relId, type: "package", target: part, external: false });
    if (this.contentTypesRoot && !this.contentTypesRoot.children.some(
      (item) => localName(item.name) === "Default" && item.attrs.Extension?.toLowerCase() === "docx",
    )) {
      this.contentTypesRoot.children.unshift({
        name: "Default",
        attrs: {
          Extension: "docx",
          ContentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        },
        children: [],
        text: "",
      });
    }
    return { relId, part };
  }

  /** Add a native ChartML part and its embedded editable workbook. */
  addChartResource(chartXml: string, workbook: Uint8Array): { relId: string; part: string } {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    while (this.pkg.has(`${docDir}charts/chart${n}.xml`) || this.pkg.has(`${docDir}embeddings/Microsoft_Excel_Worksheet${n}.xlsx`)) n++;
    const part = `${docDir}charts/chart${n}.xml`;
    const workbookPart = `${docDir}embeddings/Microsoft_Excel_Worksheet${n}.xlsx`;
    this.pkg.raw()[part] = strToU8(chartXml);
    this.pkg.raw()[workbookPart] = workbook;
    this.pkg.raw()[relsPathFor(part)] = strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${n}.xlsx"/>` +
      `</Relationships>`,
    );

    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const rel of rels.children) {
      const match = /^rId(\d+)$/.exec(rel.attrs.Id ?? "");
      if (match) maxId = Math.max(maxId, Number(match[1]));
    }
    const relId = `rId${maxId + 1}`;
    rels.children.push({
      name: "Relationship",
      attrs: {
        Id: relId,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        Target: `charts/chart${n}.xml`,
      },
      children: [],
      text: "",
    });
    this.documentRels.set(relId, { id: relId, type: "chart", target: part, external: false });

    const overrides = [
      [`/${part}`, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"],
      [`/${workbookPart}`, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
    ];
    for (const [partName, contentType] of overrides) {
      if (this.contentTypesRoot && !this.contentTypesRoot.children.some((item) => item.attrs.PartName === partName)) {
        this.contentTypesRoot.children.push({
          name: "Override",
          attrs: { PartName: partName, ContentType: contentType },
          children: [],
          text: "",
        });
      }
    }
    return { relId, part };
  }

  /** Add the SmartArt data/layout/style/color parts and its cached diagram drawing. */
  addSmartArtResources(
    layoutXml: string,
    styleXml: string,
    colorsXml: string,
    drawingXml: string,
    dataXml: (drawingRelId: string) => string,
  ): { dataRelId: string; layoutRelId: string; styleRelId: string; colorsRelId: string; drawingRelId: string } {
    const docDir = this.docPart.slice(0, this.docPart.lastIndexOf("/") + 1);
    let n = 1;
    while (
      this.pkg.has(`${docDir}diagrams/data${n}.xml`) ||
      this.pkg.has(`${docDir}diagrams/layout${n}.xml`) ||
      this.pkg.has(`${docDir}diagrams/quickStyle${n}.xml`) ||
      this.pkg.has(`${docDir}diagrams/colors${n}.xml`) ||
      this.pkg.has(`${docDir}diagrams/drawing${n}.xml`)
    ) n++;
    const parts = {
      data: `${docDir}diagrams/data${n}.xml`,
      layout: `${docDir}diagrams/layout${n}.xml`,
      style: `${docDir}diagrams/quickStyle${n}.xml`,
      colors: `${docDir}diagrams/colors${n}.xml`,
      drawing: `${docDir}diagrams/drawing${n}.xml`,
    };
    const rels = this.ensureRelsRoot();
    let maxId = 0;
    for (const rel of rels.children) {
      const match = /^rId(\d+)$/.exec(rel.attrs.Id ?? "");
      if (match) maxId = Math.max(maxId, Number(match[1]));
    }
    const dataRelId = `rId${maxId + 1}`;
    const layoutRelId = `rId${maxId + 2}`;
    const styleRelId = `rId${maxId + 3}`;
    const colorsRelId = `rId${maxId + 4}`;
    const drawingRelId = `rId${maxId + 5}`;
    const related = [
      [dataRelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData", `diagrams/data${n}.xml`, parts.data],
      [layoutRelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout", `diagrams/layout${n}.xml`, parts.layout],
      [styleRelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle", `diagrams/quickStyle${n}.xml`, parts.style],
      [colorsRelId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors", `diagrams/colors${n}.xml`, parts.colors],
      [drawingRelId, "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing", `diagrams/drawing${n}.xml`, parts.drawing],
    ] as const;
    for (const [id, type, target, part] of related) {
      rels.children.push({ name: "Relationship", attrs: { Id: id, Type: type, Target: target }, children: [], text: "" });
      this.documentRels.set(id, { id, type, target: part, external: false });
    }
    this.pkg.raw()[parts.layout] = strToU8(layoutXml);
    this.pkg.raw()[parts.style] = strToU8(styleXml);
    this.pkg.raw()[parts.colors] = strToU8(colorsXml);
    this.pkg.raw()[parts.drawing] = strToU8(drawingXml);
    this.pkg.raw()[parts.data] = strToU8(dataXml(drawingRelId));

    const overrides = [
      [`/${parts.data}`, "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"],
      [`/${parts.layout}`, "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"],
      [`/${parts.style}`, "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"],
      [`/${parts.colors}`, "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"],
      [`/${parts.drawing}`, "application/vnd.ms-office.drawingml.diagramDrawing+xml"],
    ];
    for (const [partName, contentType] of overrides) {
      if (this.contentTypesRoot && !this.contentTypesRoot.children.some((item) => item.attrs.PartName === partName)) {
        this.contentTypesRoot.children.push({ name: "Override", attrs: { PartName: partName, ContentType: contentType }, children: [], text: "" });
      }
    }
    return { dataRelId, layoutRelId, styleRelId, colorsRelId, drawingRelId };
  }

  static load(data: ArrayBuffer | Uint8Array): DocxDocument {
    return new DocxDocument(Package.from(data));
  }

  media(part: string): Uint8Array | undefined {
    return this.pkg.binary(part);
  }

  /**
   * Pending-media registry (plan doc 05 change 1 / doc 16 §6): parts whose
   * XML registration (rels + content-type + extent geometry) exists but
   * whose BYTES have not arrived yet — the out-of-band media model. Layout
   * needs no bytes (extents live in the XML), so a doc with pending parts
   * lays out pixel-identically; the renderer shows a placeholder until
   * `installMedia`. Keyed by part name; the value is the doc-16 §5.3
   * metadata (declared sha + optional E2EE iv/epoch for re-supply).
   */
  readonly pendingMedia = new Map<string, { sha: string; iv?: string; genesisId?: string }>();

  /** DISPLAY-ONLY transfer state for a pending part (doc 16 §5.2 step 4), so
   * the skeleton can say "fetching" versus "nobody online has this yet".
   * Written by the media transfer layer, read by the renderer; never part of
   * the document's identity and never serialized. */
  readonly mediaTransferState = new Map<string, "fetching" | "waiting" | "unavailable">();

  /** Per-part media metadata that PERSISTS after install (doc 16 §5.3):
   * holder duty needs the sha (and E2EE iv/epoch) of READY parts to answer
   * re-supply requests; pendingMedia above only tracks not-yet-arrived. */
  readonly mediaMeta = new Map<string, { sha: string; iv?: string; genesisId?: string }>();

  /** "ready" = bytes present; "pending" = registered, bytes absent (doc 05).
   * Unregistered parts report "pending" too — a rel pointing at a missing
   * zip entry after a save/parse round-trip IS the pending state (a hole is
   * detectable, never corrupt — doc 16 §6 round-trip rule). */
  mediaStatus(part: string): "ready" | "pending" {
    return this.pkg.binary(part) ? "ready" : "pending";
  }

  /**
   * Register an image part + relationship WITHOUT bytes (doc 16 §2: the
   * intent carries the sha; bytes travel out of band). Same deterministic
   * naming/rId scan as addImageResource so every replica applying the same
   * canonical intent derives identical registration.
   */
  registerPendingImage(
    sha: string,
    ext: string,
    meta?: { iv?: string; genesisId?: string },
    source?: XmlElement,
  ): string {
    // Register with a 0-byte placeholder entry so part-name scanning and
    // content-type bookkeeping behave identically to a real image…
    const { relId, part } = this.addImageResourceAt(new Uint8Array(0), ext, source);
    // …then remove the placeholder bytes: the zip entry must be ABSENT for
    // a pending part (doc 16 §6 — absence is the unambiguous hole).
    delete this.pkg.raw()[part];
    this.pendingMedia.set(part, { sha, ...meta });
    this.mediaMeta.set(part, { sha, ...meta });
    return relId;
  }

  /** Install fetched bytes into a pending part (doc 05). The caller has
   * ALREADY verified the sha against the intent's declaration (the
   * reservation is the hash commitment — doc 16 §1.1); this installs and
   * clears the pending record. Returns false for unknown parts. */
  installMedia(part: string, bytes: Uint8Array): boolean {
    if (!this.pendingMedia.has(part) && this.pkg.binary(part)) return false; // already ready
    this.pkg.raw()[part] = bytes;
    this.pendingMedia.delete(part);
    // Arriving bytes change what the page SHOWS without changing the model —
    // and nothing downstream would otherwise notice. Incremental layout
    // reuses its pages when the model version is unchanged, and the repaint
    // differ rebuilds a node only when an ITEM FIELD changes, so the skeleton
    // painted while the bytes were in flight would stay on screen forever
    // (observed in a real browser: pixels installed, placeholder still
    // showing). This bump is what makes "the image finally arrived" a
    // repaintable event; one image arriving is rare, so it costs nothing.
    this._modelVersion++;
    return true;
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
