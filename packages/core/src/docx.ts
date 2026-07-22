import { Package } from "./zip.js";
import { XmlElement, parseXml, serializeXml, child, children, intAttr, onOff, attr, localName, cyrb53 } from "./xml.js";
import { strToU8, zipSync } from "fflate";
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
import { parseBody, parseBlocks, parseParagraph, DocParseContext } from "./parse/document.js";
import { parseNotesPart } from "./parse/notes.js";
import { Relationships, parseRelationships, relsPathFor } from "./parse/rels.js";
import { mergeParaProps, mergeRunProps } from "./parse/properties.js";
import { extractOlePackage } from "./parse/ole.js";
import {
  buildSmartArtColorsXml,
  buildSmartArtDataXml,
  buildSmartArtDrawingXml,
  buildSmartArtLayoutXml,
  buildSmartArtStyleXml,
} from "./edit/smartart.js";

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
  /** Note content by note id (render-only; sources stripped). */
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
  /** Retained settings.xml tree. A synthetic empty root keeps history root
   * indices stable for documents that did not originally contain the part. */
  private readonly settingsPart: string;
  private readonly settingsRoot: XmlElement;
  private settingsDirty = false;
  /** Parsed document.xml root (read-only outside the class; the layout engine
   * scans it for incremental-reuse eligibility, tests walk it). */
  readonly docRoot: XmlElement;
  private readonly hfParts: { relId: string; target: string; root: XmlElement; isHeader: boolean; rels: Relationships }[] = [];
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
  private nextDocPrId = 1000;

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

    if (settings) {
      this.evenAndOddHeaders = onOff(child(settings, "evenAndOddHeaders")) ?? false;
      (this as { mirrorMargins: boolean }).mirrorMargins = onOff(child(settings, "mirrorMargins")) ?? false;
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
      this.rememberOriginalXml(rel.target, root);
      this.hydrateCorePropertyControls(root, coreProperties);
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
        const separator = root.children.find(
          (item) => localName(item.name) === "footnote" && attr(item, "type") === "separator",
        );
        if (separator) {
          this.footnoteSeparator.push(...parseBlocks(separator, { ...this.ctxBase, rels: partRels }));
        }
      }
      // Footnotes are editable (sources kept, part retained + re-serialized on
      // save); endnotes stay render-only for now.
      const notes = parseNotesPart(root, { ...this.ctxBase, rels: partRels }, isFn);
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
    // Re-derive footnote blocks from the retained tree (editable: keep source
    // refs) so an edit to a footnote's w:t re-measures. Endnotes are render-
    // only, parsed once at load, and left untouched here.
    if (this.footnotesRoot) {
      this.footnotes.clear();
      const notes = parseNotesPart(this.footnotesRoot, { ...this.ctxBase, rels: this.footnotesRels }, true);
      for (const [id, blocks] of notes) this.footnotes.set(id, blocks);
    }
    this._modelVersion++;
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

    const findBlockList = (blocks: Block[]): { blocks: Block[]; index: number } | null => {
      const index = blocks.findIndex((block) => block.type === "paragraph" && block.src === referenceSource);
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

    const findBlockList = (blocks: Block[]): { blocks: Block[]; index: number } | null => {
      const index = blocks.findIndex((block) => block.type === "paragraph" && block.src === beforeSource);
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
    const findBlockList = (blocks: Block[]): { blocks: Block[]; index: number } | null => {
      const index = blocks.findIndex((block) => block.type === "paragraph" && block.src === source);
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
    const old = location.blocks[location.index];
    if (old.type !== "paragraph" || old.sectionBreak) return null;

    const unsafe = (element: XmlElement): boolean => {
      const name = localName(element.name);
      // Bookmark ranges retain parsed Run identities in refBookmarks. A local
      // paragraph replacement would leave those references stale; other
      // fields, controls, and revisions are safe to parse one-for-one.
      if (name === "sectPr" || name === "bookmarkStart" || name === "bookmarkEnd") return true;
      return element.children.some(unsafe);
    };
    if (unsafe(source)) return null;

    const paragraph = parseParagraph(source, {
      ...this.ctxBase,
      rels: this.documentRels,
      readPart: (part: string) => this.readXmlOptional(part),
      independentTextboxStories: true,
    });
    if (paragraph.revisionHidden || paragraph.sectionBreak) return null;
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
    this.numbering = parseNumbering(this.numberingRoot ?? undefined, this.ctxBase);
    this._layoutGlobalSig = null;
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
      const notes = parseNotesPart(this.footnotesRoot, { ...this.ctxBase, rels: this.footnotesRels }, true);
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

  /** Flag the footnotes part dirty when `t` lives inside it, so save()
   * re-serializes footnotes.xml. Called by the editor after a text edit; a
   * no-op for body/header/footer targets. */
  markDirtyIfFootnote(t: XmlElement): void {
    if (!this.footnotesRoot || this.footnotesDirty) return;
    const contains = (el: XmlElement): boolean => {
      if (el === t) return true;
      for (const c of el.children) if (contains(c)) return true;
      return false;
    };
    if (contains(this.footnotesRoot)) this.footnotesDirty = true;
  }

  /** The mutable XML roots (document body, related modeled parts, settings).
   * settingsRoot is always second and always present so its history snapshot
   * index stays stable even when optional related roots are created later. */
  editableRoots(): XmlElement[] {
    const roots = [this.docRoot, this.settingsRoot, ...this.hfParts.map((p) => p.root)];
    if (this.footnotesRoot) roots.push(this.footnotesRoot);
    if (this.commentsRoot) roots.push(this.commentsRoot);
    if (this.commentsExtRoot) roots.push(this.commentsExtRoot);
    return roots;
  }

  /** Toggle the document-global facing-page margin mode in settings.xml. */
  setMirrorMargins(enabled: boolean): void {
    this.settingsRoot.children = this.settingsRoot.children.filter((c) => localName(c.name) !== "mirrorMargins");
    if (enabled) {
      // CT_Settings is a sequence. mirrorMargins follows saveFormsData and
      // precedes alignBordersAndEdges/proofState/defaultTabStop/compat. Insert
      // before the first child that is not one of its schema predecessors so
      // Word never has to repair settings.xml.
      const beforeMirror = new Set([
        "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
        "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
        "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
        "embedSystemFonts", "saveSubsetFonts", "saveFormsData",
      ]);
      const index = this.settingsRoot.children.findIndex((c) => !beforeMirror.has(localName(c.name)));
      const mirror = { name: "w:mirrorMargins", attrs: {}, children: [], text: "" };
      this.settingsRoot.children.splice(index === -1 ? this.settingsRoot.children.length : index, 0, mirror);
    }

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
    this.settingsDirty = true;
    (this as { mirrorMargins: boolean }).mirrorMargins = enabled;
  }

  /**
   * Find the parent element of `target` in any modeled XML tree (document
   * body, headers, footers). Linear scan — documents are small and this only
   * runs on structural edits (Enter, paragraph merge).
   */
  /** XML roots that can carry tracked changes: body, headers/footers, footnotes. */
  revisionRoots(): XmlElement[] {
    const roots = [this.docRoot, ...this.hfParts.map((p) => p.root)];
    if (this.footnotesRoot) roots.push(this.footnotesRoot);
    return roots;
  }

  findParentOf(target: XmlElement): XmlElement | undefined {
    const roots = [this.docRoot, ...this.hfParts.map((p) => p.root)];
    if (this.footnotesRoot) roots.push(this.footnotesRoot);
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
  private normalizePercentageTableGrids(): void {
    const setAttr = (element: XmlElement, name: string, value: string): void => {
      const key = Object.keys(element.attrs).find((item) => localName(item) === name);
      element.attrs[key ?? `${element.name.includes(":") ? element.name.split(":")[0] + ":" : ""}${name}`] = value;
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
    this.normalizePercentageTableGrids();
    const files: Record<string, Uint8Array> = { ...this.pkg.raw() };
    if (files["docProps/custom.xml"] && this.contentTypesRoot && !this.contentTypesRoot.children.some(
      (item) => localName(item.name) === "Override" && item.attrs.PartName === "/docProps/custom.xml",
    )) {
      const prefixEnd = this.contentTypesRoot.name.indexOf(":") + 1;
      const prefix = prefixEnd > 0 ? this.contentTypesRoot.name.slice(0, prefixEnd) : "";
      this.contentTypesRoot.children.push({
        name: `${prefix}Override`,
        attrs: {
          PartName: "/docProps/custom.xml",
          ContentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        },
        children: [],
        text: "",
      });
    }
    this.writeModeledXml(files, this.docPart, this.docRoot);
    for (const part of this.hfParts) {
      this.writeModeledXml(files, part.target, part.root);
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
    if (this.settingsDirty) files[this.settingsPart] = strToU8(serializeXml(this.settingsRoot, true));
    if (this.relsRoot) this.writeModeledXml(files, this.relsPath, this.relsRoot);
    if (this.contentTypesRoot) this.writeModeledXml(files, "[Content_Types].xml", this.contentTypesRoot);
    return zipSync(files);
  }

  /** Fresh unique docPr id for inserted drawings. */
  nextDrawingId(): number {
    return this.nextDocPrId++;
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
    return relId;
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
