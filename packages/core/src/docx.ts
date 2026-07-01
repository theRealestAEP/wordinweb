import { Package } from "./zip.js";
import { parseXml, child, intAttr, onOff } from "./xml.js";
import { twipsToPx } from "./units.js";
import {
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
  readonly sections: Section[];
  /** Header/footer parts keyed by relationship id from document.xml.rels. */
  readonly headers: Map<string, HeaderFooter> = new Map();
  readonly footers: Map<string, HeaderFooter> = new Map();
  readonly documentRels: Relationships;
  /** settings.xml w:evenAndOddHeaders — enables the "even" header/footer variants. */
  readonly evenAndOddHeaders: boolean = false;
  /** settings.xml w:defaultTabStop in px (Word default 0.5"). */
  readonly defaultTabStop: number = 48;

  private constructor(pkg: Package) {
    this.pkg = pkg;

    const docPart = this.findDocumentPart();
    const docDir = docPart.slice(0, docPart.lastIndexOf("/") + 1);

    const themeXml = this.readXmlOptional(docDir + "theme/theme1.xml");
    this.theme = parseTheme(themeXml);
    const ctxBase = { theme: this.theme };

    this.styles = parseStyles(this.readXmlOptional(docDir + "styles.xml"), ctxBase);
    this.numbering = parseNumbering(this.readXmlOptional(docDir + "numbering.xml"), ctxBase);

    this.documentRels = parseRelationships(
      this.readXmlOptional(relsPathFor(docPart)),
      docPart,
    );

    const docRoot = this.readXmlOptional(docPart);
    if (!docRoot) throw new Error(`Missing ${docPart} in package`);
    const body = child(docRoot, "body");
    if (!body) throw new Error("document.xml has no w:body");

    const settings = this.readXmlOptional(docDir + "settings.xml");
    if (settings) {
      this.evenAndOddHeaders = onOff(child(settings, "evenAndOddHeaders")) ?? false;
      const tabStop = intAttr(child(settings, "defaultTabStop"), "val");
      if (tabStop !== undefined && tabStop > 0) this.defaultTabStop = twipsToPx(tabStop);
    }

    const ctx: DocParseContext = { ...ctxBase, rels: this.documentRels };
    this.sections = parseBody(body, ctx);

    // Load header/footer parts referenced from the document rels.
    for (const rel of this.documentRels.values()) {
      const isHeader = rel.type.endsWith("/header");
      const isFooter = rel.type.endsWith("/footer");
      if (!isHeader && !isFooter) continue;
      const root = this.readXmlOptional(rel.target);
      if (!root) continue;
      const partRels = parseRelationships(this.readXmlOptional(relsPathFor(rel.target)), rel.target);
      const partCtx: DocParseContext = { ...ctxBase, rels: partRels };
      const hf: HeaderFooter = { blocks: parseBlocks(root, partCtx) };
      (isHeader ? this.headers : this.footers).set(rel.id, hf);
    }
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
