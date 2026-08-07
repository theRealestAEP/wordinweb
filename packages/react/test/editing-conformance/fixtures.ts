/**
 * Fixture .docx builders for the interactive-editing conformance suite.
 *
 * Each builder returns the bytes of a minimal but VALID .docx whose body puts
 * the caret contexts of the matrix (plain paragraph, empty paragraph, list
 * items, table cells, heading, tracked changes) at known text landmarks, so a
 * test can place the caret with `api.find(...)` + arrow keys — the same
 * gesture path a user takes — without positional clicks (jsdom has no layout).
 */
import { strToU8, zipSync } from "fflate";

const CONTENT_TYPES_BASE =
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`;

const RELS_ROOT =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;

const W = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;

export interface FixtureParts {
  numbering?: boolean;
  styles?: boolean;
}

/** Wrap body XML into a complete one-part-per-need .docx. */
export function docxBytes(bodyXml: string, parts: FixtureParts = {}): Uint8Array {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document ${W}><w:body>${bodyXml}</w:body></w:document>`;
  const files: Record<string, Uint8Array> = {
    "word/document.xml": strToU8(documentXml),
    "_rels/.rels": strToU8(RELS_ROOT),
  };
  let contentTypes = CONTENT_TYPES_BASE;
  const docRels: string[] = [];
  if (parts.numbering) {
    contentTypes += `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;
    docRels.push(
      `<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`,
    );
    files["word/numbering.xml"] = strToU8(NUMBERING_XML);
  }
  if (parts.styles) {
    contentTypes += `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`;
    docRels.push(
      `<Relationship Id="rIdSty" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
    );
    files["word/styles.xml"] = strToU8(STYLES_XML);
  }
  if (docRels.length > 0) {
    files["word/_rels/document.xml.rels"] = strToU8(
      `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${docRels.join("")}</Relationships>`,
    );
  }
  files["[Content_Types].xml"] = strToU8(contentTypes + `</Types>`);
  return zipSync(files);
}

const NUMBERING_XML =
  `<?xml version="1.0"?><w:numbering ${W}>` +
  `<w:abstractNum w:abstractNumId="0">` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
  `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="◦"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>` +
  `</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="1">` +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
  `</w:abstractNum>` +
  `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
  `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>` +
  `</w:numbering>`;

const STYLES_XML =
  `<?xml version="1.0"?><w:styles ${W}>` +
  `<w:style w:type="paragraph" w:styleId="Heading1">` +
  `<w:name w:val="heading 1"/><w:qFormat/>` +
  `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>` +
  `<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>` +
  `</w:style>` +
  `</w:styles>`;

export const p = (text: string, pPr = ""): string =>
  `<w:p>${pPr}${text === "" ? "" : `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`}</w:p>`;

const listPPr = (numId: number, ilvl = 0): string =>
  `<w:pPr><w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>`;

export const listItem = (text: string, opts: { numId?: number; ilvl?: number } = {}): string =>
  p(text, listPPr(opts.numId ?? 1, opts.ilvl ?? 0));

const cell = (paras: string): string =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${paras}</w:tc>`;

/** Three plain paragraphs. Landmarks: "alpha one", "bravo two", "charlie three". */
export const plainDoc = (): Uint8Array =>
  docxBytes(p("alpha one") + p("bravo two") + p("charlie three"));

/** A paragraph, an EMPTY paragraph (no runs), a paragraph. */
export const emptyParaDoc = (): Uint8Array => docxBytes(p("above empty") + p("") + p("below empty"));

/** Bullet list: first/middle/last items, then an EMPTY list item, inside plain
 * paragraphs so list boundaries are visible. */
export const listDoc = (): Uint8Array =>
  docxBytes(
    p("before list") +
      listItem("item first") +
      listItem("item middle") +
      listItem("item last") +
      p("after list"),
    { numbering: true },
  );

/** A list whose LAST item is empty (the undeletable-item shape). */
export const emptyListItemDoc = (): Uint8Array =>
  docxBytes(p("before list") + listItem("item only") + listItem("") + p("after list"), {
    numbering: true,
  });

/** 2x2 table between paragraphs; r1c1 is EMPTY. */
export const tableDoc = (): Uint8Array =>
  docxBytes(
    p("before table") +
      `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
      `<w:tr>${cell(p("cell r0c0"))}${cell(p("cell r0c1"))}</w:tr>` +
      `<w:tr>${cell(p("cell r1c0"))}${cell(p(""))}</w:tr>` +
      `</w:tbl>` +
      p("after table"),
  );

/** A Heading1 paragraph between body paragraphs. */
export const headingDoc = (): Uint8Array =>
  docxBytes(
    p("intro text") +
      p("Title Here", `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`) +
      p("body text"),
    { styles: true },
  );

/** Every context in one document — the fuzzer's arena: plain paragraphs, an
 * empty paragraph, a heading, a bullet list (with an empty item), a 2x2 table
 * (with an empty cell), and tracked changes. */
export const kitchenSinkDoc = (): Uint8Array =>
  docxBytes(
    p("alpha one") +
      p("") +
      p("Title Here", `<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>`) +
      listItem("item first") +
      listItem("item second") +
      listItem("") +
      `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p("cell a")}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p("cell b")}</w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p("cell c")}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${p("")}</w:tc></w:tr></w:tbl>` +
      `<w:p><w:r><w:t xml:space="preserve">kept </w:t></w:r>` +
      `<w:ins w:id="1" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">ins text</w:t></w:r></w:ins>` +
      `<w:del w:id="2" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">del text</w:delText></w:r></w:del>` +
      `<w:r><w:t xml:space="preserve"> tail</w:t></w:r></w:p>` +
      p("omega last"),
    { numbering: true, styles: true },
  );

/** A document that already carries tracked changes (one w:ins, one w:del). */
export const trackedDoc = (): Uint8Array =>
  docxBytes(
    `<w:p><w:r><w:t xml:space="preserve">kept start </w:t></w:r>` +
      `<w:ins w:id="1" w:author="Prior Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">inserted text</w:t></w:r></w:ins>` +
      `<w:r><w:t xml:space="preserve"> mid </w:t></w:r>` +
      `<w:del w:id="2" w:author="Prior Reviewer" w:date="2026-01-01T00:00:00Z"><w:r><w:delText xml:space="preserve">deleted text</w:delText></w:r></w:del>` +
      `<w:r><w:t xml:space="preserve"> kept end</w:t></w:r></w:p>` +
      p("second tracked-doc para"),
  );
