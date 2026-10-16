import { XmlElement, child } from "../../core/src/xml.js";
import { RevisionKind } from "../../core/src/edit/suggest.js";

/**
 * The summary page at the end of a printable copy: a page break, then one
 * bordered table with the compare's date, the two documents' names, the
 * revision counts by kind, their total, and each distinct note. Every
 * paragraph and run is formatted outright (left, 10 pt Times New Roman,
 * no spacing), so the document's own styles cannot change how it reads.
 */
export interface SummarySpec {
  originalName: string;
  revisedName: string;
  /**
   * The first row's heading. Without one the row reads "Summary report:
   * document comparison done on <date>"; with one it reads "<title>, <date>"
   * (the date as before: en-US, UTC). The caller names the product.
   */
  title?: string;
}

/** The summary as body XML, to go ahead of the body's final sectPr. */
export function summaryXml(spec: SummarySpec, date: string, revisions: Record<RevisionKind, number>, notes: string[], sectPr: XmlElement | undefined): string {
  const width = textWidth(sectPr);
  const first = Math.round(width * 0.74);
  const widths = [first, width - first];
  const when = `${new Date(date).toLocaleString("en-US", { timeZone: "UTC" })} UTC`;
  const total = Object.values(revisions).reduce((sum, n) => sum + n, 0);
  const rows: string[] = [
    row([[spec.title ? `${spec.title}, ${when}` : `Summary report: document comparison done on ${when}`, true]]),
    row([["Original:", true], [spec.originalName, false]]),
    row([["Modified:", true], [spec.revisedName, false]]),
    row([["Add", false], [String(revisions.insertion + revisions.markInsertion), false]]),
    row([["Delete", false], [String(revisions.deletion + revisions.markDeletion), false]]),
    row([["Table Insert", false], [String(revisions.rowInsertion), false]]),
    row([["Table Delete", false], [String(revisions.rowDeletion), false]]),
    row([["Format changes", false], [String(revisions.runFormat + revisions.paragraphFormat + revisions.tableFormat + revisions.rowFormat + revisions.cellFormat + revisions.sectionFormat), false]]),
    row([["Total Changes:", true], [String(total), true]]),
    ...[...new Set(notes)].map((note) => row([[`Note: ${note}`, false]])),
  ];
  const borders = ["top", "left", "bottom", "right", "insideH", "insideV"].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("");
  return (
    `<w:p>${PPR}<w:r><w:br w:type="page"/></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblBorders>${borders}</w:tblBorders><w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
    `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>${rows.join("")}</w:tbl>`
  );

  /** A row of one cell spanning both columns, or one cell per column. */
  function row(cells: [string, boolean][]): string {
    const tcs = cells.map(([text, bold], i) => {
      const w = cells.length === 1 ? width : widths[i];
      const span = cells.length === 1 ? `<w:gridSpan w:val="2"/>` : "";
      return `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span}</w:tcPr><w:p>${PPR}<w:r>${rPr(bold)}<w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p></w:tc>`;
    });
    return `<w:tr>${tcs.join("")}</w:tr>`;
  }
}

/** Paragraph properties that override whatever the Normal style says. */
const PPR = `<w:pPr><w:keepNext w:val="0"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr>`;

function rPr(bold: boolean): string {
  return `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>${bold ? "<w:b/><w:bCs/>" : '<w:b w:val="0"/><w:bCs w:val="0"/>'}<w:i w:val="0"/><w:iCs w:val="0"/><w:color w:val="000000"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:highlight w:val="none"/><w:u w:val="none"/></w:rPr>`;
}

/** The last section's text width in twips: its page width less its side margins (Letter, 1" margins when unset). */
function textWidth(sectPr: XmlElement | undefined): number {
  const twips = (el: XmlElement | undefined, attr: string, fallback: number): number => {
    const n = Number(el?.attrs[attr]);
    return Number.isFinite(n) && n > 0 ? n : fallback;
  };
  const page = twips(child(sectPr, "pgSz"), "w:w", 12240);
  const margins = child(sectPr, "pgMar");
  return Math.max(1440, page - twips(margins, "w:left", 1440) - twips(margins, "w:right", 1440));
}

function escape(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
