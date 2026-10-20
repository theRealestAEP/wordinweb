// What a reader of one of the oldest Word formats (Word for Macintosh 4/5,
// Word for Windows 1/2, Word for DOS, Windows Write) hands the package
// writer (package.ts): the document's text already in Unicode, its
// formatting as the property sets of doc/sprm.ts, which the .doc reader's
// XML writers turn into WordprocessingML.
import type { ParaProps, RunProps, SectionProps, TableProps } from "../doc/sprm.js";

export type Inline =
  | { kind: "text"; text: string }
  | { kind: "tab" }
  | { kind: "break" }
  | { kind: "page" }
  | { kind: "column" }
  | { kind: "noBreakHyphen" }
  | { kind: "softHyphen" }
  /** A field Word computes: PAGE, DATE, TIME (with its picture switch), and the text it showed. */
  | { kind: "field"; code: string; result: string }
  /** A field's marks, around its code and then its result (the text between them in the runs). */
  | { kind: "fieldBegin" }
  | { kind: "fieldSeparate" }
  | { kind: "fieldEnd" }
  /** A picture: an index into OldDocument.pictures, its size in twips. */
  | { kind: "picture"; index: number; width: number; height: number }
  /** A footnote reference in the text: the note's number (1-based, OldDocument.footnotes[id - 1]). */
  | { kind: "footnote"; id: number; customMark?: string }
  /** The note's own number at the start of its text. */
  | { kind: "footnoteMark" };

/**
 * A run: its content and its properties. The properties are direct ones,
 * over the paragraph style's: RunProps.ftc0 indexes OldDocument.fonts, and
 * a toggle is 0 or 1 (never the "as the style" values).
 */
export interface Run {
  props: RunProps;
  content: Inline[];
}

/** A paragraph: ParaProps.istd indexes OldDocument.styles (0 is Normal). */
export interface Paragraph {
  type: "paragraph";
  props: ParaProps;
  /** The paragraph mark's run properties. */
  mark: RunProps;
  runs: Run[];
}

/** A table row: TableProps.left and cells[].right are the cell edges in twips, cells its content. */
export interface Row {
  props: TableProps;
  cells: Block[][];
}

export interface Table {
  type: "table";
  rows: Row[];
}

export type Block = Paragraph | Table;

export type HeaderKind = "header" | "footer";
export type HeaderType = "default" | "first" | "even";

export interface Section {
  props: SectionProps;
  blocks: Block[];
  headers: { kind: HeaderKind; type: HeaderType; blocks: Block[] }[];
}

/** A paragraph style, its properties complete (no basedOn chain). */
export interface Style {
  name: string;
  para: ParaProps;
  run: RunProps;
  /** The style (index into OldDocument.styles) of the paragraph after one in this style. */
  next?: number;
}

export interface Picture {
  data: Uint8Array;
  /** The file extension the renderer knows it by: pict or wmf. */
  ext: "pict" | "wmf";
}

export interface OldDocument {
  fonts: string[];
  /** styles[0] is Normal. */
  styles: Style[];
  sections: Section[];
  footnotes: Block[][];
  pictures: Picture[];
  /** Twips; undefined for Word's 720. */
  defaultTab?: number;
  /** Footnotes at the end of the document rather than the foot of the page. */
  endnotes?: boolean;
  /** Different headers and footers on even pages. */
  evenAndOddHeaders?: boolean;
}
