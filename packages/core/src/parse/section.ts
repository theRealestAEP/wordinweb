import { XmlElement, attr, child, children, intAttr, localName, onOff } from "../xml.js";
import { Border, ColumnSpec, HeaderFooterRefs, SectionProps } from "../model.js";
import { ptToPx, twipsToPx } from "../units.js";

/** US Letter portrait with 1" margins — Word's fallback geometry. */
export function defaultSectionProps(): SectionProps {
  return {
    pageWidth: twipsToPx(12240),
    pageHeight: twipsToPx(15840),
    marginTop: twipsToPx(1440),
    marginRight: twipsToPx(1440),
    marginBottom: twipsToPx(1440),
    marginLeft: twipsToPx(1440),
    headerDistance: twipsToPx(720),
    footerDistance: twipsToPx(720),
    gutter: 0,
    headerRefs: {},
    footerRefs: {},
    titlePage: false,
    columns: { count: 1, space: twipsToPx(720) },
  };
}

export function parseSectionProps(sectPr: XmlElement | undefined): SectionProps {
  const props = defaultSectionProps();
  if (!sectPr) return props;

  const pgSz = child(sectPr, "pgSz");
  if (pgSz) {
    const w = intAttr(pgSz, "w");
    const h = intAttr(pgSz, "h");
    if (w) props.pageWidth = twipsToPx(w);
    if (h) props.pageHeight = twipsToPx(h);
    if (attr(pgSz, "orient") === "landscape" && props.pageHeight > props.pageWidth) {
      // Some producers only set orient; swap if inconsistent.
      [props.pageWidth, props.pageHeight] = [props.pageHeight, props.pageWidth];
    }
  }

  const pgMar = child(sectPr, "pgMar");
  if (pgMar) {
    const top = intAttr(pgMar, "top");
    const right = intAttr(pgMar, "right");
    const bottom = intAttr(pgMar, "bottom");
    const left = intAttr(pgMar, "left");
    const header = intAttr(pgMar, "header");
    const footer = intAttr(pgMar, "footer");
    const gutter = intAttr(pgMar, "gutter");
    // top/bottom can be negative (fixed margin regardless of header size)
    if (top !== undefined) props.marginTop = twipsToPx(top);
    if (right !== undefined) props.marginRight = twipsToPx(right);
    if (bottom !== undefined) props.marginBottom = twipsToPx(bottom);
    if (left !== undefined) props.marginLeft = twipsToPx(left);
    if (header !== undefined) props.headerDistance = twipsToPx(header);
    if (footer !== undefined) props.footerDistance = twipsToPx(footer);
    if (gutter !== undefined) props.gutter = twipsToPx(gutter);
  }

  for (const ref of sectPr.children) {
    const ln = localName(ref.name);
    if (ln !== "headerReference" && ln !== "footerReference") continue;
    const type = attr(ref, "type") ?? "default";
    const rid = attr(ref, "id");
    if (!rid) continue;
    const target: HeaderFooterRefs = ln === "headerReference" ? props.headerRefs : props.footerRefs;
    if (type === "default" || type === "first" || type === "even") target[type] = rid;
  }

  props.titlePage = onOff(child(sectPr, "titlePg")) ?? false;

  const pgNumType = child(sectPr, "pgNumType");
  if (pgNumType) {
    const start = intAttr(pgNumType, "start");
    if (start !== undefined) props.pageNumberStart = start;
    const fmt = attr(pgNumType, "fmt");
    if (fmt) props.pageNumberFormat = fmt;
  }

  const cols = child(sectPr, "cols");
  if (cols) {
    const num = intAttr(cols, "num") ?? 1;
    const space = intAttr(cols, "space");
    const spec: ColumnSpec = {
      count: Math.max(1, num),
      space: space !== undefined ? twipsToPx(space) : twipsToPx(720),
    };
    const colEls = children(cols, "col");
    if (colEls.length > 0) {
      spec.count = colEls.length;
      spec.widths = colEls.map((c) => twipsToPx(intAttr(c, "w") ?? 0));
      // Each w:col carries its OWN trailing space; the w:cols-level space is
      // only the fallback (probe3-columns-unequal: w:col space=360tw entries
      // under a w:cols with no space attribute — Word separates the columns
      // by 18pt, not the 36pt default).
      spec.spaces = colEls.map((c) => {
        const s = intAttr(c, "space");
        return s !== undefined ? twipsToPx(s) : spec.space;
      });
    }
    const sep = attr(cols, "sep");
    if (sep === "1" || sep === "true" || sep === "on") spec.sep = true;
    props.columns = spec;
  }

  // w:docGrid type=lines/linesAndChars: Word snaps each line's single-line
  // font height up to linePitch (twips) before applying the line-spacing
  // multiplier. type=default/snapToChars don't affect vertical line pitch.
  const docGrid = child(sectPr, "docGrid");
  if (docGrid) {
    const gtype = attr(docGrid, "type") ?? "default";
    if (gtype === "default" || gtype === "lines" || gtype === "linesAndChars" || gtype === "snapToChars") {
      props.docGridType = gtype;
    }
    const linePitch = intAttr(docGrid, "linePitch");
    if ((gtype === "lines" || gtype === "linesAndChars" || gtype === "snapToChars") && linePitch) {
      props.docGridLinePitch = twipsToPx(linePitch);
    }
    // Word's real-world "charsAndLines" grid (both a character and a line grid).
    // In compat 15 it does NOT snap line height up to linePitch - each East
    // Asian line keeps its font's natural pitch (probe3-chargrid: MS Mincho
    // 15.4pt, Chinese fallback 20.5pt, both under the 18pt linePitch). We do not
    // set docGridLinePitch (no snap); instead flag the section so line
    // measurement uses the glyphs' true grid line height rather than the tall
    // macOS substitute box.
    if (gtype === "charsAndLines") {
      props.docGridCharGrid = true;
    }
  }

  const pgBorders = child(sectPr, "pgBorders");
  if (pgBorders) {
    const side = (name: string): Border | undefined => {
      const el = child(pgBorders, name);
      if (!el) return undefined;
      const val = attr(el, "val") ?? "single";
      if (val === "none" || val === "nil") return undefined;
      const sz = intAttr(el, "sz") ?? 4; // eighth-points
      const colorAttr = attr(el, "color");
      return {
        style: val === "double" ? "double" : val === "dashed" ? "dashed" : val === "dotted" ? "dotted" : "single",
        width: Math.max((sz / 8) * (4 / 3), 0.75),
        color: colorAttr && colorAttr !== "auto" ? "#" + colorAttr : "#000000",
        space: ptToPx(intAttr(el, "space") ?? 0),
      };
    };
    const borders = { top: side("top"), bottom: side("bottom"), left: side("left"), right: side("right") };
    if (borders.top || borders.bottom || borders.left || borders.right) {
      props.pageBorders = {
        ...borders,
        offsetFrom: attr(pgBorders, "offsetFrom") === "page" ? "page" : "text",
      };
    }
  }

  const lnNum = child(sectPr, "lnNumType");
  if (lnNum) {
    const restart = attr(lnNum, "restart");
    props.lineNumbering = {
      countBy: intAttr(lnNum, "countBy") ?? 1,
      start: intAttr(lnNum, "start") ?? 1,
      // w:distance is in twips; default ~0.25in when absent.
      distance: twipsToPx(intAttr(lnNum, "distance") ?? 360),
      restart: restart === "continuous" || restart === "newSection" ? restart : "newPage",
    };
  }

  const type = attr(child(sectPr, "type"), "val");
  if (type === "continuous" || type === "nextPage" || type === "evenPage" || type === "oddPage" || type === "nextColumn") {
    props.type = type;
  }

  const vAlign = attr(child(sectPr, "vAlign"), "val");
  if (vAlign === "center" || vAlign === "both" || vAlign === "bottom" || vAlign === "top") {
    props.vAlign = vAlign;
  }

  const textDir = attr(child(sectPr, "textDirection"), "val");
  if (textDir === "tbRl" || textDir === "tbRlV") props.textDirection = "tbRl";

  const fnPr = child(sectPr, "footnotePr");
  if (fnPr) {
    const fmt = attr(child(fnPr, "numFmt"), "val");
    if (fmt) props.footnoteNumFmt = fmt;
    const start = intAttr(child(fnPr, "numStart"), "val");
    if (start !== undefined) props.footnoteNumStart = start;
  }
  const enPr = child(sectPr, "endnotePr");
  if (enPr) {
    const fmt = attr(child(enPr, "numFmt"), "val");
    if (fmt) props.endnoteNumFmt = fmt;
    const start = intAttr(child(enPr, "numStart"), "val");
    if (start !== undefined) props.endnoteNumStart = start;
  }

  return props;
}
