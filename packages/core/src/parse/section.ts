import { XmlElement, attr, child, children, intAttr, localName, onOff } from "../xml.js";
import { ColumnSpec, HeaderFooterRefs, SectionProps } from "../model.js";
import { twipsToPx } from "../units.js";

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
    }
    props.columns = spec;
  }

  const type = attr(child(sectPr, "type"), "val");
  if (type === "continuous" || type === "nextPage" || type === "evenPage" || type === "oddPage" || type === "nextColumn") {
    props.type = type;
  }

  const vAlign = attr(child(sectPr, "vAlign"), "val");
  if (vAlign === "center" || vAlign === "both" || vAlign === "bottom" || vAlign === "top") {
    props.vAlign = vAlign;
  }

  return props;
}
