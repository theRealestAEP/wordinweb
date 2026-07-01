import { XmlElement, attr, child, children, childVal, onOff } from "../xml.js";
import { ParaProps, RunProps, Style, Styles } from "../model.js";
import { ParseContext, mergeParaProps, mergeRunProps, parseParaProps, parseRunProps } from "./properties.js";

export function parseStyles(root: XmlElement | undefined, ctx: ParseContext): Styles {
  const styles: Styles = {
    byId: new Map(),
    defaultRPr: {},
    defaultPPr: {},
  };
  if (!root) return styles;

  const docDefaults = child(root, "docDefaults");
  if (docDefaults) {
    const rPrDefault = child(child(docDefaults, "rPrDefault"), "rPr");
    styles.defaultRPr = parseRunProps(rPrDefault, ctx);
    const pPrDefault = child(child(docDefaults, "pPrDefault"), "pPr");
    styles.defaultPPr = parseParaProps(pPrDefault, ctx);
  }
  // Word's implicit defaults when docDefaults omit them (10pt = 13.33px)
  if (styles.defaultRPr.size === undefined) styles.defaultRPr.size = (10 * 4) / 3;
  if (!styles.defaultRPr.font) styles.defaultRPr.font = ctx.theme?.minorFont ?? "Calibri";

  for (const s of children(root, "style")) {
    const id = attr(s, "styleId");
    const type = attr(s, "type") as Style["type"] | undefined;
    if (!id || !type) continue;
    const style: Style = {
      id,
      type,
      name: childVal(s, "name"),
      basedOn: childVal(s, "basedOn"),
      isDefault: attr(s, "default") === "1" || attr(s, "default") === "true",
    };
    const pPr = child(s, "pPr");
    if (pPr) style.pPr = parseParaProps(pPr, ctx);
    const rPr = child(s, "rPr");
    if (rPr) style.rPr = parseRunProps(rPr, ctx);
    styles.byId.set(id, style);
    if (style.isDefault && type === "paragraph") styles.defaultParagraphStyle = id;
  }
  return styles;
}

const MAX_CHAIN = 20;

/** Effective paragraph props from the style chain (no direct formatting). */
export function resolveParagraphStyleChain(styles: Styles, styleId: string | undefined): {
  pPr: ParaProps;
  rPr: RunProps;
} {
  const pChain: Style[] = [];
  let cur = styleId ?? styles.defaultParagraphStyle;
  let guard = 0;
  while (cur && guard++ < MAX_CHAIN) {
    const s = styles.byId.get(cur);
    if (!s) break;
    pChain.unshift(s);
    cur = s.basedOn;
  }
  let pPr: ParaProps = { ...styles.defaultPPr };
  let rPr: RunProps = { ...styles.defaultRPr };
  for (const s of pChain) {
    if (s.pPr) pPr = mergeParaProps(pPr, s.pPr);
    if (s.rPr) rPr = mergeRunProps(rPr, s.rPr);
  }
  return { pPr, rPr };
}

/** Effective run props contributed by a character style chain. */
export function resolveCharacterStyleChain(styles: Styles, styleId: string | undefined): RunProps {
  const chain: Style[] = [];
  let cur = styleId;
  let guard = 0;
  while (cur && guard++ < MAX_CHAIN) {
    const s = styles.byId.get(cur);
    if (!s) break;
    chain.unshift(s);
    cur = s.basedOn;
  }
  let rPr: RunProps = {};
  for (const s of chain) {
    if (s.rPr) rPr = mergeRunProps(rPr, s.rPr);
  }
  return rPr;
}
