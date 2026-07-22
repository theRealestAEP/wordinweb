import { strToU8 } from "fflate";
import type { DocxDocument } from "../docx.js";
import type { SmartArtData } from "../model.js";
import { type XmlElement, attr, localName, parseXml, serializeXml } from "../xml.js";
import { parseRelationships, relsPathFor } from "../parse/rels.js";

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const NS_DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WIDTH_EMU = 480 * EMU_PER_PX;
const HEIGHT_EMU = 240 * EMU_PER_PX;
const COLORS = ["4472C4", "ED7D31", "70AD47", "5B9BD5", "A5A5A5", "FFC000"];

interface DiagramShape {
  modelId: string;
  x: number;
  y: number;
  width: number;
  height: number;
  text?: string;
  geometry?: "roundRect" | "ellipse" | "line";
  color?: string;
}

export interface SmartArtTextFormat {
  fontFamily: string;
  fontSizePt: number;
  color: string;
  bold: boolean;
  italic: boolean;
  alignment: "left" | "center" | "right";
}

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function descendant(node: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!node) return undefined;
  if (localName(node.name) === name) return node;
  for (const item of node.children) {
    const found = descendant(item, name);
    if (found) return found;
  }
  return undefined;
}

function descendants(node: XmlElement, name: string, out: XmlElement[] = []): XmlElement[] {
  if (localName(node.name) === name) out.push(node);
  for (const child of node.children) descendants(child, name, out);
  return out;
}

function smartArtDrawingPart(doc: DocxDocument, drawing: XmlElement): string | null {
  const relIds = descendant(drawing, "relIds");
  const dataRel = relIds ? doc.documentRels.get(attr(relIds, "dm") ?? "") : undefined;
  if (!dataRel || dataRel.external) return null;
  const dataXml = doc.pkg.text(dataRel.target);
  if (!dataXml) return null;
  const drawingRelId = attr(descendant(parseXml(dataXml), "dataModelExt"), "relId");
  if (drawingRelId) {
    const hostRel = doc.documentRels.get(drawingRelId);
    if (hostRel && !hostRel.external) return hostRel.target;
    const relsXml = doc.pkg.text(relsPathFor(dataRel.target));
    if (relsXml) {
      const dataRelEntry = parseRelationships(parseXml(relsXml), dataRel.target).get(drawingRelId);
      if (dataRelEntry && !dataRelEntry.external) return dataRelEntry.target;
    }
  }
  for (const rel of doc.documentRels.values()) {
    if (!rel.external && rel.type.endsWith("/diagramDrawing")) return rel.target;
  }
  return null;
}

function smartArtNodeShapes(root: XmlElement): XmlElement[] {
  return descendants(root, "sp").filter((shape) => {
    const spPr = shape.children.find((child) => localName(child.name) === "spPr");
    if (!spPr) return false;
    const name = attr(descendant(shape, "cNvPr"), "name") ?? "";
    const geometry = (attr(descendant(spPr, "prstGeom"), "prst") ?? "").toLowerCase();
    return !/^connector\b/i.test(name) && geometry !== "line" && !geometry.includes("connector");
  });
}

function smartArtNodeShapeProperties(root: XmlElement): XmlElement[] {
  return smartArtNodeShapes(root).flatMap((shape) => {
    const spPr = shape.children.find((child) => localName(child.name) === "spPr");
    return spPr ? [spPr] : [];
  });
}

function replaceTextLeaves(root: XmlElement, value: string): boolean {
  const leaves = descendants(root, "t").filter((node) => node.children.length === 0);
  if (!leaves.length) return false;
  leaves[0].text = value;
  for (const leaf of leaves.slice(1)) leaf.text = "";
  return true;
}

function directChild(node: XmlElement, name: string): XmlElement | undefined {
  return node.children.find((child) => localName(child.name) === name);
}

function textColor(rPr: XmlElement | undefined): string {
  const color = attr(descendant(directChild(rPr ?? el("a:rPr"), "solidFill"), "srgbClr"), "val")
    ?? attr(descendant(directChild(rPr ?? el("a:rPr"), "solidFill"), "sysClr"), "lastClr");
  return `#${(color ?? "000000").toUpperCase()}`;
}

function applyTextFormat(root: XmlElement, format: SmartArtTextFormat): void {
  for (const run of descendants(root, "r")) {
    let rPr = directChild(run, "rPr");
    if (!rPr) {
      rPr = el("a:rPr");
      run.children.unshift(rPr);
    }
    rPr.attrs.sz = String(Math.round(format.fontSizePt * 100));
    rPr.attrs.b = format.bold ? "1" : "0";
    rPr.attrs.i = format.italic ? "1" : "0";
    for (const name of ["latin", "ea", "cs"] as const) {
      let font = directChild(rPr, name);
      if (!font) {
        font = el(`a:${name}`);
        rPr.children.push(font);
      }
      font.attrs.typeface = format.fontFamily;
    }
    rPr.children = rPr.children.filter((child) => {
      const name = localName(child.name);
      return !["solidFill", "noFill", "gradFill", "pattFill"].includes(name);
    });
    rPr.children.unshift(el("a:solidFill", {}, [el("a:srgbClr", { val: format.color.replace(/^#/, "").toUpperCase() })]));
  }
  for (const paragraph of descendants(root, "p")) {
    let pPr = directChild(paragraph, "pPr");
    if (!pPr) {
      pPr = el("a:pPr");
      paragraph.children.unshift(pPr);
    }
    pPr.attrs.algn = format.alignment === "center" ? "ctr" : format.alignment === "right" ? "r" : "l";
  }
}

/** Read one visible node fill from a SmartArt cached drawing. */
export function smartArtFillColor(doc: DocxDocument, drawing: XmlElement, nodeIndex = 0): string | null {
  const part = smartArtDrawingPart(doc, drawing);
  const xml = part ? doc.pkg.text(part) : undefined;
  if (!xml) return null;
  const spPr = smartArtNodeShapeProperties(parseXml(xml))[nodeIndex];
  if (!spPr || spPr.children.some((child) => localName(child.name) === "noFill")) return null;
  const solidFill = spPr.children.find((child) => localName(child.name) === "solidFill");
  if (!solidFill) return null;
  const color = attr(descendant(solidFill, "srgbClr"), "val")
    ?? attr(descendant(solidFill, "sysClr"), "lastClr");
  return `#${(color ?? COLORS[0]).toUpperCase()}`;
}

/** Apply a fill to one visible SmartArt node, or every node when omitted. */
export function setSmartArtFill(doc: DocxDocument, drawing: XmlElement, color: string | null, nodeIndex?: number): boolean {
  const part = smartArtDrawingPart(doc, drawing);
  const xml = part ? doc.pkg.text(part) : undefined;
  if (!part || !xml) return false;
  const root = parseXml(xml);
  const allNodes = smartArtNodeShapeProperties(root);
  const nodes = nodeIndex === undefined ? allNodes : allNodes.slice(nodeIndex, nodeIndex + 1);
  if (!nodes.length) return false;
  for (const spPr of nodes) {
    spPr.children = spPr.children.filter((child) => {
      const name = localName(child.name);
      return !["solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill"].includes(name);
    });
    const fill = color
      ? el("a:solidFill", {}, [el("a:srgbClr", { val: color.replace(/^#/, "").toUpperCase() })])
      : el("a:noFill");
    const lineIndex = spPr.children.findIndex((child) => localName(child.name) === "ln");
    spPr.children.splice(lineIndex === -1 ? spPr.children.length : lineIndex, 0, fill);
  }
  doc.pkg.raw()[part] = strToU8(serializeXml(root, true));
  doc.markPackageResourceChanged();
  doc.refresh();
  return true;
}

/** Replace one visible SmartArt node's text in both its editable data model and cached drawing. */
export function setSmartArtNodeText(doc: DocxDocument, drawing: XmlElement, nodeIndex: number, value: string): boolean {
  const relIds = descendant(drawing, "relIds");
  const dataRel = relIds ? doc.documentRels.get(attr(relIds, "dm") ?? "") : undefined;
  const drawingPart = smartArtDrawingPart(doc, drawing);
  const dataXml = dataRel && !dataRel.external ? doc.pkg.text(dataRel.target) : undefined;
  const drawingXml = drawingPart ? doc.pkg.text(drawingPart) : undefined;
  if (!dataRel || dataRel.external || !drawingPart || !dataXml || !drawingXml) return false;

  const drawingRoot = parseXml(drawingXml);
  const shape = smartArtNodeShapes(drawingRoot)[nodeIndex];
  const modelId = shape ? attr(shape, "modelId") : undefined;
  if (!shape || !modelId) return false;

  const dataRoot = parseXml(dataXml);
  const point = descendants(dataRoot, "pt").find((candidate) => attr(candidate, "modelId") === modelId);
  if (!point || !replaceTextLeaves(shape, value) || !replaceTextLeaves(point, value)) return false;

  doc.pkg.raw()[drawingPart] = strToU8(serializeXml(drawingRoot, true));
  doc.pkg.raw()[dataRel.target] = strToU8(serializeXml(dataRoot, true));
  doc.markPackageResourceChanged();
  doc.refresh();
  return true;
}

/** Read the visible text formatting for a SmartArt node. */
export function smartArtTextFormat(doc: DocxDocument, drawing: XmlElement, nodeIndex = 0): SmartArtTextFormat | null {
  const part = smartArtDrawingPart(doc, drawing);
  const xml = part ? doc.pkg.text(part) : undefined;
  if (!xml) return null;
  const shape = smartArtNodeShapes(parseXml(xml))[nodeIndex];
  const rPr = shape ? descendant(shape, "rPr") : undefined;
  if (!shape || !rPr) return null;
  const alignment = attr(descendant(shape, "pPr"), "algn");
  return {
    fontFamily: attr(descendant(rPr, "latin"), "typeface") ?? "Calibri",
    fontSizePt: Number(attr(rPr, "sz") ?? "1200") / 100,
    color: textColor(rPr),
    bold: attr(rPr, "b") === "1",
    italic: attr(rPr, "i") === "1",
    alignment: alignment === "r" ? "right" : alignment === "ctr" ? "center" : "left",
  };
}

/** Apply text formatting to one SmartArt node, or every node when omitted. */
export function setSmartArtTextFormat(
  doc: DocxDocument,
  drawing: XmlElement,
  format: SmartArtTextFormat,
  nodeIndex?: number,
): boolean {
  const relIds = descendant(drawing, "relIds");
  const dataRel = relIds ? doc.documentRels.get(attr(relIds, "dm") ?? "") : undefined;
  const drawingPart = smartArtDrawingPart(doc, drawing);
  const dataXml = dataRel && !dataRel.external ? doc.pkg.text(dataRel.target) : undefined;
  const drawingXml = drawingPart ? doc.pkg.text(drawingPart) : undefined;
  if (!dataRel || dataRel.external || !drawingPart || !dataXml || !drawingXml) return false;

  const drawingRoot = parseXml(drawingXml);
  const allShapes = smartArtNodeShapes(drawingRoot);
  const shapes = nodeIndex === undefined ? allShapes : allShapes.slice(nodeIndex, nodeIndex + 1);
  if (!shapes.length) return false;
  const modelIds = new Set(shapes.map((shape) => attr(shape, "modelId")).filter((value): value is string => !!value));
  const dataRoot = parseXml(dataXml);
  const points = descendants(dataRoot, "pt").filter((point) => modelIds.has(attr(point, "modelId") ?? ""));
  for (const shape of shapes) applyTextFormat(shape, format);
  for (const point of points) applyTextFormat(point, format);

  doc.pkg.raw()[drawingPart] = strToU8(serializeXml(drawingRoot, true));
  doc.pkg.raw()[dataRel.target] = strToU8(serializeXml(dataRoot, true));
  doc.markPackageResourceChanged();
  doc.refresh();
  return true;
}

function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

export function normalizeSmartArtData(data: SmartArtData): SmartArtData {
  const items = data.items.map((item) => item.trim()).filter(Boolean);
  return { layout: data.layout, items: items.length ? items : [""] };
}

function line(modelId: string, x1: number, y1: number, x2: number, y2: number): DiagramShape {
  return { modelId, x: x1, y: y1, width: x2 - x1, height: y2 - y1, geometry: "line" };
}

function diagramShapes(data: SmartArtData): DiagramShape[] {
  const n = data.items.length;
  const gap = 95_250;
  if (data.layout === "list") {
    const h = Math.max((HEIGHT_EMU - gap * (n + 1)) / n, 120_000);
    return data.items.map((text, index) => ({
      modelId: String(index + 1),
      x: gap, y: gap + index * (h + gap), width: WIDTH_EMU - gap * 2, height: h,
      text, geometry: "roundRect", color: COLORS[index % COLORS.length],
    }));
  }
  if (data.layout === "hierarchy") {
    const rootW = Math.min(WIDTH_EMU * 0.42, 1_600_000);
    const rootH = 560_000;
    const rootX = (WIDTH_EMU - rootW) / 2;
    if (n === 1) return [{ modelId: "1", x: rootX, y: (HEIGHT_EMU - rootH) / 2, width: rootW, height: rootH, text: data.items[0], geometry: "roundRect", color: COLORS[0] }];
    const childCount = n - 1;
    const childW = Math.max((WIDTH_EMU - gap * (childCount + 1)) / childCount, 260_000);
    const childH = 620_000;
    const childY = HEIGHT_EMU - childH - gap;
    const shapes: DiagramShape[] = [{ modelId: "1", x: rootX, y: gap, width: rootW, height: rootH, text: data.items[0], geometry: "roundRect", color: COLORS[0] }];
    data.items.slice(1).forEach((text, index) => {
      const x = gap + index * (childW + gap);
      shapes.push(line(String(n + index + 2), WIDTH_EMU / 2, gap + rootH, x + childW / 2, childY));
      shapes.push({ modelId: String(index + 2), x, y: childY, width: childW, height: childH, text, geometry: "roundRect", color: COLORS[(index + 1) % COLORS.length] });
    });
    return shapes;
  }
  if (data.layout === "cycle") {
    const shapeW = Math.min(1_100_000, WIDTH_EMU / Math.max(n, 3));
    const shapeH = 500_000;
    const cx = WIDTH_EMU / 2;
    const cy = HEIGHT_EMU / 2;
    const rx = Math.max((WIDTH_EMU - shapeW) / 2 - gap, 0);
    const ry = Math.max((HEIGHT_EMU - shapeH) / 2 - gap, 0);
    const nodes = data.items.map((text, index) => {
      const angle = -Math.PI / 2 + (index / n) * Math.PI * 2;
      return { modelId: String(index + 1), x: cx + Math.cos(angle) * rx - shapeW / 2, y: cy + Math.sin(angle) * ry - shapeH / 2, width: shapeW, height: shapeH, text, geometry: "ellipse" as const, color: COLORS[index % COLORS.length] };
    });
    const connectors = nodes.map((node, index) => {
      const next = nodes[(index + 1) % nodes.length];
      return line(String(n + index + 1), node.x + node.width / 2, node.y + node.height / 2, next.x + next.width / 2, next.y + next.height / 2);
    });
    return [...connectors, ...nodes];
  }
  const width = Math.max((WIDTH_EMU - gap * (n + 1)) / n, 260_000);
  const height = 850_000;
  const y = (HEIGHT_EMU - height) / 2;
  const nodes = data.items.map((text, index) => ({
    modelId: String(index + 1),
    x: gap + index * (width + gap), y, width, height, text,
    geometry: "roundRect" as const, color: COLORS[index % COLORS.length],
  }));
  const connectors = nodes.slice(0, -1).map((node, index) => line(String(n + index + 2), node.x + node.width, y + height / 2, nodes[index + 1].x, y + height / 2));
  return [...connectors, ...nodes];
}

function drawingShapeXml(shape: DiagramShape, index: number): string {
  const x = Math.round(Math.min(shape.x, shape.x + shape.width));
  const y = Math.round(Math.min(shape.y, shape.y + shape.height));
  const width = Math.max(1, Math.round(Math.abs(shape.width)));
  const height = Math.max(1, Math.round(Math.abs(shape.height)));
  const geometry = shape.geometry ?? "roundRect";
  const lineShape = geometry === "line";
  const flips = `${shape.width < 0 ? ' flipH="1"' : ""}${shape.height < 0 ? ' flipV="1"' : ""}`;
  return `<dsp:sp modelId="${shape.modelId}"><dsp:nvSpPr><dsp:cNvPr id="${index + 1}" name="${lineShape ? "Connector" : "Node"} ${index + 1}"/>` +
    `<dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm${flips}><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>` +
    `<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>` +
    (lineShape
      ? `<a:noFill/><a:ln w="19050"><a:solidFill><a:srgbClr val="7F8C8D"/></a:solidFill></a:ln>`
      : `<a:solidFill><a:srgbClr val="${shape.color ?? COLORS[0]}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`) +
    `<a:effectLst/></dsp:spPr>` +
    (shape.text === undefined ? "" : `<dsp:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>` +
      `<a:r><a:rPr lang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>` +
      `<a:latin typeface="Calibri"/></a:rPr><a:t>${escapeXml(shape.text)}</a:t></a:r><a:endParaRPr lang="en-US" sz="1200"/></a:p></dsp:txBody>`) +
    `</dsp:sp>`;
}

/** Build the cached DrawingML rendering Word stores alongside SmartArt data. */
export function buildSmartArtDrawingXml(input: SmartArtData): string {
  const data = normalizeSmartArtData(input);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<dsp:drawing xmlns:dgm="${NS_DGM}" xmlns:dsp="${NS_DSP}" xmlns:a="${NS_A}"><dsp:spTree>` +
    `<dsp:nvGrpSpPr><dsp:cNvPr id="0" name="SmartArt"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>` +
    `${diagramShapes(data).map(drawingShapeXml).join("")}</dsp:spTree></dsp:drawing>`;
}

/** Build the editable SmartArt node model related to the cached drawing. */
export function buildSmartArtDataXml(input: SmartArtData, drawingRelId: string): string {
  const data = normalizeSmartArtData(input);
  const points = data.items.map((text, index) =>
    `<dgm:pt modelId="${index + 1}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p>` +
    `<a:r><a:rPr lang="en-US"/><a:t>${escapeXml(text)}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
  ).join("");
  const connections = data.items.map((_, index) =>
    `<dgm:cxn modelId="${data.items.length + index + 1}" type="parOf" srcId="0" destId="${index + 1}" ` +
    `srcOrd="${index}" destOrd="0" presId=""/>`,
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<dgm:dataModel xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}"><dgm:ptLst>` +
    `<dgm:pt modelId="0" type="doc"><dgm:prSet loTypeId="urn:wordinweb:smartart:${data.layout}" loCatId="${data.layout}" ` +
    `qsTypeId="urn:wordinweb:smartart:style" qsCatId="simple" ` +
    `csTypeId="urn:wordinweb:smartart:colors" csCatId="accent1"/>` +
    `<dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p/></dgm:t></dgm:pt>${points}</dgm:ptLst>` +
    `<dgm:cxnLst>${connections}</dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:extLst>` +
    `<a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="${NS_DSP}" ` +
    `relId="${escapeXml(drawingRelId)}" minVer="${NS_DGM}"/></a:ext></dgm:extLst></dgm:dataModel>`;
}

export function buildSmartArtLayoutXml(input: SmartArtData): string {
  const data = normalizeSmartArtData(input);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<dgm:layoutDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="urn:wordinweb:smartart:${data.layout}">` +
    `<dgm:title val="${data.layout}"/><dgm:desc val="WordInWeb ${data.layout} diagram"/>` +
    `<dgm:catLst><dgm:cat type="${data.layout}" pri="1000"/></dgm:catLst>` +
    `<dgm:layoutNode name="diagram"><dgm:alg type="lin"/><dgm:shape/><dgm:presOf/>` +
    `<dgm:constrLst><dgm:constr type="w" for="ch" forName="node" refType="w"/>` +
    `<dgm:constr op="equ" type="h" for="ch" forName="node"/>` +
    `<dgm:constr op="equ" type="primFontSz" for="ch" forName="node" val="65"/></dgm:constrLst><dgm:ruleLst/>` +
    `<dgm:forEach axis="ch" ptType="node"><dgm:layoutNode name="node" styleLbl="node0">` +
    `<dgm:varLst><dgm:bulletEnabled val="true"/></dgm:varLst><dgm:alg type="tx"/>` +
    `<dgm:shape type="roundRect"/><dgm:presOf axis="desOrSelf" ptType="node"/>` +
    `<dgm:constrLst><dgm:constr type="tMarg" refType="primFontSz" fact="0.3"/>` +
    `<dgm:constr type="bMarg" refType="primFontSz" fact="0.3"/><dgm:constr type="lMarg" refType="primFontSz" fact="0.3"/>` +
    `<dgm:constr type="rMarg" refType="primFontSz" fact="0.3"/></dgm:constrLst>` +
    `<dgm:ruleLst><dgm:rule type="primFontSz" val="5"/></dgm:ruleLst>` +
    `</dgm:layoutNode></dgm:forEach></dgm:layoutNode></dgm:layoutDef>`;
}

export function buildSmartArtStyleXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<dgm:styleDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="urn:wordinweb:smartart:style">` +
    `<dgm:title val="WordInWeb"/><dgm:desc val="WordInWeb diagram style"/>` +
    `<dgm:catLst><dgm:cat type="simple" pri="1000"/></dgm:catLst>` +
    `<dgm:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></dgm:scene3d>` +
    `<dgm:styleLbl name="node0"><dgm:scene3d><a:camera prst="orthographicFront"/>` +
    `<a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/><dgm:style>` +
    `<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>` +
    `<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>` +
    `<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>` +
    `<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef>` +
    `</dgm:style></dgm:styleLbl></dgm:styleDef>`;
}

export function buildSmartArtColorsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<dgm:colorsDef xmlns:dgm="${NS_DGM}" xmlns:a="${NS_A}" uniqueId="urn:wordinweb:smartart:colors">` +
    `<dgm:title val="WordInWeb"/><dgm:desc val="WordInWeb diagram colors"/>` +
    `<dgm:catLst><dgm:cat type="accent1" pri="1000"/></dgm:catLst><dgm:styleLbl name="node0">` +
    `<dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/></dgm:fillClrLst>` +
    `<dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst>` +
    `<dgm:effectClrLst/><dgm:txLinClrLst/>` +
    `<dgm:txFillClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:txFillClrLst>` +
    `<dgm:txEffectClrLst/></dgm:styleLbl></dgm:colorsDef>`;
}

/** Insert native SmartArt parts and an inline diagram reference. */
export function insertSmartArtAt(doc: DocxDocument, caretT: XmlElement, input: SmartArtData): XmlElement | null {
  const caretRun = doc.findParentOf(caretT);
  const parent = caretRun && doc.findParentOf(caretRun);
  if (!caretRun || !parent || localName(caretRun.name) !== "r") return null;
  const data = normalizeSmartArtData(input);
  const rels = doc.addSmartArtResources(
    buildSmartArtLayoutXml(data),
    buildSmartArtStyleXml(),
    buildSmartArtColorsXml(),
    buildSmartArtDrawingXml(data),
    (drawingRelId) => buildSmartArtDataXml(data, drawingRelId),
  );
  const w = prefixOf(caretRun);
  const id = String(doc.nextDrawingId());
  const drawing = el(`${w}drawing`, {}, [
    el("wp:inline", { "xmlns:wp": NS_WP, distT: "0", distB: "0", distL: "0", distR: "0" }, [
      el("wp:extent", { cx: String(WIDTH_EMU), cy: String(HEIGHT_EMU) }),
      el("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" }),
      el("wp:docPr", { id, name: `Diagram ${id}` }),
      el("wp:cNvGraphicFramePr"),
      el("a:graphic", { "xmlns:a": NS_A }, [
        el("a:graphicData", { uri: NS_DGM }, [
          el("dgm:relIds", {
            "xmlns:dgm": NS_DGM, "xmlns:r": NS_R,
            "r:dm": rels.dataRelId, "r:lo": rels.layoutRelId,
            "r:qs": rels.styleRelId, "r:cs": rels.colorsRelId,
          }),
        ]),
      ]),
    ]),
  ]);
  parent.children.splice(parent.children.indexOf(caretRun) + 1, 0, el(`${w}r`, {}, [drawing]));
  doc.refresh();
  return drawing;
}

/** Replace a selected SmartArt diagram's data, layout definition, and cached drawing. */
export function setSmartArtData(doc: DocxDocument, drawing: XmlElement, input: SmartArtData): boolean {
  const relIds = descendant(drawing, "relIds");
  if (!relIds) return false;
  const dataRel = doc.documentRels.get(attr(relIds, "dm") ?? "");
  const layoutRel = doc.documentRels.get(attr(relIds, "lo") ?? "");
  if (!dataRel || dataRel.external || !layoutRel || layoutRel.external) return false;
  const dataRootXml = doc.pkg.text(dataRel.target);
  if (!dataRootXml) return false;
  const drawingRelId = attr(descendant(parseXml(dataRootXml), "dataModelExt"), "relId");
  const drawingPart = smartArtDrawingPart(doc, drawing);
  if (!drawingRelId || !drawingPart) return false;
  const data = normalizeSmartArtData(input);
  doc.pkg.raw()[dataRel.target] = strToU8(buildSmartArtDataXml(data, drawingRelId));
  doc.pkg.raw()[layoutRel.target] = strToU8(buildSmartArtLayoutXml(data));
  doc.pkg.raw()[drawingPart] = strToU8(buildSmartArtDrawingXml(data));
  doc.markPackageResourceChanged();
  doc.refresh();
  return true;
}
