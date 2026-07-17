import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

export type ShapePreset = "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox";
export type WordArtPreset = "plain" | "archUp" | "archDown" | "wave" | "chevron";
export interface InkPoint { x: number; y: number }
export type DrawingTool =
  | { kind?: "pen"; color: string; width: number }
  | { kind: "highlighter"; color: string; width: number }
  | { kind: "eraser"; size: number }
  | { kind: "lasso" };

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function descendant(node: XmlElement, name: string): XmlElement | undefined {
  if (localName(node.name) === name) return node;
  for (const child of node.children) {
    const found = descendant(child, name);
    if (found) return found;
  }
  return undefined;
}

/** Insert a floating DrawingML shape anchored to the caret paragraph. */
export function insertShapeAt(
  doc: DocxDocument,
  caretT: XmlElement,
  preset: ShapePreset,
  text = "",
): XmlElement | null {
  const caretRun = doc.findParentOf(caretT);
  const parent = caretRun && doc.findParentOf(caretRun);
  if (!caretRun || !parent || localName(caretRun.name) !== "r") return null;

  const w = prefixOf(caretRun);
  const id = String(doc.nextDrawingId());
  const width = preset === "textBox" ? 240 : 192;
  const height = preset === "textBox" ? 72 : 96;
  const cx = String(Math.round(width * EMU_PER_PX));
  const cy = String(Math.round(height * EMU_PER_PX));
  const geometry =
    preset === "roundedRectangle" ? "roundRect" : preset === "rectangle" || preset === "textBox" ? "rect" : preset;
  const isTextBox = preset === "textBox";
  const shapeName = isTextBox ? `Text Box ${id}` : `Shape ${id}`;

  const spPr = [
    el("a:xfrm", {}, [el("a:off", { x: "0", y: "0" }), el("a:ext", { cx, cy })]),
    el("a:prstGeom", { prst: geometry }, [el("a:avLst")]),
    isTextBox
      ? el("a:noFill")
      : el("a:solidFill", {}, [el("a:srgbClr", { val: "4472C4" })]),
    el("a:ln", { w: "12700" }, [el("a:solidFill", {}, [el("a:srgbClr", { val: isTextBox ? "404040" : "2F5597" })])]),
  ];
  const textRun = el(`${w}r`, {}, [
    el(`${w}rPr`, {}, [el(`${w}color`, { [`${w}val`]: isTextBox ? "202124" : "FFFFFF" }), el(`${w}sz`, { [`${w}val`]: "22" })]),
    el(`${w}t`, { "xml:space": "preserve" }, [], text),
  ]);
  const drawing = el(`${w}drawing`, {}, [
    el("wp:anchor", {
      "xmlns:wp": NS_WP,
      distT: "0", distB: "0", distL: "114300", distR: "114300",
      simplePos: "0", relativeHeight: "251658240", behindDoc: "0",
      locked: "0", layoutInCell: "1", allowOverlap: "1",
    }, [
      el("wp:simplePos", { x: "0", y: "0" }),
      el("wp:positionH", { relativeFrom: "margin" }, [el("wp:posOffset", {}, [], "0")]),
      el("wp:positionV", { relativeFrom: "paragraph" }, [el("wp:posOffset", {}, [], "0")]),
      el("wp:extent", { cx, cy }),
      el("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" }),
      el("wp:wrapSquare", { wrapText: "bothSides" }),
      el("wp:docPr", { id, name: shapeName }),
      el("wp:cNvGraphicFramePr"),
      el("a:graphic", { "xmlns:a": NS_A }, [
        el("a:graphicData", { uri: NS_WPS }, [
          el("wps:wsp", { "xmlns:wps": NS_WPS }, [
            el("wps:cNvSpPr", { txBox: "1" }),
            el("wps:spPr", {}, spPr),
            el("wps:txbx", {}, [
              el(`${w}txbxContent`, {}, [
                el(`${w}p`, {}, [
                  el(`${w}pPr`, {}, [el(`${w}jc`, { [`${w}val`]: isTextBox ? "left" : "center" })]),
                  textRun,
                ]),
              ]),
            ]),
            el("wps:bodyPr", { rot: "0", anchor: isTextBox ? "t" : "ctr" }, [el("a:noAutofit")]),
          ]),
        ]),
      ]),
    ]),
  ]);
  const run = el(`${w}r`, {}, [drawing]);
  parent.children.splice(parent.children.indexOf(caretRun) + 1, 0, run);
  doc.refresh();
  return drawing;
}

/** Insert editable DrawingML WordArt with a native preset text warp. */
export function insertWordArtAt(
  doc: DocxDocument,
  caretT: XmlElement,
  text: string,
  preset: WordArtPreset = "plain",
): XmlElement | null {
  if (!text) return null;
  const drawing = insertShapeAt(doc, caretT, "textBox", text);
  if (!drawing) return null;

  const warp = {
    plain: "textNoShape",
    archUp: "textArchUp",
    archDown: "textArchDown",
    wave: "textWave1",
    chevron: "textChevron",
  }[preset];
  const docPr = descendant(drawing, "docPr");
  if (docPr) docPr.attrs.name = `WordArt ${docPr.attrs.id}`;

  if (preset === "plain") {
    const height = String(Math.round(40 * EMU_PER_PX));
    const extent = descendant(drawing, "extent");
    const transformExtent = descendant(drawing, "ext");
    if (extent) extent.attrs.cy = height;
    if (transformExtent) transformExtent.attrs.cy = height;
  }

  const spPr = descendant(drawing, "spPr");
  if (spPr) {
    spPr.children = spPr.children.filter((child) => localName(child.name) !== "ln");
    spPr.children.push(el("a:ln", {}, [el("a:noFill")]));
  }

  const rPr = descendant(drawing, "rPr");
  if (rPr) {
    rPr.children.unshift(el(`${prefixOf(rPr)}b`));
    const color = rPr.children.find((child) => localName(child.name) === "color");
    const size = rPr.children.find((child) => localName(child.name) === "sz");
    if (color) color.attrs[Object.keys(color.attrs).find((key) => localName(key) === "val") ?? `${prefixOf(color)}val`] = "2E74B5";
    if (size) size.attrs[Object.keys(size.attrs).find((key) => localName(key) === "val") ?? `${prefixOf(size)}val`] = "40";
  }

  const bodyPr = descendant(drawing, "bodyPr");
  if (bodyPr) bodyPr.children.unshift(el("a:prstTxWarp", { prst: warp }, [el("a:avLst")]));
  doc.refresh();
  return drawing;
}

export function isDrawingWordArt(drawing: XmlElement): boolean {
  return (descendant(drawing, "docPr")?.attrs.name ?? "").startsWith("WordArt ");
}

export function drawingWordArtText(drawing: XmlElement): string {
  return isDrawingWordArt(drawing) ? descendant(drawing, "t")?.text ?? "" : "";
}

export function setDrawingWordArtText(doc: DocxDocument, drawing: XmlElement, text: string): boolean {
  const textElement = isDrawingWordArt(drawing) ? descendant(drawing, "t") : undefined;
  if (!textElement) return false;
  textElement.text = text;
  textElement.attrs["xml:space"] = "preserve";
  doc.refresh();
  return true;
}

/** Insert a freehand stroke as editable, movable DrawingML custom geometry. */
export function insertInkAt(
  doc: DocxDocument,
  caretT: XmlElement,
  points: readonly InkPoint[],
  color = "#202124",
  widthPx = 2,
  opacity = 1,
  refreshModel = true,
): XmlElement | null {
  if (points.length < 2) return null;
  const caretRun = doc.findParentOf(caretT);
  const parent = caretRun && doc.findParentOf(caretRun);
  if (!caretRun || !parent || localName(caretRun.name) !== "r") return null;

  const w = prefixOf(caretRun);
  const pad = Math.max(widthPx, 1) + 2;
  const minX = Math.max(0, Math.min(...points.map((point) => point.x)) - pad);
  const minY = Math.max(0, Math.min(...points.map((point) => point.y)) - pad);
  const maxX = Math.max(...points.map((point) => point.x)) + pad;
  const maxY = Math.max(...points.map((point) => point.y)) + pad;
  const width = Math.max(maxX - minX, 4);
  const height = Math.max(maxY - minY, 4);
  const cx = Math.round(width * EMU_PER_PX);
  const cy = Math.round(height * EMU_PER_PX);
  const id = String(doc.nextDrawingId());
  const pt = (point: InkPoint) => el("a:pt", {
    x: String(Math.round((point.x - minX) * EMU_PER_PX)),
    y: String(Math.round((point.y - minY) * EMU_PER_PX)),
  });
  const commands = [
    el("a:moveTo", {}, [pt(points[0])]),
    ...points.slice(1).map((point) => el("a:lnTo", {}, [pt(point)])),
  ];
  const drawing = el(`${w}drawing`, {}, [
    el("wp:anchor", {
      "xmlns:wp": NS_WP,
      distT: "0", distB: "0", distL: "0", distR: "0",
      simplePos: "0", relativeHeight: "251658240", behindDoc: "0",
      locked: "0", layoutInCell: "1", allowOverlap: "1",
    }, [
      el("wp:simplePos", { x: "0", y: "0" }),
      el("wp:positionH", { relativeFrom: "page" }, [el("wp:posOffset", {}, [], String(Math.round(minX * EMU_PER_PX)))]),
      el("wp:positionV", { relativeFrom: "page" }, [el("wp:posOffset", {}, [], String(Math.round(minY * EMU_PER_PX)))]),
      el("wp:extent", { cx: String(cx), cy: String(cy) }),
      el("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" }),
      el("wp:wrapNone"),
      el("wp:docPr", { id, name: `Ink ${id}`, descr: "WordInWeb ink" }),
      el("wp:cNvGraphicFramePr"),
      el("a:graphic", { "xmlns:a": NS_A }, [
        el("a:graphicData", { uri: NS_WPS }, [
          el("wps:wsp", { "xmlns:wps": NS_WPS }, [
            el("wps:cNvSpPr"),
            el("wps:spPr", {}, [
              el("a:xfrm", {}, [el("a:off", { x: "0", y: "0" }), el("a:ext", { cx: String(cx), cy: String(cy) })]),
              el("a:custGeom", {}, [
                el("a:avLst"), el("a:gdLst"), el("a:ahLst"), el("a:cxnLst"),
                el("a:rect", { l: "0", t: "0", r: "r", b: "b" }),
                el("a:pathLst", {}, [el("a:path", { w: String(cx), h: String(cy), fill: "none", stroke: "1" }, commands)]),
              ]),
              el("a:noFill"),
              el("a:ln", { w: String(Math.max(Math.round(widthPx * EMU_PER_PX), 1)), cap: "rnd" }, [
                el("a:solidFill", {}, [el(
                  "a:srgbClr",
                  { val: color.replace(/^#/, "").toUpperCase() },
                  opacity < 1 ? [el("a:alpha", { val: String(Math.round(opacity * 100000)) })] : [],
                )]),
                el("a:round"),
              ]),
            ]),
            el("wps:bodyPr"),
          ]),
        ]),
      ]),
    ]),
  ]);
  parent.children.splice(parent.children.indexOf(caretRun) + 1, 0, el(`${w}r`, {}, [drawing]));
  if (refreshModel) doc.refresh();
  return drawing;
}
