import { DocxDocument } from "../docx.js";
import { XmlElement, localName } from "../xml.js";

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

export type ShapePreset = "line" | "verticalLine" | "rectangle" | "roundedRectangle" | "ellipse" | "diamond" | "textBox";
export type WordArtPreset = "plain" | "archUp" | "archDown" | "wave" | "chevron";
export type DrawingLineDash = "solid" | "dashed" | "dotted";
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

function vmlLengthPx(raw: string | undefined): number {
  if (!raw) return 0;
  const match = /^(-?[\d.]+)\s*(pt|in|px)?$/.exec(raw.trim());
  if (!match) return 0;
  const value = parseFloat(match[1]);
  if (match[2] === "pt") return value * 4 / 3;
  if (match[2] === "in") return value * 96;
  return value;
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
  const isLine = preset === "line" || preset === "verticalLine";
  const isVerticalLine = preset === "verticalLine";
  const width = isVerticalLine ? 2 : isLine ? 240 : preset === "textBox" ? 240 : 192;
  const height = isVerticalLine ? 240 : isLine ? 2 : preset === "textBox" ? 72 : 96;
  const cx = String(Math.round(width * EMU_PER_PX));
  const cy = String(Math.round(height * EMU_PER_PX));
  // Keep the line geometry horizontal while the outer extent retains enough
  // height to provide a usable selection target and resize handles.
  const shapeCx = isVerticalLine ? "0" : cx;
  const shapeCy = isLine && !isVerticalLine ? "0" : cy;
  const geometry =
    isLine ? "line" : preset === "roundedRectangle" ? "roundRect" : preset === "rectangle" || preset === "textBox" ? "rect" : preset;
  const isTextBox = preset === "textBox";
  const shapeName = isLine ? `Line ${id}` : isTextBox ? `Text Box ${id}` : `Shape ${id}`;

  const spPr = [
    el("a:xfrm", {}, [
      el("a:off", { x: "0", y: "0" }),
      el("a:ext", { cx: shapeCx, cy: shapeCy }),
    ]),
    el("a:prstGeom", { prst: geometry }, [el("a:avLst")]),
    isTextBox || isLine
      ? el("a:noFill")
      : el("a:solidFill", {}, [el("a:srgbClr", { val: "4472C4" })]),
    el("a:ln", { w: "12700" }, [el("a:solidFill", {}, [el("a:srgbClr", { val: isTextBox || isLine ? "404040" : "2F5597" })])]),
  ];
  const textRun = el(`${w}r`, {}, [
    el(`${w}rPr`, {}, [el(`${w}color`, { [`${w}val`]: isTextBox ? "202124" : "FFFFFF" }), el(`${w}sz`, { [`${w}val`]: "22" })]),
    el(`${w}t`, { "xml:space": "preserve" }, [], text),
  ]);
  const shapeChildren = isLine
    ? [
        el("wps:cNvCnPr", {}, [el("a:cxnSpLocks", { noChangeShapeType: "1" })]),
        el("wps:spPr", { bwMode: "auto" }, spPr),
        el("wps:bodyPr"),
      ]
    : [
        el("wps:cNvSpPr", isTextBox ? { txBox: "1" } : {}),
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
      ];
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
      isLine ? el("wp:wrapNone") : el("wp:wrapSquare", { wrapText: "bothSides" }),
      el("wp:docPr", { id, name: shapeName }),
      el("wp:cNvGraphicFramePr"),
      el("a:graphic", { "xmlns:a": NS_A }, [
        el("a:graphicData", { uri: NS_WPS }, [
          el("wps:wsp", { "xmlns:wps": NS_WPS }, shapeChildren),
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

  // WordArt is decorative page content. Let it sit over the document instead
  // of carving a 240 px channel through every line near its anchor. Several
  // square-wrapped WordArt objects otherwise reflow the same body text after
  // each drag, which also makes Word lay out the saved file very differently.
  const anchor = descendant(drawing, "anchor");
  if (anchor) {
    const wrap = anchor.children.findIndex((child) => localName(child.name).startsWith("wrap"));
    if (wrap !== -1) anchor.children[wrap] = el("wp:wrapNone");
    anchor.attrs.behindDoc = "0";
  }

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

  const paragraphProperties = descendant(drawing, "pPr");
  const justification = paragraphProperties ? descendant(paragraphProperties, "jc") : undefined;
  if (justification) {
    const valueKey =
      Object.keys(justification.attrs).find((key) => localName(key) === "val") ??
      `${prefixOf(justification)}val`;
    justification.attrs[valueKey] = "center";
  }

  const bodyPr = descendant(drawing, "bodyPr");
  if (bodyPr) {
    bodyPr.attrs = {
      rot: "0",
      spcFirstLastPara: "1",
      vertOverflow: "overflow",
      horzOverflow: "overflow",
      vert: "horz",
      wrap: "none",
      lIns: "91440",
      tIns: "45720",
      rIns: "91440",
      bIns: "45720",
      numCol: "1",
      spcCol: "0",
      rtlCol: "0",
      fromWordArt: "0",
      anchor: "t",
      anchorCtr: "0",
      forceAA: "0",
      compatLnSpc: "1",
    };
    // Keep the authored box size stable in WordInWeb. Word's gallery uses
    // spAutoFit, which recomputes the box height and breaks repeated drag.
    bodyPr.children = [
      el("a:prstTxWarp", { prst: warp }, [el("a:avLst")]),
      el("a:noAutofit"),
    ];
  }
  doc.refresh();
  return drawing;
}

export function isDrawingWordArt(drawing: XmlElement): boolean {
  return (descendant(drawing, "docPr")?.attrs.name ?? "").startsWith("WordArt ") ||
    (["shape", "rect"].includes(localName(drawing.name)) && Boolean(descendant(drawing, "textpath")));
}

export function isDrawingLine(drawing: XmlElement): boolean {
  if (localName(drawing.name) === "line") return true;
  const preset = descendant(drawing, "prstGeom")?.attrs.prst ?? "";
  return preset === "line" || preset.startsWith("straightConnector");
}

export function drawingWordArtText(drawing: XmlElement): string {
  if (!isDrawingWordArt(drawing)) return "";
  return descendant(drawing, "textpath")?.attrs.string ?? descendant(drawing, "t")?.text ?? "";
}

export function setDrawingWordArtText(doc: DocxDocument, drawing: XmlElement, text: string): boolean {
  if (!isDrawingWordArt(drawing)) return false;
  const textPath = descendant(drawing, "textpath");
  if (textPath) {
    textPath.attrs.string = text;
    doc.refresh();
    return true;
  }
  const textElement = descendant(drawing, "t");
  if (!textElement) return false;
  textElement.text = text;
  textElement.attrs["xml:space"] = "preserve";
  doc.refresh();
  return true;
}

/** Set the visible glyph fill and alpha of DrawingML or VML WordArt. */
export function setDrawingWordArtStyle(
  doc: DocxDocument,
  drawing: XmlElement,
  color: string,
  opacity = 1,
): boolean {
  if (!isDrawingWordArt(drawing)) return false;
  const hex = color.replace(/^#/, "").toUpperCase();
  const alpha = Math.round(Math.max(0, Math.min(1, opacity)) * 100000);
  const textPath = descendant(drawing, "textpath");
  if (textPath) {
    drawing.attrs.filled = "t";
    drawing.attrs.fillcolor = `#${hex}`;
    let fill = descendant(drawing, "fill");
    if (!fill) {
      fill = el(`${prefixOf(drawing)}fill`);
      drawing.children.push(fill);
    }
    fill.attrs.opacity = String(Math.round(alpha / 100) / 1000);
    doc.refresh();
    return true;
  }

  const rPr = descendant(drawing, "rPr");
  const spPr = descendant(drawing, "spPr");
  if (!rPr || !spPr) return false;

  let fallback = rPr.children.find((child) => localName(child.name) === "color");
  if (!fallback) {
    fallback = el(`${prefixOf(rPr)}color`);
    rPr.children.push(fallback);
  }
  const fallbackHex = [0, 2, 4].map((offset) => {
    const channel = parseInt(hex.slice(offset, offset + 2), 16);
    return Math.round(255 + (channel - 255) * Math.max(0, Math.min(1, opacity))).toString(16).padStart(2, "0");
  }).join("").toUpperCase();
  // Word uses w14:alpha below. LibreOffice ignores that extension and paints
  // the legacy w:color instead, so give the fallback the same visual weight
  // on the white page that repeating watermarks normally sit over.
  fallback.attrs[`${prefixOf(fallback)}val`] = fallbackHex;
  rPr.children = rPr.children.filter((child) => localName(child.name) !== "textFill");
  rPr.children.push(el("w14:textFill", { "xmlns:w14": NS_W14 }, [
    el("w14:solidFill", {}, [
      el("w14:srgbClr", { "w14:val": hex }, alpha < 100000 ? [el("w14:alpha", { "w14:val": String(alpha) })] : []),
    ]),
  ]));

  spPr.children = spPr.children.filter((child) => !["solidFill", "noFill"].includes(localName(child.name)));
  const lineIndex = spPr.children.findIndex((child) => localName(child.name) === "ln");
  spPr.children.splice(lineIndex === -1 ? spPr.children.length : lineIndex, 0, el("a:noFill"));
  doc.refresh();
  return true;
}

export function drawingLineStyle(drawing: XmlElement): { color: string; width: number; dash: DrawingLineDash } | null {
  const spPr = descendant(drawing, "spPr");
  if (!spPr) {
    if (!["line", "shape", "rect"].includes(localName(drawing.name))) return null;
    if (drawing.attrs.stroked === "f" || drawing.attrs.stroked === "false") return null;
    const dashValue = descendant(drawing, "stroke")?.attrs.dashstyle?.toLowerCase() ?? "solid";
    return {
      color: drawing.attrs.strokecolor ?? "#000000",
      width: vmlLengthPx(drawing.attrs.strokeweight) || 0.75,
      dash: dashValue.includes("dot") ? "dotted" : dashValue !== "solid" ? "dashed" : "solid",
    };
  }
  const line = descendant(drawing, "ln");
  if (!line) return { color: "#000000", width: 0.75, dash: "solid" };
  if (line.children.some((child) => localName(child.name) === "noFill")) return null;
  const color = descendant(line, "srgbClr")?.attrs.val ?? "000000";
  const width = Math.max((parseInt(line.attrs.w ?? "0", 10) || 0) / EMU_PER_PX, 0.75);
  const dashValue = line.children.find((child) => localName(child.name) === "prstDash")?.attrs.val;
  const dash = dashValue === "dot" || dashValue === "sysDot"
    ? "dotted"
    : dashValue && dashValue !== "solid"
      ? "dashed"
      : "solid";
  return { color: `#${color.toUpperCase()}`, width, dash };
}

export function drawingFillColor(drawing: XmlElement): string | null {
  const spPr = descendant(drawing, "spPr");
  if (!spPr) {
    if (!["shape", "rect"].includes(localName(drawing.name)) || drawing.attrs.filled === "f") return null;
    return drawing.attrs.fillcolor ?? null;
  }
  if (spPr.children.some((child) => localName(child.name) === "noFill")) return null;
  const solidFill = spPr.children.find((child) => localName(child.name) === "solidFill");
  const color = solidFill ? descendant(solidFill, "srgbClr")?.attrs.val : undefined;
  return color ? `#${color.toUpperCase()}` : null;
}

export function setDrawingFill(doc: DocxDocument, drawing: XmlElement, color: string | null): boolean {
  const spPr = descendant(drawing, "spPr");
  if (!spPr) {
    if (!["shape", "rect"].includes(localName(drawing.name))) return false;
    drawing.attrs.filled = color ? "t" : "f";
    if (color) drawing.attrs.fillcolor = color;
    else delete drawing.attrs.fillcolor;
    doc.refresh();
    return true;
  }
  spPr.children = spPr.children.filter((child) => {
    const name = localName(child.name);
    return name !== "solidFill" && name !== "noFill";
  });
  const fill = color
    ? el("a:solidFill", {}, [el("a:srgbClr", { val: color.replace(/^#/, "").toUpperCase() })])
    : el("a:noFill");
  const lineIndex = spPr.children.findIndex((child) => localName(child.name) === "ln");
  spPr.children.splice(lineIndex === -1 ? spPr.children.length : lineIndex, 0, fill);
  doc.refresh();
  return true;
}

export function setDrawingLineStyle(
  doc: DocxDocument,
  drawing: XmlElement,
  color: string | null,
  widthPx = 0.75,
  dash: DrawingLineDash = "solid",
): boolean {
  const spPr = descendant(drawing, "spPr");
  if (!spPr) {
    if (!["line", "shape", "rect"].includes(localName(drawing.name))) return false;
    if (color === null) {
      drawing.attrs.stroked = "f";
      delete drawing.attrs.strokecolor;
      delete drawing.attrs.strokeweight;
      doc.refresh();
      return true;
    }
    drawing.attrs.stroked = "t";
    drawing.attrs.strokecolor = color;
    drawing.attrs.strokeweight = `${Math.round(widthPx * 75) / 100}pt`;
    let stroke = descendant(drawing, "stroke");
    if (!stroke) {
      stroke = el(`${prefixOf(drawing)}stroke`);
      drawing.children.push(stroke);
    }
    stroke.attrs.dashstyle = dash === "dotted" ? "dot" : dash === "dashed" ? "dash" : "solid";
    doc.refresh();
    return true;
  }
  let line = spPr.children.find((child) => localName(child.name) === "ln");
  if (!line) {
    line = el("a:ln");
    spPr.children.push(line);
  }
  if (color === null) {
    line.children = line.children.filter((child) => {
      const name = localName(child.name);
      return name !== "solidFill" && name !== "noFill" && name !== "prstDash";
    });
    line.children.unshift(el("a:noFill"));
    doc.refresh();
    return true;
  }
  line.attrs.w = String(Math.max(1, Math.round(widthPx * EMU_PER_PX)));
  line.children = line.children.filter((child) => {
    const name = localName(child.name);
    return name !== "solidFill" && name !== "noFill" && name !== "prstDash";
  });
  line.children.unshift(el("a:solidFill", {}, [el("a:srgbClr", { val: color.replace(/^#/, "").toUpperCase() })]));
  line.children.splice(1, 0, el("a:prstDash", { val: dash === "dotted" ? "dot" : dash === "dashed" ? "dash" : "solid" }));
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
