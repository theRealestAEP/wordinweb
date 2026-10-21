/**
 * WordprocessingML for the pictures and shapes `drawings.ts` describes: an
 * inline `<w:drawing><wp:inline>` for a picture, an anchored
 * `<w:drawing><wp:anchor>` for a floating picture, shape or group. The
 * caller owns the package: it allocates relationship ids for the media
 * through `addMedia` and fills text boxes through `textBoxContent`.
 */
import type { ArtImage } from "./officeart.js";
import type { DocCrop, DocCustomGeometry, DocFill, DocFloatingShape, DocInlinePicture, DocLine, DocShape } from "./drawings.js";

export interface DrawingXmlOptions {
  /** Store an image as a media part and return the relationship id that names it from the current part. */
  addMedia: (image: ArtImage) => string;
  /** The next wp:docPr id (unique within the document). */
  nextId: () => number;
  /** The block content of a text box (w:p / w:tbl elements), or undefined for an empty paragraph. */
  textBoxContent?: (shape: DocShape) => string | undefined;
  /** Added to a text box story's number to make a linked chain's id unique in the document (header stories count apart). */
  textBoxIdBase?: number;
}

const NS =
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function escape(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

const emu = (v: number): string => String(Math.round(v));

/** An inline picture run body: `<w:drawing>` with a `wp:inline`. */
export function inlinePictureXml(picture: DocInlinePicture, options: DrawingXmlOptions): string {
  const id = options.nextId();
  const name = picture.shape?.name ?? `Picture ${id}`;
  const shape = picture.shape;
  const line = shape?.line ?? picfBorderLine(picture);
  const frame = { x: 0, y: 0, cx: picture.size.cx, cy: picture.size.cy };
  // An inline shape with no picture (a horizontal rule): drawn as a shape.
  const asShape = !picture.image && shape && shape.kind === "shape";
  const graphic = asShape
    ? `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">${wspXml(shape, frame, false, options)}</a:graphicData>`
    : `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">${picXml(picture.image, picture.crop, frame, shape?.rotation ?? 0, !!shape?.flipH, !!shape?.flipV, line, name, options)}</a:graphicData>`;
  return (
    `<w:drawing ${NS}><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${emu(picture.size.cx)}" cy="${emu(picture.size.cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${id}" name="${escape(name)}"${shape?.description ? ` descr="${escape(shape.description)}"` : ""}/>` +
    `<wp:cNvGraphicFramePr>${asShape ? "" : '<a:graphicFrameLocks noChangeAspect="1"/>'}</wp:cNvGraphicFramePr>` +
    `<a:graphic>${graphic}</a:graphic>` +
    `</wp:inline></w:drawing>`
  );
}

/** The PICF's own borders (Word 97 pictures) as a line: the widest of the four. */
function picfBorderLine(picture: DocInlinePicture): DocLine | null {
  const borders = Object.values(picture.picf.borders).filter((b) => b.style > 0 && b.width > 0);
  if (!borders.length) return null;
  const widest = borders.reduce((a, b) => (b.width > a.width ? b : a));
  return { color: widest.color, width: widest.width * 12700 / 8, dash: widest.style === 6 || widest.style === 7 ? "dash" : "solid", compound: "sng" };
}

/** A floating picture, shape or group: `<w:drawing>` with a `wp:anchor`. */
export function anchoredShapeXml(floating: DocFloatingShape, options: DrawingXmlOptions): string {
  const { shape } = floating;
  if (shape.wordArt) return wordArtPictXml(floating, options.nextId());
  const id = options.nextId();
  const name = shape.name ?? `${shape.kind === "picture" ? "Picture" : shape.kind === "group" ? "Group" : "Shape"} ${id}`;
  const position = (axis: "horizontal" | "vertical"): string => {
    const p = floating[axis];
    const tag = axis === "horizontal" ? "positionH" : "positionV";
    const body = p.align ? `<wp:align>${p.align}</wp:align>` : `<wp:posOffset>${emu(p.offset)}</wp:posOffset>`;
    return `<wp:${tag} relativeFrom="${p.relativeFrom}">${body}</wp:${tag}>`;
  };
  const frame = { x: 0, y: 0, cx: floating.extent.cx, cy: floating.extent.cy };
  let graphic: string;
  let uri: string;
  if (shape.kind === "group") {
    uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    graphic = groupXml(shape, true, options);
  } else if (shape.kind === "picture" && !shape.text) {
    uri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    graphic = picXml(shape.picture?.image ?? null, shape.picture?.crop, frame, shape.rotation, shape.flipH, shape.flipV, shape.line, name, options);
  } else {
    uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    graphic = wspXml(shape, frame, false, options);
  }
  const relativeHeight = 251658240 + floating.zIndex + (floating.behindText ? 0 : 1024);
  return (
    `<w:drawing ${NS}><wp:anchor distT="${emu(floating.distance.top)}" distB="${emu(floating.distance.bottom)}" distL="${emu(floating.distance.left)}" distR="${emu(floating.distance.right)}" ` +
    `simplePos="0" relativeHeight="${relativeHeight}" behindDoc="${floating.behindText ? 1 : 0}" locked="${floating.anchorLock ? 1 : 0}" layoutInCell="${floating.layoutInCell ? 1 : 0}" allowOverlap="1"${shape.hidden ? ' hidden="1"' : ""}>` +
    `<wp:simplePos x="0" y="0"/>${position("horizontal")}${position("vertical")}` +
    `<wp:extent cx="${emu(frame.cx)}" cy="${emu(frame.cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    wrapXml(floating) +
    `<wp:docPr id="${id}" name="${escape(name)}"${shape.description ? ` descr="${escape(shape.description)}"` : ""}${shape.hidden ? ' hidden="1"' : ""}/>` +
    `<wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="${uri}">${graphic}</a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing>`
  );
}

/**
 * A floating WordArt shape as Word writes legacy WordArt: a VML v:shape whose
 * v:textpath text is drawn in the fill colour and stretched to the shape's
 * box. The fill belongs to the letters, not to a box behind them.
 */
function wordArtPictXml(floating: DocFloatingShape, id: number): string {
  const { shape } = floating;
  const pt = (value: number): string => `${Math.round(value / 127) / 100}pt`;
  const relH: Record<string, string> = { margin: "margin", page: "page", column: "text", character: "char" };
  const relV: Record<string, string> = { margin: "margin", page: "page", paragraph: "text", line: "line" };
  const style = [
    "position:absolute",
    `margin-left:${pt(floating.horizontal.offset)}`,
    `margin-top:${pt(floating.vertical.offset)}`,
    `width:${pt(floating.extent.cx)}`,
    `height:${pt(floating.extent.cy)}`,
    ...(shape.rotation ? [`rotation:${shape.rotation}`] : []),
    // Word draws a header's drawings under the body text, so a header WordArt (a watermark) goes behind it.
    `z-index:${floating.behindText || floating.dgglbl === 1 ? -(floating.zIndex + 1) : floating.zIndex + 1}`,
    ...(floating.horizontal.align ? [`mso-position-horizontal:${floating.horizontal.align}`] : []),
    `mso-position-horizontal-relative:${relH[floating.horizontal.relativeFrom] ?? "text"}`,
    ...(floating.vertical.align ? [`mso-position-vertical:${floating.vertical.align}`] : []),
    `mso-position-vertical-relative:${relV[floating.vertical.relativeFrom] ?? "text"}`,
  ].join(";");
  const fill = shape.fill;
  const font = shape.wordArt?.font ? `font-family:&quot;${escape(shape.wordArt.font)}&quot;` : "";
  return (
    `<w:pict xmlns:v="urn:schemas-microsoft-com:vml">` +
    `<v:shape id="WordArt ${id}" style="${style}" fillcolor="#${fill?.color ?? "808080"}" stroked="f">` +
    (fill?.opacity !== undefined ? `<v:fill opacity="${fill.opacity}"/>` : "") +
    `<v:textpath on="t" fitshape="t" style="${font}" string="${escape(shape.wordArt?.text ?? "")}"/>` +
    `</v:shape></w:pict>`
  );
}

function wrapXml(floating: DocFloatingShape): string {
  switch (floating.wrap) {
    case "none": return "<wp:wrapNone/>";
    case "topAndBottom": return "<wp:wrapTopAndBottom/>";
    case "square": return `<wp:wrapSquare wrapText="${floating.wrapSide}"/>`;
    case "tight":
    case "through": {
      const polygon = floating.wrapPolygon && floating.wrapPolygon.length >= 3 ? floating.wrapPolygon : [{ x: 0, y: 0 }, { x: 21600, y: 0 }, { x: 21600, y: 21600 }, { x: 0, y: 21600 }];
      const [start, ...rest] = polygon;
      const tag = floating.wrap === "tight" ? "wrapTight" : "wrapThrough";
      return `<wp:${tag} wrapText="${floating.wrapSide}"><wp:wrapPolygon edited="0"><wp:start x="${start.x}" y="${start.y}"/>${rest.map((p) => `<wp:lineTo x="${p.x}" y="${p.y}"/>`).join("")}</wp:wrapPolygon></wp:${tag}>`;
    }
  }
}

function xfrmXml(frame: { x: number; y: number; cx: number; cy: number }, rotation: number, flipH: boolean, flipV: boolean, extra = ""): string {
  const rot = rotation ? ` rot="${Math.round(((rotation % 360) + 360) % 360 * 60000)}"` : "";
  return `<a:xfrm${rot}${flipH ? ' flipH="1"' : ""}${flipV ? ' flipV="1"' : ""}><a:off x="${emu(frame.x)}" y="${emu(frame.y)}"/><a:ext cx="${emu(Math.max(0, frame.cx))}" cy="${emu(Math.max(0, frame.cy))}"/>${extra}</a:xfrm>`;
}

function picXml(image: ArtImage | null, crop: DocCrop | undefined, frame: { x: number; y: number; cx: number; cy: number }, rotation: number, flipH: boolean, flipV: boolean, line: DocLine | null, name: string, options: DrawingXmlOptions): string {
  const rId = image ? options.addMedia(image) : undefined;
  const pct = (v: number): string => String(Math.round(v * 100000));
  const srcRect = crop && (crop.left || crop.top || crop.right || crop.bottom) ? `<a:srcRect l="${pct(crop.left)}" t="${pct(crop.top)}" r="${pct(crop.right)}" b="${pct(crop.bottom)}"/>` : "";
  const blip = rId ? `<a:blip r:embed="${rId}"/>` : "";
  return (
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${escape(name)}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill>${blip}${srcRect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr>${xfrmXml(frame, rotation, flipH, flipV)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${lineXml(line)}</pic:spPr></pic:pic>`
  );
}

function groupXml(group: DocShape, topLevel: boolean, options: DrawingXmlOptions): string {
  // The children's frames are absolute EMU (placeShape mapped them), so the
  // child space is the group's own frame.
  const f = group.frame;
  const off = topLevel ? { x: 0, y: 0, cx: f.cx, cy: f.cy } : f;
  const xfrm = xfrmXml(off, group.rotation, group.flipH, group.flipV, `<a:chOff x="${emu(f.x)}" y="${emu(f.y)}"/><a:chExt cx="${emu(f.cx)}" cy="${emu(f.cy)}"/>`);
  const children = (group.children ?? []).map((child) => childXml(child, options)).join("");
  if (topLevel) return `<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrm}</wpg:grpSpPr>${children}</wpg:wgp>`;
  return `<wpg:grpSp><wpg:cNvPr id="${options.nextId()}" name="${escape(group.name ?? "Group")}"/><wpg:cNvGrpSpPr/><wpg:grpSpPr>${xfrm}</wpg:grpSpPr>${children}</wpg:grpSp>`;
}

function childXml(shape: DocShape, options: DrawingXmlOptions): string {
  if (shape.kind === "group") return groupXml(shape, false, options);
  if (shape.kind === "picture" && !shape.text) {
    return picXml(shape.picture?.image ?? null, shape.picture?.crop, shape.frame, shape.rotation, shape.flipH, shape.flipV, shape.line, shape.name ?? "Picture", options);
  }
  return wspXml(shape, shape.frame, true, options);
}

function wspXml(shape: DocShape, frame: { x: number; y: number; cx: number; cy: number }, inGroup: boolean, options: DrawingXmlOptions): string {
  const text = shape.text;
  const isTextBox = !!text || !!shape.wordArt;
  const cNvPr = inGroup ? `<wps:cNvPr id="${options.nextId()}" name="${escape(shape.name ?? "Shape")}"/>` : "";
  let body = "";
  // A linked chain: the first box holds the whole story and the layout
  // flows its overflow into the boxes that continue it (the .doc keeps no
  // split of its own).
  const chainId = text ? (options.textBoxIdBase ?? 0) + text.story + 1 : 0;
  if (text && text.sequence > 0) {
    body = `<wps:linkedTxbx id="${chainId}" seq="${text.sequence}"/>`;
  } else if (isTextBox) {
    const content = shape.wordArt ? `<w:p><w:r><w:t xml:space="preserve">${escape(shape.wordArt.text)}</w:t></w:r></w:p>` : options.textBoxContent?.(shape) ?? "<w:p/>";
    body = `<wps:txbx${text?.next ? ` id="${chainId}"` : ""}><w:txbxContent>${content}</w:txbxContent></wps:txbx>`;
  }
  const insets = text?.insets ?? { left: 91440, top: 45720, right: 91440, bottom: 45720 };
  const bodyPr =
    `<wps:bodyPr${text?.vert ? ` vert="${text.vert}"` : ""} wrap="${text && !text.wrap ? "none" : "square"}" lIns="${emu(insets.left)}" tIns="${emu(insets.top)}" rIns="${emu(insets.right)}" bIns="${emu(insets.bottom)}" anchor="${text?.anchor ?? "t"}"${text?.anchorCenter ? ' anchorCtr="1"' : ""}>` +
    (text?.fitShapeToText ? "<a:spAutoFit/>" : "<a:noAutofit/>") +
    `</wps:bodyPr>`;
  const geometry = shape.custom ? custGeomXml(shape.custom) : `<a:prstGeom prst="${shape.preset ?? "rect"}">${avLstXml(shape)}</a:prstGeom>`;
  return (
    `<wps:wsp>${cNvPr}<wps:cNvSpPr${isTextBox ? ' txBox="1"' : ""}/>` +
    `<wps:spPr>${xfrmXml(frame, shape.rotation, shape.flipH, shape.flipV)}${geometry}${fillXml(shape.fill, options)}${lineXml(shape.line)}${shadowXml(shape)}</wps:spPr>` +
    body + bodyPr + `</wps:wsp>`
  );
}

/** Adjust values with a known unit: the rounded rectangle's corner (a 16.16 fraction of the shorter side, halved). */
function avLstXml(shape: DocShape): string {
  if (shape.preset === "roundRect" && shape.adjust[0] !== undefined) {
    return `<a:avLst><a:gd name="adj" fmla="val ${Math.round(shape.adjust[0] * 100000 / 65536)}"/></a:avLst>`;
  }
  return "<a:avLst/>";
}

function custGeomXml(geometry: DocCustomGeometry): string {
  const pt = (x: number, y: number): string => `<a:pt x="${Math.round(x)}" y="${Math.round(y)}"/>`;
  const path = geometry.commands.map((c) => {
    switch (c.op) {
      case "moveTo": return `<a:moveTo>${pt(c.x, c.y)}</a:moveTo>`;
      case "lnTo": return `<a:lnTo>${pt(c.x, c.y)}</a:lnTo>`;
      case "cubicBezTo": return `<a:cubicBezTo>${c.points.map(([x, y]) => pt(x, y)).join("")}</a:cubicBezTo>`;
      case "close": return "<a:close/>";
    }
  }).join("");
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${Math.max(1, geometry.width)}" h="${Math.max(1, geometry.height)}">${path}</a:path></a:pathLst></a:custGeom>`;
}

function color(hex: string, opacity?: number): string {
  const alpha = opacity !== undefined && opacity < 1 ? `<a:alpha val="${Math.round(Math.max(0, opacity) * 100000)}"/>` : "";
  return `<a:srgbClr val="${hex}">${alpha}</a:srgbClr>`;
}

function fillXml(fill: DocFill | null, options: DrawingXmlOptions): string {
  if (!fill) return "<a:noFill/>";
  switch (fill.kind) {
    case "picture": {
      if (!fill.image) return `<a:solidFill>${color(fill.color, fill.opacity)}</a:solidFill>`;
      const rId = options.addMedia(fill.image);
      return `<a:blipFill><a:blip r:embed="${rId}"/>${fill.tile ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' : "<a:stretch><a:fillRect/></a:stretch>"}</a:blipFill>`;
    }
    case "gradient": {
      // OfficeArt angles run counter-clockwise from "left to right"; DrawingML's lin ang is clockwise from the same start.
      const ang = Math.round((((-(fill.angle ?? 0)) % 360) + 360) % 360 * 60000);
      const focus = fill.focus ?? 0;
      const stops = focus === 50 || focus === -50
        ? `<a:gs pos="0">${color(fill.backColor ?? fill.color)}</a:gs><a:gs pos="50000">${color(fill.color)}</a:gs><a:gs pos="100000">${color(fill.backColor ?? fill.color)}</a:gs>`
        : focus < 0
          ? `<a:gs pos="0">${color(fill.backColor ?? fill.color)}</a:gs><a:gs pos="100000">${color(fill.color)}</a:gs>`
          : `<a:gs pos="0">${color(fill.color)}</a:gs><a:gs pos="100000">${color(fill.backColor ?? fill.color)}</a:gs>`;
      return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${ang}" scaled="0"/></a:gradFill>`;
    }
    case "pattern":
    case "solid":
    default:
      return `<a:solidFill>${color(fill.color, fill.opacity)}</a:solidFill>`;
  }
}

function lineXml(line: DocLine | null): string {
  if (!line) return "<a:ln><a:noFill/></a:ln>";
  const join = line.join === "bevel" ? "<a:bevel/>" : line.join === "miter" ? '<a:miter lim="800000"/>' : line.join === "round" ? "<a:round/>" : "";
  const end = (tag: "headEnd" | "tailEnd", arrow: DocLine["headEnd"]): string => (arrow && arrow.type !== "none" ? `<a:${tag} type="${arrow.type}" w="${arrow.width}" len="${arrow.length}"/>` : "");
  return (
    `<a:ln w="${emu(line.width)}"${line.compound !== "sng" ? ` cmpd="${line.compound}"` : ""}${line.cap ? ` cap="${line.cap}"` : ""}>` +
    `<a:solidFill>${color(line.color, line.opacity)}</a:solidFill><a:prstDash val="${line.dash}"/>${join}${end("headEnd", line.headEnd)}${end("tailEnd", line.tailEnd)}</a:ln>`
  );
}

function shadowXml(shape: DocShape): string {
  const shadow = shape.shadow;
  if (!shadow) return "";
  const dist = Math.hypot(shadow.dx, shadow.dy);
  const dir = Math.round((((Math.atan2(shadow.dy, shadow.dx) * 180) / Math.PI + 360) % 360) * 60000);
  return `<a:effectLst><a:outerShdw dist="${emu(dist)}" dir="${dir}" algn="tl" rotWithShape="0">${color(shadow.color, shadow.opacity)}</a:outerShdw></a:effectLst>`;
}
