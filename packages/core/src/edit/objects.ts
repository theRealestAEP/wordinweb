import { DocxDocument } from "../docx.js";
import { buildOlePackage } from "../parse/ole.js";
import { type XmlElement, localName } from "../xml.js";

const EMU_PER_PX = 9525;
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_AM3D = "http://schemas.microsoft.com/office/drawing/2017/model3d";
const NS_WP15 = "http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing";

export interface Model3DInsert {
  data: Uint8Array;
  poster: Uint8Array;
  posterExt?: "png" | "jpeg";
  alt?: string;
}

export interface Model3DRotation {
  x: number;
  y: number;
  z: number;
}

export interface WebVideoInsert {
  url: string;
  poster: Uint8Array;
  posterExt?: "png" | "jpeg";
  width?: number;
  height?: number;
}

export interface EmbeddedObjectInsert {
  data: Uint8Array;
  filename: string;
  poster: Uint8Array;
  posterExt?: "png" | "jpeg";
}

function el(name: string, attrs: Record<string, string> = {}, children: XmlElement[] = [], text = ""): XmlElement {
  return { name, attrs, children, text };
}

function prefixOf(node: XmlElement): string {
  return node.name.includes(":") ? node.name.slice(0, node.name.indexOf(":") + 1) : "";
}

function insertAfterCaretRun(doc: DocxDocument, caretT: XmlElement, content: XmlElement): boolean {
  const caretRun = doc.findParentOf(caretT);
  const parent = caretRun && doc.findParentOf(caretRun);
  if (!caretRun || !parent || localName(caretRun.name) !== "r") return false;
  const w = prefixOf(caretRun);
  parent.children.splice(parent.children.indexOf(caretRun) + 1, 0, el(`${w}r`, {}, [content]));
  doc.refresh();
  return true;
}

function pictureGraphic(relId: string, id: string, cx: string, cy: string, webVideo?: string): XmlElement {
  const blipChildren = webVideo ? [
    el("a:extLst", {}, [
      el("a:ext", { uri: "{9E29B28B-C9C2-4B96-BA92-6C9A3D8F4A6A}" }, [
        el("wp15:webVideoPr", {
          "xmlns:wp15": NS_WP15,
          embeddedHtml: webVideo,
          w: "640",
          h: "360",
        }),
      ]),
    ]),
  ] : [];
  return el("a:graphic", { "xmlns:a": NS_A }, [
    el("a:graphicData", { uri: NS_PIC }, [
      el("pic:pic", { "xmlns:pic": NS_PIC }, [
        el("pic:nvPicPr", {}, [
          el("pic:cNvPr", { id, name: `Picture ${id}` }),
          el("pic:cNvPicPr", {}, [el("a:picLocks", { noChangeAspect: "1" })]),
        ]),
        el("pic:blipFill", {}, [
          el("a:blip", { "xmlns:r": NS_R, "r:embed": relId }, blipChildren),
          el("a:stretch", {}, [el("a:fillRect")]),
        ]),
        el("pic:spPr", {}, [
          el("a:xfrm", {}, [el("a:off", { x: "0", y: "0" }), el("a:ext", { cx, cy })]),
          el("a:prstGeom", { prst: "rect" }, [el("a:avLst")]),
        ]),
      ]),
    ]),
  ]);
}

function inlineDrawing(id: string, name: string, descr: string, cx: string, cy: string, graphic: XmlElement): XmlElement {
  return el("w:drawing", {}, [
    el("wp:inline", { "xmlns:wp": NS_WP, distT: "0", distB: "0", distL: "0", distR: "0" }, [
      el("wp:extent", { cx, cy }),
      el("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" }),
      el("wp:docPr", { id, name, descr }),
      el("wp:cNvGraphicFramePr", {}, [el("a:graphicFrameLocks", { "xmlns:a": NS_A, noChangeAspect: "1" })]),
      graphic,
    ]),
  ]);
}

/** Insert an editable native Office 2019 3D model with a compatibility poster. */
export function insertModel3DAt(doc: DocxDocument, caretT: XmlElement, input: Model3DInsert): boolean {
  const posterRelId = doc.addImageResource(input.poster, input.posterExt ?? "png");
  const model = doc.addModel3DResource(input.data);
  const id = String(doc.nextDrawingId());
  const width = 320;
  const height = 220;
  const cx = String(width * EMU_PER_PX);
  const cy = String(height * EMU_PER_PX);
  const modelGraphic = el("a:graphic", { "xmlns:a": NS_A }, [
    el("a:graphicData", { uri: NS_AM3D }, [
      el("am3d:model3d", { "xmlns:am3d": NS_AM3D, "xmlns:r": NS_R, "r:embed": model.relId }, [
        el("am3d:spPr", {}, [
          el("a:xfrm", {}, [el("a:off", { x: "0", y: "0" }), el("a:ext", { cx, cy })]),
          el("a:prstGeom", { prst: "rect" }, [el("a:avLst")]),
        ]),
        el("am3d:camera", {}, [
          el("am3d:pos", { x: "0", y: "0", z: "77741944" }),
          el("am3d:up", { dx: "0", dy: "36000000", dz: "0" }),
          el("am3d:lookAt", { x: "0", y: "0", z: "0" }),
          el("am3d:perspective", { fov: "2700000" }),
        ]),
        el("am3d:trans", {}, [
          el("am3d:meterPerModelUnit", { n: "1", d: "1" }),
          el("am3d:preTrans", { dx: "0", dy: "0", dz: "0" }),
          el("am3d:scale", {}, [
            el("am3d:sx", { n: "1000000", d: "1000000" }),
            el("am3d:sy", { n: "1000000", d: "1000000" }),
            el("am3d:sz", { n: "1000000", d: "1000000" }),
          ]),
          el("am3d:rot", { ax: "0", ay: "0", az: "0" }),
          el("am3d:postTrans", { dx: "0", dy: "0", dz: "0" }),
        ]),
        el("am3d:raster", { rName: "Office3DRenderer", rVer: "16.0" }, [
          el("am3d:blip", { "r:embed": posterRelId }),
        ]),
        el("am3d:objViewport", { viewportSz: cx }),
        el("am3d:ambientLight", {}, [
          el("am3d:clr", {}, [el("a:scrgbClr", { r: "50000", g: "50000", b: "50000" })]),
          el("am3d:illuminance", { n: "500000", d: "1000000" }),
        ]),
        el("am3d:ptLight", { rad: "0" }, [
          el("am3d:clr", {}, [el("a:scrgbClr", { r: "100000", g: "100000", b: "100000" })]),
          el("am3d:intensity", { n: "1000000", d: "1000000" }),
          el("am3d:pos", { x: "30000000", y: "50000000", z: "50000000" }),
        ]),
      ]),
    ]),
  ]);
  const alt = input.alt ?? "3D model";
  const choice = inlineDrawing(id, `3D Model ${id}`, alt, cx, cy, modelGraphic);
  const fallback = inlineDrawing(id, `3D Model ${id}`, alt, cx, cy, pictureGraphic(posterRelId, id, cx, cy));
  const alternate = el("mc:AlternateContent", { "xmlns:mc": NS_MC }, [
    el("mc:Choice", { Requires: "am3d", "xmlns:am3d": NS_AM3D }, [choice]),
    el("mc:Fallback", {}, [fallback]),
  ]);
  return insertAfterCaretRun(doc, caretT, alternate);
}

/** Save a 3D model's native X/Y/Z orientation in DrawingML angle units. */
export function setModel3DRotation(
  doc: DocxDocument,
  drawingEl: XmlElement,
  rotation: Model3DRotation,
): boolean {
  let model: XmlElement | undefined;
  const visit = (node: XmlElement): void => {
    if (model) return;
    if (localName(node.name) === "model3d") model = node;
    else for (const child of node.children) visit(child);
  };
  visit(drawingEl);
  if (!model) return false;
  const trans = model.children.find((child) => localName(child.name) === "trans");
  if (!trans) return false;
  let rot = trans.children.find((child) => localName(child.name) === "rot");
  if (!rot) {
    rot = el("am3d:rot");
    const post = trans.children.findIndex((child) => localName(child.name) === "postTrans");
    trans.children.splice(post === -1 ? trans.children.length : post, 0, rot);
  }
  const angle = (degrees: number): string => {
    const normalized = ((degrees % 360) + 360) % 360;
    return String(Math.round(normalized * 60000));
  };
  rot.attrs.ax = angle(rotation.x);
  rot.attrs.ay = angle(rotation.y);
  rot.attrs.az = angle(rotation.z);
  doc.refresh();
  return true;
}

export function normalizeWebVideoUrl(value: string): string | null {
  try {
    const url = new URL(value.trim());
    if (url.protocol !== "http:" && url.protocol !== "https:") return null;
    if (url.hostname === "youtu.be") return `https://www.youtube.com/embed/${url.pathname.slice(1)}`;
    if (url.hostname.endsWith("youtube.com") && url.pathname === "/watch") {
      const id = url.searchParams.get("v");
      if (id) return `https://www.youtube.com/embed/${id}`;
    }
    if (url.hostname.endsWith("vimeo.com") && /^\/\d+\/?$/.test(url.pathname)) {
      return `https://player.vimeo.com/video/${url.pathname.replace(/\//g, "")}`;
    }
    return url.href;
  } catch {
    return null;
  }
}

/** Insert Word's native online-video picture extension. */
export function insertWebVideoAt(doc: DocxDocument, caretT: XmlElement, input: WebVideoInsert): boolean {
  const url = normalizeWebVideoUrl(input.url);
  if (!url) return false;
  const posterRelId = doc.addImageResource(input.poster, input.posterExt ?? "png");
  const width = input.width ?? 320;
  const height = input.height ?? 180;
  const cx = String(Math.round(width * EMU_PER_PX));
  const cy = String(Math.round(height * EMU_PER_PX));
  const id = String(doc.nextDrawingId());
  const embeddedHtml = `<iframe width="640" height="360" src="${url}" frameborder="0" allowfullscreen=""></iframe>`;
  return insertAfterCaretRun(
    doc,
    caretT,
    inlineDrawing(id, `Online Video ${id}`, url, cx, cy, pictureGraphic(posterRelId, id, cx, cy, embeddedHtml)),
  );
}

/** Insert an arbitrary file as a native OLE Package with a VML preview. */
export function insertEmbeddedObjectAt(doc: DocxDocument, caretT: XmlElement, input: EmbeddedObjectInsert): boolean {
  const posterRelId = doc.addImageResource(input.poster, input.posterExt ?? "png");
  const wordDocument = input.filename.toLowerCase().endsWith(".docx");
  const object = wordDocument
    ? doc.addEmbeddedWordDocumentResource(input.data)
    : doc.addEmbeddedObjectResource(buildOlePackage(input.data, input.filename));
  const id = String(doc.nextDrawingId());
  const vmlId = 1024 + Number(id);
  const shapeTypeId = `_x0000_t${vmlId}`;
  const shapeId = `_x0000_i${vmlId}`;
  const width = 320;
  const height = 180;
  const objectEl = el("w:object", {
    "xmlns:w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "xmlns:v": "urn:schemas-microsoft-com:vml",
    "xmlns:o": "urn:schemas-microsoft-com:office:office",
    "xmlns:r": NS_R,
    "w:dxaOrig": String(width * 15),
    "w:dyaOrig": String(height * 15),
  }, [
    el("v:shapetype", {
      id: shapeTypeId,
      coordsize: "21600,21600",
      "o:spt": "75",
      "o:preferrelative": "t",
      path: "m@4@5l@4@11@9@11@9@5xe",
      filled: "f",
      stroked: "f",
    }, [
      el("v:stroke", { joinstyle: "miter" }),
      el("v:formulas", {}, [
        el("v:f", { eqn: "if lineDrawn pixelLineWidth 0" }),
        el("v:f", { eqn: "sum @0 1 0" }),
        el("v:f", { eqn: "sum 0 0 @1" }),
        el("v:f", { eqn: "prod @2 1 2" }),
        el("v:f", { eqn: "prod @3 21600 pixelWidth" }),
        el("v:f", { eqn: "prod @3 21600 pixelHeight" }),
        el("v:f", { eqn: "sum @0 0 1" }),
        el("v:f", { eqn: "prod @6 1 2" }),
        el("v:f", { eqn: "prod @7 21600 pixelWidth" }),
        el("v:f", { eqn: "sum @8 21600 0" }),
        el("v:f", { eqn: "prod @7 21600 pixelHeight" }),
        el("v:f", { eqn: "sum @10 21600 0" }),
      ]),
      el("v:path", { "o:extrusionok": "f", gradientshapeok: "t", "o:connecttype": "rect" }),
      el("o:lock", { "v:ext": "edit", aspectratio: "t" }),
    ]),
    el("v:shape", {
      id: shapeId,
      type: `#${shapeTypeId}`,
      style: `width:${width * 0.75}pt;height:${height * 0.75}pt`,
      "o:ole": "",
    }, [el("v:imagedata", { "r:id": posterRelId, "o:title": input.filename })]),
    el("o:OLEObject", {
      Type: "Embed",
      ProgID: wordDocument ? "Word.Document.12" : "Package",
      ShapeID: shapeId,
      DrawAspect: "Content",
      ObjectID: `_${1000000000 + Number(id)}`,
      "r:id": object.relId,
    }, wordDocument ? [el("o:FieldCodes", {}, [], "\\s")] : []),
  ]);
  return insertAfterCaretRun(doc, caretT, objectEl);
}
