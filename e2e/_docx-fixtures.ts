/**
 * Shared .docx fixtures for the object-interaction specs.
 *
 * Both #153 (a dragged object landing where it was dropped) and #161
 * (selecting an object below the last line of text) need the SAME awkward
 * document: a floating 3D model inside an indented sidebar table, and a
 * floating picture as the control. They differ only in where the picture is
 * anchored, so the builder takes that as an option rather than being copied.
 */
import { deflateSync } from "node:zlib";
import { zipSync, strToU8 } from "fflate";

/** A real, decodable PNG — the poster is decoded by the browser. */
export function makePng(width: number, height: number): Buffer {
  const crcTable: number[] = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crcTable[n] = c >>> 0;
  }
  const crc32 = (buf: Buffer): number => {
    let c = 0xffffffff;
    for (const byte of buf) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type: string, data: Buffer): Buffer => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([len, body, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.alloc(height * (1 + width * 3));
  for (let y = 0; y < height; y++) {
    const row = y * (1 + width * 3);
    raw[row] = 0;
    for (let x = 0; x < width; x++) {
      const p = row + 1 + x * 3;
      raw[p] = (x * 255) / width;
      raw[p + 1] = (y * 255) / height;
      raw[p + 2] = 128;
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const EMU_PER_PX = 9525;
const px = (n: number) => String(Math.round(n * EMU_PER_PX));

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"',
  'mc:Ignorable="w14 wp14"',
].join(" ");

/**
 * A wp:anchor around `graphic`. `hRel` is the horizontal frame: Word anchors
 * a 3D model to the COLUMN and a dropped picture to the MARGIN, and the two
 * origins differ — which is exactly the difference the move path has to
 * absorb, so the fixture keeps them different.
 */
function anchor(graphic: string, opts: { hRel: string; x: number; y: number; w: number; h: number; wrap: string }): string {
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0"` +
    ` relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="${opts.hRel}"><wp:posOffset>${px(opts.x)}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${px(opts.y)}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${px(opts.w)}" cy="${px(opts.h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `${opts.wrap}` +
    `<wp:docPr id="${opts.hRel === "column" ? 11 : 12}" name="Object ${opts.hRel}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `${graphic}` +
    `</wp:anchor></w:drawing>`
  );
}

const MODEL_GRAPHIC =
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">` +
  `<am3d:model3d r:embed="rId5">` +
  `<am3d:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${px(144)}" cy="${px(144)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></am3d:spPr>` +
  `<am3d:raster><am3d:blip r:embed="rId4"/></am3d:raster>` +
  `</am3d:model3d></a:graphicData></a:graphic>`;

const PIC_GRAPHIC =
  `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="control.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${px(96)}" cy="${px(96)}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
  `</a:graphicData></a:graphic>`;

/**
 * One page holding a floating 3D MODEL in an indented sidebar table and a
 * floating PICTURE anchored in ordinary body text above it — the object that
 * reproduces, and the control that does not.
 */
export function modelAndPictureDocx({ pictureAfterTable = false } = {}): Buffer {
  const para = (text: string, extra = "") =>
    `<w:p>${extra ? `<w:r>${extra}</w:r>` : ""}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  const model = anchor(MODEL_GRAPHIC, { hRel: "column", x: 6, y: 0, w: 144, h: 144, wrap: "<wp:wrapNone/>" });
  // Clear ABOVE and to the right of the sidebar, so a click aimed at one
  // object can never land on the other, before or after three +60px drags.
  const picture = anchor(PIC_GRAPHIC, { hRel: "margin", x: 380, y: 0, w: 96, h: 96, wrap: '<wp:wrapSquare wrapText="bothSides"/>' });
  const filler = Array.from({ length: 10 }, (_, i) => para(`Body line ${i} of the document.`));
  // THE MODEL LIVES IN A TABLE CELL, as Word's cover-letter template puts it
  // (an indented sidebar). That is what makes this document reproduce and a
  // plain body float not: the cell gives the anchor its own origin on BOTH
  // axes — indented from the page margin, and down at the cell's paragraph —
  // and a page-relative write lands that far off on each.
  const sidebar =
    `<w:tbl><w:tblPr><w:tblW w:w="4320" w:type="dxa"/><w:tblInd w:w="360" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>` +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>` +
    para("Sidebar", model) +
    `</w:tc></w:tr></w:tbl>`;
  // WHERE THE PICTURE IS ANCHORED IS THE WHOLE VARIABLE. After the sidebar it
  // sits BELOW the last line of body text on the page, which is the condition
  // #161 turns on — Word's click-and-type reads a press there as a click in
  // empty space. Before the sidebar there are later lines and it does not.
  const body =
    para("Curriculum vitae") +
    (pictureAfterTable
      ? filler.slice(0, 5).join("") + sidebar + filler.slice(5, 8).join("") +
        para("Anchor for the control picture.", picture) + filler.slice(8).join("")
      : para("Anchor for the control picture.", picture) + filler.slice(0, 3).join("") +
        sidebar + filler.slice(3).join("")) +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>`;

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Default Extension="png" ContentType="image/png"/>` +
        `<Default Extension="glb" ContentType="model/gltf-binary"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
        `</Types>`,
    ),
    "_rels/.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
        `</Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`,
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
        `<Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/model3d" Target="media/model3d1.glb"/>` +
        `</Relationships>`,
    ),
    "word/media/image1.png": new Uint8Array(makePng(64, 64)),
    // A minimal glTF-binary header: enough for the renderer to mount a viewer.
    "word/media/model3d1.glb": new Uint8Array([0x67, 0x6c, 0x54, 0x46, 2, 0, 0, 0, 12, 0, 0, 0]),
  };
  return Buffer.from(zipSync(files));
}

