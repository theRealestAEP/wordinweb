/**
 * The .doc converter's hooks for pictures and shapes: `special` answers the
 * picture character (U+0001) and the floating shape anchor (U+0008) with a
 * `<w:r><w:drawing>` run, and `parts` collects the media parts those runs
 * refer to.
 */
import { parsePlc } from "./bytes.js";
import type { ConvertHooks, DocReader, ExtraParts } from "./convert.js";
import { FC, tableBlob } from "./fib.js";
import {
  type DocFspa,
  type DocOfficeArt,
  type DocShape,
  type DocTextBox,
  floatingShape,
  parseInlinePicture,
  parseOfficeArtContent,
  parsePlcfSpa,
} from "./drawings.js";
import { anchoredShapeXml, inlinePictureXml } from "./drawings-xml.js";
import type { ArtImage } from "./officeart.js";

const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
/** Extensions the converter's [Content_Types].xml already declares. */
const DEFAULT_EXTENSIONS = new Set(["png", "jpeg", "jpg", "gif", "bmp", "tiff", "emf", "wmf"]);

/** The hooks that write a .doc's pictures and shapes into the converted package. */
export function drawingHooks(reader: DocReader): ConvertHooks {
  const { streams, fib, parts: starts } = reader;
  const art = documentArt(reader);
  const anchors = new Map<number, { fspa: DocFspa; header: boolean }>();
  const mom = tableBlob(fib, streams.table, FC.plcSpaMom);
  for (const { cp, fspa } of parsePlcfSpa(mom, 0, mom.length)) anchors.set(starts.main + cp, { fspa, header: false });
  const hdr = tableBlob(fib, streams.table, FC.plcSpaHdr);
  for (const { cp, fspa } of parsePlcfSpa(hdr, 0, hdr.length)) anchors.set(starts.headers + cp, { fspa, header: true });

  const extra: ExtraParts = { files: {}, documentRels: [], contentTypes: [] };
  const mediaIds = new Map<string, string>();
  let docPrId = 0;
  const addMedia = (image: ArtImage): string => {
    const key = `${image.type}:${image.data.length}:${fnv(image.data)}`;
    let id = mediaIds.get(key);
    if (id) return id;
    const n = mediaIds.size + 1;
    const name = `media/image${n}.${image.ext}`;
    id = `rIdDocImage${n}`;
    extra.files[`word/${name}`] = image.data;
    extra.documentRels.push({ id, type: IMAGE_REL, target: name });
    if (!DEFAULT_EXTENSIONS.has(image.ext)) extra.contentTypes.push({ partName: `/word/${name}`, contentType: image.contentType });
    mediaIds.set(key, id);
    return id;
  };
  const nextId = (): number => ++docPrId;
  const inCode = fieldCodePictures(reader);

  const special: ConvertHooks["special"] = (ch, cp, props, rPr, paragraphs) => {
    const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
    if (ch === 0x01) {
      // A picture in a field's code (Word's QUOTE field keeps an equation's
      // picture there and again in the result) shows only in the result.
      if (props.fData || props.picLocation === undefined || !streams.data || inCode.has(cp)) return undefined;
      const picture = parseInlinePicture(streams.data, props.picLocation, { store: art, delayStream: streams.wordDocument });
      if (!picture || (!picture.image && !picture.shape)) return undefined;
      return `<w:r>${rPrXml}${inlinePictureXml(picture, { addMedia, nextId })}</w:r>`;
    }
    if (ch === 0x08) {
      const anchor = anchors.get(cp);
      if (!anchor) return undefined;
      const floating = floatingShape(art, anchor.fspa, streams.wordDocument, anchor.header ? 1 : 0);
      if (!floating) return undefined;
      const textBoxContent = (shape: DocShape): string | undefined => {
        const range = shape.text && textBoxRange(reader, shape.text, anchor.header);
        return range && range.end > range.start ? paragraphs(range.start, range.end) : undefined;
      };
      const continuedTextBoxStories = art.drawings.find((d) => d.dgglbl === (anchor.header ? 1 : 0))?.continuedTextBoxStories;
      return `<w:r>${rPrXml}${anchoredShapeXml(floating, { addMedia, nextId, textBoxContent, textBoxIdBase: anchor.header ? 0x10000 : 0, continuedTextBoxStories })}</w:r>`;
    }
    return undefined;
  };

  return { special, parts: () => extra };
}

/** The CPs of the picture characters that sit in a field's code (between its begin and separator marks). */
function fieldCodePictures(reader: DocReader): Set<number> {
  const text = reader.pieces.text(0, Math.min(reader.pieces.end, reader.parts.end));
  const out = new Set<number>();
  const inCode: boolean[] = [];
  for (let cp = 0; cp < text.length; cp++) {
    const ch = text.charCodeAt(cp);
    if (ch === 0x13) inCode.push(true);
    else if (ch === 0x14) inCode[inCode.length - 1] = false;
    else if (ch === 0x15) inCode.pop();
    else if (ch === 0x01 && inCode[inCode.length - 1]) out.add(cp);
  }
  return out;
}

/** The document's OfficeArt (fcDggInfo of the Table stream), with the BLIP store's pictures read from the WordDocument stream. */
export function documentArt(reader: DocReader): DocOfficeArt {
  const dgg = tableBlob(reader.fib, reader.streams.table, FC.dggInfo);
  return parseOfficeArtContent(dgg, 0, dgg.length, reader.streams.wordDocument);
}

/**
 * The CP range of a text box story in the document's CP space, without the
 * story's guard paragraph mark. A linked chain shares one story.
 */
export function textBoxRange(reader: DocReader, text: DocTextBox, header: boolean): { start: number; end: number } | undefined {
  const { fib, streams, parts: starts } = reader;
  const stories = parsePlc(tableBlob(fib, streams.table, header ? FC.plcfHdrtxbxTxt : FC.plcftxbxTxt), 22);
  if (text.story < 0 || text.story + 1 >= stories.cps.length) return undefined;
  const base = header ? starts.headerTextboxes : starts.textboxes;
  const start = stories.cps[text.story];
  const end = stories.cps[text.story + 1];
  // The story's last character is its guard paragraph mark.
  return { start: base + start, end: base + Math.max(start, end - 1) };
}

function fnv(bytes: Uint8Array): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) h = Math.imul(h ^ bytes[i], 0x01000193) >>> 0;
  return h.toString(16);
}
