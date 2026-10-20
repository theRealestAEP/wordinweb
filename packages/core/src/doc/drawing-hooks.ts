/**
 * The .doc converter's hooks for pictures and shapes: `special` answers the
 * picture character (U+0001) and the floating shape anchor (U+0008) with a
 * `<w:r><w:drawing>` run, and `parts` collects the media parts those runs
 * refer to.
 */
import { parsePlc, u16 } from "./bytes.js";
import type { ConvertHooks, DocReader } from "./convert.js";
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

export interface DrawingHookOptions {
  /**
   * The block XML (w:p, w:tbl) of the text at [start, end) in the
   * document's CP space, for a text box's content. Without it a text box
   * holds one empty paragraph.
   */
  storyXml?: (start: number, end: number) => string;
}

export interface DrawingHooks extends Pick<ConvertHooks, "special" | "parts"> {
  parts: NonNullable<ConvertHooks["parts"]>;
  /** The document's OfficeArt, parsed once. */
  art: DocOfficeArt;
  /** The CP range of a text box's text in the document's CP space, when its story exists. */
  textBoxRange: (text: DocTextBox, header: boolean) => { start: number; end: number } | undefined;
}

const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
/** Extensions the converter's [Content_Types].xml already declares. */
const DEFAULT_EXTENSIONS = new Set(["png", "jpeg", "jpg", "gif", "bmp", "tiff", "emf", "wmf"]);

export function drawingHooks(reader: DocReader, options: DrawingHookOptions = {}): DrawingHooks {
  const { streams, fib, parts: starts } = reader;
  const dgg = tableBlob(fib, streams.table, FC.dggInfo);
  const art = parseOfficeArtContent(dgg, 0, dgg.length, streams.wordDocument);
  const anchors = new Map<number, { fspa: DocFspa; header: boolean }>();
  const mom = tableBlob(fib, streams.table, FC.plcSpaMom);
  for (const { cp, fspa } of parsePlcfSpa(mom, 0, mom.length)) anchors.set(starts.main + cp, { fspa, header: false });
  const hdr = tableBlob(fib, streams.table, FC.plcSpaHdr);
  for (const { cp, fspa } of parsePlcfSpa(hdr, 0, hdr.length)) anchors.set(starts.headers + cp, { fspa, header: true });

  const parts: DrawingHooks["parts"] = { files: {}, documentRels: [], contentTypes: [] };
  const mediaIds = new Map<string, string>();
  let mediaCount = 0;
  let docPrId = 0;
  const addMedia = (image: ArtImage): string => {
    const key = `${image.type}:${image.data.length}:${fnv(image.data)}`;
    let id = mediaIds.get(key);
    if (id) return id;
    mediaCount++;
    const name = `media/image${mediaCount}.${image.ext}`;
    id = `rIdDocImage${mediaCount}`;
    parts.files[`word/${name}`] = image.data;
    parts.documentRels.push({ id, type: IMAGE_REL, target: name });
    if (!DEFAULT_EXTENSIONS.has(image.ext)) parts.contentTypes.push({ partName: `/word/${name}`, contentType: image.contentType });
    mediaIds.set(key, id);
    return id;
  };

  const textBoxRange = (text: DocTextBox, header: boolean): { start: number; end: number } | undefined => {
    const stories = parsePlc(tableBlob(fib, streams.table, header ? FC.plcfHdrtxbxTxt : FC.plcftxbxTxt), 22);
    if (text.story < 0 || text.story + 1 >= stories.cps.length) return undefined;
    const base = header ? starts.headerTextboxes : starts.textboxes;
    let start = stories.cps[text.story];
    let end = stories.cps[text.story + 1];
    // A linked chain's story is split per box by the break table (Tbkd.itxbxs names the story).
    const breaks = parsePlc(tableBlob(fib, streams.table, header ? FC.plcfTxbxHdrBkd : FC.plcfTxbxBkd), 10);
    const own: number[] = [];
    for (let i = 0; i < breaks.data.length; i++) if (u16(breaks.data[i], 0) === text.story) own.push(i);
    if (own.length > 1 && text.sequence < own.length) {
      const i = own[text.sequence];
      start = Math.max(start, breaks.cps[i]);
      end = Math.min(end, breaks.cps[i + 1]);
    }
    // The story's last character is its guard paragraph mark.
    return { start: base + start, end: base + Math.max(start, end - 1) };
  };

  const xmlOptions = {
    addMedia,
    nextId: () => ++docPrId,
    textBoxContent: (shape: DocShape): string | undefined => {
      if (!shape.text || !options.storyXml) return undefined;
      const range = textBoxRange(shape.text, currentHeader);
      return range ? options.storyXml(range.start, range.end) : undefined;
    },
  };
  let currentHeader = false;

  const special: ConvertHooks["special"] = (ch, cp, props, rPr) => {
    const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : "";
    if (ch === 0x01) {
      if (props.fData || props.picLocation === undefined || !streams.data) return undefined;
      const picture = parseInlinePicture(streams.data, props.picLocation, { store: art, delayStream: streams.wordDocument });
      if (!picture || (!picture.image && !picture.shape)) return undefined;
      currentHeader = cp >= starts.headers && cp < starts.comments;
      return `<w:r>${rPrXml}${inlinePictureXml(picture, xmlOptions)}</w:r>`;
    }
    if (ch === 0x08) {
      const anchor = anchors.get(cp);
      if (!anchor) return undefined;
      currentHeader = anchor.header;
      const floating = floatingShape(art, anchor.fspa, streams.wordDocument, anchor.header ? 1 : 0);
      if (!floating) return undefined;
      return `<w:r>${rPrXml}${anchoredShapeXml(floating, xmlOptions)}</w:r>`;
    }
    return undefined;
  };

  return { special, parts, art, textBoxRange };
}

function fnv(bytes: Uint8Array): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < bytes.length; i++) h = Math.imul(h ^ bytes[i], 0x01000193) >>> 0;
  return h.toString(16);
}
