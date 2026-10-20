/**
 * Test-side helpers over the .doc reader: the document's OfficeArt, its
 * shape anchors, and every picture character's Data-stream offset.
 */
import { type DocReader, openDoc } from "../src/doc/convert.js";
import { type DocFspa, type DocOfficeArt, parseOfficeArtContent, parsePlcfSpa } from "../src/doc/drawings.js";
import { FC, tableBlob } from "../src/doc/fib.js";
import { applyRunSprms, newRunProps, toggleValue } from "../src/doc/sprm.js";

export { openDoc };

export function officeArt(reader: DocReader): DocOfficeArt {
  const blob = tableBlob(reader.fib, reader.streams.table, FC.dggInfo);
  return parseOfficeArtContent(blob, 0, blob.length, reader.streams.wordDocument);
}

/** The shape anchors of the main document or the header story, with their CP in the story. */
export function shapeAnchors(reader: DocReader, which: "main" | "header" = "main"): { cp: number; fspa: DocFspa }[] {
  const blob = tableBlob(reader.fib, reader.streams.table, which === "main" ? FC.plcSpaMom : FC.plcSpaHdr);
  return parsePlcfSpa(blob, 0, blob.length);
}

/** The Data-stream offset of every picture character (U+0001 with fSpec, no fData) in the document, in text order. */
export function pictureLocations(reader: DocReader): { cp: number; offset: number; ole: boolean }[] {
  const out: { cp: number; offset: number; ole: boolean }[] = [];
  const end = Math.min(reader.pieces.end, reader.parts.end);
  const text = reader.pieces.text(0, end);
  for (const run of reader.props.charRuns(0, end)) {
    const props = newRunProps();
    applyRunSprms(run.grpprl, props);
    if (props.picLocation === undefined || props.fData || toggleValue(props, "fSpec", reader.ctx, 0) !== true) continue;
    for (let cp = run.cp; cp < run.end; cp++) {
      if (text.charCodeAt(cp) === 0x01) out.push({ cp, offset: props.picLocation, ole: !!props.fOle2 });
    }
  }
  return out;
}
