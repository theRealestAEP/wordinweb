// Sweep page sizes with a comment to find the pagination loop (finding 1). Temporary; untracked.
import { Writable } from "node:stream";
import { renderDocxToPdf } from "../src/pdf/render.js";
import { bytesSource } from "../src/stream/source.js";
import { W_NS, makeDocx } from "../test/helpers.js";

const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
function doc(wTw: number, hTw: number, repeats: number, margin: number): Uint8Array {
  const body =
    `<w:p><w:r><w:t xml:space="preserve">Hi </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>there</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="${wTw}" w:h="${hTw}"/><w:pgMar w:top="${margin}" w:bottom="${margin}" w:left="${margin}" w:right="${margin}"/></w:sectPr>`;
  return makeDocx({
    "word/document.xml": `<?xml version="1.0"?>\n<w:document ${W_NS}><w:body>${body}</w:body></w:document>`,
    "word/comments.xml": `<?xml version="1.0"?>\n<w:comments ${W_NS} ${W14}><w:comment w:id="0" w:author="Ada Reviewer" w:initials="AR" w:date="2026-06-01T10:00:00Z"><w:p><w:r><w:t xml:space="preserve">${"long comment words ".repeat(repeats)}</w:t></w:r></w:p></w:comment></w:comments>`,
  });
}

const sizes = process.argv.slice(2).map(Number);
const [w0, h0, hStep, hEnd, repeats, margin] = sizes;
for (let h = h0; h <= hEnd; h += hStep) {
  let bytes = 0;
  const out = new Writable({
    write(chunk, _enc, cb) {
      bytes += chunk.length;
      if (bytes > 5e6) {
        console.log(`h=${h} runaway: ${bytes} bytes`);
        process.exit(9);
      }
      cb();
    },
  });
  try {
    const result = await renderDocxToPdf(bytesSource(doc(w0, h, repeats, margin)), out, { comments: true });
    console.log(`h=${h} pages ${result.pages}, ${bytes} bytes`);
  } catch (e) {
    console.log(`h=${h} error: ${(e as Error).message}`);
  }
}
