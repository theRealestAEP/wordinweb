// A separate process measures conversion's synchronous allocations in the
// OS RSS high-water mark, followed by parsing, layout and picture writing.
import { Writable } from "node:stream";
import { inflateSync } from "node:zlib";
import { renderDocxToPdf } from "../src/pdf/render.js";

const [control, target] = process.argv.slice(2);
async function render(path: string) {
  let images = 0;
  let content = false;
  const out = new Writable({ write(chunk: Buffer, _encoding, done) {
    if (content) {
      // The generated WMFs each draw this line in their local coordinates.
      images += (inflateSync(chunk).toString("latin1").match(/10 10 m 90 40 l/g) ?? []).length;
      content = false;
    } else content = chunk.toString("latin1").startsWith("<<  /Filter /FlateDecode /Length ");
    done();
  } });
  return { result: await renderDocxToPdf(path, out), images };
}
await render(control);
(globalThis as { gc?: () => void }).gc!();
const start = process.memoryUsage().rss;
const rendered = await render(target);
console.log(JSON.stringify({
  rssRiseMB: Math.round((process.resourceUsage().maxRSS * 1024 - start) / (1024 * 1024)),
  pages: rendered.result.pages,
  unsupportedImages: rendered.result.unsupportedImages,
  images: rendered.images,
  ...process.memoryUsage(),
}));
