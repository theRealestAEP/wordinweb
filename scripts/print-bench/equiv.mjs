// Output equivalence: every fixture printed the old way (renderForPrint) and
// the new way (openPrintStream), each rendered page's DOM hashed and compared.
//   node equiv.mjs [--jobs N] [--batch 25] [--out results.json] <docx or dir>...
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { firstDiff, launch, runNew, runOld } from "./lib.mjs";

const args = process.argv.slice(2);
const opt = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const value = args[i + 1];
  args.splice(i, 2);
  return value;
};
const jobs = Number(opt("jobs", "4"));
const batch = Number(opt("batch", "25"));
const out = opt("out", null);
// A mismatched page's two DOMs are written here for a look.
const diffDir = join(tmpdir(), "print-bench-diffs");
mkdirSync(diffDir, { recursive: true });

const files = [];
for (const arg of args) {
  if (statSync(arg).isDirectory()) {
    for (const name of readdirSync(arg).sort()) if (name.endsWith(".docx")) files.push(join(arg, name));
  } else files.push(arg);
}

async function compareOne(browser, file) {
  const bytes = new Uint8Array(readFileSync(file));
  const name = basename(file);
  const result = { name, oldPages: null, newPages: null, mismatches: [], error: null, oldMs: 0, newMs: 0, newCountMs: 0 };
  const oldContext = await browser.newContext();
  const newContext = await browser.newContext();
  try {
    const oldPage = await oldContext.newPage();
    const old = await runOld(oldPage, bytes, { batch });
    result.oldPages = old.pageCount;
    result.oldMs = old.layoutMs;
    const newPage = await newContext.newPage();
    const t0 = performance.now();
    const fresh = await runNew(newPage, bytes, {
      batch,
      onBatch: async (first, hashes, api) => {
        for (let i = 0; i < hashes.length; i++) {
          const index = first + i;
          if (old.hashes[index] === hashes[i]) continue;
          if (result.mismatches.length >= 3) {
            result.mismatches.push({ page: index + 1 });
            continue;
          }
          const [a, b] = [await old.pageHtml(index), await api.pageHtml(index)];
          const diff = a !== null && b !== null ? firstDiff(a, b) : { at: -1, old: String(a).slice(0, 100), new: String(b).slice(0, 100) };
          writeFileSync(join(diffDir, `${name}.p${index + 1}.old.html`), a ?? "");
          writeFileSync(join(diffDir, `${name}.p${index + 1}.new.html`), b ?? "");
          result.mismatches.push({ page: index + 1, ...diff });
        }
      },
    });
    result.newMs = Math.round(performance.now() - t0);
    result.newCountMs = fresh.countMs;
    result.newPages = fresh.pageCount;
    if (fresh.hashes.length !== fresh.pageCount) result.error = `streamed ${fresh.hashes.length} pages, counted ${fresh.pageCount}`;
  } catch (error) {
    result.error = String(error && error.message ? error.message : error).slice(0, 300);
  } finally {
    await oldContext.close();
    await newContext.close();
  }
  return result;
}

const browser = await launch();
const results = [];
let next = 0;
const worker = async () => {
  while (next < files.length) {
    const file = files[next++];
    const r = await compareOne(browser, file);
    results.push(r);
    const status = r.error ? `ERROR ${r.error}` : r.oldPages !== r.newPages ? `PAGES old ${r.oldPages} new ${r.newPages}` : r.mismatches.length ? `DIFF pages ${r.mismatches.map((m) => m.page).join(",")}` : "ok";
    console.log(`${r.name.padEnd(48)} ${String(r.oldPages).padStart(4)} pages  old ${String(r.oldMs).padStart(6)}ms  new ${String(r.newMs).padStart(6)}ms  ${status}`);
  }
};
await Promise.all(Array.from({ length: Math.min(jobs, files.length) }, worker));
await browser.close();
results.sort((a, b) => a.name.localeCompare(b.name));
const ok = results.filter((r) => !r.error && r.oldPages === r.newPages && r.mismatches.length === 0).length;
console.log(`\n${ok}/${results.length} identical`);
for (const r of results) {
  if (r.error) console.log(`  ${r.name}: ${r.error}`);
  else if (r.oldPages !== r.newPages) console.log(`  ${r.name}: ${r.oldPages} pages old, ${r.newPages} new`);
  for (const m of r.mismatches) if (m.old !== undefined) console.log(`  ${r.name} p${m.page} @${m.at}\n    old: ${JSON.stringify(m.old)}\n    new: ${JSON.stringify(m.new)}`);
}
if (out) writeFileSync(out, JSON.stringify(results, null, 1));
