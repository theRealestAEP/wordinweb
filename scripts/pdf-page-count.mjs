#!/usr/bin/env node
/**
 * The PDF renderer's two page totals, for every document, rendered plain and
 * with { comments: true, markup: true }:
 *
 *   node scripts/pdf-page-count.mjs [--jobs N] [--out FILE] PATH...
 *
 * PATH is a .docx or .doc, or a directory of them (not recursed). The
 * counting pass's total (RenderPdfResult.countedPages) is what NUMPAGES
 * prints, the Y of "Page X of Y"; it must equal the pages the rendering pass
 * lays out (RenderPdfResult.pages). The report lists each render whose totals
 * differ or that failed, then the totals; the exit code is 1 when any totals
 * differ. A failed render (a script the renderer cannot shape, a damaged
 * file) is listed but does not fail the check. Needs packages/core built (dist).
 */
import { spawn } from "node:child_process";
import { readdirSync, statSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";
import { createInterface } from "node:readline";
import { Writable } from "node:stream";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = new Map();
const paths = [];
const argv = process.argv.slice(2);
for (let i = 0; i < argv.length; i++) {
  if (!argv[i].startsWith("--")) paths.push(argv[i]);
  else {
    const next = argv[i + 1];
    args.set(argv[i].slice(2), next && !next.startsWith("--") && (argv[i] === "--jobs" || argv[i] === "--out") ? argv[++i] : "true");
  }
}

const MODES = { plain: {}, markup: { comments: true, markup: true } };

async function check(path) {
  const { renderDocxToPdf } = await import(join(root, "packages/core/dist/pdf/render.js"));
  const out = {};
  for (const [mode, options] of Object.entries(MODES)) {
    const discard = new Writable({ write: (_chunk, _encoding, callback) => callback() });
    try {
      const result = await renderDocxToPdf(path, discard, options);
      out[mode] = { pages: result.pages, counted: result.countedPages };
    } catch (error) {
      out[mode] = { error: `${error.code ?? "failed"}: ${error.message}` };
    }
  }
  return { path, ...out };
}

async function worker() {
  console.log = console.error;
  const lines = createInterface({ input: process.stdin });
  for await (const line of lines) process.stdout.write(JSON.stringify(await check(line)) + "\n");
}

async function main() {
  const files = [];
  for (const path of paths) {
    const full = resolve(path);
    if (!statSync(full).isDirectory()) {
      files.push(full);
      continue;
    }
    for (const name of readdirSync(full).sort()) if (/\.docx?$/i.test(name) && statSync(join(full, name)).isFile()) files.push(join(full, name));
  }
  if (files.length === 0) {
    console.error("nothing to check: pass .docx or .doc files or directories of them");
    process.exit(2);
  }
  const jobs = Math.max(1, Number(args.get("jobs") ?? 6));
  const results = [];
  let next = 0;
  const run = () =>
    new Promise((resolveWorker, reject) => {
      const child = spawn(process.execPath, [fileURLToPath(import.meta.url), "--worker"], { stdio: ["pipe", "pipe", "inherit"] });
      const feed = () => {
        if (next < files.length) child.stdin.write(files[next++] + "\n");
        else child.stdin.end();
      };
      createInterface({ input: child.stdout }).on("line", (line) => {
        const result = JSON.parse(line);
        results.push(result);
        for (const mode of Object.keys(MODES)) {
          const r = result[mode];
          if (r.error) console.error(`${basename(result.path)} ${mode}: ${r.error}`);
          else if (r.pages !== r.counted) console.error(`${basename(result.path)} ${mode}: counted ${r.counted} pages, rendered ${r.pages}`);
        }
        if (results.length % 100 === 0) console.error(`${results.length}/${files.length}`);
        feed();
      });
      child.on("error", reject);
      child.on("exit", (code) => (code ? reject(new Error(`worker exited ${code}`)) : resolveWorker()));
      feed();
    });
  await Promise.all(Array.from({ length: Math.min(jobs, files.length) }, run));
  results.sort((a, b) => a.path.localeCompare(b.path));
  const renders = results.flatMap((result) => Object.keys(MODES).map((mode) => result[mode]));
  const differ = renders.filter((r) => !r.error && r.pages !== r.counted).length;
  const failed = renders.filter((r) => r.error).length;
  console.log(`${results.length} documents, ${renders.length} renders: ${differ} with a counted total other than the pages rendered; ${failed} failed`);
  if (args.has("out")) writeFileSync(args.get("out"), JSON.stringify(results, null, 1));
  process.exitCode = differ ? 1 : 0;
}

if (args.has("worker")) worker();
else main();
