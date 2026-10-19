#!/usr/bin/env node
/**
 * Every comment and reply in a document's comments.xml against the cards the
 * PDF renderer prints with { comments: true, markup: true }:
 *
 *   node scripts/pdf-comment-cards.mjs [--jobs N] [--out FILE] [--verbose] PATH...
 *
 * PATH is a .docx or a directory of them (not recursed). Every comment some
 * part of the document marks (a range or a reference) is expected; one
 * counts as printed when pdftotext finds its text (whitespace aside) and
 * its author's name in the PDF. The report lists each document with a
 * shortfall or an error, then the totals; the exit code is 1 when any
 * comment is missing. Needs packages/core built (dist) and pdftotext on PATH.
 */
import { execFileSync, spawn } from "node:child_process";
import { createWriteStream, mkdtempSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";
import { strFromU8, unzipSync } from "fflate";

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

/** The comments the document marks (a comments.xml entry no part refers to
 * is not in the document; Word shows none): id, author, text, and whether
 * it is a reply. */
function commentsOf(docx) {
  const parts = unzipSync(docx, { filter: (file) => /^word\/(comments|commentsExtended|people|document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(file.name) });
  const xml = (name) => (parts[name] ? strFromU8(parts[name]) : "");
  const marked = new Set();
  for (const name of Object.keys(parts)) {
    if (/^word\/(comments|commentsExtended|people)\.xml$/.test(name)) continue;
    for (const m of xml(name).matchAll(/<w:(?:commentRangeStart|commentReference)\b[^>]*\bw:id="([^"]*)"/g)) marked.add(m[1]);
  }
  const decode = (s) => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n))).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, "&");
  const replyParas = new Set([...xml("word/commentsExtended.xml").matchAll(/<w15:commentEx\b([^>]*)>/g)].filter((m) => /paraIdParent=/.test(m[1])).map((m) => /w15:paraId="([^"]*)"/.exec(m[1])?.[1]));
  // The card names an author as people.xml does (see comments.ts).
  const names = new Map();
  for (const m of xml("word/people.xml").matchAll(/<w15:person\b([^>]*)>([\s\S]*?)<\/w15:person>/g)) {
    const author = decode(/w15:author="([^"]*)"/.exec(m[1])?.[1] ?? "");
    const userId = decode(/w15:userId="([^"]*)"/.exec(m[2])?.[1] ?? "");
    if (!author) continue;
    names.set(author, author);
    if (userId) names.set(userId, author);
    const address = /^S::([^:]+)::/.exec(userId)?.[1];
    if (address) names.set(address, author);
  }
  const out = [];
  for (const m of xml("word/comments.xml").matchAll(/<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g)) {
    const attrs = m[1];
    const body = m[2];
    const id = /w:id="([^"]*)"/.exec(attrs)?.[1] ?? "";
    if (!marked.has(id)) continue;
    const carried = decode(/w:author="([^"]*)"/.exec(attrs)?.[1] ?? "");
    const author = names.get(carried) ?? carried;
    const text = [...body.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>|<w:t\b[^>]*\/>/g)].map((t) => decode(t[1] ?? "")).join("");
    const paraIds = [...body.matchAll(/w14:paraId="([^"]*)"/g)].map((p) => p[1]);
    out.push({ id, author, text, reply: paraIds.some((p) => replyParas.has(p)) });
  }
  return out;
}

async function check(path) {
  const { renderDocxToPdf } = await import(join(root, "packages/core/dist/pdf/render.js"));
  const dir = mkdtempSync(join(process.env.TMPDIR ?? tmpdir(), "comment-cards-"));
  const pdf = join(dir, "doc.pdf");
  // A render refused before its first write leaves the stream to open after the directory is gone.
  const out = createWriteStream(pdf);
  out.on("error", () => undefined);
  let comments = [];
  try {
    comments = commentsOf(readFileSync(path));
    if (comments.length === 0) return { path, comments: 0, replies: 0, printed: 0, missing: [] };
    const result = await renderDocxToPdf(path, out, { comments: true, markup: true });
    const text = execFileSync("pdftotext", ["-raw", pdf, "-"], { encoding: "utf8", maxBuffer: 1 << 28, stdio: ["ignore", "pipe", "ignore"] });
    const squeezed = text.replace(/\s+/g, "");
    const missing = [];
    for (const comment of comments) {
      // A long author name is clipped on the card with an ellipsis: its start is enough.
      const author = comment.author.replace(/\s+/g, "").slice(0, 10);
      const found = squeezed.includes(comment.text.replace(/\s+/g, "")) && squeezed.includes(author);
      if (!found) missing.push(`${comment.id}${comment.reply ? " (reply)" : ""} by ${comment.author}: "${comment.text.slice(0, 40)}"`);
    }
    return { path, pages: result.pages, comments: comments.length, replies: comments.filter((c) => c.reply).length, printed: comments.length - missing.length, missing };
  } catch (error) {
    return { path, comments: comments.length, replies: comments.filter((c) => c.reply).length, printed: 0, missing: [], error: `${error.code ?? "failed"}: ${error.message}` };
  } finally {
    out.destroy();
    rmSync(dir, { recursive: true, force: true });
  }
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
    for (const name of readdirSync(full).sort()) if (name.endsWith(".docx") && statSync(join(full, name)).isFile()) files.push(join(full, name));
  }
  if (files.length === 0) {
    console.error("nothing to check: pass .docx files or directories of them");
    process.exit(2);
  }
  const jobs = Math.max(1, Number(args.get("jobs") ?? 6));
  const results = [];
  let next = 0;
  let done = 0;
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
        done++;
        if (args.has("verbose") || result.missing.length || result.error) report(result);
        else if (done % 100 === 0) console.error(`${done}/${files.length}`);
        feed();
      });
      child.on("error", reject);
      child.on("exit", (code) => (code ? reject(new Error(`worker exited ${code}`)) : resolveWorker()));
      feed();
    });
  await Promise.all(Array.from({ length: Math.min(jobs, files.length) }, run));
  results.sort((a, b) => a.path.localeCompare(b.path));
  const sum = (key) => results.reduce((n, r) => n + r[key], 0);
  const short = results.filter((r) => r.missing.length);
  const errors = results.filter((r) => r.error);
  console.log(`${results.length} documents, ${results.filter((r) => r.comments).length} with comments: ${sum("comments")} comments (${sum("replies")} replies) expected, ${sum("printed")} printed, ${sum("comments") - sum("printed")} missing in ${short.length} documents; ${errors.length} render errors`);
  if (args.has("out")) writeFileSync(args.get("out"), JSON.stringify(results, null, 1));
  process.exitCode = short.length || errors.length ? 1 : 0;
}

function report(result) {
  const name = basename(result.path);
  if (result.error) console.error(`${name}: ${result.error}`);
  else if (result.missing.length) console.error(`${name}: ${result.printed} of ${result.comments} printed; missing ${result.missing.join("; ")}`);
  else console.error(`${name}: ${result.printed} of ${result.comments} printed`);
}

if (args.has("worker")) worker();
else main();
