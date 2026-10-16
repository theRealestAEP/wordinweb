/**
 * The printable copy's invariants (packages/compare-cli/test/invariants.ts)
 * over the compare benchmark's pairs and over fuzzed pairs
 * (packages/compare-cli/test/fuzz.ts), in parallel.
 *
 *   node_modules/.bin/vite-node scripts/printable-invariants.ts --bench [--variants DIR] [--fixtures DIR] [--only a,b]
 *   node_modules/.bin/vite-node scripts/printable-invariants.ts --fuzz 2000 [--seed 1] [--shrink] [--save DIR]
 *
 * Common flags: --jobs N (default 6), --no-pdf (invariants 1 to 3 only),
 * --out FILE (the results as JSON). A bench pair is a variant
 * <fixture>--<level>--s<seed>.docx against its fixture. A fuzz case is the
 * pair generateCase(seed) makes, for seeds seed .. seed + N - 1; with
 * --shrink a failing case is reduced to a small one, and with --save its
 * documents and spec are written to DIR. The exit code is 1 when a pair
 * fails an invariant or errors.
 */
import { spawn } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import { createInterface } from "node:readline";
import { fileURLToPath } from "node:url";
import { strFromU8, unzipSync } from "fflate";
import { Failure, PairReport, comparePair, formatFailure } from "../packages/compare-cli/test/invariants.js";
import { Case, generateCase, serialize, shrink } from "../packages/compare-cli/test/fuzz.js";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const args = new Map<string, string>();
const argv = process.argv.slice(2);
for (let i = 0; i < argv.length; i++) {
  if (!argv[i].startsWith("--")) continue;
  const next = argv[i + 1];
  args.set(argv[i].slice(2), next && !next.startsWith("--") ? argv[++i] : "true");
}

type Item = { kind: "bench"; name: string; original: string; revised: string } | { kind: "fuzz"; seed: number };
interface Result {
  item: Item;
  report: PairReport;
  /** Revision markup the inputs themselves carried into the compare (an input's unaccepted change). */
  inputMarkup: string[];
  shrunk?: { seed: number; edits: string[]; failures: Failure[]; saved?: string };
}

const INPUT_MARKUP = /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|trPrChange|tcPrChange|numberingChange|cellIns|cellDel|cellMerge)[\s>/]/g;

/** A failure's shape, free of its particulars: the same bug fails the same way. */
export function signature(f: Failure): string {
  return `${f.invariant}:${f.detail.replace(/"(?:[^"\\]|\\.)*"/g, '""').replace(/\d+/g, "N")}`;
}

// ---------------------------------------------------------------------------
// Worker: one item at a time from stdin, one result per line on stdout
// ---------------------------------------------------------------------------

async function worker(): Promise<void> {
  const pdf = !args.has("no-pdf");
  const saveDir = args.get("save");
  const work = mkdtempSync(join(process.env.TMPDIR ?? tmpdir(), "printable-invariants-"));
  const lines = createInterface({ input: process.stdin });
  for await (const line of lines) {
    const item = JSON.parse(line) as Item;
    const dir = join(work, "case");
    rmSync(dir, { recursive: true, force: true });
    mkdirSync(dir);
    let result: Result;
    try {
      if (item.kind === "bench") {
        const report = await comparePair(item.original, item.revised, dir, { pdf });
        result = { item, report, inputMarkup: inputMarkup(item.original, item.revised) };
      } else {
        const c = generateCase(item.seed);
        const report = await runCase(c, dir, pdf);
        result = { item, report, inputMarkup: [] };
        if (report.failures.length && args.has("shrink")) {
          const want = signature(report.failures[0]);
          const needsPdf = pdf && report.failures[0].invariant === 4;
          const small = await shrink(c, async (candidate) => (await runCase(candidate, dir, needsPdf)).failures.some((f) => signature(f) === want));
          const final = await runCase(small, dir, needsPdf);
          result.shrunk = { seed: c.seed, edits: small.edits.map((e) => e.kind), failures: final.failures };
          if (saveDir) {
            mkdirSync(saveDir, { recursive: true });
            const prefix = join(saveDir, `seed-${c.seed}`);
            writeFileSync(`${prefix}.json`, JSON.stringify({ seed: c.seed, edits: small.edits, original: small.original, revised: small.revised }, null, 1));
            writeFileSync(`${prefix}-original.docx`, serialize(small.original));
            writeFileSync(`${prefix}-revised.docx`, serialize(small.revised));
            result.shrunk.saved = prefix;
          }
        }
      }
    } catch (error) {
      result = { item, report: { failures: [], error: `worker: ${(error as Error).stack ?? String(error)}`, ms: 0 }, inputMarkup: [] };
    }
    process.stdout.write(JSON.stringify(result) + "\n");
  }
  rmSync(work, { recursive: true, force: true });
}

async function runCase(c: Case, dir: string, pdf: boolean): Promise<PairReport> {
  const original = join(dir, "original.docx");
  const revised = join(dir, "revised.docx");
  writeFileSync(original, serialize(c.original));
  writeFileSync(revised, serialize(c.revised));
  return comparePair(original, revised, dir, { pdf });
}

function inputMarkup(...paths: string[]): string[] {
  const found = new Set<string>();
  for (const path of paths) {
    const files = unzipSync(new Uint8Array(readFileSync(path)));
    for (const name of Object.keys(files)) {
      if (!/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(name)) continue;
      for (const m of strFromU8(files[name]).matchAll(INPUT_MARKUP)) found.add(m[1]);
    }
  }
  return [...found].sort();
}

// ---------------------------------------------------------------------------
// Main: the items, the workers, the report
// ---------------------------------------------------------------------------

function items(): Item[] {
  const out: Item[] = [];
  if (args.has("bench")) {
    const variants = resolve(args.get("variants") ?? join(root, "bench/compare/variants"));
    const fixtures = resolve(args.get("fixtures") ?? join(root, "../wordinweb-parity/apps/demo/public/fixtures"));
    const only = args.has("only") ? new Set((args.get("only") as string).split(",")) : null;
    for (const name of readdirSync(variants).sort()) {
      const m = /^(.+?)--([a-z-]+)--s(\d+)\.docx$/.exec(name);
      if (!m || (only && !only.has(m[1]))) continue;
      const original = join(fixtures, `${m[1]}.docx`);
      if (!existsSync(original)) continue;
      out.push({ kind: "bench", name: basename(name, ".docx"), original, revised: join(variants, name) });
    }
  }
  if (args.has("fuzz")) {
    const n = Number(args.get("fuzz"));
    const seed = Number(args.get("seed") ?? 1);
    for (let i = 0; i < n; i++) out.push({ kind: "fuzz", seed: seed + i });
  }
  return out;
}

async function main(): Promise<void> {
  const queue = items();
  if (queue.length === 0) {
    console.error("nothing to do: pass --bench and/or --fuzz N");
    process.exit(2);
  }
  const jobs = Math.max(1, Math.min(Number(args.get("jobs") ?? 6), queue.length));
  const results: Result[] = [];
  const total = queue.length;
  const started = Date.now();
  const passthrough = argv.filter((a) => a !== "--bench" && !/^--(fuzz|seed|jobs|out|only|variants|fixtures)$/.test(a) && !/^\d+$/.test(a));
  await Promise.all(
    Array.from({ length: jobs }, () => {
      return new Promise<void>((done, fail) => {
        const child = spawn(process.execPath, [join(root, "node_modules/.bin/vite-node"), fileURLToPath(import.meta.url), "--worker", ...passthrough], { stdio: ["pipe", "pipe", "inherit"] });
        const feed = (): void => {
          const next = queue.shift();
          if (!next) {
            child.stdin.end();
            return;
          }
          child.stdin.write(JSON.stringify(next) + "\n");
        };
        createInterface({ input: child.stdout }).on("line", (line) => {
          const result = JSON.parse(line) as Result;
          results.push(result);
          console.error(progress(result, results.length, total));
          feed();
        });
        child.on("exit", (code) => (code === 0 ? done() : fail(new Error(`worker exited with ${code}`))));
        feed();
      });
    }),
  );
  report(results, Date.now() - started);
  const out = args.get("out");
  if (out) {
    mkdirSync(dirname(resolve(out)), { recursive: true });
    writeFileSync(out, JSON.stringify(results, null, 1));
    console.log(`results: ${out}`);
  }
  process.exit(results.some((r) => r.report.failures.length || r.report.error) ? 1 : 0);
}

function label(item: Item): string {
  return item.kind === "bench" ? item.name : `fuzz seed ${item.seed}`;
}

function progress(r: Result, n: number, total: number): string {
  const { report } = r;
  const status = report.error ? `ERROR ${report.error.split("\n")[0].slice(0, 120)}` : report.failures.length ? `FAIL ${report.failures.length} (${[...new Set(report.failures.map((f) => f.invariant))].join(",")})` : "ok";
  const carried = r.inputMarkup.length ? ` [input markup: ${r.inputMarkup.join(",")}]` : "";
  return `[${n}/${total}] ${label(r.item)}: ${status}${report.pdfSkipped ? ` (pdf skipped: ${report.pdfSkipped})` : ""} ${report.ms} ms${carried}`;
}

function report(results: Result[], ms: number): void {
  const failing = results.filter((r) => r.report.failures.length);
  const errors = results.filter((r) => r.report.error);
  console.log(`\n${results.length} pairs in ${(ms / 1000).toFixed(0)} s: ${results.length - failing.length - errors.length} clean, ${failing.length} failing, ${errors.length} errors, ${results.filter((r) => r.report.pdfSkipped).length} without a PDF`);
  const byInvariant = new Map<number, number>();
  for (const r of failing) for (const inv of new Set(r.report.failures.map((f) => f.invariant))) byInvariant.set(inv, (byInvariant.get(inv) ?? 0) + 1);
  if (byInvariant.size) console.log(`pairs failing by invariant: ${[...byInvariant].sort().map(([k, v]) => `${k}=${v}`).join(" ")}`);
  // Each distinct failure once, with the pairs it hit.
  const groups = new Map<string, { pairs: string[]; example: Failure; carried: number }>();
  for (const r of failing) {
    for (const sig of new Set(r.report.failures.map(signature))) {
      const group = groups.get(sig) ?? { pairs: [], example: r.report.failures.find((f) => signature(f) === sig) as Failure, carried: 0 };
      group.pairs.push(label(r.item));
      if (r.inputMarkup.length) group.carried++;
      groups.set(sig, group);
    }
  }
  for (const [sig, group] of [...groups].sort((a, b) => b[1].pairs.length - a[1].pairs.length)) {
    console.log(`\n${group.pairs.length} pairs${group.carried ? ` (${group.carried} whose inputs carried revision markup)` : ""}: ${sig}`);
    console.log(`  e.g. ${group.pairs.slice(0, 6).join(", ")}${group.pairs.length > 6 ? ", …" : ""}`);
    console.log(`  ${formatFailure(group.example).replace(/\n/g, "\n  ")}`);
  }
  for (const r of errors) console.log(`\nERROR ${label(r.item)}: ${r.report.error?.split("\n").slice(0, 3).join("\n  ")}`);
  for (const r of results) {
    if (!r.shrunk) continue;
    console.log(`\nshrunk fuzz seed ${r.shrunk.seed} (edits ${r.shrunk.edits.join(",")})${r.shrunk.saved ? `: ${r.shrunk.saved}*` : ""}`);
    for (const f of r.shrunk.failures.slice(0, 4)) console.log(`  ${formatFailure(f).replace(/\n/g, "\n  ")}`);
  }
}

if (args.has("worker")) await worker();
else await main();
