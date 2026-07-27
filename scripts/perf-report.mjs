#!/usr/bin/env node
/**
 * PERF REPORT for the collab editor.
 *
 * The stress/swarm E2E suites are the measurement instrument: every number they
 * take is printed as one grep-able line
 *
 *     STRESS-METRIC <scenario> <key>=<value> <key>=<value> ...
 *
 * This script is the *record*. It runs (or ingests) those suites, parses every
 * metric line generically, appends one record per run to a JSONL history, and
 * regenerates `internal/perf/PERF-REPORT.md` from that history — a dated,
 * regenerable document that shows every measured area, compares the latest run
 * against a rolling baseline, and flags where performance DROPPED.
 *
 *   node scripts/perf-report.mjs run [--only=<pattern>] [--no-wait]
 *   node scripts/perf-report.mjs ingest <logfile> [<logfile> ...]
 *   node scripts/perf-report.mjs render
 *
 * Design rules that matter:
 *   - The parser NEVER hardcodes a scenario list. Any `STRESS-METRIC <name>`
 *     line becomes a record; unknown scenarios and unknown keys survive into
 *     the report untouched. New suites need no change here.
 *   - Rendering is a pure function of the history file: same history in, same
 *     bytes out. No "generated at" stamp, no locale-dependent formatting, no
 *     set-iteration ordering. That is what makes a diff of this report mean
 *     something.
 *   - Numbers are only comparable on the same machine, so trends are computed
 *     WITHIN an environment signature (platform + cpu count) and the report
 *     says so when the signature changed.
 */

import { execFileSync, spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const PERF_DIR = path.join(ROOT, "internal", "perf");
const HISTORY = path.join(PERF_DIR, "history.jsonl");
const REPORT = path.join(PERF_DIR, "PERF-REPORT.md");

/* ------------------------------------------------------------------ parse -- */

/**
 * Strip ANSI colour and any reporter prefix ahead of the marker.
 *
 * The ESC is OPTIONAL on purpose. A real terminal capture has it, but a log that
 * has been through a CI artifact viewer or a naive control-character filter
 * arrives with a bare `[32m` still wrapped around the line. Both must parse, or a
 * colourised scenario name silently becomes a DIFFERENT scenario (`hud` vs
 * `hud[39m`) and its trend history splits in two.
 */
function metricPayload(line) {
  const clean = line.replace(/\u001b?\[[0-9;]*m/g, "");
  const at = clean.indexOf("STRESS-METRIC ");
  return at === -1 ? null : clean.slice(at + "STRESS-METRIC ".length).trim();
}

function coerce(raw) {
  return /^-?\d+(\.\d+)?$/.test(raw) ? Number(raw) : raw;
}

/**
 * Parse a whole log into `{ scenarios, flags }`.
 *
 * A scenario emitted twice in one run (e.g. a retried round) is NOT collapsed —
 * the second becomes `name#2` — because silently keeping only the last value
 * would drop a measurement. Bare tokens with no `=` (the suites emit
 * `SOFT-BOUND-EXCEEDED` this way) are kept as per-scenario flags.
 */
export function parseLog(text) {
  const scenarios = Object.create(null);
  const flags = Object.create(null);
  const seen = Object.create(null);

  for (const line of text.split(/\r?\n/)) {
    const payload = metricPayload(line);
    if (payload === null) continue;
    const tokens = payload.split(/\s+/).filter(Boolean);
    if (tokens.length === 0) continue;

    const base = tokens[0];
    seen[base] = (seen[base] ?? 0) + 1;
    const name = seen[base] === 1 ? base : `${base}#${seen[base]}`;

    const fields = Object.create(null);
    const bare = [];
    for (const tok of tokens.slice(1)) {
      const eq = tok.indexOf("=");
      if (eq <= 0) bare.push(tok);
      else fields[tok.slice(0, eq)] = coerce(tok.slice(eq + 1));
    }
    scenarios[name] = fields;
    if (bare.length) flags[name] = bare;
  }
  return { scenarios, flags };
}

/* ---------------------------------------------------------------- history -- */

function gitDescribe() {
  try {
    const head = execFileSync("git", ["rev-parse", "--short", "HEAD"], { cwd: ROOT, encoding: "utf8" }).trim();
    const status = execFileSync("git", ["status", "--porcelain"], { cwd: ROOT, encoding: "utf8" }).trim();
    return status ? `${head}+dirty` : head;
  } catch {
    return "unknown";
  }
}

function readHistory() {
  if (!fs.existsSync(HISTORY)) return [];
  return fs
    .readFileSync(HISTORY, "utf8")
    .split("\n")
    .filter((l) => l.trim())
    .map((l, i) => {
      try {
        return JSON.parse(l);
      } catch {
        console.warn(`[perf] skipping unparseable history line ${i + 1}`);
        return null;
      }
    })
    .filter(Boolean);
}

function appendRecord({ scenarios, flags }, source) {
  const count = Object.keys(scenarios).length;
  if (count === 0) {
    console.error("[perf] no STRESS-METRIC lines found — nothing appended to history.");
    return null;
  }
  const rec = {
    ts: new Date().toISOString(),
    gitDescribe: gitDescribe(),
    node: process.version,
    platform: `${process.platform}-${process.arch}`,
    cpus: os.cpus().length,
    source,
    scenarios,
    ...(Object.keys(flags).length ? { flags } : {}),
  };
  fs.mkdirSync(PERF_DIR, { recursive: true });
  fs.appendFileSync(HISTORY, `${JSON.stringify(rec)}\n`);
  console.log(`[perf] recorded ${count} scenario(s) from ${source} -> ${path.relative(ROOT, HISTORY)}`);
  return rec;
}

/* ------------------------------------------------------- metric semantics -- */

/**
 * Direction per key. Explicit tables first (they beat the suffix heuristics),
 * then the defaults the brief asks for: `*Ms`/`*ms` lower-is-better,
 * `fps`/`opsPerSec` higher-is-better, everything else neutral.
 *
 * NEUTRAL is not "unimportant" — it is "a bigger number here is not a worse
 * editor". Scenario inputs (client counts, paragraph counts, ops submitted)
 * and pure state readings (seq, docBytes) live there so they never raise a
 * false regression when a suite is retuned.
 */
const LOWER_BETTER = new Set([
  "rttAvgMs", "rttMaxMs", "keyboardEchoMs", "keyboardMs", "remoteVisibleMs", "msPerChar",
  "submitMs", "ackedMs", "totalMs", "avgRoundMs", "maxRoundMs",
  "convergeMs", "convergedMs", "settleAvgMs", "settleMaxMs", "unsettledCheckpoints",
  "quietConvergeMs", "timeToQuietConvergenceMs", "finalMs",
  "editorReadyMs", "coldJoinMs", "lateJoinMs", "joinMs", "seedMs", "allClientsMs",
  "msPerParaQ1", "msPerParaQ2", "msPerParaQ3", "msPerParaQ4", "growthQ4overQ1",
  "typistDrift", "localDrift", "driftChars", "drifted",
  "hotDriftedClients", "hotMaxDriftChars", "driftedDuringChurn", "stillDrifted",
  "selfHeals", "selfHealsObserved",
  "worstFrameMs", "repaintAvgMs", "longTasks",
]);
const HIGHER_BETTER = new Set(["fps", "opsPerSec"]);
const NEUTRAL = new Set([
  "clients", "typists", "bursters", "paragraphs", "ops", "rounds", "typedChars", "intents",
  "totalOps", "typedTokens", "burstOps", "churnMs", "docChars", "docBytes", "hudParagraphs",
  "seq", "opsIn", "opsOut", "observerSeq", "observerOpsIn", "roundMarkers", "roundsKept",
  "hashGroups", "hotHashGroups", "targetParas", "healPath", "selfHealPath", "converged",
  "typedSurvived", "burstSurvived",
  // BIMODAL-BY-DESIGN swarm fields (B6a classifier finding, A8): the hot
  // census reads AT the instant churn stops, so raw drift and settle time
  // depend on whether anyone happened to hold an un-echoed optimistic op at
  // that instant — healthy runs legitimately read 0 OR 1-2 clients /
  // 6ms OR ~2s. Trending them flags the optimistic model working as a
  // regression every other run. The number that MEANS drift-is-wrong is
  // unexplainedDrift (below, lower-better); these stay recorded but
  // unflagged.
  "driftedClients", "maxDriftChars", "inFlightClients", "selfSettleMs",
]);
/** Per-scenario neutral overrides for keys whose meaning varies by scenario:
 * swarm-clash-final's convergeMs is echo-drain time under flood (bimodal on
 * in-flight state, see NEUTRAL note) while every other scenario's convergeMs
 * is a genuine latency. */
const SCENARIO_NEUTRAL = new Map([["swarm-clash-final", new Set(["convergeMs"])]]);
LOWER_BETTER.add("unexplainedDrift");

export function direction(key, scenario) {
  if (scenario && SCENARIO_NEUTRAL.get(scenario)?.has(key)) return "neutral";
  if (NEUTRAL.has(key)) return "neutral";
  if (LOWER_BETTER.has(key)) return "lower";
  if (HIGHER_BETTER.has(key)) return "higher";
  if (/^fps$|PerSec$/i.test(key)) return "higher";
  if (/(Ms|ms)$/.test(key)) return "lower";
  if (/drift|stuck|unsettled|longTask|slowdown|growth/i.test(key)) return "lower";
  return "neutral";
}

/**
 * Areas group the same metric key wherever it is emitted, so "latency" is one
 * table across every scenario instead of a number buried in six. Keys that
 * belong to no area are not lost: every scenario also gets its own complete
 * table further down the report.
 */
const AREAS = [
  ["Latency", ["rttAvgMs", "rttMaxMs", "keyboardEchoMs", "keyboardMs", "remoteVisibleMs", "msPerChar"]],
  ["Throughput", ["opsPerSec", "submitMs", "ackedMs", "totalMs", "avgRoundMs", "maxRoundMs"]],
  ["Convergence", ["convergeMs", "settleAvgMs", "settleMaxMs", "unsettledCheckpoints", "quietConvergeMs", "timeToQuietConvergenceMs", "finalMs"]],
  ["Join / scale", ["editorReadyMs", "convergedMs", "coldJoinMs", "lateJoinMs", "joinMs", "seedMs", "allClientsMs", "msPerParaQ1", "msPerParaQ2", "msPerParaQ3", "msPerParaQ4", "growthQ4overQ1"]],
  ["Client health", ["typistDrift", "localDrift", "driftChars", "drifted", "driftedClients", "maxDriftChars", "hotDriftedClients", "hotMaxDriftChars", "driftedDuringChurn", "stillDrifted", "selfHeals", "selfHealsObserved"]],
  ["Rendering", ["fps", "worstFrameMs", "repaintAvgMs", "longTasks"]],
];

/** `-1` is the suites' "this reading was not available" sentinel (an absent HUD,
 * a feature-detected counter that isn't exposed). It is displayed as-is but must
 * never enter a baseline or raise a regression. */
const isValue = (v) => typeof v === "number" && Number.isFinite(v) && v !== -1;

/* ------------------------------------------------------------ formatting -- */

/**
 * Value -> table cell. Escaping lives HERE rather than at the call sites
 * because several suite metrics legitimately use `|` as an internal separator
 * (`selfHeals=2|0|1`, `targetParas=b7|b91|b184`); an unescaped one silently
 * shreds the markdown row it lands in.
 */
function fmt(v) {
  if (v === undefined || v === null) return "—";
  if (typeof v !== "number") return escapeCell(v);
  if (!Number.isFinite(v)) return String(v);
  return Number.isInteger(v) ? String(v) : v.toFixed(1);
}

function pct(latest, prev) {
  if (!isValue(latest) || !isValue(prev)) return "—";
  if (latest === prev) return "0.0%"; // covers the 0 -> 0 case a ratio cannot
  if (prev === 0) return "new";
  const d = ((latest - prev) / Math.abs(prev)) * 100;
  const s = `${d >= 0 ? "+" : ""}${d.toFixed(1)}%`;
  return Math.abs(d) < 0.05 ? "0.0%" : s;
}

/** UTC everywhere: a report rendered on two machines must be byte-identical. */
function shortTs(iso) {
  const d = new Date(iso);
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}`;
}
function longTs(iso) {
  const d = new Date(iso);
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())} UTC`;
}

function median(nums) {
  if (nums.length === 0) return undefined;
  const s = [...nums].sort((a, b) => a - b);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

const envSig = (r) => `${r.platform} · ${r.cpus} cpu`;
const table = (rows) => rows.map((r) => `| ${r.join(" | ")} |`).join("\n");
const escapeCell = (s) => String(s).replace(/\|/g, "\\|");

/* ---------------------------------------------------------------- render -- */

const THRESHOLD = 0.2; // >20% off the baseline, per the report's contract
const MIN_ABS = 1; // ignore sub-unit wobble on tiny numbers (0.4ms -> 0.9ms)

/** Compare one metric against its rolling baseline. */
function assess(key, latest, baselineRuns, scenario) {
  const dir = direction(key, scenario);
  const samples = baselineRuns.filter(isValue);
  const baseline = median(samples);
  const prev = baselineRuns.length ? baselineRuns[baselineRuns.length - 1] : undefined;

  let flag = "";
  const comparable = dir !== "neutral" && isValue(latest) && baseline !== undefined;
  if (comparable && baseline === 0 && dir === "lower" && latest >= MIN_ABS) {
    // A zero baseline has no ratio, but 0 -> 3 on a drift or self-heal counter
    // is the single most important regression this report can catch. Percentage
    // thinking would drop it on the floor; flag it outright.
    flag = "▲ WORSE";
  } else if (comparable && baseline > 0 && Math.abs(latest - baseline) >= MIN_ABS) {
    const ratio = latest / baseline;
    if (dir === "lower" && ratio > 1 + THRESHOLD) flag = "▲ WORSE";
    else if (dir === "higher" && ratio < 1 - THRESHOLD) flag = "▲ WORSE";
    else if (dir === "lower" && ratio < 1 - THRESHOLD) flag = "▼ better";
    else if (dir === "higher" && ratio > 1 + THRESHOLD) flag = "▼ better";
  }
  return { dir, baseline, prev, flag, samples: samples.length };
}

/** All runs sharing the latest run's environment signature, oldest first. */
function comparableRuns(history) {
  const latest = history[history.length - 1];
  return history.filter((r) => envSig(r) === envSig(latest));
}

function scenarioNames(history) {
  const names = new Set();
  for (const r of history) for (const n of Object.keys(r.scenarios ?? {})) names.add(n);
  return [...names].sort();
}

/** Values of `scenario.key` in the given runs (missing runs contribute nothing). */
function seriesOf(runs, scenario, key) {
  const out = [];
  for (const r of runs) {
    const v = r.scenarios?.[scenario]?.[key];
    if (typeof v === "number") out.push(v);
  }
  return out;
}

export function render(history) {
  if (history.length === 0) {
    return "# Collab perf report\n\nNo runs recorded yet. Run `npm run perf:report`.\n";
  }

  const latest = history[history.length - 1];
  const peers = comparableRuns(history); // same platform+cpus, oldest first
  const priorPeers = peers.slice(0, -1); // everything before the latest run
  const baselineWindow = priorPeers.slice(-5); // rolling baseline: up to 5 prior runs
  const sigChanged = peers.length !== history.length;

  const out = [];
  const L = (s = "") => out.push(s);

  /* ---- header ---- */
  L("# Collab perf report");
  L();
  L("Every performance number the collab editor's browser suites measure, recorded over time,");
  L("with the latest run compared against a rolling baseline so a **drop shows up as a flag**");
  L("instead of as a number nobody re-read.");
  L();
  L("Re-run (runs the suites, appends one history record, regenerates this file):");
  L();
  L("```");
  L("npm run perf:report");
  L("```");
  L();
  L("Other modes: `node scripts/perf-report.mjs ingest <logfile>` (record a CI log without");
  L("re-running) and `node scripts/perf-report.mjs render` (rebuild this file from history only).");
  L();
  L("**Where the numbers come from.** The suites (`e2e/stress.spec.ts`, `e2e/swarm.spec.ts`,");
  L("`scripts/swarm.mjs`) print one grep-able line per measurement:");
  L();
  L("```");
  L("STRESS-METRIC <scenario> <key>=<value> <key>=<value> ...");
  L("```");
  L();
  L("The parser is generic — any scenario name and any key is recorded, so a new suite needs no");
  L("change here. Nothing is dropped: keys outside the known areas still appear in the");
  L("per-scenario tables.");
  L();
  L("**Reading the flags.** `▲ WORSE` = the latest value is more than 20% worse than the");
  L("baseline (the median of up to the 5 preceding runs on this machine); `▼ better` = more than");
  L("20% better. Direction is per key: `*Ms` durations and drift/heal counts are lower-is-better,");
  L("`fps` and `opsPerSec` are higher-is-better, and scenario inputs (client counts, paragraph");
  L("counts, `seq`) are neutral and never flagged. `-1` is the suites' \"not available\" sentinel");
  L("and is excluded from baselines. Sub-unit movements are not flagged. A lower-is-better");
  L("counter that was 0 and is now non-zero (drift appearing where there was none) is flagged");
  L("outright — no percentage applies, and it is the regression that matters most.");
  L();
  L("**Machines are not comparable.** Trends are computed only across runs with the same");
  L("platform + cpu-count signature as the latest run.");
  L();

  /* ---- latest run banner ---- */
  L(`## Latest run — ${longTs(latest.ts)}`);
  L();
  L(`- **Commit:** \`${latest.gitDescribe}\` · **Node:** ${latest.node} · **Source:** ${latest.source}`);
  L(`- **Environment:** ${envSig(latest)}`);
  L(
    `- **Trend basis:** ${baselineWindow.length} prior run(s) on this environment` +
      (baselineWindow.length === 0 ? " — no baseline yet, so nothing can be flagged." : ` (of ${history.length} total recorded).`),
  );
  if (sigChanged) {
    const others = [...new Set(history.filter((r) => envSig(r) !== envSig(latest)).map(envSig))].sort();
    L(`- **Environment signature changed:** ${others.join(", ")} also appear in history and are EXCLUDED from every comparison below.`);
  }
  L(`- **Scenarios in this run:** ${Object.keys(latest.scenarios).sort().join(", ")}`);
  if (latest.flags) {
    for (const [sc, fl] of Object.entries(latest.flags).sort()) L(`- **Suite flag** on \`${sc}\`: ${fl.join(" ")}`);
  }
  L();

  /* ---- regressions summary ---- */
  const regressions = [];
  const improvements = [];
  const names = scenarioNames(history);
  for (const sc of names) {
    const fields = latest.scenarios?.[sc];
    if (!fields) continue;
    for (const key of Object.keys(fields)) {
      const a = assess(key, fields[key], seriesOf(baselineWindow, sc, key), sc);
      if (a.flag === "▲ WORSE") regressions.push([sc, key, fmt(fields[key]), fmt(a.baseline), pct(fields[key], a.baseline)]);
      else if (a.flag === "▼ better") improvements.push([sc, key, fmt(fields[key]), fmt(a.baseline), pct(fields[key], a.baseline)]);
    }
  }

  L(`## Regressions — ${longTs(latest.ts).slice(0, 10)}`);
  L();
  if (baselineWindow.length === 0) {
    L("No baseline on this environment yet (this is the first comparable run), so no regression can be");
    L("computed. The numbers below are the starting point.");
  } else if (regressions.length === 0) {
    L(`**None.** No metric is more than 20% worse than its baseline across ${names.length} scenario(s).`);
  } else {
    L(`**${regressions.length} metric(s) dropped more than 20% below baseline.**`);
    L();
    L(table([["Scenario", "Metric", "Latest", "Baseline", "Δ vs baseline"], ["---", "---", "---", "---", "---"], ...regressions]));
  }
  if (improvements.length) {
    L();
    L(`<details><summary>${improvements.length} metric(s) improved by more than 20%</summary>`);
    L();
    L(table([["Scenario", "Metric", "Latest", "Baseline", "Δ vs baseline"], ["---", "---", "---", "---", "---"], ...improvements]));
    L();
    L("</details>");
  }
  L();

  /* ---- area sections ---- */
  const claimed = new Set();
  for (const [area, keys] of AREAS) {
    for (const k of keys) claimed.add(k);
    const rows = [];
    for (const sc of names) {
      const fields = latest.scenarios?.[sc];
      if (!fields) continue;
      for (const key of keys) {
        if (!(key in fields)) continue;
        const series = seriesOf(baselineWindow, sc, key);
        const a = assess(key, fields[key], series, sc);
        rows.push([
          `\`${sc}\``,
          key,
          fmt(fields[key]),
          fmt(a.prev),
          pct(fields[key], a.prev),
          fmt(a.baseline),
          a.flag || (a.dir === "neutral" ? "·" : ""),
        ]);
      }
    }
    L(`## ${area}`);
    L();
    if (rows.length === 0) {
      L("_Not measured in the latest run._");
    } else {
      L(table([
        ["Scenario", "Metric", "Latest", "Previous", "Δ vs prev", "Baseline", "Flag"],
        ["---", "---", "---", "---", "---", "---", "---"],
        ...rows,
      ]));
    }
    L();
  }

  /* ---- per-scenario full tables (nothing dropped) ---- */
  L("## Per-scenario detail");
  L();
  L("Every key the suites emitted, in emit order, including any this report has no area for.");
  L();
  for (const sc of names) {
    const fields = latest.scenarios?.[sc];
    L(`### \`${sc}\``);
    L();
    if (!fields) {
      const last = [...history].reverse().find((r) => r.scenarios?.[sc]);
      L(`_Not emitted in the latest run (last seen ${last ? shortTs(last.ts) : "—"} UTC)._`);
      L();
      continue;
    }
    const rows = Object.keys(fields).map((key) => {
      const a = assess(key, fields[key], seriesOf(baselineWindow, sc, key), sc);
      const unknown = claimed.has(key) ? "" : " ¹";
      return [
        `${key}${unknown}`,
        fmt(fields[key]),
        fmt(a.prev),
        pct(fields[key], a.prev),
        fmt(a.baseline),
        a.dir === "neutral" ? "·" : a.dir === "lower" ? "↓ better" : "↑ better",
        a.flag || "",
      ];
    });
    L(table([
      ["Metric", "Latest", "Previous", "Δ vs prev", "Baseline", "Direction", "Flag"],
      ["---", "---", "---", "---", "---", "---", "---"],
      ...rows,
    ]));
    if (rows.some((r) => r[0].endsWith("¹"))) {
      L();
      L("¹ key has no area grouping in this report — recorded verbatim.");
    }
    if (latest.flags?.[sc]) {
      L();
      L(`Suite flag: \`${latest.flags[sc].join(" ")}\``);
    }
    L();
  }

  /* ---- history tails ---- */
  L("## History");
  L();
  L("Up to the last 8 comparable runs, headline metrics only (times UTC).");
  L();
  for (const sc of names) {
    const runs = peers.filter((r) => r.scenarios?.[sc]).slice(-8);
    if (runs.length === 0) continue;
    const lastFields = runs[runs.length - 1].scenarios[sc];
    // Headline = the first up-to-3 keys with a known direction; a scenario that
    // is all-neutral falls back to its first numeric keys so the tail is never empty.
    let headline = Object.keys(lastFields).filter((k) => direction(k) !== "neutral" && typeof lastFields[k] === "number").slice(0, 3);
    if (headline.length === 0) headline = Object.keys(lastFields).filter((k) => typeof lastFields[k] === "number").slice(0, 3);
    if (headline.length === 0) continue;
    L(`### \`${sc}\``);
    L();
    L(table([
      ["When", ...headline, "Commit"],
      ["---", ...headline.map(() => "---"), "---"],
      ...runs.map((r) => [shortTs(r.ts), ...headline.map((k) => fmt(r.scenarios[sc][k])), `\`${r.gitDescribe}\``]),
    ]));
    L();
  }

  /* ---- run ledger ---- */
  L("## Runs");
  L();
  L(table([
    ["When (UTC)", "Commit", "Environment", "Node", "Source", "Scenarios"],
    ["---", "---", "---", "---", "---", "---"],
    ...[...history].reverse().map((r) => [
      shortTs(r.ts),
      `\`${r.gitDescribe}\``,
      envSig(r) === envSig(latest) ? envSig(r) : `${envSig(r)} ⚠︎`,
      r.node,
      escapeCell(r.source),
      String(Object.keys(r.scenarios ?? {}).length),
    ]),
  ]));
  L();
  L("⚠︎ = different environment signature from the latest run; excluded from all comparisons.");
  L();

  return `${out.join("\n").replace(/\n{3,}/g, "\n\n")}`;
}

function writeReport(history) {
  fs.mkdirSync(PERF_DIR, { recursive: true });
  fs.writeFileSync(REPORT, render(history));
  console.log(`[perf] wrote ${path.relative(ROOT, REPORT)} (${history.length} run(s) in history)`);
}

/* ------------------------------------------------------------------- run -- */

function portBusy(port) {
  const r = spawnSync("lsof", ["-nP", `-iTCP:${port}`, "-sTCP:LISTEN"], { encoding: "utf8" });
  return r.status === 0 && r.stdout.trim().length > 0;
}

/**
 * The playwright `webServer` is configured `reuseExistingServer: false`, so a
 * stray stack on 5399/1399 does not merely collide — it fails the whole run.
 * Wait it out rather than trampling another process's ports.
 */
function waitForPorts(ports, budgetMs) {
  const deadline = Date.now() + budgetMs;
  let announced = false;
  for (;;) {
    const busy = ports.filter(portBusy);
    if (busy.length === 0) return true;
    if (Date.now() > deadline) {
      console.error(`[perf] port(s) ${busy.join(", ")} still busy after ${Math.round(budgetMs / 1000)}s — aborting.`);
      console.error("[perf] fall back to: node scripts/perf-report.mjs ingest <saved-log>");
      return false;
    }
    if (!announced) {
      console.log(`[perf] waiting for port(s) ${busy.join(", ")} to free up...`);
      announced = true;
    }
    // Coarse sleep without pulling in a dependency.
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 5000);
  }
}

function doRun(argv) {
  const only = (argv.find((a) => a.startsWith("--only=")) ?? "").slice("--only=".length);
  const suites = only
    ? [only]
    : ["stress", "swarm"].filter((s) => fs.existsSync(path.join(ROOT, "e2e", `${s}.spec.ts`)));
  if (suites.length === 0) {
    console.error("[perf] no stress/swarm specs found under e2e/ — nothing to run.");
    process.exit(1);
  }
  console.log(`[perf] suites: ${suites.join(", ")}`);

  if (!argv.includes("--no-wait") && !waitForPorts([5399, 1399], 5 * 60_000)) process.exit(2);

  // One playwright invocation for all suites: the webServer boots once, which
  // is both faster and far less likely to race another agent for the ports.
  const args = ["playwright", "test", ...suites];
  console.log(`[perf] npx ${args.join(" ")}`);
  const r = spawnSync("npx", args, { cwd: ROOT, encoding: "utf8", maxBuffer: 256 * 1024 * 1024 });
  const log = `${r.stdout ?? ""}\n${r.stderr ?? ""}`;

  fs.mkdirSync(PERF_DIR, { recursive: true });
  const logPath = path.join(PERF_DIR, "last-run.log");
  fs.writeFileSync(logPath, log);
  process.stdout.write(log);
  if (r.status !== 0) {
    console.warn(`[perf] playwright exited ${r.status} — recording whatever metrics it printed before failing.`);
  }
  return appendRecord(parseLog(log), `run:${suites.join("+")}${r.status === 0 ? "" : `(exit ${r.status})`}`);
}

/* ------------------------------------------------------------------ main -- */

function main() {
  const [mode = "run", ...rest] = process.argv.slice(2);

  if (mode === "render") {
    writeReport(readHistory());
    return;
  }
  if (mode === "ingest") {
    const files = rest.filter((a) => !a.startsWith("--"));
    if (files.length === 0) {
      console.error("usage: node scripts/perf-report.mjs ingest <logfile> [<logfile> ...]");
      process.exit(1);
    }
    for (const f of files) {
      const p = path.resolve(f);
      if (!fs.existsSync(p)) {
        console.error(`[perf] no such log: ${p}`);
        process.exit(1);
      }
      appendRecord(parseLog(fs.readFileSync(p, "utf8")), `ingest:${path.basename(p)}`);
    }
    writeReport(readHistory());
    return;
  }
  if (mode === "run") {
    doRun(rest);
    writeReport(readHistory());
    return;
  }
  console.error(`unknown mode: ${mode}\nusage: node scripts/perf-report.mjs run|ingest <log>|render`);
  process.exit(1);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) main();
