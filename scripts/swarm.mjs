#!/usr/bin/env node
/**
 * MANY-EDITOR SWARM SIMULATOR — "what happens when 30 people open the same
 * document and all start typing?"
 *
 *   node scripts/swarm.mjs --n 20 --seconds 45 --paras 300 --typists 6
 *   node scripts/swarm.mjs --n 30 --seconds 60 --headed
 *   node scripts/swarm.mjs --n 12 --seconds 20 --paras 100 --no-build
 *
 * It boots its OWN stack (server + Vite) on dedicated ports so it never
 * collides with the Playwright runner (5399/1399) or the dev demo
 * (5817/1234), launches N real Chromium contexts, seeds a document, then
 * runs a clash mix for `--seconds`: `--typists` clients on real keyboards
 * (frequently on the SAME line, which is where the conflicts are) and the
 * rest flooding intents into one run with jitter.
 *
 * Every 5s it prints a live status table — per client: seq, pending, document
 * length, the short hash of its canonical bytes, and the connection's
 * self-heal counter when the build exposes one. Clients holding DIFFERENT
 * hashes are the whole point: that is drift, live.
 *
 * The verdict is EVENTUAL convergence. Drift during the churn is expected
 * (stress.spec documents a typist's optimistic replica ending a few chars
 * short under a concurrent flood); drift that outlives the churn is a defect.
 * Exit code is non-zero if the clients never agree within 60s of quiet.
 *
 * Flags: --n, --seconds, --paras, --typists, --headed, --no-build,
 *        --vite-port, --port, --keep (leave the stack up on exit).
 */
import { spawn } from "node:child_process";
import { setTimeout as delay } from "node:timers/promises";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { chromium } from "@playwright/test";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");

/* ---------------------------------------------------------------- args -- */

function parseArgs(argv) {
  const o = {
    n: 12,
    seconds: 30,
    paras: 100,
    typists: 4,
    headed: false,
    build: true,
    keep: false,
    vitePort: process.env.SWARM_VITE_PORT ?? "5601",
    port: process.env.SWARM_PORT ?? "1601",
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    const num = () => Number(argv[++i]);
    if (a === "--n" || a === "-n") o.n = num();
    else if (a === "--seconds" || a === "-s") o.seconds = num();
    else if (a === "--paras" || a === "-p") o.paras = num();
    else if (a === "--typists" || a === "-t") o.typists = num();
    else if (a === "--vite-port") o.vitePort = String(argv[++i]);
    else if (a === "--port") o.port = String(argv[++i]);
    else if (a === "--headed") o.headed = true;
    else if (a === "--no-build") o.build = false;
    else if (a === "--keep") o.keep = true;
    else if (a === "--help" || a === "-h") {
      console.log(
        "usage: node scripts/swarm.mjs [--n 20] [--seconds 45] [--paras 300] [--typists 6]\n" +
          "                             [--headed] [--no-build] [--vite-port 5601] [--port 1601] [--keep]",
      );
      process.exit(0);
    } else {
      console.error(`unknown flag: ${a}`);
      process.exit(2);
    }
  }
  o.n = Math.max(2, Math.min(60, o.n | 0));
  o.typists = Math.max(0, Math.min(o.n - 1, o.typists | 0));
  return o;
}

const opts = parseArgs(process.argv.slice(2));
const BASE = `http://localhost:${opts.vitePort}`;
const LANDING = `${BASE}/?server=localhost:${opts.port}&perf=1`;
/** How long after the churn stops the clients still have to agree. */
const QUIET_BUDGET_MS = 60_000;

const log = (...a) => console.log(...a);
const metric = (scenario, fields) =>
  log(
    `STRESS-METRIC ${scenario} ` +
      Object.entries(fields)
        .map(([k, v]) => `${k}=${typeof v === "number" && !Number.isInteger(v) ? v.toFixed(1) : v}`)
        .join(" "),
  );

/* --------------------------------------------------------------- stack -- */

let stack = null;
let browser = null;
let shuttingDown = false;

async function shutdown(code) {
  if (shuttingDown) return;
  shuttingDown = true;
  try {
    await browser?.close();
  } catch {
    /* the browser may already be gone */
  }
  if (stack && !opts.keep) {
    stack.kill("SIGTERM");
    // dev.mjs forwards SIGTERM to the server + Vite; give them a tick to go.
    await delay(400);
    if (stack.exitCode === null) stack.kill("SIGKILL");
  }
  process.exit(code);
}
for (const sig of ["SIGINT", "SIGTERM"]) process.on(sig, () => void shutdown(130));

/** Boot server + Vite on OUR ports and wait until the demo answers. */
async function bootStack() {
  log(`[swarm] booting stack — demo :${opts.vitePort}, server :${opts.port}${opts.build ? "" : " (--no-build)"}`);
  stack = spawn("node", ["scripts/dev.mjs", "--no-open", ...(opts.build ? [] : ["--no-build"])], {
    cwd: ROOT,
    // WW_ROOM_CAP: the hub caps rooms at 10 participants (owner decision);
    // the swarm deliberately exceeds that, so raise it for this harness only.
    env: { ...process.env, VITE_PORT: opts.vitePort, PORT: opts.port, WW_ROOM_CAP: String(Math.max(10, Number(opts.n) + 2)) },
    stdio: ["ignore", "pipe", "pipe"],
  });
  const tail = [];
  for (const s of [stack.stdout, stack.stderr]) {
    s.setEncoding("utf8");
    s.on("data", (d) => {
      tail.push(d);
      if (tail.length > 40) tail.shift();
    });
  }
  stack.on("exit", (c) => {
    if (!shuttingDown) {
      console.error(`[swarm] the stack exited early (code ${c}):\n${tail.join("")}`);
      void shutdown(1);
    }
  });

  const deadline = Date.now() + 180_000; // a cold run builds the packages first
  for (;;) {
    try {
      const r = await fetch(BASE, { signal: AbortSignal.timeout(2000) });
      if (r.ok) break;
    } catch {
      /* not listening yet */
    }
    if (Date.now() > deadline) {
      console.error(`[swarm] the demo never came up on ${BASE}:\n${tail.join("")}`);
      await shutdown(1);
    }
    await delay(500);
  }
  // The server answers on its own port before the app is worth driving.
  for (;;) {
    try {
      await fetch(`http://localhost:${opts.port}/stats`, { signal: AbortSignal.timeout(2000) });
      break;
    } catch {
      if (Date.now() > deadline) break;
      await delay(300);
    }
  }
  log(`[swarm] stack up`);
}

/* ------------------------------------------------------------- clients -- */

const WW = "window.__ww";

/** Wait until the demo's dev hook is live and its core members are callable. */
async function waitHook(page, timeout = 60_000) {
  await page.waitForFunction(
    () => {
      const w = window.__ww;
      return !!w && typeof w.submitOp === "function" && typeof w.allocIds === "function" && typeof w.saveB64 === "function";
    },
    undefined,
    { timeout },
  );
}

async function createDoc(page) {
  await page.goto(LANDING);
  await page.getByTestId("make-collaborative").click();
  await page.waitForURL(/[?&]doc=/, { timeout: 30_000 });
  await page.locator(".dxw-page").first().waitFor({ state: "visible", timeout: 30_000 });
  await waitHook(page);
  return page.url();
}

async function joinDoc(page, url) {
  await page.goto(url);
  await page.locator(".dxw-page").first().waitFor({ state: "visible", timeout: 60_000 });
  await waitHook(page);
}

const text = (page) => page.evaluate(() => window.__ww.text());
const saveB64 = (page) => page.evaluate(() => window.__ww.saveB64());
const perf = (page) => page.evaluate(() => window.__ww.perf?.() ?? null);
/** The connection's self-heal counter, from either exposure the build offers
 * (`__ww.selfHeals()` or the perf snapshot); null when it has neither. */
const healCount = (page) =>
  page.evaluate(() => {
    const w = window.__ww;
    if (typeof w?.selfHeals === "function") return w.selfHeals();
    const s = w?.perf?.();
    return typeof s?.selfHeals === "number" ? s.selfHeals : null;
  });

/** 32-bit digest of a client's canonical bytes — a readable stand-in for the
 * whole base64 blob when N clients have to be compared at a glance. */
function shortHash(s) {
  if (s === null || s === undefined) return "null----";
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

/** Seed `n` paragraphs through the real connection; returns one rung per
 * paragraph ({blockId, runId, len}) so a later edit can address it. */
async function seedParagraphs(page, n) {
  return page.evaluate(async (count) => {
    const ww = window.__ww;
    const rungs = [];
    let cur = { blockId: 1, runId: 2 };
    for (let i = 0; i < count; i++) {
      const t = `Paragraph ${i} - the quick brown fox jumps over the lazy dog.`;
      ww.submitOp({ kind: "insertText", at: { ...cur, offset: 0 }, text: t });
      const [b, r] = ww.allocIds(2);
      ww.submitOp({ kind: "splitParagraph", at: { ...cur, offset: t.length }, newBlockId: b, newRunId: r });
      rungs.push({ ...cur, len: t.length });
      cur = { blockId: b, runId: r };
      await new Promise((res) => setTimeout(res, 2));
    }
    rungs.push({ ...cur, len: 0 });
    return rungs;
  }, n);
}

/** Runs IN THE BROWSER: a click point on a random mounted line of text (or
 * the first line when `first`). The layout engine emits a positioned span per
 * measured chunk, so lines are found geometrically, never by their text. */
function lineClickPoint(first) {
  const leaves = [];
  for (const p of document.querySelectorAll(".dxw-page")) {
    for (const el of p.querySelectorAll("*")) {
      if (!el.children.length && (el.textContent ?? "").trim()) leaves.push(el);
    }
  }
  if (!leaves.length) {
    const box = document.querySelector(".dxw-page")?.getBoundingClientRect();
    return box ? { x: box.x + 30, y: box.y + 25 } : null;
  }
  const el = first ? leaves[0] : leaves[Math.floor(Math.random() * leaves.length)];
  const r = el.getBoundingClientRect();
  if (r.width === 0 || r.top < 0 || r.bottom > window.innerHeight) return null;
  return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
}

/**
 * A real keyboard until `until`. Two thirds of the re-clicks land on the SAME
 * first line as every other typist (maximum conflict); the rest wander.
 *
 * The caret is placed once up front and MOVED only occasionally, because a
 * click is the expensive part: on a client receiving a heavy flood, one click
 * can take seconds to be serviced (the swarm reports `clickMs`), and a typist
 * that re-clicks every token would spend the whole run waiting instead of
 * typing.
 */
async function typistLoop(page, tag, until, stats) {
  let tokens = 0;
  let clicks = 0;
  let clickMs = 0;
  while (Date.now() < until && !stats.stop) {
    try {
      if (clicks === 0 || Math.random() < 0.25) {
        const c0 = Date.now();
        const spot = await page.evaluate(lineClickPoint, Math.random() < 0.66);
        if (spot) await page.mouse.click(spot.x, spot.y);
        clickMs += Date.now() - c0;
        clicks++;
      }
      await page.keyboard.type(`${tag}${tokens % 10}`, { delay: 12 });
      tokens++;
      stats.typed++;
    } catch (e) {
      if (!stats.stop) stats.errors.push(String(e).slice(0, 120));
      break;
    }
    await delay(120 + Math.random() * 260);
  }
  stats.clicks += clicks;
  stats.clickMs += clickMs;
  return tokens;
}

/**
 * An intent flood into ONE run, driven from inside the page so the ops leave
 * faster than a CDP round-trip allows.
 *
 * Offsets stay in range WITHOUT reading the document: this client alone has
 * put `base` characters into the target run and nothing here deletes, so the
 * run is at least that long on every version of the document, however the
 * other clients have reshaped it.
 */
async function burstLoop(page, at, base, tag, ms) {
  return page.evaluate(
    async (a) => {
      const ww = window.__ww;
      const t0 = Date.now();
      let mine = a.base;
      let ops = 0;
      while (Date.now() - t0 < a.ms) {
        const token = `${a.tag}${ops % 10}`;
        ww.submitOp({
          kind: "insertText",
          at: { ...a.at, offset: Math.floor(Math.random() * (mine + 1)) },
          text: token,
        });
        mine += token.length;
        ops++;
        await new Promise((r) => setTimeout(r, 25 + Math.random() * 60));
      }
      return ops;
    },
    { at, base, tag, ms },
  );
}

/** One snapshot of every client: seq/pending/self-heals from the HUD, plus
 * document length and canonical hash. */
async function survey(pages) {
  return Promise.all(
    pages.map(async (p) => {
      try {
        const [s, t, b, h] = await Promise.all([perf(p), text(p), saveB64(p), healCount(p)]);
        return {
          seq: s?.seq ?? -1,
          pending: s?.pending ?? -1,
          opsOut: s?.opsOut ?? 0,
          heals: h,
          chars: t.length,
          hash: shortHash(b),
        };
      } catch (e) {
        return { seq: -1, pending: -1, opsOut: 0, heals: null, chars: -1, hash: "gone----", err: String(e).slice(0, 60) };
      }
    }),
  );
}

/** hash → client indices, majority first. */
function groupsOf(rows) {
  const g = new Map();
  rows.forEach((r, i) => g.set(r.hash, [...(g.get(r.hash) ?? []), i]));
  return [...g.entries()].sort((a, b) => b[1].length - a[1].length);
}

function printTable(rows, roles, elapsedS, opsPerSec) {
  const g = groupsOf(rows);
  const majority = g[0][0];
  log(
    `\n[t=${String(elapsedS).padStart(3)}s] ops/s=${opsPerSec.toFixed(1)}  groups=${g.length}  ` +
      g.map(([h, m]) => `${m.length}×${h}`).join(" "),
  );
  log("  client role   seq  pend   chars  hash      heal");
  rows.forEach((r, i) => {
    log(
      `  c${String(i).padStart(2, "0")}    ${roles[i]}    ${String(r.seq).padStart(5)} ${String(r.pending).padStart(5)} ` +
        `${String(r.chars).padStart(7)}  ${r.hash}${r.hash === majority ? " " : "*"} ${r.heals === null ? "  -" : String(r.heals).padStart(3)}`,
    );
  });
}

/* ---------------------------------------------------------------- main -- */

async function main() {
  await bootStack();

  log(
    `[swarm] launching ${opts.n} browser contexts (${opts.typists} typists, ${opts.n - opts.typists - 1} intent floods, 1 observer)`,
  );
  browser = await chromium.launch({ headless: !opts.headed });
  const contexts = [];
  const pages = [];
  for (let i = 0; i < opts.n; i++) {
    const ctx = await browser.newContext();
    contexts.push(ctx);
    pages.push(await ctx.newPage());
  }
  const roles = pages.map((_, i) => (i < opts.typists ? "T" : i === opts.n - 1 ? "O" : "B"));

  const t0 = Date.now();
  const url = await createDoc(pages[0]);
  log(`[swarm] doc created; joining ${opts.n - 1} clients…`);
  // Joins go in small waves: a simultaneous stampede of 30 cold joins measures
  // the join path, not the swarm.
  for (let i = 1; i < opts.n; i += 4) {
    await Promise.all(pages.slice(i, i + 4).map((p) => joinDoc(p, url)));
  }
  const joinMs = Date.now() - t0;
  log(`[swarm] ${opts.n} clients live in ${joinMs}ms`);

  let rungs = [{ blockId: 1, runId: 2, len: 0 }];
  let seedMs = 0;
  if (opts.paras > 0) {
    log(`[swarm] seeding ${opts.paras} paragraphs (broadcast to all ${opts.n} clients)…`);
    const st = Date.now();
    rungs = await seedParagraphs(pages[0], opts.paras);
    const submitMs = Date.now() - st;
    const seeded = (await text(pages[0])).length;
    // Everyone must actually HAVE the seed before the clash starts, or the
    // clash is measuring the tail of the seed.
    let spread = null;
    for (const deadline = Date.now() + 120_000; Date.now() < deadline; ) {
      const lens = await Promise.all(pages.map((p) => text(p).then((t) => t.length)));
      if (lens.every((l) => l === seeded)) {
        spread = Date.now() - st;
        break;
      }
      await delay(500);
    }
    seedMs = Date.now() - st;
    metric("swarm-seed", {
      clients: opts.n,
      paragraphs: opts.paras,
      ops: opts.paras * 2,
      submitMs,
      allClientsMs: spread ?? -1,
      opsPerSec: (opts.paras * 2 * 1000) / seedMs,
      docChars: seeded,
    });
    if (spread === null) log("[swarm] WARNING: not every client had the full seed after 120s");
  }

  /* ------------------------------------------------------------ clash -- */

  const churnMs = opts.seconds * 1000;
  const stats = { stop: false, typed: 0, clicks: 0, clickMs: 0, errors: [] };
  // The baseline survey is taken BEFORE the clock starts — with 30 clients it
  // is a second's worth of round-trips, and it must not eat the churn window.
  const before = await survey(pages);
  const healsBefore = before.map((r) => r.heals);
  const healPresent = healsBefore.some((h) => h !== null);
  log(`[swarm] clash for ${opts.seconds}s…`);
  const until = Date.now() + churnMs;

  // Targets: three far-apart paragraphs, so the floods are not all rebasing
  // against the same neighbourhood.
  const targets = [0.1, 0.5, 0.9].map((f) => rungs[Math.min(rungs.length - 1, Math.floor(rungs.length * f))]);
  const tags = "abcdefghijklmnopqrstuvwxyz";

  const workers = [];
  for (let i = 0; i < opts.n; i++) {
    if (roles[i] === "T") workers.push(typistLoop(pages[i], tags[i % 26].toUpperCase(), until, stats));
    else if (roles[i] === "B") {
      const t = targets[i % 3];
      workers.push(burstLoop(pages[i], { blockId: t.blockId, runId: t.runId }, t.len, tags[i % 26], churnMs));
    } else workers.push(Promise.resolve(0));
  }

  // Live status every 5s while the swarm runs.
  let lastOpsOut = before.reduce((s, r) => s + r.opsOut, 0);
  let lastTick = Date.now();
  const ticker = setInterval(() => {
    void (async () => {
      const rows = await survey(pages);
      const now = Date.now();
      const outs = rows.reduce((s, r) => s + r.opsOut, 0);
      printTable(rows, roles, Math.round((now - (until - churnMs)) / 1000), ((outs - lastOpsOut) * 1000) / (now - lastTick));
      lastOpsOut = outs;
      lastTick = now;
    })().catch(() => {});
  }, 5000);

  const counts = await Promise.all(workers);
  stats.stop = true;
  clearInterval(ticker);
  const churnActualMs = Date.now() - (until - churnMs);
  const totalOps = counts.reduce((x, y) => x + y, 0);

  /* ------------------------------------------------------------ quiet -- */

  log(`\n[swarm] churn over (${totalOps} ops submitted). Waiting for quiet convergence…`);
  const hot = await survey(pages);
  const hotGroups = groupsOf(hot);
  const hotMajority = hotGroups[0][0];
  const hotDrifted = hot.map((r, i) => (r.hash === hotMajority ? -1 : i)).filter((i) => i >= 0);
  const majorityChars = hot[hotGroups[0][1][0]].chars;
  const hotMaxDrift = Math.max(0, ...hot.map((r) => Math.abs(r.chars - majorityChars)));
  printTable(hot, roles, Math.round(churnActualMs / 1000), 0);

  const qt = Date.now();
  let convergedMs = null;
  for (;;) {
    const bs = await Promise.all(pages.map((p) => saveB64(p).catch(() => null)));
    if (bs[0] !== null && bs.every((b) => b === bs[0])) {
      convergedMs = Date.now() - qt;
      break;
    }
    if (Date.now() - qt > QUIET_BUDGET_MS) break;
    await delay(500);
  }

  const after = await survey(pages);
  const healDelta = after.reduce((s, r, i) => s + ((r.heals ?? 0) - (healsBefore[i] ?? 0)), 0);
  const finalGroups = groupsOf(after);
  const healed = hotDrifted.filter((i) => after[i].hash === finalGroups[0][0]);

  metric("swarm-clash", {
    clients: opts.n,
    typists: opts.typists,
    bursters: opts.n - opts.typists - 1,
    paragraphs: opts.paras,
    churnMs: churnActualMs,
    ops: totalOps,
    opsPerSec: (totalOps * 1000) / churnActualMs,
    keyboardTokens: stats.typed,
    // What a human would feel as "the editor froze": the average time one
    // click took to be serviced while the swarm was running.
    avgClickMs: stats.clicks ? stats.clickMs / stats.clicks : -1,
    hotHashGroups: hotGroups.length,
    hotDriftedClients: hotDrifted.length,
    hotMaxDriftChars: hotMaxDrift,
    quietConvergeMs: convergedMs ?? -1,
    selfHealPath: healPresent ? "exposed" : "absent",
    selfHealsObserved: healPresent ? healDelta : -1,
    docChars: after[0].chars,
  });
  metric("swarm-verdict", {
    converged: convergedMs !== null ? "yes" : "NO",
    timeToQuietConvergenceMs: convergedMs ?? -1,
    driftedDuringChurn: hotDrifted.length,
    driftedHealed: healed.length,
    stillDrifted: finalGroups.length - 1 === 0 ? 0 : opts.n - finalGroups[0][1].length,
    joinMs,
    seedMs,
  });

  if (convergedMs === null) {
    log(`\n[swarm] FAILED: clients never agreed within ${QUIET_BUDGET_MS / 1000}s of quiet.`);
    printTable(after, roles, Math.round((Date.now() - (until - churnMs)) / 1000), 0);
    const stuck = after.map((r, i) => (r.hash === groupsOf(after)[0][0] ? -1 : i)).filter((i) => i >= 0);
    const withPending = stuck.filter((i) => after[i].pending > 0);
    if (withPending.length) {
      log(
        `[swarm] ${withPending.length} of ${stuck.length} drifted clients still hold UNDRAINED pending intents ` +
          `(${withPending.map((i) => `c${i}:${after[i].pending}`).join(" ")}). A pending entry that never resolves also ` +
          `blocks the quiescent self-heal, which only runs when pending is empty.`,
      );
    }
    // Reload the stragglers: rebuilding from the stored bundle + the server's
    // tail is what heals an optimistic replica. If that converges them, the
    // canonical log was intact all along and the drift was purely local.
    log(`[swarm] probing: reloading ${stuck.length} drifted client(s) to see whether the canonical log is intact…`);
    for (const i of stuck) {
      await pages[i].reload();
      await pages[i].locator(".dxw-page").first().waitFor({ state: "visible", timeout: 60_000 });
      await waitHook(pages[i]);
    }
    let healedMs = null;
    for (const t = Date.now(); Date.now() - t < 30_000; ) {
      const bs = await Promise.all(pages.map((p) => saveB64(p).catch(() => null)));
      if (bs[0] !== null && bs.every((b) => b === bs[0])) {
        healedMs = Date.now() - t;
        break;
      }
      await delay(500);
    }
    const healedText = await text(pages[0]).catch(() => "");
    // Did the stuck pending edits reach the document, or vanish? The reload
    // rebuilds from the stored bundle, which still holds them — so a document
    // that GREW means the reload flushed them in, and one that did not means
    // they were dropped. Report the measurement, not an assumption.
    const majorityBefore = after[groupsOf(after)[0][1][0]].chars;
    const delta = healedText.length - majorityBefore;
    metric("swarm-reload-probe", {
      reloaded: stuck.length,
      convergedAfterReload: healedMs === null ? "no" : "yes",
      reloadConvergeMs: healedMs ?? -1,
      docCharsBefore: majorityBefore,
      docCharsAfter: healedText.length,
      charsRecoveredByReload: delta,
    });
    log(
      healedMs === null
        ? "[swarm] reload did NOT converge them — worse than local drift; the canonical log itself disagrees."
        : `[swarm] reload converged all clients in ${healedMs}ms — the canonical log was intact and the drift was ` +
          `local optimistic state. ` +
          (delta > 0
            ? `The document GREW by ${delta} chars: the reload flushed the stuck pending edits into it, so they were ` +
              `stranded, not lost — but only a reload delivered them.`
            : `The document did not grow, so the stuck pending edits never reached it.`),
    );
  } else {
    log(
      `\n[swarm] OK: ${opts.n} clients byte-identical ${convergedMs}ms after the churn stopped ` +
        `(${hotDrifted.length} drifted during the churn` +
        `${healPresent ? `, ${healDelta} self-heal${healDelta === 1 ? "" : "s"} observed` : ""}).`,
    );
  }
  if (stats.errors.length) log(`[swarm] typist errors: ${stats.errors.slice(0, 3).join(" | ")}`);

  await shutdown(convergedMs === null ? 1 : 0);
}

main().catch(async (e) => {
  console.error("[swarm]", e);
  await shutdown(1);
});
