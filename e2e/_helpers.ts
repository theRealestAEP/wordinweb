import { expect, type Page, type Browser, type BrowserContext } from "@playwright/test";

/**
 * Shared helpers for the browser E2E suite (plan doc 09). The demo learns
 * the collab server's port from `?server=`; the Playwright config boots it
 * on 1399 (see playwright.config.ts). The demo opens to a LOCAL editor;
 * "Make collaborative" seals the current doc and goes live (encrypted).
 *
 * The demo exposes a dev-only `window.__ww` hook (defined in
 * examples/anon-share/src/app.tsx) — DO NOT rely on any member not listed
 * in `WwHook` below:
 *   _session   the raw CollabSession (for deep inspection)
 *   submitOp   (intent) => void        submit through the real connection
 *   admin      (action) => void        owner admin op (no-op unless owner)
 *   allocIds   (n) => number[]         real disjoint carried ids
 *   ready      boolean
 *   roster     () => RosterEntry[]
 *   saveB64    () => string | null     canonical docx, byte-identical check
 *   text       () => string            concatenated w:t text
 */

export const SERVER = "localhost:1399";
export const LANDING = `/?server=${SERVER}`;
export const PAGE = ".dxw-page";

/**
 * The code every board scenario goes live with.
 *
 * A share code is MANDATORY now (owner's decision): a link is a capability,
 * and the code is the second factor that stops a leaked link from being one.
 * That makes it part of every scenario's setup rather than a feature one spec
 * exercises, so it lives here — a spec that doesn't care about codes says
 * nothing about them and gets this one.
 *
 * Deliberately NOT in the URL: putting it there would hand the board a
 * document that opens from the link alone, which is precisely the property
 * the code exists to remove, and every join-side assertion would then be
 * passing for the wrong reason.
 */
export const BOARD_CODE = "board-code";

export interface WwHook {
  _session: unknown;
  submitOp(i: unknown): void;
  admin(a: unknown): void;
  allocIds(n: number): number[];
  ready: boolean;
  roster(): unknown[];
  saveB64(): string | null;
  text(): string;
}

/** Wait until the demo's __ww hook is live AND its core members are callable
 * (guards against a Vite pre-bundle race that can briefly serve a partial
 * session). Call after the collab editor has mounted. */
export async function waitHook(page: Page): Promise<void> {
  await expect
    .poll(() =>
      page.evaluate(() => {
        const w = (window as unknown as { __ww?: WwHook }).__ww;
        return !!w && typeof w.submitOp === "function" && typeof w.allocIds === "function" && typeof w.saveB64 === "function";
      }),
    )
    .toBe(true);
}

/** Open the demo's LOCAL editor (no `?doc=`); returns the page ready to type
 * locally. */
export async function openLocal(page: Page): Promise<void> {
  await page.goto(LANDING);
  await expect(page.getByTestId("local-editor")).toBeVisible();
}

/**
 * From the LOCAL editor, go collaborative and wait for the session to mount.
 * Returns the share URL (with `?doc=` and `#k=`).
 *
 * The flow is: "Collab" opens a dialog (a REQUIRED share code + what a code
 * does), then "Start Collab" seals and goes live. This helper is the single
 * choke point every spec goes through, so the dialog lives here rather than in
 * a dozen specs. Pass `shareCode` for a scenario that needs a specific one;
 * everything else gets BOARD_CODE, because going live without a code is no
 * longer a thing the product can do.
 */
export async function goLive(page: Page, shareCode: string = BOARD_CODE): Promise<string> {
  await page.getByTestId("make-collaborative").click();
  await expect(page.getByTestId("collab-modal")).toBeVisible();
  await page.getByTestId("share-code").fill(shareCode);
  await page.getByTestId("start-collab").click();
  await expect(page).toHaveURL(/[?&]doc=/);
  await expect(page).toHaveURL(/#k=/);
  await expect(page.getByTestId("toolbar")).toBeVisible(); // collab chrome mounted
  await expect(page.locator(PAGE)).toBeVisible();
  await waitHook(page);
  return page.url();
}

/** Create a fresh collaborative doc from a blank local editor. */
export async function createCollabDoc(page: Page, shareCode: string = BOARD_CODE): Promise<string> {
  await openLocal(page);
  return goLive(page, shareCode);
}

/**
 * Supply the share code IF this browser is asked for it.
 *
 * "If" is load-bearing and cannot be assumed either way: the creator's own
 * browser (and any tab sharing its localStorage) already holds the code and
 * joins straight through, while a fresh context is refused `code-required`
 * and prompted. Both are correct, and which one a given spec produces depends
 * on whether it opened a new context — so this waits for WHICHEVER arrives,
 * the document or the prompt, instead of assuming.
 */
export async function enterCodeIfPrompted(page: Page, shareCode: string = BOARD_CODE): Promise<void> {
  const prompt = page.getByTestId("join-share-code");
  await expect
    .poll(async () => (await prompt.count()) > 0 || (await page.locator(PAGE).first().count()) > 0,
      { message: "neither the document nor the share-code prompt appeared", timeout: 15_000 })
    .toBe(true);
  if ((await prompt.count()) === 0) return;
  await prompt.fill(shareCode);
  await page.getByTestId("join-submit").click();
}

/** Join an existing collab doc via its share URL (the `?doc=#k=` link), and
 * the code that link is not enough without. */
export async function joinCollab(page: Page, url: string, shareCode: string = BOARD_CODE): Promise<void> {
  await page.goto(url);
  await enterCodeIfPrompted(page, shareCode);
  await expect(page.locator(PAGE)).toBeVisible();
  await waitHook(page);
}

/** Submit an intent through the real connection (`session.submitOp`). */
export async function submitOp(page: Page, intent: Record<string, unknown>): Promise<void> {
  await page.evaluate((i) => (window as unknown as { __ww: WwHook }).__ww.submitOp(i), intent);
}

/** Allocate real disjoint carried ids from the connection. */
export async function allocIds(page: Page, n: number): Promise<number[]> {
  return page.evaluate((k) => (window as unknown as { __ww: WwHook }).__ww.allocIds(k), n);
}

/** Owner admin op (kick / readOnly / setRole) — no-op unless this page's
 * connection proved the owner token. */
export async function admin(page: Page, action: Record<string, unknown>): Promise<void> {
  await page.evaluate((a) => (window as unknown as { __ww: WwHook }).__ww.admin(a), action);
}

/**
 * `__ww` EXISTS ONLY WHILE A SESSION IS MOUNTED. The demo sets the hook from a
 * `useEffect` keyed on the session and clears it to null in between, so any
 * read taken while a client is (re)mounting — a takeover, a revival, a slow
 * first paint under a ten-client flood — hits null.
 *
 * These readers therefore return null rather than dereferencing it. Every
 * caller is a POLL (settleAll, expect.poll, census), so a null simply means
 * "not yet, ask again", which is what the caller already does for a client
 * that has not converged. Dereferencing threw
 * `Cannot read properties of undefined (reading 'saveB64')` and failed the
 * whole run instead — a harness race presenting as a product failure, seen
 * once in a full-board run and indistinguishable from real drift in the
 * report.
 */
export async function saveB64(page: Page): Promise<string | null> {
  return page.evaluate(() => {
    const w = (window as unknown as { __ww?: WwHook }).__ww;
    return w ? w.saveB64() : null;
  });
}

export async function docText(page: Page): Promise<string> {
  return page.evaluate(() => {
    const w = (window as unknown as { __ww?: WwHook }).__ww;
    return w ? w.text() : "";
  });
}

/** The text painted on screen (from the live DocxView). */
export async function paintedText(page: Page): Promise<string> {
  return (await page.locator(PAGE).first().textContent()) ?? "";
}

/** Poll until every page holds byte-identical docx (`saveB64`). */
export async function converge(pages: Page[], label = "convergence"): Promise<void> {
  await expect
    .poll(
      async () => {
        const bs = await Promise.all(pages.map(saveB64));
        return bs[0] !== null && bs.every((b) => b === bs[0]);
      },
      { message: `clients did not converge (${label})`, timeout: 8000 },
    )
    .toBe(true);
}

/** Click precisely on the first text line (not the page center — a center
 * click on a one-line doc lands in empty space and misplaces the caret). */
export async function clickTextStart(page: Page): Promise<void> {
  const box = await page.locator(PAGE).first().boundingBox();
  if (box) await page.mouse.click(box.x + 30, box.y + 25);
  else await page.locator(PAGE).first().click();
}

/** Type into the editor via the real keydown path (precise click first). */
export async function typeInEditor(page: Page, text: string): Promise<void> {
  await clickTextStart(page);
  await page.keyboard.type(text, { delay: 15 });
}

/** Fresh isolated contexts + pages (each mints its own per-browser clientId). */
export async function newClients(browser: Browser, n: number): Promise<{ pages: Page[]; contexts: BrowserContext[] }> {
  const contexts = await Promise.all(Array.from({ length: n }, () => browser.newContext()));
  const pages = await Promise.all(contexts.map((c) => c.newPage()));
  return { pages, contexts };
}

/* ------------------------------------------------------------------ *
 * SWARM additions (e2e/swarm.spec.ts + scripts/swarm.mjs).
 *
 * Everything below is ADDITIVE — a many-client run needs three things the
 * small specs never did: a census (who agrees with whom, not just "do all
 * agree"), a document big enough to be built by a ladder of carried ids, and
 * a feature-detected read of the connection's self-heal counter.
 * ------------------------------------------------------------------ */

/** One grep-able measurement line, same convention as stress.spec.ts. */
export function metric(scenario: string, fields: Record<string, string | number>): void {
  const body = Object.entries(fields)
    .map(([k, v]) => `${k}=${typeof v === "number" ? (Number.isInteger(v) ? v : v.toFixed(1)) : v}`)
    .join(" ");
  console.log(`STRESS-METRIC ${scenario} ${body}`);
}

/** A stable 32-bit digest of a base64 document, as 8 hex chars. Comparing
 * whole `saveB64` strings across 30 clients is correct but unreadable; the
 * short hash is what goes in a status table. */
export function shortHash(s: string | null): string {
  if (s === null) return "null";
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

export interface Census {
  /** Per-client short hash, index-aligned with the pages passed in. */
  hashes: string[];
  /** hash → client indices holding it. */
  groups: Map<string, number[]>;
  /** The hash held by the most clients (ties broken by first occurrence). */
  majority: string;
  /** Clients not in the majority group. */
  drifted: number[];
  /** Per-client document text length. */
  lengths: number[];
  /** Largest |client chars − majority chars| across all clients. */
  maxDriftChars: number;
}

/** Who agrees with whom. With two clients "converged" is a boolean; with ten
 * it is a distribution, and a single straggler must be nameable. */
export async function census(pages: Page[]): Promise<Census> {
  const hashes = (await Promise.all(pages.map(saveB64))).map(shortHash);
  const lengths = await Promise.all(pages.map((p) => docText(p).then((t) => t.length)));
  const groups = new Map<string, number[]>();
  hashes.forEach((h, i) => groups.set(h, [...(groups.get(h) ?? []), i]));
  let majority = hashes[0];
  for (const [h, members] of groups) if (members.length > (groups.get(majority)?.length ?? 0)) majority = h;
  const drifted = hashes.map((h, i) => (h === majority ? -1 : i)).filter((i) => i >= 0);
  const majorityLen = lengths[groups.get(majority)![0]];
  const maxDriftChars = Math.max(0, ...lengths.map((l) => Math.abs(l - majorityLen)));
  return { hashes, groups, majority, drifted, lengths, maxDriftChars };
}

/** Non-throwing convergence probe: ms until every client is byte-identical,
 * or null if they never were within `timeout`. (stress.spec has a private
 * twin; the swarm needs it from the shared module and from the CLI tool.) */
export async function settleAll(pages: Page[], timeout: number): Promise<number | null> {
  const t0 = Date.now();
  for (;;) {
    const bs = await Promise.all(pages.map(saveB64));
    if (bs[0] !== null && bs.every((b) => b === bs[0])) return Date.now() - t0;
    if (Date.now() - t0 > timeout) return null;
    await pages[0].waitForTimeout(200);
  }
}

/** The demo's live perf snapshot — null unless the HUD is armed (`?perf=1`).
 * Typed loosely on purpose: the swarm only reads a few fields and must not
 * break when the HUD grows new ones (e.g. `selfHeals`). */
export async function perfSnapshot(page: Page): Promise<Record<string, unknown> | null> {
  return page.evaluate(
    () => (window as unknown as { __ww?: { perf?: () => unknown } }).__ww?.perf?.() as Record<string, unknown> | null,
  );
}

/**
 * The encrypted connection's self-heal counter (quiescent replica-vs-mirror
 * check → rebuild), or null when this build doesn't expose one. A number,
 * including 0, means the self-heal path exists; null means legacy
 * reload-to-heal.
 *
 * BOTH exposures are probed: the demo hook grew `__ww.selfHeals()` directly,
 * and the counter may also surface on the perf snapshot. Checking only one
 * silently reports "absent" on a build that has it.
 */
export async function selfHeals(page: Page): Promise<number | null> {
  const direct = await page.evaluate(() => {
    const w = (window as unknown as { __ww?: { selfHeals?: () => number } }).__ww;
    return typeof w?.selfHeals === "function" ? w.selfHeals() : null;
  });
  if (typeof direct === "number") return direct;
  const s = await perfSnapshot(page);
  return typeof s?.selfHeals === "number" ? (s.selfHeals as number) : null;
}

/** The client's un-acked local intent count (HUD only, -1 without it). A
 * client whose pending never drains is stuck, not merely slow. */
export async function pendingCount(page: Page): Promise<number> {
  const s = await perfSnapshot(page);
  return typeof s?.pending === "number" ? (s.pending as number) : -1;
}

/** Reload a client and wait for its session to come back up — rebuilding from
 * the stored bundle + the server's tail is what heals a drifted replica. */
export async function reloadClient(page: Page): Promise<void> {
  await page.reload();
  await expect(page.locator(PAGE).first()).toBeVisible();
  await waitHook(page);
}

/** `joinCollab` waits on `.dxw-page` strictly, which a paginated (big)
 * document turns into a strict-mode violation. Same join, first page only. */
export async function joinCollabBig(page: Page, url: string, shareCode: string = BOARD_CODE): Promise<void> {
  await page.goto(url);
  await enterCodeIfPrompted(page, shareCode);
  await expect(page.locator(PAGE).first()).toBeVisible();
  await waitHook(page);
}

/** One rung of the seeding ladder: the paragraph's block + run, how many
 * characters that run holds (so a later edit can pick an IN-RANGE offset), and
 * how long into the seed it was submitted — the ladder is the only place the
 * demo lets you watch per-op cost change as the document grows. */
export interface Rung {
  blockId: number;
  runId: number;
  len: number;
  tMs: number;
}

/**
 * Seed `n` paragraphs through the real connection and return one rung per
 * paragraph. The ladder: fill the current run, split at its end, continue in
 * the run the split created — the blank document's first run is
 * `{blockId:1, runId:2}`. Returned rungs are the ONLY safe way to address a
 * paragraph later: nothing else in the demo exposes ids.
 */
export async function seedParagraphs(page: Page, n: number, pacingMs = 3): Promise<Rung[]> {
  return page.evaluate(
    async ([count, pace]) => {
      const ww = (window as unknown as { __ww: { submitOp(i: unknown): void; allocIds(k: number): number[] } }).__ww;
      const rungs: { blockId: number; runId: number; len: number; tMs: number }[] = [];
      const t0 = Date.now();
      let cur = { blockId: 1, runId: 2 };
      for (let i = 0; i < count; i++) {
        const text = `Paragraph ${i} - the quick brown fox jumps over the lazy dog.`;
        ww.submitOp({ kind: "insertText", at: { ...cur, offset: 0 }, text });
        const [newBlockId, newRunId] = ww.allocIds(2);
        ww.submitOp({ kind: "splitParagraph", at: { ...cur, offset: text.length }, newBlockId, newRunId });
        rungs.push({ ...cur, len: text.length, tMs: Date.now() - t0 });
        cur = { blockId: newBlockId, runId: newRunId };
        if (pace > 0) await new Promise((r) => setTimeout(r, pace));
      }
      rungs.push({ ...cur, len: 0, tMs: Date.now() - t0 });
      return rungs;
    },
    [n, pacingMs] as const,
  );
}

/** Wait until this client's HUD reports every local intent acked. Returns
 * false when no HUD is armed (nothing to wait on). */
export async function waitAcked(page: Page, timeout = 60_000): Promise<boolean> {
  const t0 = Date.now();
  for (;;) {
    const s = await perfSnapshot(page);
    if (!s) return false;
    if (s.pending === 0) return true;
    if (Date.now() - t0 > timeout) return true;
    await page.waitForTimeout(150);
  }
}

/**
 * Runs IN THE BROWSER: a click point on the LAST piece of text in the
 * document (the visually lowest, right-most text leaf of the last mounted
 * page), scrolled into view. Null while that page is unmounted — the view
 * only mounts pages near the viewport. Searching by TEXT would not work: the
 * layout engine emits a positioned span per measured chunk, so no single
 * element holds a whole sentence.
 */
export function tailClickPoint(): { x: number; y: number } | null {
  let last: HTMLElement | null = null;
  for (const page of Array.from(document.querySelectorAll<HTMLElement>(".dxw-page"))) {
    const leaves = Array.from(page.querySelectorAll<HTMLElement>("*")).filter(
      (el) => !el.children.length && (el.textContent ?? "").trim(),
    );
    if (leaves.length) last = leaves[leaves.length - 1];
  }
  if (!last) return null;
  last.scrollIntoView({ block: "center" });
  const r = last.getBoundingClientRect();
  if (r.width === 0 || r.top < 0 || r.bottom > window.innerHeight) return null;
  return { x: r.right - 2, y: r.y + r.height / 2 };
}

/** Scroll the editor's scrolling ancestor to the bottom (mounts the last
 * page of a paginated document so its text can be clicked). */
export async function scrollToEnd(page: Page): Promise<void> {
  await page.evaluate(() => {
    let el: HTMLElement | null = document.querySelector<HTMLElement>(".dxw-pages");
    while (el && el.scrollHeight <= el.clientHeight + 4) el = el.parentElement;
    (el ?? document.scrollingElement)?.scrollTo({ top: 10 ** 7 });
  });
}
