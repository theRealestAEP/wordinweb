import { useEffect, useState } from "react";
import type { BundleStore, StoredDocSummary } from "wordinweb/collab";
import { fmtSize } from "./file-menu";

/**
 * Startup reclaim for this browser's bundle store.
 *
 * THE DEBT THIS PAYS OFF: until e8c3188 every "Save version" wrote a full
 * document copy to IndexedDB and nothing ever deleted one — `.slice(-25)`
 * trimmed only the on-screen strip. version-retention.ts stops NEW saves from
 * leaking, but the copies that already accumulated stay wedged in the quota
 * forever: a hard refresh does not clear IndexedDB, so an owner who leaked
 * gigabytes sees "storage is full or blocked" on every load with no way out.
 * This runs ON LOAD, before the app needs to write anything, and reclaims
 * what is provably redundant.
 *
 * WHAT IS SAFE TO RECLAIM, IN ORDER — the ordering is the whole design:
 *
 *  1. `superseded` copies — banked after a fast-forward; their content is
 *     almost always still present in the live bundle. Reclaimed first.
 *  2. `version` entries — deliberate restore points, but by construction a
 *     copy of a bundle that also existed live. Reclaimed oldest-first, and
 *     NEVER a document's newest version (same rule as pruneVersions).
 *
 * NEVER automatically: the live bundle, `draft` entries, `local:` archives,
 * or `unknown` keys. A draft is the ONLY copy of someone's diverged offline
 * work; a `local:` archive the only copy of a document they never went live
 * with; an unknown key is a future build's data this build cannot judge.
 * Under zero custody the browser holds the only durable copy — deleting any
 * of those is data loss, not housekeeping. If reclaiming everything safe
 * still leaves the store over budget, that is SAID (overBudget), and the way
 * out is the manual two-step delete in File > Saved — never a silent delete
 * of unique work to make a number fit.
 *
 * Reclaim stops as soon as the store is under budget: even a redundant copy
 * is kept when there is room for it (bias toward keeping).
 *
 * THE BUDGET: measured bytes come from `list()` — real sizes, independent of
 * `navigator.storage.estimate()`, which may be absent or lie. The estimate
 * only sets the THRESHOLD: half the reported quota, capped at 1 GiB (Chromium
 * reports ~60% of free disk; "half" of that must not mean gigabytes of
 * orphans are fine), falling back to a conservative 256 MiB where the quota
 * is unknowable (the browsers with the stingiest quotas).
 */

const STORE_QUOTA_SHARE = 0.5;
const STORE_BUDGET_CAP = 1024 * 1024 * 1024;
const STORE_BUDGET_FALLBACK = 256 * 1024 * 1024;

/** The whole-store byte threshold above which startup reclaim runs. */
export async function storeByteBudget(): Promise<number> {
  try {
    const est = await globalThis.navigator?.storage?.estimate?.();
    if (est?.quota && Number.isFinite(est.quota)) {
      return Math.min(STORE_BUDGET_CAP, est.quota * STORE_QUOTA_SHARE);
    }
  } catch {
    // estimate() itself failing degrades to the fallback below.
  }
  return STORE_BUDGET_FALLBACK;
}

export interface ReclaimReport {
  /** What was deleted, in deletion order (superseded first, then versions oldest-first). */
  reclaimed: StoredDocSummary[];
  reclaimedBytes: number;
  /** What remains in the store after reclaim. */
  keptCount: number;
  keptBytes: number;
  budgetBytes: number;
  /** Still over budget after everything safe was reclaimed — the rest is
   * unique work, and only the user may delete it (File > Saved). */
  overBudget: boolean;
}

const age = (s: StoredDocSummary) => s.versionSavedAt ?? s.savedAt;

/** Reclaim redundant stored copies down to `budgetBytes`. See module doc for
 * the order and the never-touch list. Deletions are RETURNED, not swallowed:
 * the caller must tell the user what was dropped. */
export async function reclaimStorage(store: BundleStore, budgetBytes: number): Promise<ReclaimReport> {
  const all = await store.list();
  let total = all.reduce((n, s) => n + s.byteLength, 0);
  const reclaimed: StoredDocSummary[] = [];
  let freed = 0;
  if (total > budgetBytes) {
    const superseded = all.filter((s) => s.kind === "superseded").sort((a, b) => age(a) - age(b));
    // A document's NEWEST version is never a candidate, whatever the budget.
    const versions = all.filter((s) => s.kind === "version");
    const newest = new Map<string, StoredDocSummary>();
    for (const v of versions) {
      const cur = newest.get(v.docId);
      if (!cur || age(v) > age(cur)) newest.set(v.docId, v);
    }
    const olderVersions = versions.filter((v) => newest.get(v.docId) !== v).sort((a, b) => age(a) - age(b));
    for (const c of [...superseded, ...olderVersions]) {
      if (total <= budgetBytes) break;
      await store.delete(c.key);
      reclaimed.push(c);
      freed += c.byteLength;
      total -= c.byteLength;
    }
  }
  return {
    reclaimed,
    reclaimedBytes: freed,
    keptCount: all.length - reclaimed.length,
    keptBytes: total,
    budgetBytes,
    overBudget: total > budgetBytes,
  };
}

/** What a reclaimed entry is called in the notice (same naming rule as the
 * File menu: never the docId — it is the room capability). */
function reclaimName(s: StoredDocSummary): string {
  const when = new Date(age(s)).toLocaleString([], { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  if (s.kind === "version") return s.label ? `version “${s.label}”` : `version (${when})`;
  return `superseded copy (${when})`;
}

/** The on-screen report for a reclaim, or null when there is nothing to say.
 * Same rule as version-retention: reclaiming silently is how people lose
 * track of their restore points. */
export function reclaimMessage(r: ReclaimReport): string | null {
  const parts: string[] = [];
  if (r.reclaimed.length) {
    const names = r.reclaimed.map(reclaimName);
    const shown = names.length > 4 ? `${names.slice(0, 3).join(", ")} and ${names.length - 3} more` : names.join(", ");
    parts.push(
      `Freed ${fmtSize(r.reclaimedBytes)} of this browser’s storage by removing ${r.reclaimed.length} redundant stored ` +
        `cop${r.reclaimed.length === 1 ? "y" : "ies"}: ${shown}. Drafts, local documents and each document’s newest version were kept.`,
    );
  }
  if (r.overBudget) {
    parts.push(
      `This browser’s storage is still holding ${fmtSize(r.keptBytes)} across ${r.keptCount} saved ` +
        `item${r.keptCount === 1 ? "" : "s"}. The rest is unique work (drafts, local documents), so nothing more is removed ` +
        `automatically — review and delete what you no longer need under File > Saved in this browser.`,
    );
  }
  return parts.length ? parts.join(" ") : null;
}

/** The message when the store cannot even be listed — degrade to words, not
 * a crash: nothing was deleted, and the manual route still gets named. */
export const RECLAIM_LIST_FAILED_MESSAGE =
  "Couldn’t check this browser’s stored documents — storage may be blocked. Nothing was deleted; if saving fails, manage copies under File > Saved.";

/**
 * Run reclaim once on mount and hold the dismissible notice. Both screens
 * (local editor and collab) mount over the same store; a second run after a
 * screen switch is a cheap no-op — the store is already under budget.
 */
export function useStartupReclaim(store: BundleStore): { notice: string | null; dismiss: () => void } {
  const [notice, setNotice] = useState<string | null>(null);
  useEffect(() => {
    let alive = true;
    storeByteBudget()
      .then((budget) => reclaimStorage(store, budget))
      .then(
        (report) => {
          if (alive) setNotice(reclaimMessage(report));
        },
        () => {
          if (alive) setNotice(RECLAIM_LIST_FAILED_MESSAGE);
        },
      );
    return () => {
      alive = false;
    };
  }, [store]);
  return { notice, dismiss: () => setNotice(null) };
}
