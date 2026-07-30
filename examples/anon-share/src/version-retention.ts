import type { BundleStore, StoredDocSummary } from "wordinweb/collab";

/**
 * Retention for stored versions ("Save version" restore points).
 *
 * THE BUG THIS FIXES: saveVersion wrote a new `<docId>#version-…` bundle on
 * every save and trimmed only the in-memory strip (`.slice(-25)`). Nothing
 * ever deleted a stored version, so every save left another full copy of the
 * document in IndexedDB forever — at 500 pages that is tens of megabytes per
 * click, and a handful of saves exhausts the origin quota ("it said it was
 * out of space").
 *
 * POLICY — bytes first, count second:
 *
 *  - The primary bound is TOTAL BYTES, because versions are full document
 *    copies and their cost scales with the document, not their number.
 *    25 versions of a 50 KB memo is ~1 MB; 25 versions of a 30 MB filing is
 *    ~750 MB. A count cap alone is no bound at all.
 *  - The budget comes from `navigator.storage.estimate()` where available:
 *    versions may use up to a quarter of the origin quota, capped at 512 MiB.
 *    A quarter leaves the rest of the quota for what versions must never
 *    crowd out — the live bundles, drafts, superseded banks and local
 *    archives that hold UNIQUE work (a version is by construction a copy of
 *    a bundle that also exists live). The 512 MiB cap matters on Chromium,
 *    where the quota is ~60% of free disk and "a quarter" would otherwise be
 *    tens of gigabytes of restore points. On Firefox (~10% of disk, ≤10 GiB)
 *    the cap usually still binds; on Safari (~1 GiB without a prompt) the
 *    quarter (~256 MiB) binds instead.
 *  - Where `estimate()` is absent (older Safari, some private modes — the
 *    stingiest quotas, and unknowable) the fallback is a conservative
 *    64 MiB: room for a couple of restore points of a big document, dozens
 *    of a small one.
 *  - The count cap (25) stays as the secondary bound so a tiny document
 *    cannot accumulate hundreds of chips.
 *
 * WHAT IS PRUNED: only `kind === "version"` entries of THIS docId, oldest
 * first, as one contiguous recent window — never the live bundle, never the
 * newest version, never drafts/superseded/local archives (those hold work
 * that exists nowhere else; they are listed under File > Saved with a manual
 * delete instead). A window rather than best-fit packing because "your last
 * N saves, newest back" is a retention rule a person can predict; "we kept
 * March 3rd but dropped March 9th because it was bigger" is not.
 *
 * Deletions are RETURNED, not swallowed: the caller must tell the user what
 * was dropped. Silently deleting a restore point is worse than refusing to
 * make a new one.
 */

export const VERSION_COUNT_CAP = 25;
const QUOTA_SHARE = 0.25;
const BUDGET_CAP = 512 * 1024 * 1024;
const BUDGET_FALLBACK = 64 * 1024 * 1024;

/** The version byte budget for this browser (see policy above). */
export async function versionByteBudget(): Promise<number> {
  try {
    const est = await globalThis.navigator?.storage?.estimate?.();
    if (est?.quota && Number.isFinite(est.quota)) {
      return Math.min(BUDGET_CAP, est.quota * QUOTA_SHARE);
    }
  } catch {
    // estimate() itself failing degrades to the fallback below.
  }
  return BUDGET_FALLBACK;
}

/**
 * Delete this doc's stored versions down to the byte budget and count cap,
 * keeping the newest `keepNewest` versions unconditionally (1 = the newest
 * version is never pruned; 0 is for making room BEFORE a save that could
 * not land). Returns what was deleted, oldest first.
 */
export async function pruneVersions(
  store: BundleStore,
  docId: string,
  budgetBytes: number,
  { countCap = VERSION_COUNT_CAP, keepNewest = 1 }: { countCap?: number; keepNewest?: number } = {},
): Promise<StoredDocSummary[]> {
  const versions = (await store.list())
    .filter((s) => s.kind === "version" && s.docId === docId)
    .sort((a, b) => (a.versionSavedAt ?? a.savedAt) - (b.versionSavedAt ?? b.savedAt)); // oldest → newest
  let bytes = 0;
  let cut = 0; // versions[0..cut) are dropped
  for (let i = versions.length - 1; i >= 0; i--) {
    const kept = versions.length - 1 - i; // how many newer ones are already kept
    bytes += versions[i].byteLength;
    if (kept < keepNewest) continue; // unconditional keeps
    if (bytes > budgetBytes || kept + 1 > countCap) {
      cut = i + 1; // this one and everything older goes (contiguous window)
      break;
    }
  }
  const dropped = versions.slice(0, cut);
  for (const v of dropped) await store.delete(v.key);
  return dropped;
}
