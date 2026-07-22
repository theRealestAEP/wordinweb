/**
 * Sources for edit-time values that are nondeterministic by default: the
 * wall clock (w:date attributes) and random ids (w14:paraId). Replicated
 * editing requires every replica to install byte-identical XML for the same
 * edit, so these values must be generated once — by the originating client —
 * and carried with the edit rather than re-derived at apply time. Local
 * editing uses the ambient defaults below.
 */
export interface EditProvenance {
  /** ISO-8601 timestamp written into w:date attributes. */
  date(): string;
  /** Fresh 8-hex-digit uppercase w14:paraId candidate. Callers still check
   * uniqueness against ids already present in the document and may call
   * again on collision. */
  paraId(): string;
}

/** Ambient defaults for local (non-replicated) editing. */
export function defaultProvenance(): EditProvenance {
  return {
    date: () => new Date().toISOString(),
    paraId: () =>
      Math.floor(Math.random() * 0xfffffff0 + 1)
        .toString(16)
        .toUpperCase()
        .padStart(8, "0"),
  };
}

/** Fixed-value provenance: replays recorded values in order. Used when
 * applying a replicated edit that carries the originator's values, and by
 * tests that need deterministic output. Throws if more values are drawn
 * than were recorded — a determinism bug, not a recoverable condition. */
export function recordedProvenance(values: { dates?: string[]; paraIds?: string[] }): EditProvenance {
  const dates = [...(values.dates ?? [])];
  const paraIds = [...(values.paraIds ?? [])];
  return {
    date: () => {
      const v = dates.shift();
      if (v === undefined) throw new Error("recordedProvenance: no date recorded for this draw");
      return v;
    },
    paraId: () => {
      const v = paraIds.shift();
      if (v === undefined) throw new Error("recordedProvenance: no paraId recorded for this draw");
      return v;
    },
  };
}
