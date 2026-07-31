/**
 * Presence selection ranges cross a trust boundary: the hub RELAYS presence
 * verbatim (it never validated caret coordinates either, and a blind E2EE
 * sequencer could not), so `sanitizePresencePosition` is the client-side
 * filter both connections apply on send and on receive.
 */
import { describe, expect, it } from "vitest";
import { PRESENCE_MAX_RANGES, sanitizePresencePosition, type PresencePosition } from "../src/protocol.js";

const anchor = { blockId: 1, runId: 2, offset: 3 };

describe("sanitizePresencePosition", () => {
  it("passes through the pre-ranges shape untouched (old clients keep working)", () => {
    const pos: PresencePosition = { anchor };
    expect(sanitizePresencePosition(pos)).toBe(pos); // same object: no allocation
    expect(sanitizePresencePosition(null)).toBeNull();
  });

  it("keeps well-formed ranges as-is", () => {
    const pos: PresencePosition = { anchor, ranges: [{ blockId: 1, runId: 2, start: 0, end: 4 }] };
    expect(sanitizePresencePosition(pos)).toBe(pos);
  });

  it("drops malformed ranges but keeps the caret", () => {
    const out = sanitizePresencePosition({
      anchor,
      ranges: [
        { blockId: 1, runId: 2, start: 0, end: 4 }, // good
        { blockId: 1, runId: 2, start: 4, end: 4 }, // empty
        { blockId: 1, runId: 2, start: 9, end: 2 }, // inverted
        { blockId: 1, runId: 2, start: -1, end: 2 }, // negative
        { blockId: 1, runId: 2, start: NaN, end: 2 },
        { blockId: 1, runId: 2, start: 0, end: Infinity },
        { blockId: 1, runId: 2, start: "0", end: "4" } as never, // wrong types
        null as never,
        "nope" as never,
      ],
    });
    expect(out!.anchor).toEqual(anchor);
    expect(out!.ranges).toEqual([{ blockId: 1, runId: 2, start: 0, end: 4 }]);
  });

  it("clamps a flood to PRESENCE_MAX_RANGES", () => {
    const ranges = Array.from({ length: 5000 }, () => ({ blockId: 1, runId: 2, start: 0, end: 4 }));
    expect(sanitizePresencePosition({ anchor, ranges })!.ranges).toHaveLength(PRESENCE_MAX_RANGES);
  });

  it("drops `ranges` entirely when nothing survives", () => {
    const out = sanitizePresencePosition({ anchor, ranges: [{ blockId: 1, runId: 2, start: 5, end: 1 }] });
    expect(out!.ranges).toBeUndefined();
    expect(out!.anchor).toEqual(anchor);
  });
});
