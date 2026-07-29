import { afterEach, describe, expect, it, vi } from "vitest";
import { DEFAULT_LIMITS, envLimits, envStr, normalizeLimits } from "../src/limits.js";

/** Doc 16 §4/§8 tier knobs: RAM budget, disk budget, spill dir. This pass is
 * plumbing only — diskBytes defaults to 0 (spill disabled) until the
 * eviction ladder lands, and these tests pin exactly that. */

const ENV_KEYS = ["WW_MEDIA_RAM_BYTES", "WW_MEDIA_DISK_BYTES", "WW_MEDIA_SPILL_DIR"];

afterEach(() => {
  for (const k of ENV_KEYS) delete process.env[k];
  vi.restoreAllMocks();
});

describe("media tier limits (doc 16 §8)", () => {
  it("defaults: 128 MB RAM, spill DISABLED (0), spill dir under the OS tmpdir", () => {
    expect(DEFAULT_LIMITS.media.ramBytes).toBe(128 * 1024 * 1024);
    expect(DEFAULT_LIMITS.media.diskBytes).toBe(0);
    expect(DEFAULT_LIMITS.media.spillDir.endsWith("wordinweb-spill")).toBe(true);
  });

  it("reads the three env vars through the ordinary conventions", () => {
    process.env.WW_MEDIA_RAM_BYTES = String(256 * 1024 * 1024);
    process.env.WW_MEDIA_DISK_BYTES = String(2 * 1024 * 1024 * 1024);
    process.env.WW_MEDIA_SPILL_DIR = "/scratch/spill";
    const l = envLimits();
    expect(l.media.ramBytes).toBe(256 * 1024 * 1024);
    expect(l.media.diskBytes).toBe(2 * 1024 * 1024 * 1024);
    expect(l.media.spillDir).toBe("/scratch/spill");
  });

  it("envStr: unset and blank fall back; a value is trimmed", () => {
    expect(envStr("WW_MEDIA_SPILL_DIR", "/fallback")).toBe("/fallback");
    process.env.WW_MEDIA_SPILL_DIR = "   ";
    expect(envStr("WW_MEDIA_SPILL_DIR", "/fallback")).toBe("/fallback");
    process.env.WW_MEDIA_SPILL_DIR = "  /scratch/spill  ";
    expect(envStr("WW_MEDIA_SPILL_DIR", "/fallback")).toBe("/scratch/spill");
  });

  it("ceiling relation: ramBytes below maxBlobBytes is raised to it; 0 stays disabled", () => {
    const l = normalizeLimits({ media: { maxBlobBytes: 4 * 1024 * 1024, ramBytes: 1024 } });
    expect(l.media.ramBytes).toBe(4 * 1024 * 1024);
    const off = normalizeLimits({ media: { ramBytes: 0 } });
    expect(off.media.ramBytes).toBe(0);
    const ample = normalizeLimits({ media: { ramBytes: 512 * 1024 * 1024 } });
    expect(ample.media.ramBytes).toBe(512 * 1024 * 1024);
  });

  it("warns when a nonzero disk budget is smaller than one room's media cap", () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    normalizeLimits({ media: { diskBytes: 1024 } });
    expect(warn).toHaveBeenCalledOnce();
    expect(String(warn.mock.calls[0][0])).toContain("WW_MEDIA_DISK_BYTES");

    warn.mockClear();
    normalizeLimits({ media: { diskBytes: 0 } }); // disabled: silent
    normalizeLimits({ media: { diskBytes: 20 * 1024 * 1024 * 1024 } }); // ample: silent
    expect(warn).not.toHaveBeenCalled();
  });
});
