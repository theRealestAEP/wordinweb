// Shard 4 of 8. One file per shard is the whole point: vitest hands work
// to workers per FILE, so the shards must be separate files to run in parallel.
// Generated shape, kept explicit rather than generated at run time so the pool
// can see all 8 units before it starts scheduling.
import { describe, it } from "vitest";
import { AUDIT_ENABLED, auditCorpusShard } from "./fixture-audit.js";

describe.runIf(AUDIT_ENABLED)("agent representation parity corpus (shard 4/8)", () => {
  it("represents every semantic component and every laid-out page in its slice", async () => {
    await auditCorpusShard(4, 8);
  });
});
