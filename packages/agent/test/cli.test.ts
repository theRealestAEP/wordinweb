import { mkdtempSync, symlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import { describe, expect, it } from "vitest";

describe("installed CLI", () => {
  it("starts through an npm-style bin symlink", () => {
    const directory = mkdtempSync(join(tmpdir(), "wordinweb-agent-cli-"));
    const binary = join(directory, "wordinweb-agent");
    symlinkSync(fileURLToPath(new URL("../dist/cli.js", import.meta.url)), binary);

    const result = spawnSync(process.execPath, [binary], { encoding: "utf8" });

    expect(result.status).toBe(1);
    expect(result.stderr).toContain("Usage: wordinweb-agent connect");
  });
});
