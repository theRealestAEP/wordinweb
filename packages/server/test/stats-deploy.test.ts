import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

const example = new URL("../../../examples/anon-share/", import.meta.url);

describe("operations dashboard stats", () => {
  it("enables /stats inside both Compose stacks without exposing it through Caddy", () => {
    for (const file of ["compose.yml", "compose.cloudflare.yml"]) {
      const compose = readFileSync(new URL(file, example), "utf8");
      expect(compose).toContain('WW_OBS: "${WW_OBS:-1}"');
    }

    for (const file of ["deploy/Caddyfile", "deploy/Caddyfile.tunnel"]) {
      const caddy = readFileSync(new URL(file, example), "utf8");
      expect(caddy).not.toMatch(/@collab path [^\n]*\/stats/);
    }
  });
});
