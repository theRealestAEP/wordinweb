#!/usr/bin/env node
/**
 * Publish gate: a built tarball must not reference a package that consumers
 * cannot install. Wired into `prepack` of every published package.
 *
 * The bug this exists for: `noExternal` bundles the JS, but the .d.ts rollup
 * resolves separately and inlines a workspace package only when tsconfig
 * `paths` maps that exact specifier to source. `@wordinweb/core` was mapped;
 * `@wordinweb/collab/client` was not — so the shipped declarations carried
 * `import { ... } from '@wordinweb/collab/client'`, a package that is
 * private:true and never published. Consumers with skipLibCheck:false got a
 * hard TS2307; consumers with skipLibCheck:true (the common setting) got
 * something worse because it is silent: every re-exported type degraded to
 * `any`, so wrong code type-checked clean.
 *
 * Nothing in the 1213-test battery could catch it. Those suites run against
 * SOURCE, where the workspace symlink resolves fine. The defect exists only
 * in the packaged artifact, so the check has to look at the artifact.
 *
 * Rule: every bare specifier in dist/ must be a declared dependency or
 * peerDependency (or a node builtin). Declared-but-unpublished is out of
 * scope here — npm install fails loudly on that, which is the easy case.
 */
import { readFileSync, readdirSync } from "node:fs";
import { builtinModules } from "node:module";
import path from "node:path";

// Runs from any package's `prepack` (cwd is the package dir), or takes an
// explicit package dir as the first argument.
const pkgDir = process.argv[2] ? path.resolve(process.argv[2]) : process.cwd();
const pkg = JSON.parse(readFileSync(path.join(pkgDir, "package.json"), "utf8"));
const distDir = path.join(pkgDir, "dist");

const allowed = new Set([
  ...Object.keys(pkg.dependencies ?? {}),
  ...Object.keys(pkg.peerDependencies ?? {}),
  ...builtinModules,
  // A package may reference itself by name through its own `exports` map;
  // that always resolves for a consumer, since they installed it.
  pkg.name,
]);

/** Comments hold prose like `import { DocxView } from "wordinweb"` — not imports. */
const stripComments = (src) => src.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^[ \t]*\/\/.*$/gm, "");

// A real bare specifier is a package name plus optional subpath. Requiring
// that shape drops parse artifacts that a text scan inevitably picks up out
// of ordinary code — e.g. `attrs["from"] ?? "0,0"` reads as `from"…"` and
// yielded a "package" named `] ?? `. Anything genuinely importable matches.
const PACKAGE_SPECIFIER = /^(?:@[a-z0-9-~][a-z0-9-._~]*\/)?[a-z0-9-~][a-z0-9-._~]*(?:\/[^\s"']+)*$/i;

/** `@scope/name/sub/path` -> `@scope/name`; `name/sub` -> `name`. */
const packageNameOf = (spec) => {
  const parts = spec.split("/");
  return spec.startsWith("@") ? parts.slice(0, 2).join("/") : parts[0];
};

// `from "x"`, `from'x'`, `import("x")`, `require("x")` — covers both the
// declaration re-exports and the runtime chunks. The whitespace class is
// `[ \t]`, deliberately NOT `\s`: `\s` crosses newlines, so a stray `from`
// at the end of one line paired with a quoted string at the start of the
// next produced phantom "packages" like `margin` out of ordinary code.
const SPECIFIER = /(?:\bfrom[ \t]*|\bimport[ \t]*\([ \t]*|\brequire[ \t]*\([ \t]*)(["'])([^"'\n]+)\1/g;

let files;
try {
  // `recursive`: tsc-built packages (core, collab, server) mirror their src
  // directory tree into dist; tsup-built ones (react) emit flat.
  files = readdirSync(distDir, { recursive: true }).filter(
    (f) => f.endsWith(".d.ts") || (f.endsWith(".js") && !f.endsWith(".map")),
  );
} catch {
  console.error(`check-dts-selfcontained: no build output at ${distDir} — run the build first.`);
  process.exit(1);
}
if (files.length === 0) {
  console.error(`check-dts-selfcontained: ${distDir} is empty — run the build first.`);
  process.exit(1);
}

const violations = [];
for (const file of files) {
  const text = stripComments(readFileSync(path.join(distDir, file), "utf8"));
  for (const [, , spec] of text.matchAll(SPECIFIER)) {
    if (spec.startsWith(".") || spec.startsWith("/") || spec.startsWith("node:")) continue;
    if (!PACKAGE_SPECIFIER.test(spec)) continue;
    const name = packageNameOf(spec);
    if (allowed.has(name)) continue;
    violations.push({ file, spec, name });
  }
}

if (violations.length > 0) {
  console.error("check-dts-selfcontained: dist/ references packages a consumer cannot install:\n");
  for (const v of violations) {
    console.error(`  ${v.file}: "${v.spec}"  (package "${v.name}" is not a dependency or peerDependency)`);
  }
  console.error(
    "\nEither declare the package as a dependency/peerDependency, or (for the\n" +
      "tsup-bundled `wordinweb` package) map its EXACT specifier — including any\n" +
      "subpath — to source in the package's tsconfig `paths` so the .d.ts rollup\n" +
      "inlines it. `dts.resolve` in tsup.config.ts looks like the knob but is\n" +
      "not: tsup turns every `paths` key into an ignore-list for it.\n",
  );
  process.exit(1);
}

console.log(`check-dts-selfcontained: OK — ${files.length} built files, no undeclared package references.`);
