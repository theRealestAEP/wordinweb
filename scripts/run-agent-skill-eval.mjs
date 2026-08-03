import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const evalDirectory = resolve(root, "packages/agent/evals");
const fixtures = readdirSync(evalDirectory)
  .filter((name) => name.endsWith(".fixture.json"))
  .sort()
  .map((name) => JSON.parse(readFileSync(resolve(evalDirectory, name), "utf8")));
const requested = process.argv.slice(2);
const selected = requested.length === 0 ? fixtures : fixtures.filter((fixture) => requested.includes(fixture.name));
const bundledCodex = "/Applications/ChatGPT.app/Contents/Resources/codex";
const codexBin = process.env.CODEX_BIN ?? (existsSync(bundledCodex) ? bundledCodex : "codex");

for (const name of requested) {
  if (!fixtures.some((fixture) => fixture.name === name)) throw new Error(`Unknown agent evaluation fixture: ${name}`);
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    cwd: options.cwd ?? root,
    env: { ...process.env, ...options.env },
    encoding: "utf8",
    stdio: options.stdio ?? "inherit",
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`${command} exited with status ${result.status}`);
}

run("npm", ["run", "build", "-w", "@wordinweb/core", "-w", "@wordinweb/collab", "-w", "@wordinweb/agent"]);

for (const fixture of selected) {
  const workspace = mkdtempSync(join(tmpdir(), `wordinweb-agent-eval-${fixture.name}-`));
  const output = resolve(workspace, "output.docx");
  let passed = false;
  process.stdout.write(`Agent evaluation fixture: ${fixture.name}\nAgent evaluation workspace: ${workspace}\n`);

  try {
    const skillTarget = resolve(workspace, ".agents/skills/wordinweb-documents");
    mkdirSync(dirname(skillTarget), { recursive: true });
    symlinkSync(resolve(root, "skills/wordinweb-documents"), skillTarget, "dir");

    const packageTarget = resolve(workspace, "node_modules/@wordinweb/agent");
    mkdirSync(dirname(packageTarget), { recursive: true });
    symlinkSync(resolve(root, "packages/agent"), packageTarget, "dir");
    writeFileSync(resolve(workspace, "package.json"), JSON.stringify({ private: true, type: "module" }, null, 2));

    const prompt = [
      "Use the $wordinweb-documents skill in this workspace.",
      "Create output.docx with the public @wordinweb/agent API.",
      "Use compact composition for the initial document and progressive inspection for verification.",
      "Write and run a JavaScript authoring program inside this workspace.",
      "Inspect the semantic and spatial result before the program saves output.docx.",
      "Keep all files inside this workspace.",
      "",
      fixture.prompt,
    ].join("\n");
    const codexArgs = [
      "exec",
      "--ephemeral",
      "--ignore-user-config",
      "--ignore-rules",
      "--skip-git-repo-check",
      "--sandbox",
      "workspace-write",
      "--cd",
      workspace,
      "--output-last-message",
      resolve(workspace, "codex-final.txt"),
    ];
    if (process.env.WORDINWEB_AGENT_EVAL_MODEL) codexArgs.push("--model", process.env.WORDINWEB_AGENT_EVAL_MODEL);
    codexArgs.push(prompt);
    run(codexBin, codexArgs, { cwd: workspace });
    if (!existsSync(output)) throw new Error(`Codex did not create ${output}`);

    run("npx", ["vitest", "run", "--root", "packages/agent", "test/authoring-eval.test.ts"], {
      env: { WORDINWEB_AGENT_EVAL_OUTPUT: output, WORDINWEB_AGENT_EVAL_FIXTURE: fixture.name },
    });
    passed = true;
  } catch (error) {
    process.stderr.write(`Agent evaluation workspace: ${workspace}\n`);
    throw error;
  } finally {
    if (passed && process.env.WORDINWEB_KEEP_AGENT_EVAL !== "1") rmSync(workspace, { recursive: true, force: true });
  }
}
