// The memory table: every fixture given, old and new, each in its own Chromium.
//   node memory-table.mjs [--modes old,new] [--out table.json] <docx>...
import { execFile } from "node:child_process";
import { writeFileSync } from "node:fs";
import { promisify } from "node:util";

const args = process.argv.slice(2);
const opt = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const value = args[i + 1];
  args.splice(i, 2);
  return value;
};
const modes = opt("modes", "old,new").split(",");
const out = opt("out", null);
const script = new URL("./memory.mjs", import.meta.url).pathname;
const rows = [];
for (const file of args) {
  for (const mode of modes) {
    let row;
    try {
      const { stdout } = await promisify(execFile)(process.execPath, [script, "--mode", mode, file], { maxBuffer: 1 << 24, timeout: 60 * 60 * 1000 });
      row = JSON.parse(stdout.trim().split("\n").pop());
    } catch (error) {
      row = { name: file.split("/").pop(), mode, error: String(error.message).slice(0, 200) };
    }
    rows.push(row);
    console.log(JSON.stringify(row));
  }
}
if (out) writeFileSync(out, JSON.stringify(rows, null, 1));
console.log("\nfixture                                      mode  pages  peak RSS MB  peak JS heap MB   total s");
for (const r of rows) {
  console.log(`${r.name.padEnd(44)} ${r.mode.padEnd(5)} ${String(r.pages ?? "-").padStart(5)}  ${String(r.peakRssMB ?? r.error).padStart(11)}  ${String(r.peakHeapMB ?? "-").padStart(15)}  ${String(Math.round((r.ms ?? 0) / 1000)).padStart(8)}`);
}
