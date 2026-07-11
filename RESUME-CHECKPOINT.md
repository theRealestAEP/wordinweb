# Resume checkpoint — 2026-07-11 (spend-limit interrupt)

All six fix agents died simultaneously on the monthly Fable-5 spend limit,
mid-task. Their work is preserved as WIP commits in their worktrees (nothing
was merged to main; main is clean and green at the last committed state).

## The goal in flight
Drive every newly-promoted probe fixture below 1% (then the whole suite
<0.05%). Suite is now **1155 pages** (original 1108 all <1%, worst 0.82; the
29 pages ≥1% are all `probe2-*`/`probe3-*` new fixtures). Fresh full-run
stats are in `parity/out/results.json` (run took 378s parallel).

## Main branch state (safe baseline)
- HEAD has: 28 fixtures promoted into `apps/demo/public/fixtures`, parallel
  runner (raster cache + page-range sharding + provenance column), sub-1%
  campaign complete on the original corpus.
- 232 unit tests green. Dev server on :5299 serves the main checkout.
- Run the suite: `DXW_PARITY_FAST=1 node scripts/parity-parallel.mjs` (~6 min).

## Per-agent state — each in its own worktree, WIP-committed, NOT merged

| Agent | Worktree branch | WIP sha | Progress | Was about to |
|---|---|---|---|---|
| **watermark** | worktree-agent-ae897de18d9095ddf | 3e36a03 | 220 lines, 6 files — FURTHEST ALONG ("pixel values match reference nearly exactly") | Visually confirm all 6 pages, run gates |
| **matrices** | worktree-agent-af8761e484a260912 | 787f882 | 136 lines, math.ts+model+parse | Add new construct constants after the DLM ladder block |
| **boxes** | worktree-agent-aefe4ab6458369130 | 613b31c | 176 lines, 4 files | Implement cell spacing in paintRow + placeTable |
| **columns** | worktree-agent-a2793813df458b54d | bf2aae8 | 93 lines, engine+model+section | Band tracking + separator emission in engine.ts |
| **vertical** | worktree-agent-ab67cee75609eec71 | (none) | reading only, no src edits | Read run()/paragraph entry/layoutFrame then start |
| **scripts** | worktree-agent-a04d1b6e64ad18d19 | (none) | baselining only, no src edits | Analyze shape-autofit fixture |

## Target pages each agent owns (from `parity/out/results.json`, worst-first)
- **watermark**: probe2-picture-watermark p1-p6 (18.5–93.9; p6 spawns an EXTRA
  page — Word=6, engine=7). VML header image paints full-opacity/top-anchored
  in header flow instead of washed-out/page-centered/behind-text.
- **vertical**: probe2-ruby-vertical p2 (88.4 tbRl vertical page), p1 (34.9
  btLr cells); probe2-mixed-orientation p3 (72.4 section vAlign=center).
- **columns**: probe3-columns-unequal p1 (67.8 unequal w:cols), p2 (1.0);
  probe3-chargrid p1 (34.5 charsAndLines grid, spawns an EXTRA page).
- **matrices**: probe2-math-matrices p1 (49.4 — m:m matrices, eqArr cases,
  m:acc/groupChr, radical degree, limLow/limUpp).
- **boxes**: probe3-linked-textboxes p1 (32.5), probe3-table-exotics p2
  (30.1)/p1 (21.4 float-overlap/cellSpacing/diagonal), probe2-run-borders p1
  (30.8 w:bdr).
- **scripts**: probe3-emoji p1 (36.5 color emoji), probe3-shape-autofit p1
  (16.6), probe3-indic p1 (17.2), probe3-kashida p1 (14.2), probe2-arabic-rtl
  p1 (13.0), probe3-mirror-book p2 (9.6), probe2-content-controls p1 (9.4);
  tail: thai 3.8, dropcaps 2.9, tracked-changes 2.8, form-checkboxes 1.2.
- **DEFERRED** (Word is priority): probe3-lo-provenance p1 (57.1,
  LibreOffice-authored — provenance-tagged separately).

## How to resume an agent
The full task briefs are in each agent's transcript, but you can relaunch
fresh: `cd` into the worktree, `git log -1` to see the WIP, then continue.
The watermark and the three other WIP agents have real code to build on —
`git show <sha>` in each worktree shows exactly what was written. Merge to
main only after: own gates pass + full parallel suite confirms no regression
(the msa/ieee lesson: partial gates miss cross-fixture opposite-pins).

## Gotchas banked this session
- NEVER commit anything under `fixtures/` or `parity/` (both gitignored;
  a stray fixtures **symlink** merge once deleted all 83 fixtures — recovered).
- Inspect `git diff --stat main...branch` for mode-120000 (symlink) entries
  before merging any agent branch.
- Fixture backup tarball: `~/DocxInWeb-fixtures-backup-20260710.tar.gz`.
- Word-for-Mac PDF export can't show tracked-changes markup (scripting limit).
- Full memory + rule history: `.claude/.../memory/docxinweb-project.md`.
