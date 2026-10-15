# Print bench

`lib.mjs` drives the web app's print path (`renderForPrint`, then the page
printed to PDF in slices) in headless Chromium through Playwright, with the
renderer's memory sampled: `launch`, `runOld` and `sampled`. The print bundle
is built from this checkout's `packages/core` each run. A browser-free PDF
comes from `docx-pdf` (packages/core, `pdf/render.ts`), which this bench's
page counts and page renders are compared against.

Longer fixtures, the body repeated k times:

    node scripts/print-bench/repeat.mjs <in.docx> <k> <out.docx>

Fixtures used so far: the 166 in `wordinweb-parity/apps/demo/public/fixtures`,
and the NIH contract and Yiddish novel repeated.
