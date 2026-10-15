# Print benchmark

Prints .docx files to PDF in headless Chromium two ways and compares them:
the whole-document `renderForPrint`, and the streaming `openPrintStream`
(batches of 25 pages by default). The print bundle is built from this
checkout's `packages/core` each run, as `vendor/wordinweb/rebuild.sh` builds
DiscoveryTool's `print.js`.

Equivalence: page counts, and each rendered page's DOM hashed and compared.
A mismatched page's two DOMs go to `$TMPDIR/print-bench-diffs`.

    node scripts/print-bench/equiv.mjs [--jobs 4] [--batch 25] [--out results.json] <docx or folder>...

Memory: peak renderer RSS and JS heap while laying out and printing every
batch to PDF, each fixture and mode in its own Chromium, with the peak RSS of
the browser process (where the page's Blobs live) and of the Node process.
The new mode serves the file from disk a range at a time, as DiscoveryTool
does, so the Node figure is what serving the file costs.

    node scripts/print-bench/memory-table.mjs [--modes old,new] [--out table.json] <docx>...
    node scripts/print-bench/memory.mjs --mode old|new [--batch 25] <docx>

Longer fixtures, the body repeated k times:

    node scripts/print-bench/repeat.mjs <in.docx> <k> <out.docx>

Times are wall clock; run on a quiet machine. Fixtures used so far: the 166
in `wordinweb-parity/apps/demo/public/fixtures`, the pairs in
`bench/memory/pairs`, and the NIH contract and Yiddish novel repeated.
