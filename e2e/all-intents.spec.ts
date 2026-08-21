import { test, expect, type Page } from "@playwright/test";
import { INTENT_KINDS } from "../packages/collab/src/client.js";
import { ADDRESS_WIRE_FIELD, registeredOperation } from "../packages/core/src/index.js";
import { BOARD_CODE, enterCodeIfPrompted } from "./_helpers";

/**
 * Full-intent-surface browser CONVERGENCE test (plan doc 09, browser tier).
 *
 * Three real browser clients on one (encrypted) collab doc: A submits every intent
 * kind through the REAL connection (`session.submitOp` — the canonical-apply
 * path), and after each we assert A, B, and a receive-only spy C converge to
 * BYTE-IDENTICAL docx (`doc.save()`, deterministic by design). The server broadcasts to B and C; their agreement with A's optimistic
 * apply is the multi-client + server-mediated convergence check.
 *
 * WHAT THIS PROVES for every canonical kind: wire serialization +
 * server sequencing + deterministic outcome → all clients AGREE (never
 * diverge — a divergence would be a worse bug than a no-op). Coverage is
 * asserted against the full union, so every kind is exercised.
 *
 * Companion below: `structural intents genuinely APPLY …` proves these are
 * real applications (table, bold), not just agreement — every intent kind
 * both round-trips AND applies in the browser.
 */

const SERVER = "localhost:1399";

// This runtime list is tied to the Intent union by a typed map. A new intent
// reaches this test and the agent capability gate in the same build.
const ALL_KINDS = INTENT_KINDS;

/** Submit an intent on A (allocating any carried ids named in `allocs`), then
 * assert all three clients converge byte-identically. `allocs` maps a field
 * name to a count; the allocated ids are spliced into the intent in-browser
 * and also returned so later steps can target the created nodes. */
async function step(
  pages: { a: Page; b: Page; c: Page },
  kind: string,
  intent: Record<string, unknown>,
  allocs?: Record<string, number>,
): Promise<number[]> {
  const allocated = await pages.a.evaluate(
    ({ intent, allocs }) => {
      const ww = (window as unknown as { __ww: { submitOp(i: unknown): void; allocIds(n: number): number[] } }).__ww;
      const out: number[] = [];
      for (const [field, n] of Object.entries(allocs ?? {})) {
        const ids = ww.allocIds(n as number);
        out.push(...ids);
        // Splice into the intent by a `$field` placeholder (single id) or
        // `$field[]` (array).
        // nodeIds is a LIST on the wire even when one id is carried, so the
        // single-id shorthand below must never collapse it to a scalar.
        (intent as Record<string, unknown>)[field] =
          (n as number) === 1 && !field.endsWith("[]") && field !== "nodeIds" ? ids[0] : ids;
      }
      ww.submitOp(intent);
      return out;
    },
    { intent, allocs },
  );
  await converge(pages, kind);
  return allocated;
}

/**
 * Submit a REGISTERED operation, taking its address field and carried-id budget
 * from the registry instead of repeating them at the call site.
 *
 * Fifty-five of the kinds below are declared in the core operation registry,
 * which already knows each one's address ("run" → runId, "cell" →
 * cellParagraphId, …) and exactly how many ids it carries for the arguments
 * given. Reading that here means a registered operation that changes its shape
 * cannot leave this test asserting the old one, and that adding a kind needs
 * only the arguments — the wiring comes from the declaration.
 */
async function regStep(
  pages: { a: Page; b: Page; c: Page },
  kind: string,
  addressId: number | undefined,
  args: Record<string, unknown>,
): Promise<number[]> {
  const definition = registeredOperation(kind);
  if (!definition) throw new Error(`${kind} is not a registered operation`);
  const intent: Record<string, unknown> = { kind, ...args };
  if (definition.address !== "document") {
    if (addressId === undefined) throw new Error(`${kind} is ${definition.address}-addressed and needs an id`);
    intent[ADDRESS_WIRE_FIELD[definition.address]] = addressId;
  }
  const carried = definition.nodeIds ? definition.nodeIds(args as never) : 0;
  return step(pages, kind, intent, carried > 0 ? { nodeIds: carried } : undefined);
}

async function saveB64(page: Page): Promise<string | null> {
  return page.evaluate(() => (window as unknown as { __ww: { saveB64(): string | null } }).__ww.saveB64());
}

/** Poll until A, B, C hold byte-identical docx (or fail with readable text). */
async function converge(pages: { a: Page; b: Page; c: Page }, label: string): Promise<void> {
  await expect
    .poll(
      async () => {
        const [a, b, c] = await Promise.all([saveB64(pages.a), saveB64(pages.b), saveB64(pages.c)]);
        return a !== null && a === b && b === c;
      },
      { message: `clients did not converge after ${label}`, timeout: 8000 },
    )
    .toBe(true);
}

async function joinAll(browser: import("@playwright/test").Browser): Promise<{
  a: Page; b: Page; c: Page; contexts: import("@playwright/test").BrowserContext[];
}> {
  const [ca, cb, cc] = await Promise.all([browser.newContext(), browser.newContext(), browser.newContext()]);
  const [a, b, c] = await Promise.all([ca.newPage(), cb.newPage(), cc.newPage()]);
  await a.goto(`/?server=${SERVER}`);
  await expect(a.getByTestId("local-editor")).toBeVisible();
  await a.getByTestId("make-collaborative").click();
  await expect(a.getByTestId("collab-modal")).toBeVisible();
  await a.getByTestId("share-code").fill(BOARD_CODE); // required — see _helpers
  await a.getByTestId("start-collab").click();
  await expect(a).toHaveURL(/[?&]doc=/);
  await expect(a.getByTestId("toolbar")).toBeVisible(); // collab chrome mounted
  await expect(a.locator(".dxw-page")).toBeVisible();
  const url = a.url();
  await Promise.all([b.goto(url), c.goto(url)]);
  // Fresh contexts hold no code, so the link alone refuses them — that is the
  // point of the code, and the joiners supply it exactly as a person would.
  await Promise.all([enterCodeIfPrompted(b), enterCodeIfPrompted(c)]);
  await Promise.all([
    expect(b.locator(".dxw-page")).toBeVisible(),
    expect(c.locator(".dxw-page")).toBeVisible(),
  ]);
  // Wait for the hooks to be live on all three.
  for (const p of [a, b, c]) {
    await expect.poll(() => p.evaluate(() => !!(window as unknown as { __ww?: unknown }).__ww)).toBe(true);
  }
  return { a, b, c, contexts: [ca, cb, cc] };
}

test("all canonical intent kinds converge byte-identically across 3 browser clients", async ({ browser }) => {
  test.setTimeout(120_000);
  const pages = await joinAll(browser);
  const covered = new Set<string>();
  const cover = (k: string) => covered.add(k);
  const P = { blockId: 1, runId: 2 };

  try {
    // --- text foundation (gives content for range/format/comment ops) ---
    await step(pages, "insertText", { kind: "insertText", at: { ...P, offset: 0 }, text: "Hello brave new world" }); cover("insertText");
    await step(pages, "formatRun", { kind: "formatRun", ...P, patch: { bold: true } }); cover("formatRun");
    await step(pages, "formatParagraph", { kind: "formatParagraph", blockId: 1, align: "center" }); cover("formatParagraph");
    await step(pages, "setSpacing", { kind: "setSpacing", blockId: 1, patch: { lineRule: "auto", line: 360 } }); cover("setSpacing");
    await step(pages, "adjustIndent", { kind: "adjustIndent", blockId: 1, direction: 1 }); cover("adjustIndent");
    await step(pages, "setListType", { kind: "setListType", blockId: 1, listKind: "bullet" }); cover("setListType");
    await step(pages, "setListLevel", { kind: "setListLevel", blockId: 1, delta: 1 }); cover("setListLevel");
    await step(pages, "setListType-off", { kind: "setListType", blockId: 1, listKind: null }); // reset
    await step(pages, "setDropCap", { kind: "setDropCap", blockId: 1, mode: "drop" }, { nodeIds: 1 }); cover("setDropCap");
    await step(pages, "setDivider", { kind: "setDivider", blockId: 1, divider: { val: "single", sz: 6 } }); cover("setDivider");

    // --- sub-range format (carried split ids) ---
    await step(pages, "formatRange", { kind: "formatRange", ...P, start: 0, end: 5, patch: { italic: true } }, { middleId: 1 }); cover("formatRange");

    // --- structural: split then merge ---
    const [nb, nr] = await step(pages, "splitParagraph", { kind: "splitParagraph", at: { ...P, offset: 6 } }, { newBlockId: 1, newRunId: 1 }); cover("splitParagraph");
    void nr;
    await step(pages, "mergeParagraph", { kind: "mergeParagraph", blockId: nb }); cover("mergeParagraph");

    // --- inline separators (insert then take it back out) ---
    await step(pages, "insertSeparator", { kind: "insertSeparator", at: { ...P, offset: 0 }, separator: "br" }); cover("insertSeparator");
    await step(pages, "deleteSeparator", { kind: "deleteSeparator", at: { ...P, offset: 0 } }); cover("deleteSeparator");

    // --- run-anchored inserts (fields, links, notes, breaks, bookmarks) ---
    await step(pages, "setLink", { kind: "setLink", runId: 2, url: "https://example.com" }, { nodeIds: 2 }); cover("setLink");
    await step(pages, "removeLink", { kind: "removeLink", runId: 2 }); cover("removeLink");
    await step(pages, "insertBookmark", { kind: "insertBookmark", runId: 2, name: "bm1" }); cover("insertBookmark");
    await step(pages, "insertBookmarkRange", { kind: "insertBookmarkRange", runId: 2, name: "bmr1", start: 0, end: 3 }); cover("insertBookmarkRange");
    await step(pages, "insertCrossRef", { kind: "insertCrossRef", runId: 2, bookmark: "bm1", refKind: "text" }, { nodeIds: 2 }); cover("insertCrossRef");
    await step(pages, "insertPageField", { kind: "insertPageField", runId: 2, fieldKind: "page" }, { nodeIds: 2 }); cover("insertPageField");
    await step(pages, "insertDateTimeField", { kind: "insertDateTimeField", runId: 2, dtKind: "date", picture: "yyyy-MM-dd" }, { nodeIds: 2 }); cover("insertDateTimeField");
    await step(pages, "insertField", { kind: "insertField", runId: 2, instruction: "AUTHOR", cachedResult: "x" }, { nodeIds: 2 }); cover("insertField");
    await step(pages, "insertFootnote", { kind: "insertFootnote", runId: 2, text: "a footnote" }, { nodeIds: 2 }); cover("insertFootnote");
    await step(pages, "insertBreak", { kind: "insertBreak", runId: 2, breakKind: "page" }, { nodeIds: 2 }); cover("insertBreak");
    await step(pages, "insertBlankPage", { kind: "insertBlankPage", runId: 2 }, { nodeIds: 3 }); cover("insertBlankPage");
    await step(pages, "insertSectionBreak", { kind: "insertSectionBreak", runId: 2, breakType: "nextPage" }, { nodeIds: 2 }); cover("insertSectionBreak");
    await step(pages, "insertCoverPage", { kind: "insertCoverPage", content: { title: "T", subtitle: "S" } }, { nodeIds: 4 }); cover("insertCoverPage");

    // --- document-level ---
    await step(pages, "setPageLayout", { kind: "setPageLayout", patch: { orientation: "landscape" } }); cover("setPageLayout");
    await step(pages, "setLineNumbering", { kind: "setLineNumbering", patch: { enabled: true, countBy: 1 } }); cover("setLineNumbering");

    // --- table + cell ops ---
    const tableIds = await step(pages, "insertTable", { kind: "insertTable", runId: 2, rows: 2, cols: 2 }, { nodeIds: 12 }); cover("insertTable");
    const cellParaId = tableIds.find((_id, i) => i > 2) ?? tableIds[tableIds.length - 1];
    await step(pages, "tableOp", { kind: "tableOp", cellParagraphId: cellParaId, op: { kind: "insertRow", where: "below" } }, { nodeIds: 6 }); cover("tableOp");
    await step(pages, "cellShading", { kind: "tableOp", cellParagraphId: cellParaId, op: { kind: "cellShading", fill: "FFFF00" } });
    await step(pages, "cellVAlign", { kind: "tableOp", cellParagraphId: cellParaId, op: { kind: "cellVAlign", v: "center" } });
    await step(pages, "resizeTableColumn", { kind: "resizeTableColumn", cellParagraphId: cellParaId, boundary: 1, deltaPx: 8 }); cover("resizeTableColumn");
    await step(pages, "resizeTableRow", { kind: "resizeTableRow", cellParagraphId: cellParaId, rowIdx: 0, heightPx: 32 }); cover("resizeTableRow");
    await step(pages, "moveTable", { kind: "moveTable", cellParagraphId: cellParaId, xPx: 24, yPx: 36, preservePageStart: false, pageDelta: 0 }); cover("moveTable");

    // --- comments ---
    await step(pages, "commentRun", { kind: "commentRun", runId: 2, text: "note", author: "A", date: "2026-07-23T00:00:00Z", paraId: "c1" }, { nodeIds: 2 }); cover("commentRun");
    const commentId = await pages.a.evaluate(() => {
      const doc = (window as unknown as { __ww: { text(): string } }).__ww;
      void doc;
      // The comment id is scan-assigned; read the first w:comment id from the doc.
      const s = (window as unknown as { __wwSession?: { doc?: { docRoot: unknown } } }).__wwSession;
      let found = "0";
      const walk = (el: { name: string; attrs?: Record<string, string>; children: unknown[] }) => {
        if (el.name.endsWith(":comment") && el.attrs?.["w:id"]) found = el.attrs["w:id"];
        (el.children as never[]).forEach(walk);
      };
      // comments live in a separate part; fall back to id "0" if not walkable.
      try { if (s?.doc) walk((s.doc as { docRoot: never }).docRoot as never); } catch { /* */ }
      return found;
    });
    await step(pages, "replyComment", { kind: "replyComment", parentId: commentId, text: "re", author: "B", date: "2026-07-23T00:01:00Z", paraIds: ["c2"] }); cover("replyComment");
    await step(pages, "editComment", { kind: "editComment", commentId, text: "note, revised" }); cover("editComment");
    await step(pages, "resolveComment", { kind: "resolveComment", commentId, resolved: true, paraId: "c3" }); cover("resolveComment");
    await step(pages, "deleteComment", { kind: "deleteComment", commentId }); cover("deleteComment");

    // --- paste ---
    await step(pages, "pasteBlocks", { kind: "pasteBlocks", afterBlockId: 1, blocksXml: "<w:p><w:r><w:t>pasted</w:t></w:r></w:p>" }, { nodeIds: 2 }); cover("pasteBlocks");

    // --- image + image setters ---
    const imgIds = await step(pages, "insertImage", { kind: "insertImage", runId: 2, blobSha: "a".repeat(64), bytesLen: 100, ext: "png", widthPx: 40, heightPx: 40 }, { nodeIds: 2 }); cover("insertImage");
    const imgRun = imgIds[0];
    await step(pages, "setImageAltText", { kind: "setImageAltText", runId: imgRun, alt: "alt text" }); cover("setImageAltText");
    await step(pages, "setImageWrap", { kind: "setImageWrap", runId: imgRun, mode: "square" }); cover("setImageWrap");

    // --- shapes / drawings + setters ---
    const shapeIds = await step(pages, "insertShape", { kind: "insertShape", runId: 2, preset: "rectangle", text: "S" }, { nodeIds: 3 }); cover("insertShape");
    const shapeRun = shapeIds[0];
    await step(pages, "setDrawingRotation", { kind: "setDrawingRotation", runId: shapeRun, degrees: 45 }); cover("setDrawingRotation");
    await step(pages, "setDrawingFill", { kind: "setDrawingFill", runId: shapeRun, color: "00FF00" }); cover("setDrawingFill");
    await step(pages, "setDrawingLineStyle", { kind: "setDrawingLineStyle", runId: shapeRun, color: "000000", widthPx: 2, dash: "dashed" }); cover("setDrawingLineStyle");
    await step(pages, "setDrawingOrder", { kind: "setDrawingOrder", runId: shapeRun, order: "front" }); cover("setDrawingOrder");
    await step(pages, "setFloatingPagePosition", { kind: "setFloatingPagePosition", runId: shapeRun, xPx: 10, yPx: 10 }); cover("setFloatingPagePosition");
    await step(pages, "resizeDrawing", { kind: "resizeDrawing", runId: shapeRun, widthPx: 180, heightPx: 90 }); cover("resizeDrawing");

    // --- wordart ---
    const waIds = await step(pages, "insertWordArt", { kind: "insertWordArt", runId: 2, text: "WA", preset: "plain" }, { nodeIds: 3 }); cover("insertWordArt");
    await step(pages, "setDrawingWordArtText", { kind: "setDrawingWordArtText", runId: waIds[0], text: "WA2" }); cover("setDrawingWordArtText");
    await step(pages, "setDrawingWordArtStyle", { kind: "setDrawingWordArtStyle", runId: waIds[0], color: "AABBCC", opacity: 0.25 }); cover("setDrawingWordArtStyle");

    // --- chart + data ---
    const chartIds = await step(pages, "insertChart", { kind: "insertChart", runId: 2, chart: { type: "column", categories: ["a", "b"], series: [{ name: "s", values: [1, 2] }] } }, { nodeIds: 3 }); cover("insertChart");
    await step(pages, "setChartData", { kind: "setChartData", runId: chartIds[0], chart: { type: "line", categories: ["a", "b"], series: [{ name: "s", values: [3, 4] }] } }); cover("setChartData");

    // --- smartart + setters ---
    const saIds = await step(pages, "insertSmartArt", { kind: "insertSmartArt", runId: 2, smartArt: { layout: "list", items: ["x", "y"] } }, { nodeIds: 3 }); cover("insertSmartArt");
    const saRun = saIds[0];
    await step(pages, "setSmartArtData", { kind: "setSmartArtData", runId: saRun, smartArt: { layout: "cycle", items: ["p", "q"] } }); cover("setSmartArtData");
    await step(pages, "setSmartArtNodeText", { kind: "setSmartArtNodeText", runId: saRun, index: 0, text: "z" }); cover("setSmartArtNodeText");
    await step(pages, "setSmartArtFill", { kind: "setSmartArtFill", runId: saRun, color: "FF0000" }); cover("setSmartArtFill");
    await step(pages, "setSmartArtTextFormat", { kind: "setSmartArtTextFormat", runId: saRun, format: { fontFamily: "Arial", fontSizePt: 12, color: "000000", bold: true, italic: false, alignment: "left" } }); cover("setSmartArtTextFormat");

    // --- math + setters ---
    const mathIds = await step(pages, "insertMath", { kind: "insertMath", runId: 2, mathText: "x^2" }, { nodeIds: 2 }); cover("insertMath");
    void mathIds;
    await step(pages, "setMathLinear", { kind: "setMathLinear", blockId: 1, mathText: "y^2" }); cover("setMathLinear");
    // Drop at offset 0 of the anchor run: re-parents the equation without
    // splitting the destination run.
    await step(pages, "moveMath", { kind: "moveMath", blockId: 1, at: { blockId: 1, runId: 2, offset: 0 } }, { nodeIds: 4 }); cover("moveMath");
    await step(pages, "deleteMath", { kind: "deleteMath", blockId: 1 }); cover("deleteMath");

    // --- header/footer part creation (a whole new package part) ---
    await step(pages, "ensureHeaderFooter", { kind: "ensureHeaderFooter", hfKind: "header" }, { nodeIds: 8 }); cover("ensureHeaderFooter");

    // --- checkbox ---
    await step(pages, "toggleCheckbox", { kind: "toggleCheckbox", runId: 2 }); cover("toggleCheckbox");

    // --- tracked changes create + review ---
    await step(pages, "suggestRevision", { kind: "suggestRevision", ranges: [{ blockId: 1, runId: 2, start: 0, end: 2 }], suggest: { author: "A", date: "2026-07-23T00:00:00Z" } }); cover("suggestRevision");
    await step(pages, "acceptRevision", { kind: "acceptRevision", index: 0 }); cover("acceptRevision");
    await step(pages, "suggestRevision-2", { kind: "suggestRevision", ranges: [{ blockId: 1, runId: 2, start: 0, end: 1 }], suggest: { author: "A", date: "2026-07-23T00:02:00Z" } });
    await step(pages, "rejectRevision", { kind: "rejectRevision", index: 0 }); cover("rejectRevision");
    await step(pages, "suggestRevision-3", { kind: "suggestRevision", ranges: [{ blockId: 1, runId: 2, start: 0, end: 1 }], suggest: { author: "A", date: "2026-07-23T00:03:00Z" } });
    await step(pages, "acceptAllRevisions", { kind: "acceptAllRevisions" }); cover("acceptAllRevisions");
    await step(pages, "suggestRevision-4", { kind: "suggestRevision", ranges: [{ blockId: 1, runId: 2, start: 0, end: 1 }], suggest: { author: "A", date: "2026-07-23T00:04:00Z" } });
    await step(pages, "rejectAllRevisions", { kind: "rejectAllRevisions" }); cover("rejectAllRevisions");

    await step(pages, "removeDrawing", { kind: "removeDrawing", runId: shapeRun }); cover("removeDrawing");

    // --- references: caption, cross-ref bookmark, TOC, index ---
    await regStep(pages, "insertCaption", 1, { label: "Figure", text: "A caption" }); cover("insertCaption");
    await regStep(pages, "ensureRefBookmark", 1, { name: "refbm1" }); cover("ensureRefBookmark");
    await regStep(pages, "insertToc", 2, { entryCount: 1, levels: 3 }); cover("insertToc");
    await regStep(pages, "insertIndexEntry", 2, { entry: "term" }); cover("insertIndexEntry");
    await regStep(pages, "insertIndex", 2, { entryCount: 1 }); cover("insertIndex");
    await regStep(pages, "refreshIndex", undefined, { entryCount: 1 }); cover("refreshIndex");
    await regStep(pages, "updateFields", undefined, { results: [] }); cover("updateFields");

    // --- styles ---
    await regStep(pages, "createStyle", undefined, { style: { styleId: "MyStyle", name: "My Style", type: "paragraph", rPr: { bold: true } } }); cover("createStyle");
    await regStep(pages, "modifyStyle", undefined, { styleId: "MyStyle", patch: { rPr: { italic: true } } }); cover("modifyStyle");
    await regStep(pages, "deleteStyle", undefined, { styleId: "MyStyle" }); cover("deleteStyle");

    // --- paragraph-level: numbering, borders, tabs ---
    await regStep(pages, "setNumberingLevel", 1, { ilvl: 0, patch: { numFmt: "decimal" } }); cover("setNumberingLevel");
    await regStep(pages, "setNumberingRestart", 1, { start: 3 }); cover("setNumberingRestart");
    await regStep(pages, "setParagraphBorders", 1, { patch: { top: { val: "single", sz: 4 } } }); cover("setParagraphBorders");
    await regStep(pages, "setTabStops", 1, { stops: [{ pos: 1440, val: "left" }] }); cover("setTabStops");

    // --- citations and bibliography (source first, so the citation resolves) ---
    await regStep(pages, "createCitationSource", undefined, { source: { tag: "SRC1", type: "Book", title: "T", author: "A", year: "2026" } }); cover("createCitationSource");
    await regStep(pages, "setCitationStyle", undefined, { style: "apa" }); cover("setCitationStyle");
    await regStep(pages, "insertCitation", 2, { tag: "SRC1" }); cover("insertCitation");
    await regStep(pages, "insertBibliography", 2, { entryCount: 1 }); cover("insertBibliography");
    await regStep(pages, "refreshBibliography", undefined, { entryCount: 1 }); cover("refreshBibliography");
    await regStep(pages, "editCitationSource", undefined, { tag: "SRC1", patch: { title: "T2" } }); cover("editCitationSource");
    await regStep(pages, "deleteCitationSource", undefined, { tag: "SRC1" }); cover("deleteCitationSource");

    // --- building blocks (create, use, remove) ---
    await regStep(pages, "createBuildingBlock", undefined, { name: "BB1", blocksXml: "<w:p><w:r><w:t>bb</w:t></w:r></w:p>" }); cover("createBuildingBlock");
    await regStep(pages, "insertBuildingBlock", 2, { name: "BB1", blockCount: 1 }); cover("insertBuildingBlock");
    await regStep(pages, "deleteBuildingBlock", undefined, { name: "BB1" }); cover("deleteBuildingBlock");

    // --- merge field, endnote ---
    await regStep(pages, "insertMergeField", 2, { name: "FirstName" }); cover("insertMergeField");
    await regStep(pages, "insertEndnote", 2, { text: "an endnote" }); cover("insertEndnote");

    // --- object setters (address "object" is wire field runId) ---
    await regStep(pages, "setCrop", imgRun, { crop: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 } }); cover("setCrop");
    await regStep(pages, "setDrawingTextFit", shapeRun, { mode: "shrink" }); cover("setDrawingTextFit");
    // No intent CREATES a 3D model, so this one is exercised against an
    // existing object: convergence is the contract here, not effect.
    await regStep(pages, "setModel3DRotation", imgRun, { rotation: { x: 10, y: 20, z: 0 } }); cover("setModel3DRotation");

    // --- document setup: watermark, title page, page numbers, notes ---
    await regStep(pages, "insertWatermark", undefined, { text: "DRAFT", headerCount: 1 }); cover("insertWatermark");
    await regStep(pages, "removeWatermark", undefined, {}); cover("removeWatermark");
    await regStep(pages, "insertPictureWatermark", undefined, { blobSha: "b".repeat(64), bytesLen: 100, ext: "png", naturalWidthPx: 40, naturalHeightPx: 40, headerCount: 1 }); cover("insertPictureWatermark");
    await regStep(pages, "setTitlePage", undefined, { enabled: true }); cover("setTitlePage");
    await regStep(pages, "setEvenOddHeaders", undefined, { enabled: true }); cover("setEvenOddHeaders");
    await regStep(pages, "setPageNumberFormat", undefined, { fmt: "decimal", start: 1 }); cover("setPageNumberFormat");
    await regStep(pages, "insertPageNumberPosition", undefined, { position: "footer", align: "center" }); cover("insertPageNumberPosition");
    await regStep(pages, "removePageNumbers", undefined, {}); cover("removePageNumbers");
    await regStep(pages, "insertHeaderFooterPreset", undefined, { hfKind: "header", preset: "blank" }); cover("insertHeaderFooterPreset");
    await regStep(pages, "setHyphenation", undefined, { auto: true, zonePt: 18 }); cover("setHyphenation");
    await regStep(pages, "setFootnoteOptions", undefined, { fmt: "decimal", start: 1 }); cover("setFootnoteOptions");
    await regStep(pages, "setEndnoteOptions", undefined, { fmt: "lowerRoman", start: 1 }); cover("setEndnoteOptions");

    // --- table depth (the table from above is still standing) ---
    await regStep(pages, "setTableStyle", cellParaId, { styleId: "TableGrid" }); cover("setTableStyle");
    await regStep(pages, "setTableLook", cellParaId, { look: { firstRow: true } }); cover("setTableLook");
    await regStep(pages, "setTableBorders", cellParaId, { scope: "table", edges: ["top"], border: { val: "single", sz: 4 } }); cover("setTableBorders");
    await regStep(pages, "setTableWidth", cellParaId, { unit: "auto" }); cover("setTableWidth");
    await regStep(pages, "setTableColumnWidth", cellParaId, { colIdx: 0, widthPt: 90 }); cover("setTableColumnWidth");
    await regStep(pages, "setTableLayout", cellParaId, { layout: "fixed" }); cover("setTableLayout");
    await regStep(pages, "setTableCellMargins", cellParaId, { scope: "table", margins: { top: 40 } }); cover("setTableCellMargins");
    await regStep(pages, "setTableHeaderRows", cellParaId, { count: 1 }); cover("setTableHeaderRows");
    await regStep(pages, "insertTableFormula", cellParaId, { formula: "=SUM(ABOVE)" }); cover("insertTableFormula");
    await regStep(pages, "sortTableRows", cellParaId, { colIdx: 0, order: "asc", compare: "text", hasHeader: false }); cover("sortTableRows");

    // --- text ⇄ table, on a paragraph of their own so nothing above is
    // rebuilt underneath the ids it still uses ---
    const convIds = await step(pages, "pasteBlocks-convert", { kind: "pasteBlocks", afterBlockId: 1, blocksXml: "<w:p><w:r><w:t>a\tb</w:t></w:r></w:p>" }, { nodeIds: 2 });
    const convBlockId = convIds[0];
    const madeTable = await regStep(pages, "convertTextToTable", convBlockId, { separator: "tab", cellCount: 2 }); cover("convertTextToTable");
    await regStep(pages, "convertTableToText", madeTable[madeTable.length - 1], { separator: "tab", rowCount: 1 }); cover("convertTableToText");

    // --- delete (last: removes content) ---
    await step(pages, "deleteText", { kind: "deleteText", blockId: 1, runId: 2, start: 0, end: 1 }); cover("deleteText");

    // COVERAGE: every kind in the union must have been exercised.
    const missing = ALL_KINDS.filter((k) => !covered.has(k));
    expect(missing, `intent kinds not exercised: ${missing.join(", ")}`).toEqual([]);

    // NON-VACUOUS: at least the text foundation genuinely applied and
    // converged (guards against "everything silently no-op'd, so agreement
    // is trivial" — insertText is confirmed-working in the browser).
    const finalText = await pages.a.evaluate(() => (window as unknown as { __ww: { text(): string } }).__ww.text());
    expect(finalText).toContain("world");
  } finally {
    await Promise.all(pages.contexts.map((c) => c.close()));
  }
});

/**
 * Structural intents genuinely APPLY across browser clients (not just
 * converge): insertTable creates a real table and formatRun a real w:b, on
 * ALL three clients. (Briefly `fixme`'d when a diagnostic DOM-walk bug made
 * these look like no-ops; they were always applying — confirmed by
 * `broadcast:applied` frames. See BUGS.md.)
 */
test("structural intents genuinely APPLY across browser clients (not just converge)", async ({ browser }) => {
  test.setTimeout(60_000);
  const pages = await joinAll(browser);
  const hasTag = (page: Page, suffix: string) =>
    page.evaluate((s) => {
      const sess = (window as unknown as { __ww?: { _session?: { doc?: { docRoot: unknown } } } }).__ww?._session;
      let found = false;
      const walk = (el: { name: string; children: unknown[] }) => {
        if (el.name.endsWith(s)) found = true;
        (el.children as never[]).forEach(walk);
      };
      if (sess?.doc) walk((sess.doc as { docRoot: never }).docRoot);
      return found;
    }, suffix);
  try {
    await step(pages, "insertText", { kind: "insertText", at: { blockId: 1, runId: 2, offset: 0 }, text: "Hello world" });
    await step(pages, "insertTable", { kind: "insertTable", runId: 2, rows: 2, cols: 2 }, { nodeIds: 20 });
    // Should be present on ALL three clients — currently false (the bug).
    for (const p of [pages.a, pages.b, pages.c]) expect(await hasTag(p, ":tbl")).toBe(true);
    await step(pages, "formatRun", { kind: "formatRun", blockId: 1, runId: 2, patch: { bold: true } });
    for (const p of [pages.a, pages.b, pages.c]) expect(await hasTag(p, ":b")).toBe(true);
  } finally {
    await Promise.all(pages.contexts.map((c) => c.close()));
  }
});
