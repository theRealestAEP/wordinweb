import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer, type TextMeasurer } from "../src/layout/measure.js";
import { makeDocx, wrapDocument, p } from "./helpers.js";
import type { FontSpec } from "../src/layout/types.js";

/**
 * EMERGENCY CHARACTER WRAP — the maximal-fit property (perf B10).
 *
 * When a token has no break opportunity (no spaces), the breaker fills each
 * line with the LARGEST prefix that still fits and carries the rest to the
 * next line. That boundary used to be found by walking the candidate length
 * down one character at a time, measuring a prefix each step — O(token²)
 * measurement, which a collab flood into a single space-free run turned into
 * seconds of main-thread time per repaint. It is now found by bisection.
 *
 * Bisection only agrees with the walk while measurement is monotone in prefix
 * length, which is an assumption about fonts rather than about code — so the
 * breaker verifies its own answer against the boundary property and falls
 * back to the exact walk when verification fails. These tests pin that
 * PROPERTY rather than the algorithm: every line fits, and no line could have
 * taken one more character. An off-by-one in the search, or a measurer that
 * violates monotonicity without the net catching it, breaks one of the two
 * halves — a test written against the old implementation would not.
 */

const measurer = new ApproxMeasurer();

interface WrappedLine {
  text: string;
  font: FontSpec;
}

/** Lay out one paragraph holding a single unbreakable token, and read back
 * its visual lines (grouped by baseline, ordered left to right). */
function wrapToken(token: string, pageWidthTwips: number): { lines: WrappedLine[]; contentWidth: number } {
  const sectPr =
    `<w:sectPr><w:pgSz w:w="${pageWidthTwips}" w:h="15840"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>`;
  const doc = DocxDocument.load(
    makeDocx({ "word/document.xml": wrapDocument(`${p(token)}${sectPr}`) }),
  );
  const result = layoutDocument(doc, { measurer });
  const byLine = new Map<number, { x: number; text: string; font: FontSpec }[]>();
  for (const page of result.pages) {
    for (const item of page.items) {
      if (item.kind !== "text" || item.text.length === 0) continue;
      const key = Math.round(item.baseline * 10);
      const row = byLine.get(key) ?? [];
      row.push({ x: item.x, text: item.text, font: item.font });
      byLine.set(key, row);
    }
  }
  const lines = [...byLine.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([, row]) => {
      const sorted = row.sort((a, b) => a.x - b.x);
      return { text: sorted.map((i) => i.text).join(""), font: sorted[0].font };
    });
  const props = doc.sections[0].props;
  return { lines, contentWidth: props.pageWidth - props.marginLeft - props.marginRight };
}

describe("hard character wrap fills each line maximally", () => {
  for (const [label, token] of [
    ["ascii", "abcdefghij".repeat(40)],
    ["digits", "0123456789".repeat(30)],
    ["mixed-case", "aB3xY9zQ7w".repeat(35)],
    ["punctuation-free flood shape", "a0b1c2d3e4f5g6h7i8j9".repeat(20)],
  ] as const) {
    it(`${label}: every line fits and none could take another character`, () => {
      const { lines, contentWidth } = wrapToken(token, 12240);
      expect(lines.length, "the token must actually wrap").toBeGreaterThan(1);
      expect(lines.map((l) => l.text).join(""), "no character may be lost or duplicated").toBe(token);

      lines.forEach((line, i) => {
        const w = measurer.width(line.text, line.font);
        expect(w, `line ${i} must fit the content width`).toBeLessThanOrEqual(contentWidth + 0.01);
        // The maximal-fit half: a line that is not the last must be unable to
        // take the first character of the following line. This is the half a
        // bisection off-by-one breaks — it would leave a line one char short.
        if (i < lines.length - 1 && lines[i + 1].text.length > 0) {
          const wider = measurer.width(line.text + lines[i + 1].text[0], line.font);
          expect(wider, `line ${i} should have taken one more character`).toBeGreaterThan(contentWidth);
        }
      });
    });
  }

  it("fills maximally across a range of page widths", () => {
    const token = "xyzw".repeat(120);
    for (const pageWidth of [7000, 8640, 10080, 12240, 15840]) {
      const { lines, contentWidth } = wrapToken(token, pageWidth);
      expect(lines.map((l) => l.text).join(""), `width ${pageWidth}: characters preserved`).toBe(token);
      lines.forEach((line, i) => {
        expect(
          measurer.width(line.text, line.font),
          `width ${pageWidth} line ${i} fits`,
        ).toBeLessThanOrEqual(contentWidth + 0.01);
        if (i < lines.length - 1 && lines[i + 1].text.length > 0) {
          expect(
            measurer.width(line.text + lines[i + 1].text[0], line.font),
            `width ${pageWidth} line ${i} is maximal`,
          ).toBeGreaterThan(contentWidth);
        }
      });
    }
  });

  /**
   * The COST property, pinned independently of the flood harness. The old
   * walk measured one prefix per character it stepped over, so fitting a
   * 2000-character token cost tens of thousands of measurements; bisection
   * costs ~log2(token) per line. A generous ceiling catches a regression to
   * linear scanning without being sensitive to line-count arithmetic.
   */
  it("a long unbreakable token is laid out with a bounded number of measurements", () => {
    const inner = new ApproxMeasurer();
    let calls = 0;
    const counting: TextMeasurer = {
      width: (text, font, letterSpacing) => {
        calls++;
        return inner.width(text, font, letterSpacing);
      },
      metrics: (font) => inner.metrics(font),
      paintBox: (font) => inner.paintBox?.(font),
      inkBox: (text, font) => inner.inkBox?.(text, font),
    };
    const token = "qwertyuiop".repeat(200); // 2000 chars, no break opportunity
    const doc = DocxDocument.load(
      makeDocx({ "word/document.xml": wrapDocument(p(token)) }),
    );
    layoutDocument(doc, { measurer: counting });
    // ~28 lines x ~11 bisection probes + per-line bookkeeping. The pre-fix
    // walk needed >25,000 measurements for this token.
    expect(calls, `measurement calls for a 2000-char token (was >25,000)`).toBeLessThan(3000);
  });

  /**
   * DIFFERENTIAL: the bisection and the walk it replaced must choose the same
   * boundary. Both are reproduced here against the engine's own measurer and
   * run over a randomized corpus of token shapes, fonts and capacities. The
   * Word-parity fixture bed is the usual guard for a line-breaking change, but
   * its fixtures and reference PDFs are gitignored and absent from this
   * checkout, so this stands in for it on the one decision the change touches.
   */
  it("bisection chooses the same boundary as the walk it replaced", () => {
    const inner = new ApproxMeasurer();
    let seed = 20260725;
    const rnd = (n: number): number => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % n;
    };
    const alphabets = ["abcdefghij", "0123456789", "WMil.,;:", "xyzw", "AaBbCc0189"];
    let cases = 0;
    for (let trial = 0; trial < 400; trial++) {
      const alpha = alphabets[rnd(alphabets.length)];
      const len = 1 + rnd(300);
      let text = "";
      for (let i = 0; i < len; i++) text += alpha[rnd(alpha.length)];
      const font: FontSpec = {
        family: rnd(2) ? "Calibri" : "Times New Roman",
        size: 8 + rnd(24),
        bold: rnd(2) === 1,
        italic: rnd(2) === 1,
      };
      const letterSpacing = rnd(4) === 0 ? rnd(3) : 0;
      const capacity = 1 + rnd(700);
      const fits = (n: number): boolean => inner.width(text.slice(0, n), font, letterSpacing) <= capacity;

      let lo = 1;
      let hi = text.length;
      while (lo < hi) {
        const mid = lo + Math.ceil((hi - lo) / 2);
        if (fits(mid)) lo = mid;
        else hi = mid - 1;
      }
      let walk = text.length;
      while (walk > 1 && !fits(walk)) walk--;

      expect(lo, `trial ${trial}: len=${text.length} capacity=${capacity}`).toBe(walk);
      cases++;
    }
    expect(cases).toBe(400);
  });

  /**
   * THE VERIFY NET, on a case where bisection genuinely diverges.
   *
   * Bisection is equivalent to the walk only while measurement grows with
   * prefix length — an assumption about fonts, not about code. So the breaker
   * checks its own answer against the boundary property and falls back to the
   * exact walk when the check fails.
   *
   * The hostile measurer here makes ONE length — the whole remainder after the
   * first line — measure far narrower than the prefixes below it. The walk
   * starts at that length, finds it fits, and puts the entire remainder on one
   * line. Bisection probes the middle first, is told it does not fit, and
   * converges to a short line instead. Measured both ways: with the net the
   * layout is 97 + 303 characters (the walk's answer); with the fallback
   * removed it is 97, 97, 98, 98, 10. This test fails without the net.
   */
  it("falls back to the walk when a non-monotone measurer misleads the search", () => {
    const inner = new ApproxMeasurer();
    const token = "abcdefghij".repeat(40);
    const build = (measurerToUse: TextMeasurer): number[] => {
      const doc = DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p(token)) }));
      const result = layoutDocument(doc, { measurer: measurerToUse });
      return result.pages
        .flatMap((page) => page.items)
        .filter((i) => i.kind === "text" && i.text.length > 0)
        .map((i) => (i.kind === "text" ? i.text.length : 0));
    };

    const plainLines = build(inner);
    expect(plainLines.length, "the token must wrap normally first").toBeGreaterThan(2);
    const remainderAfterFirstLine = token.length - plainLines[0];

    const misleading: TextMeasurer = {
      width: (text, font, letterSpacing) => {
        const w = inner.width(text, font, letterSpacing);
        return text.length === remainderAfterFirstLine ? w * 0.2 : w;
      },
      metrics: (font) => inner.metrics(font),
      paintBox: (font) => inner.paintBox?.(font),
      inkBox: (text, font) => inner.inkBox?.(text, font),
    };

    const misledLines = build(misleading);
    expect(misledLines.reduce((a, b) => a + b, 0), "no character lost").toBe(token.length);
    expect(misledLines, "the whole remainder must land on the second line, as the walk would").toEqual([
      plainLines[0],
      remainderAfterFirstLine,
    ]);
  });

  it("a single character that cannot fit still advances (no infinite loop)", () => {
    // A page so narrow that the content width is under one glyph: the breaker
    // must still consume the token rather than spin.
    const { lines } = wrapToken("mmmmmmmmmm", 2900);
    expect(lines.map((l) => l.text).join("")).toBe("mmmmmmmmmm");
  });
});
