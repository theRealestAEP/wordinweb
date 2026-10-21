import { describe, expect, it } from "vitest";
import { createRequire } from "node:module";
import { dirname } from "node:path";
import { DocxDocument } from "../src/docx.js";
import { layoutDocument } from "../src/layout/engine.js";
import { ApproxMeasurer } from "../src/layout/measure.js";
import { FontSet } from "../src/pdf/fonts.js";
import { NodeMeasurer } from "../src/pdf/measure.js";
import { makeDocx, p, wrapDocument } from "./helpers.js";

function documentFor(text: string) {
  return DocxDocument.load(makeDocx({ "word/document.xml": wrapDocument(p(text)) }));
}

function textFor(length: number, words: boolean) {
  const pattern = words ? "alpha bravo charlie delta echo foxtrot golf hotel india juliet " : "abcdefghij";
  return pattern.repeat(Math.ceil(length / pattern.length)).slice(0, length);
}

function lineStarts(result: ReturnType<typeof layoutDocument>): number[] {
  const positions: number[] = [];
  for (const page of result.pages) {
    let lastTop: number | undefined;
    for (const item of page.items) {
      if (item.kind !== "text" || !item.src) continue;
      if (item.lineTop !== lastTop) positions.push(item.src.offset);
      lastTop = item.lineTop;
    }
  }
  return positions;
}

for (const words of [true, false]) {
  describe(words ? "one paragraph of words and spaces" : "one paragraph containing a long word", () => {
    it("keeps the original 2k-character break positions", () => {
      const result = layoutDocument(documentFor(textFor(2000, words)), { measurer: new ApproxMeasurer() });
      // Captured from the original breaker before the performance change.
      expect(lineStarts(result)).toEqual(words
        ? [0, 102, 209, 315, 417, 524, 630, 732, 839, 945, 1047, 1154, 1260, 1362, 1469, 1575, 1677, 1784, 1890, 1992]
        : [0, 97, 194, 292, 390, 487, 584, 682, 780, 877, 974, 1072, 1170, 1267, 1364, 1462, 1560, 1657, 1754, 1852, 1950]);
    });

    it("lays out 40k characters in less than three times the time for 20k", () => {
      const small = documentFor(textFor(20000, words));
      const large = documentFor(textFor(40000, words));
      const timed = (doc: DocxDocument): number => {
        // Each call gets a fresh measurer, so the paragraph cache cannot skip layout.
        const measurer = new ApproxMeasurer();
        // Process CPU time excludes scheduling delays from other test workers.
        const start = process.cpuUsage();
        layoutDocument(doc, { measurer });
        const elapsed = process.cpuUsage(start);
        return (elapsed.user + elapsed.system) / 1000;
      };
      // Warm up both sizes. Alternate their order to spread JIT and GC costs.
      for (let i = 0; i < 3; i++) {
        timed(small);
        timed(large);
      }
      const smallTimes: number[] = [];
      const largeTimes: number[] = [];
      for (let i = 0; i < 5; i++) {
        if (i % 2 === 0) {
          smallTimes.push(timed(small));
          largeTimes.push(timed(large));
        } else {
          largeTimes.push(timed(large));
          smallTimes.push(timed(small));
        }
      }
      // Compare the fastest samples to reduce pauses from GC.
      // The measurement-volume test below also checks growth deterministically.
      const smallMs = Math.min(...smallTimes);
      const largeMs = Math.min(...largeTimes);
      expect(largeMs / smallMs, `20k: ${smallMs.toFixed(1)} ms; 40k: ${largeMs.toFixed(1)} ms`).toBeLessThan(3);
    }, 30000);
  });
}

it("bounds the total characters measured while wrapping a long word", () => {
  class CountingMeasurer extends ApproxMeasurer {
    characters = 0;

    override width(...args: Parameters<ApproxMeasurer["width"]>): number {
      this.characters += args[0].length;
      return super.width(...args);
    }
  }
  const measured = (length: number): number => {
    const measurer = new CountingMeasurer();
    layoutDocument(documentFor(textFor(length, false)), { measurer });
    return measurer.characters;
  };
  // Count input volume as well as calls: a few whole-remainder probes per
  // line caused the quadratic cost even though call counts stayed small.
  expect(measured(40000)).toBeLessThan(measured(20000) * 2.2);
});

it("keeps PDF line breaks identical to the original verified search", () => {
  const require = createRequire(import.meta.url);
  const fonts = new FontSet([dirname(require.resolve("@fontsource/carlito/files/carlito-latin-400-normal.woff"))]);
  for (const words of [true, false]) {
    const doc = documentFor(textFor(2000, words));
    const original = new NodeMeasurer(fonts);
    const expected = layoutDocument(doc, {
      measurer: {
        width: (text, font, spacing) => original.width(text, font, spacing),
        metrics: (font) => original.metrics(font),
      },
    });
    const actual = layoutDocument(doc, { measurer: new NodeMeasurer(fonts) });
    expect(actual.pages).toEqual(expected.pages);
  }
});
