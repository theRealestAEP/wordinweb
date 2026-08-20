import { describe, expect, it } from "vitest";
import { DocxDocument } from "../src/docx.js";
import { compareDocuments } from "../src/edit/compare/index.js";
import { collectRevisions } from "../src/edit/suggest.js";
import { AUTHOR, DATE, kindCounts, kindsIn, loadBody, para, revisionTexts } from "./compare-helpers.js";

/**
 * THE QUALITY GATE.
 *
 * The round-trip gate (compare-roundtrip.test.ts) proves the output can be
 * reviewed back to either input. It says nothing about whether the output is
 * WORTH reviewing: delete-the-whole-paragraph-and-reinsert-it round-trips
 * perfectly and is useless to a human. That is what this file pins down.
 *
 * Every fixture asserts an exact revision count and kind distribution, never a
 * rendered snapshot — counts survive refactoring and a diff feature will be
 * refactored. Each one names the failure mode it exists to catch.
 *
 * ## Calibration
 *
 * Three numbers decide the output, all in COMPARE_TUNING. Only the first came
 * with outside evidence, and that evidence was gathered on syntax trees rather
 * than prose, so all three were settled here:
 *
 *  - **minDice 0.5** — kept as GumTree's default. Nothing in the corpus wanted
 *    it moved: the gap bound already stops distant look-alikes from pairing,
 *    which is the failure 0.5 would otherwise have to prevent on its own.
 *  - **shortParagraphTokens 5 / minShortRatio 0.8** — kept. It is what keeps
 *    "Alpha item" from matching "Bravo item" in fixture 7, which is the
 *    reordered-short-list failure. Loosening the ratio to 0.6 makes those two
 *    pair and turns a moved list item into a nonsense in-place rewrite.
 *  - **coalesceWindow 3** — kept as the absolute floor, and NOT enough on its
 *    own. Fixture 3 is the proof: a rewritten sentence keeps three matched
 *    words ("shortage of test") in the middle, the floor lets them through,
 *    and the reviewer gets two strike/insert pairs where they made one edit.
 *    **coalesceRatio 2 was added for that**, and the two calibration fixtures
 *    at the bottom are what stop it from over-merging.
 */

const DIFF = { author: AUTHOR, date: DATE } as const;

function compare(originalBody: string, revisedBody: string, options: object = {}): DocxDocument {
  return compareDocuments(loadBody(originalBody), loadBody(revisedBody), { ...DIFF, ...options });
}

const HEADING = `<w:pStyle w:val="Heading1"/>`;

function table(rows: string[][]): string {
  const grid = rows[0].map(() => `<w:gridCol w:w="2000"/>`).join("");
  const body = rows.map((cells) => `<w:tr>${cells.map((c) => `<w:tc><w:tcPr/>${para(c)}</w:tc>`).join("")}</w:tr>`).join("");
  return `<w:tbl><w:tblPr/><w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

// A hundred words of ordinary prose, so a one-word fix has somewhere to hide.
const HUNDRED =
  "The committee will recieve the quarterly statement before the annual meeting, and every member is " +
  "expected to read it in full. The statement covers revenue, operating costs, capital expenditure and " +
  "the reserve position as at the end of the period. Where a figure has moved by more than five per cent " +
  "against the prior year, the finance team has added a short note explaining the movement. Questions " +
  "should be sent to the secretary at least three working days before the meeting so that answers can be " +
  "prepared in advance and circulated with the agenda.";

describe("compare corpus — 1: a typo fixed inside a hundred words", () => {
  it("strikes one word and inserts one, and touches nothing else", () => {
    // FAILURE MODE: the whole paragraph deleted and reinserted.
    const doc = compare(para(HUNDRED), para(HUNDRED.replace("recieve", "receive")));
    expect(kindCounts(doc)).toEqual({ deletion: 1, insertion: 1 });
    expect(revisionTexts(doc, "del")).toEqual(["recieve "]);
    expect(revisionTexts(doc, "ins")).toEqual(["receive "]);
  });
});

describe("compare corpus — 2: two words changed two hundred words apart", () => {
  it("makes two localised pairs, not one span from the first change to the last", () => {
    // FAILURE MODE: one giant region swallowing everything between the edits.
    const filler = "filler word here ".repeat(40) + "omega beta " + "more filler text ".repeat(40);
    const doc = compare(para("Alpha " + filler + "final gamma"), para("ALPHA " + filler + "final GAMMA"));
    expect(kindCounts(doc)).toEqual({ deletion: 2, insertion: 2 });
    expect(revisionTexts(doc, "del")).toEqual(["Alpha ", "gamma"]);
    expect(revisionTexts(doc, "ins")).toEqual(["ALPHA ", "GAMMA"]);
  });
});

describe("compare corpus — 3: one sentence rewritten out of five", () => {
  const S1 = "The project began in March with a small team. ";
  const S2 = "Early progress was steady and the first milestone arrived on time. ";
  const OLD = "The second phase was delayed by a shortage of test hardware, which took six weeks to resolve. ";
  const NEW = "A shortage of test hardware held up the second phase for six weeks before it was resolved. ";
  const S4 = "By the summer the schedule had been recovered. ";
  const S5 = "The final release is expected before the end of the year.";

  it("strikes the sentence as ONE run and inserts the new one as ONE run", () => {
    // FAILURE MODE: confetti. The two wordings share "shortage of test", "six
    // weeks" and "the second phase", so a raw LCS matches them and returns a
    // dozen interleaved fragments. This is the fixture the coalescing ratio
    // exists for.
    const doc = compare(para(S1 + S2 + OLD + S4 + S5), para(S1 + S2 + NEW + S4 + S5));
    expect(kindCounts(doc)).toEqual({ deletion: 1, insertion: 1 });
    expect(revisionTexts(doc, "del")).toEqual([OLD]);
    expect(revisionTexts(doc, "ins")).toEqual([NEW]);
  });
});

describe("compare corpus — 4: a paragraph moved from position 2 to position 7", () => {
  const line = (n: string): string => para(`${n} paragraph text that is reasonably long so it fingerprints well.`);
  const names = ["Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel"];
  const moved = ["Alpha", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Bravo", "Hotel"];

  it("shows one deletion and one insertion, and does not cascade", () => {
    // FAILURE MODE: the alignment shifts and paragraphs 3-6 all read as
    // changed. Moves themselves are NOT detected — RevisionKind has no
    // moveFrom/moveTo and nothing writes those elements — so delete-plus-insert
    // is the correct output here, and is what Word usually produces too.
    const doc = compare(names.map(line).join(""), moved.map(line).join(""));
    expect(kindCounts(doc)).toEqual({ deletion: 1, insertion: 1, markDeletion: 1, markInsertion: 1 });

    // The untouched paragraphs carry no TEXT revision. The two mark revisions
    // are on paragraphs 0 and 5 by construction: an added or removed paragraph
    // is recorded by striking the pilcrow of the paragraph BEFORE it, which is
    // where OOXML puts that difference. Nothing about paragraphs 1-5's own
    // content is touched.
    // Block 1 of the RESULT is the struck copy of "Bravo"; 2-6 and 8 are the
    // paragraphs that merely stood still.
    for (const at of [2, 3, 4, 5, 6, 8]) {
      expect(kindsIn(doc, at).filter((k) => k === "insertion" || k === "deletion")).toEqual([]);
    }
    expect(revisionTexts(doc, "del")).toEqual(["Bravo paragraph text that is reasonably long so it fingerprints well."]);
    expect(revisionTexts(doc, "ins")).toEqual(["Bravo paragraph text that is reasonably long so it fingerprints well."]);
  });
});

describe("compare corpus — 5: a paragraph split in two", () => {
  it("records ONE inserted paragraph mark and no text revision at all", () => {
    // FAILURE MODE: delete the whole paragraph, insert two new ones.
    const doc = compare(para("First sentence. Second sentence."), para("First sentence. ") + para("Second sentence."));
    expect(kindCounts(doc)).toEqual({ markInsertion: 1 });
  });
});

describe("compare corpus — 6: two paragraphs merged into one", () => {
  it("records ONE deleted paragraph mark and no text revision at all", () => {
    // FAILURE MODE: delete two paragraphs, insert one.
    const doc = compare(para("First sentence. ") + para("Second sentence."), para("First sentence. Second sentence."));
    expect(kindCounts(doc)).toEqual({ markDeletion: 1 });
  });
});

describe("compare corpus — 7: a short list reordered", () => {
  it("moves one item and leaves the first alone", () => {
    // FAILURE MODE: the short-paragraph threshold matches "Alpha item" against
    // "Bravo item" (they share half their characters) and all three items come
    // out rewritten.
    const doc = compare(
      para("Alpha item") + para("Bravo item") + para("Charlie item"),
      para("Alpha item") + para("Charlie item") + para("Bravo item"),
    );
    expect(revisionTexts(doc, "del")).toEqual(["Bravo item"]);
    expect(revisionTexts(doc, "ins")).toEqual(["Bravo item"]);
    expect(kindsIn(doc, 0).filter((k) => k === "insertion" || k === "deletion")).toEqual([]);
  });
});

describe("compare corpus — 8: a table row inserted mid-table", () => {
  it("inserts one row and leaves every other row untouched", () => {
    // FAILURE MODE: cascade — every row below the insert reads as changed
    // because the rows were aligned by position instead of by content.
    const doc = compare(
      table([
        ["Region", "Revenue"],
        ["North", "1200"],
        ["South", "980"],
        ["East", "1450"],
      ]),
      table([
        ["Region", "Revenue"],
        ["North", "1200"],
        ["Central", "610"],
        ["South", "980"],
        ["East", "1450"],
      ]),
    );
    expect(kindCounts(doc)).toEqual({ rowInsertion: 1, insertion: 2 });
    expect(revisionTexts(doc, "ins")).toEqual(["Central", "610"]);
    expect(revisionTexts(doc, "del")).toEqual([]);
  });

  it("deletes one row and leaves every other row untouched", () => {
    const doc = compare(
      table([["Region"], ["North"], ["Central"], ["South"]]),
      table([["Region"], ["North"], ["South"]]),
    );
    expect(kindCounts(doc)).toEqual({ rowDeletion: 1, deletion: 1 });
    expect(revisionTexts(doc, "del")).toEqual(["Central"]);
  });
});

describe("compare corpus — 9: formatting only", () => {
  it("records a run format revision and NO text revision", () => {
    // FAILURE MODE: the text is struck and reinserted just to carry the bold.
    const doc = compare(para("Same words throughout"), para("Same words throughout", { rPr: "<w:b/>" }));
    expect(kindCounts(doc)).toEqual({ runFormat: 1 });
  });

  it("records a paragraph format revision and NO text revision", () => {
    const doc = compare(para("Same words throughout"), para("Same words throughout", { pPr: HEADING }));
    expect(kindCounts(doc)).toEqual({ paragraphFormat: 1 });
  });

  it("records nothing at all with formatting comparison off", () => {
    const doc = compare(para("Same words"), para("Same words", { rPr: "<w:b/>" }), { formatting: false });
    expect(collectRevisions(doc)).toHaveLength(0);
  });
});

describe("compare corpus — 10: whitespace only", () => {
  it("keeps the paragraph anchored and reports the change by default", () => {
    // FAILURE MODE: the paragraph fails to anchor and renders as a rewrite.
    const doc = compare(para("one  two three"), para("one two three"));
    expect(kindCounts(doc)).toEqual({ deletion: 1, insertion: 1 });
    expect(revisionTexts(doc, "del")).toEqual(["one  "]);
  });

  it("records nothing with whitespace comparison off", () => {
    const doc = compare(para("one  two three"), para("one two three"), { whitespace: false });
    expect(collectRevisions(doc)).toHaveLength(0);
  });
});

describe("compare corpus — 11: the same heading in three chapters", () => {
  const chapter = (body: string): string => para("Introduction", { pPr: HEADING }) + para(body);
  const one = "Chapter one opens with a summary of the year and its main themes.";
  const three = "Chapter three sets out the recommendations and their expected cost.";

  it("edits only chapter two's body, and leaves chapters one and three alone", () => {
    // FAILURE MODE: with no unique anchor among three identical "Introduction"
    // headings, patience diff gives up on the whole region and the chapters
    // scramble. Histogram still prefers the least-repeated candidate.
    const doc = compare(
      chapter(one) + chapter("Chapter two describes the process used to collect the survey data.") + chapter(three),
      chapter(one) + chapter("Chapter two describes the method used to collect the survey data.") + chapter(three),
    );
    expect(kindCounts(doc)).toEqual({ deletion: 1, insertion: 1 });
    expect(revisionTexts(doc, "del")).toEqual(["process "]);
    for (const at of [0, 1, 2, 4, 5]) expect(kindsIn(doc, at)).toEqual([]);
  });
});

describe("compare corpus — 12: identical documents", () => {
  it("finds nothing, which catches whole classes of normalization bug", () => {
    const body =
      para("A heading", { pPr: HEADING }) +
      para(HUNDRED) +
      table([
        ["Region", "Revenue"],
        ["North", "1200"],
      ]) +
      para("A closing line.");
    const doc = compare(body, body);
    expect(collectRevisions(doc)).toHaveLength(0);
  });
});

describe("compare corpus — calibration counter-cases for the coalescing ratio", () => {
  it("two one-word edits three words apart stay two edits", () => {
    // If coalesceRatio swallowed this, every pair of nearby edits would merge
    // and the reviewer would be told words changed that did not.
    const doc = compare(
      para("alpha bravo charlie delta echo foxtrot golf"),
      para("alpha BRAVO charlie delta echo FOXTROT golf"),
    );
    expect(kindCounts(doc)).toEqual({ deletion: 2, insertion: 2 });
    expect(revisionTexts(doc, "del")).toEqual(["bravo ", "foxtrot "]);
  });

  it("a long untouched stretch between two big rewrites is not swallowed", () => {
    const middle = "the middle stretch stays exactly as it was and is fairly long ";
    const doc = compare(
      para("one two three four five six seven eight nine ten " + middle + "aa bb cc dd ee ff gg hh ii jj"),
      para("ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN " + middle + "AA BB CC DD EE FF GG HH II JJ"),
    );
    expect(kindCounts(doc)).toEqual({ deletion: 2, insertion: 2 });
    expect(revisionTexts(doc, "ins")).toEqual([
      "ONE TWO THREE FOUR FIVE SIX SEVEN EIGHT NINE TEN ",
      "AA BB CC DD EE FF GG HH II JJ",
    ]);
  });
});
