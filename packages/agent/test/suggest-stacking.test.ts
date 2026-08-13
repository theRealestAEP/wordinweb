import { describe, expect, it } from "vitest";
import { AgentDocument, LocalDocumentSession } from "../src/index.js";
import { body, makeDocx } from "./helpers.js";
import { collectRevisions, rejectAllRevisions } from "@wordinweb/core";

/**
 * SUGGESTIONS STACK: A SECOND ONE READS THE FIRST AS THE READER DOES (#136).
 *
 * Reported against the desktop app: "you can[']t really stack the AI
 * suggestions? things get weird".
 *
 * The AI panel turns suggesting on for its turn, which forces the revision
 * view to markup so the reader can watch the suggestion appear. The agent's
 * projection is derived from the same model, and in markup view that model
 * still holds the deleted runs. So the second turn was handed the pending
 * insertion and the pending deletion INTERLEAVED — "The quick brown fox
 * leajumps over the lazy dog" for a document that reads "leaps" — and, since
 * deleted runs are not editable text, a rewrite spanning one was refused
 * outright with "cannot be rewritten across a non-text atom".
 *
 * Both symptoms are the same cause, so both are pinned here: what the second
 * turn READS, and that its edit lands.
 */

const SEED = body(`<w:p><w:r><w:t>The quick brown fox jumps over the lazy dog</w:t></w:r></w:p>`);

function connect() {
  const session = new LocalDocumentSession(makeDocx(SEED));
  const agent = AgentDocument.connect(session, {
    provenance: { author: "AI", now: () => "2026-08-12T00:00:00Z" },
  });
  return { session, agent };
}

/** Rewrite line 1 as a suggestion, the way the AI panel's patch tool does. */
async function suggestLine(agent: AgentDocument, newText: string): Promise<void> {
  const window = agent.project({ mode: "text" });
  await agent.patch({
    revision: window.revision,
    mode: "text",
    suggest: true,
    edits: [{ startLine: 1, endLine: 1, newText }],
  });
}

const reads = (agent: AgentDocument) => agent.project({ mode: "text" }).text;

describe("stacked suggestions", () => {
  it("shows the second turn the text a reader sees, not the markup", async () => {
    const { session, agent } = connect();
    // The panel forces markup for the duration of its turn.
    session.doc.setRevisionView("markup");

    await suggestLine(agent, "The quick brown fox leaps over the lazy dog");
    expect(collectRevisions(session.doc).length, "one delete + one insert").toBe(2);
    // The reader sees "leaps". So must the next turn — NOT "leajumps".
    expect(reads(agent)).toBe("The quick brown fox leaps over the lazy dog");
    expect(session.doc.revisionView, "and the screen still shows the markup").toBe("markup");
  });

  it("lets a second suggestion land on top of the first", async () => {
    const { session, agent } = connect();
    session.doc.setRevisionView("markup");

    await suggestLine(agent, "The quick brown fox leaps over the lazy dog");
    await suggestLine(agent, "The swift brown fox leaps over the lazy dog");

    expect(collectRevisions(session.doc).length, "two changes, tracked separately").toBe(4);
    expect(reads(agent)).toBe("The swift brown fox leaps over the lazy dog");
    // Both are still suggestions: rejecting everything restores the original.
    rejectAllRevisions(session.doc);
    expect(reads(agent)).toBe("The quick brown fox jumps over the lazy dog");
  });

  it("reads the same whichever view the reader happens to be in", async () => {
    // The projection is what the model edits against, so it must not depend on
    // a display setting the reader can toggle at any moment.
    const { session, agent } = connect();
    await suggestLine(agent, "The quick brown fox leaps over the lazy dog");

    session.doc.setRevisionView("final");
    const inFinal = reads(agent);
    session.doc.setRevisionView("markup");
    const inMarkup = reads(agent);
    expect(inMarkup).toBe(inFinal);
    expect(session.doc.revisionView, "reading did not disturb the view").toBe("markup");
  });
});
