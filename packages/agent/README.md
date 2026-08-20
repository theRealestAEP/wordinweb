# `@wordinweb/agent`

A complete DOCX creation, inspection, and editing toolkit for AI agents. It
gives agents progressive, schema-enforced control over document content,
formatting, review data, tables, equations, drawings, and page layout.

The repository includes the installable
[`wordinweb-documents` agent skill](https://github.com/theRealestAEP/wordinweb/tree/main/skills/wordinweb-documents).
It teaches agents compact composition and progressive editing. It documents
the full interface, operation schemas, result types, and composition rules.

Install the skill from GitHub for Codex:

```bash
gh skill install theRealestAEP/wordinweb wordinweb-documents --agent codex --scope user
```

## Install

```bash
npm install @wordinweb/agent
```

## Live AI invitation

An AI invitation link can bootstrap a live shell session without an MCP
installation:

```bash
npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-6' wordinweb-agent connect '<short invitation URL>'
```

The command starts a detached local bridge and returns a `sessionId`. For Codex,
the bridge starts one dedicated resident document thread. Each private message
starts a turn on that thread, so its document context stays warm. The bridge
waits without model turns. Claude uses its captured session for each private
message. Each started turn supplies a `wakeId`. Include it in every document
command for that turn:

```bash
npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-6' wordinweb-agent session '<sessionId>' '{"command":"sync","wakeId":"<wakeId>"}'
```

The bridge supports `sync`, `capabilities`, `inspect`, `edit`, `project`,
`patch`, `chat`, and `close`. `wait` remains available for manual clients. Every
edit includes the inspected revision. When the room advances, the package
compares the edit's inspected targets with their current state. Unchanged
targets proceed across the newer revision. Changed targets return `needs_sync`.
A `patch` against a stale window returns `needs_sync` the same way. In
suggestion mode the bridge applies every patch as tracked changes. The bridge
places the agent's visible collaboration cursor after its latest edit. A
successful turn ends with a `chat` command for the current `wakeId`. The bridge
interrupts a turn after 60 seconds and reports the failure in the private
document chat.

For a bulk prose rewrite, project the story once and send the changed lines
back:

```bash
npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-6' wordinweb-agent session '<sessionId>' '{"command":"project","wakeId":"<wakeId>","request":{"mode":"md"}}'
npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-6' wordinweb-agent session '<sessionId>' '{"command":"patch","wakeId":"<wakeId>","request":{"revision":"<projection revision>","mode":"md","edits":[{"startLine":4,"endLine":4,"newText":"Adopt the managed platform."}]}}'
```

Inspect the current wake state with a short status command:

```bash
npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-6' wordinweb-agent session '<sessionId>' '{"daemon":"status"}'
```

For a broad text task, read all non-empty stories in one bounded compact call:

```bash
npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-6' wordinweb-agent session '<sessionId>' '{"command":"inspect","wakeId":"<wakeId>","request":{"kind":"context"}}'
```

## Headless use

```ts
import { readFile, writeFile } from "node:fs/promises";
import { AgentDocument } from "@wordinweb/agent";

const agentDoc = AgentDocument.load(await readFile("input.docx"), {
  provenance: { author: "Document agent" },
});

const overview = agentDoc.inspect({ kind: "overview" });
const first = agentDoc.inspect({
  kind: "read",
  story: "body",
  maxBlocks: 20,
  maxCharacters: 12_000,
});

if (!("blocks" in first) || first.blocks[0].type !== "paragraph") {
  throw new Error("Expected a paragraph");
}

const paragraph = first.blocks[0];
const run = paragraph.runs[0];
await agentDoc.edit({
  revision: overview.revision,
  operations: [{
    kind: "insertText",
    at: { blockRef: paragraph.ref, runRef: run.ref, offset: 0 },
    text: "Drafted without a browser. ",
  }],
});

await writeFile("output.docx", agentDoc.save());
```

For new documents, compose the complete structure in one schema-enforced call:

```ts
const document = AgentDocument.create();
const result = await document.compose({
  revision: document.revision,
  page: { margins: { top: 0.75, right: 0.75, bottom: 0.75, left: 0.75 } },
  body: [
    { type: "paragraph", text: "Systems Decision Brief", styleId: "Title" },
    { type: "heading", level: 1, text: "Context" },
    { type: "paragraph", text: "The team must select a delivery model." },
    {
      type: "table",
      headerRows: 1,
      rows: [["Option", "Strength"], ["Managed", "Fast delivery"], ["Custom", "Maximum control"]],
    },
    {
      type: "chart",
      widthPx: 520,
      heightPx: 280,
      align: "center",
      chart: {
        type: "column",
        title: "Decision profile",
        categories: ["Control", "Speed"],
        series: [{ name: "Managed", values: [8, 9] }, { name: "Custom", values: [10, 6] }],
      },
    },
  ],
  footer: [{
    type: "pageNumber",
    fieldKind: "pageOfTotal",
    align: "center",
    color: "627D98",
    fontSizePt: 8,
  }],
});

console.log(result.overview.outline, result.createdObjects);
```

`AgentDocument.create()` starts a blank DOCX. `tools()` returns eight portable
tool objects for composition, capabilities, inspection, edits, text projection,
projection patches, asset reads, and DOCX saves.

## Text projection

`project` renders a story as deterministic text — plain paragraphs, structural
markdown, or the outline — with an anchor map that carries every line back to
its block, its runs, and their wire offsets. `patch` takes line-range hunks or
a unified diff written against that projection and compiles them into the same
intents an editor emits, so a bulk rewrite costs one window and one call.

```ts
const projection = agentDoc.project({ mode: "md" });
await agentDoc.patch({
  revision: projection.revision,
  mode: "md",
  edits: [{ startLine: 4, endLine: 4, newText: "Adopt the managed platform." }],
});
```

Agents read the text and never see a wire offset: the anchor map is built in
the same pass as the text and owns that translation. Projecting never changes
the document. The projection contract, the atom placeholders, and the patch
rules live in the skill's `references/interface.md`.

## Progressive inspection

For a broad text task, start with `context`. It returns text and edit references
from all non-empty stories in one response. The default global budget is 100
blocks and 24,000 characters. It omits formatting, empty metadata, bookmarks,
and object summaries. Add `include: ["bookmarks", "objects"]` only when the
task needs those fields.

Use `overview` when the task needs page setup, the outline, story counts, or
object counts. A compose call already returns this overview plus compact
summaries for created objects. Then use these requests:

- `context` returns bounded bulk editing context. Its story cursor can continue
  through `read` when one story exceeds the global budget.
- `read` returns a bounded block and character window. Use its cursor for the
  next window. It includes detailed formatting, components, bookmarks, and
  table cell references.
- `search` returns bounded matches with block and run references.
- `object` expands one equation, image, shape, chart, SmartArt diagram, field,
  note reference, or other inline component.
- `spatial` returns bounded page geometry. It also computes rotated-object
  intersections, overlap area, and paint order.
- `asset` reads bytes for an `assetRef` from an object result.

Browser layout uses canvas metrics and reports `quality: "exact"`. Headless
layout uses deterministic approximate metrics and reports
`quality: "approximate"`. Both modes use the same page layout model.

## References and edit safety

The interface exposes opaque stable references for document targets.

- `block:*` and `run:*` references identify editable nodes.
- `object:*` references identify inspectable components and supported drawing edit targets.
- `view:*` references identify render-only content, such as cached SmartArt
  text or endnote content that the current editor cannot mutate.
- An object result contains `editRef` only when an edit target exists.

Every edit includes the revision that the agent inspected. A newer revision
triggers a fingerprint comparison for the referenced paragraphs, tables, and
objects. Unchanged targets proceed, while changed targets request a new
inspection. Document-wide operations use the global revision. Operation schemas
reject extra fields, internal IDs, bad reference types, and invalid canonical
intents. Local batches apply to a clone first, so a failed operation leaves the
source document unchanged.

Call `capabilities(category, kind)` or the capabilities tool before an edit.
This returns a closed JSON Schema for the selected operation. The package
exposes every canonical WordInWeb edit intent.

The portable edit tool embeds every operation schema as a discriminated
union. Nested patches, chart data, SmartArt data, table operations, and page
layout values use closed schemas.

Objects placed in a header or footer story repeat on the applicable section
pages. Insert a line with `insertShape`, then use its `editRef` with position,
size, rotation, line-style, order, or removal operations. The same operations
apply to Word-authored VML lines.

## Solo browser session

Use `LocalDocumentSession` when a human and an agent edit the same local page.
The session owns one `DocxDocument`. Agent edits and `DocxView` edits pass
through the same canonical apply path.

```tsx
import { useMemo } from "react";
import { DocxView, useAgentDocumentSession } from "wordinweb";
import {
  AgentDocument,
  LocalDocumentSession,
  localDocumentViewBinding,
} from "@wordinweb/agent";

export function LocalEditor({ bytes }: { bytes: Uint8Array }) {
  const session = useMemo(() => new LocalDocumentSession(bytes), [bytes]);
  const binding = useMemo(() => localDocumentViewBinding(session), [session]);
  const view = useAgentDocumentSession(binding);

  // Give this object, or agentDoc.tools(), to the host agent runtime.
  const agentDoc = useMemo(() => AgentDocument.connect(session), [session]);
  void agentDoc;

  return <DocxView {...view} editable style={{ height: "100vh" }} />;
}
```

## Collaborative and offline sessions

The existing React collaboration session already owns the live document,
stable ID allocator, media relay, and offline tail. Adapt that session without
adding React to this package:

```ts
import {
  AgentDocument,
  collaborativeAgentTarget,
} from "@wordinweb/agent";

const agentDoc = AgentDocument.connect(
  collaborativeAgentTarget(() => currentCollabSession),
  { provenance: { author: "Document agent" } },
);
```

The callback must return the latest `useCollab` session after each React
render. The agent API stays the same in all modes.

| Mode | Document owner | Edit route | Browser required |
| --- | --- | --- | --- |
| Headless | `AgentDocument` | Local atomic clone | No |
| Solo browser | `LocalDocumentSession` | Canonical local session | Yes, for review and human edits |
| Live collaboration | `useCollab` session | Sequenced room intent | Yes, for the room client |
| Offline collaboration | `useCollab` session | Durable offline intent tail | Yes, for the offline client |

An offline agent edit applies to the local replica and enters the same durable
tail as a human edit. Reconnect uses the existing replay and reconciliation
flow. A live agent edit passes through the existing server sequence and reaches
all collaborators.

## Fixture and coverage gates

`npm run test:agent-fixtures` audits every DOCX in the WordInWeb parity corpus.
It checks semantic components, object detail, assets, JSON safety, every story,
every page, and overlap geometry.

Typed exhaustive gates cover canonical intents, agent capabilities, run
components, layout primitives, canonical apply, canonical validation, and DOM
rendering. The browser convergence test consumes the same canonical intent
list. A new edit or component must reach each surface before the build passes.

The repository also includes actual Codex skill evaluations. Each natural
document request starts a fresh Codex process in a temporary workspace,
installs this repository skill, creates a DOCX through the public package, and
scores the output through semantic, spatial, asset, story, section, and OOXML
evidence. The requests cover a technical brief, branded media guide, field
inspection packet, policy review copy, and visual strategy report.

```bash
npm run test:agent-skill:codex
```

Pass fixture names after `--` to run a focused evaluation:

```bash
npm run test:agent-skill:codex -- media-brand-guide policy-review-copy
```

Set `WORDINWEB_AGENT_EVAL_MODEL` to select a model. Set
`WORDINWEB_KEEP_AGENT_EVAL=1` to retain the temporary workspace for review.

`npm run test:agent-skill` checks that the GitHub skill documents every runtime
tool, inspection mode, reference family, edit operation, and agent-facing
field. It also validates every portable authoring fixture. Supply a DOCX and
its fixture name to run the complete scorer:

```bash
WORDINWEB_AGENT_EVAL_OUTPUT=/path/to/output.docx \
WORDINWEB_AGENT_EVAL_FIXTURE=media-brand-guide \
npm run test:agent-skill
```
