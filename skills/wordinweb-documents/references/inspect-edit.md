# Inspect and edit a DOCX

Use this reference for a loaded document, a solo browser session, or a
collaborative session. The complete type catalog remains in
[interface.md](interface.md).

## Progressive inspection

For a broad text task, start with one request:

```json
{ "kind": "context" }
```

The result includes text and edit references from all non-empty stories. The
global default budget is 100 blocks and 24,000 characters. It omits formatting,
empty metadata, bookmarks, and object summaries. Add
`"include": ["bookmarks", "objects"]` only when the task needs them.

Use the smallest detailed request when required:

```json
{ "kind": "overview" }
{ "kind": "read", "story": "body", "maxBlocks": 20, "maxCharacters": 12000 }
{ "kind": "read", "story": "body", "cursor": { "value": "returned cursor" } }
{ "kind": "search", "query": "revenue", "maxResults": 50 }
{ "kind": "object", "ref": "object:12:0" }
{ "kind": "spatial", "pages": { "start": 1, "count": 5 }, "includeOverlaps": true }
{ "kind": "fit", "pages": { "start": 1, "count": 5 } }
```

After inserting or resizing a drawing, check `fit`. It reports each
text-bearing drawing's box against the extent its text laid into, whether the
text overflows, how many lines the box hides, and the shape's autofit mode,
plus how full each page is. Resolve an overflow with `setDrawingTextFit`
(`resizeShape`) or by resizing the drawing.

`overview` returns story counts, section geometry, the outline, component
counts, and semantic `objectCounts`. `read` returns formatting, hyperlinks,
component summaries, bookmarks, table cells, and other detailed fields. A
`context` story cursor can continue through a `read` request for that story.

Each run returns compact component summaries with `ref`, `editRef`, `type`,
and an optional `label`. Expand an object when its summary lacks a required
fact. WordArt details include its text, fill, opacity, story, and geometry.

Use cursors for large stories. Reuse only a cursor from the same revision.
Headless spatial inspection reports deterministic approximate metrics.
Browser spatial inspection reports canvas-based exact metrics.

## Bulk text work

For a task that is mostly "change this prose", project the story as text with
`word_document_project` and send the changed lines back through
`word_document_patch`:

```json
{ "mode": "md" }
```

```json
{
  "revision": "17",
  "mode": "md",
  "edits": [{ "startLine": 4, "endLine": 4, "newText": "Adopt the managed platform." }]
}
```

One projection window plus one patch replaces dozens of structured calls for
rewrites, insertions, paragraph splits, and merges. Keep the structured
operations for formatting, tables, drawings, equations, review data, and page
layout.

The projection contract, the atom placeholders, and the patch rules are in
[interface.md](interface.md#text-projection).

## References

- `block:*` identifies an editable paragraph or table.
- `run:*` identifies an editable run.
- `object:*` identifies an inspectable component.
- `asset:*` identifies readable or insertable binary media.
- `spatial:*` identifies one laid-out occurrence.
- `view:*` identifies revision-scoped inspection content.

Treat every reference as opaque. Use a component's `editRef` for a drawing
edit. Re-inspect after a structural edit creates or removes references.

## Edit workflow

1. Query the exact operation schema.
2. Copy the inspected revision into the edit request.
3. Apply up to 100 related operations as one transaction.
4. Re-inspect the affected content and page range.
5. Save after the checks pass.
6. Render and inspect every page when the host supplies a DOCX renderer.

Query one operation:

```json
{ "kind": "setDrawingLineStyle" }
```

Apply one transaction:

```json
{
  "revision": "17",
  "operations": [
    {
      "kind": "insertText",
      "at": { "blockRef": "block:1", "runRef": "run:2", "offset": 0 },
      "text": "Quarterly report"
    }
  ]
}
```

The capability result supplies the closed JSON Schema. Query by exact `kind`
to keep the response small. The runtime allocates internal IDs, collaboration
sequence data, and review provenance.

For a borderless drawing, call `setDrawingLineStyle` with `color: null`.
Supply `widthPx` and `dash` when `color` contains an RGB value.

## Session behavior

The inspection and edit contracts stay constant across these modes:

- `AgentDocument.load`: headless local document.
- `LocalDocumentSession`: solo browser editor and agent share one document.
- `collaborativeAgentTarget`: live, reconnecting, and durable offline session.

A local transaction validates on a clone before commit. A collaborative
transaction validates against a clone, then submits canonical operations to
the existing room session.
