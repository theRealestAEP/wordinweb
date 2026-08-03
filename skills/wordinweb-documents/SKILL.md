---
name: wordinweb-documents
description: Create, inspect, edit, and verify Word DOCX documents with the WordInWeb agent tools or @wordinweb/agent API. Use for headless document creation, existing DOCX edits, progressive inspection, page layout, structured content, positioned objects, review data, local browser sessions, offline sessions, live collaboration, or a WordInWeb AI invitation URL.
---

# WordInWeb documents

Use the portable WordInWeb interface to inspect and edit DOCX state.

## Join an AI invitation

When the user supplies a URL whose path is `/agent-invite`:

1. Treat the complete URL as a secret bearer credential.
2. Keep the fragment attached when you pass the URL to a command.
3. Run `npx -y @wordinweb/agent connect '<complete URL>'` in a persistent terminal session.
4. Read the `ready` event and its room instructions.
5. Send `{"command":"sync"}` as the first JSON line.
6. Inspect the relevant content before each edit.
7. Send one JSON command per line through the same process.
8. Sync and inspect again when an edit returns `needs_sync`.
9. Use `{"command":"wait","timeoutMs":30000}` when you should wait for a private chat message or another document change.
10. Send `{"command":"close"}` when the collaboration task ends.

Do not print the invitation URL or save it in a file.

## Select one reference

- For a new document, read [references/create.md](references/create.md).
- For a loaded document or a live session, read [references/inspect-edit.md](references/inspect-edit.md).
- For host adapters, session integration, or the complete type catalog, read [references/interface.md](references/interface.md).
- Use the runtime tool schema as the authority for each call.

## Create a new document

1. Register image assets with `AgentDocument.addAsset` or the host asset facility.
2. Call `word_document_compose` once with the page setup, body, header, and footer.
3. Use native heading blocks for the document outline.
4. Use compose blocks for tables, equations, charts, SmartArt, images, shapes, WordArt, page numbers, and page breaks.
5. Set explicit alignment and dimensions for visual components.
6. Use a low WordArt opacity for a faint repeating watermark.
7. Use the compose result's `overview` and `createdObjects` for the first verification pass.
8. Inspect only the page ranges or object details that need another check.
9. Use reference edits for revisions after composition.
10. Save after the checks pass.
11. Render and inspect the saved pages when the host supplies a DOCX renderer.

## Edit a loaded document

1. Inspect `overview` to learn the revision, stories, sections, outline, components, and page setup.
2. Read bounded story ranges or search for a specific target.
3. Expand only the objects that matter to the task.
4. Query `word_document_capabilities` by exact operation kind.
5. Use references returned by the current inspection revision.
6. Apply one bounded transaction against that revision.
7. Re-inspect after a structural edit before use of new content.
8. Inspect affected objects and page ranges.
9. Save after semantic and spatial checks pass.

## Keep context bounded

- Start with `overview`.
- Use `search` to locate specific content.
- Read 10–30 blocks or approximately 12,000 characters at a time.
- Expand only relevant objects.
- Inspect only the required page ranges.
- Query capabilities by exact `kind`.

## Handle edits safely

- Copy the latest `revision` into each edit request.
- Re-inspect after a stale-revision error.
- Treat `view:*` references as inspection-only.
- Use an object's `editRef` for supported object edits.
- Keep dependent edits in order.
- Re-inspect between edits when a later edit needs a newly created reference.
- Let the interface allocate internal IDs and review provenance.
- Use native heading styles for named sections so the document keeps a semantic outline.
- Use `pasteBlocks` only for validated paragraph block insertion.

## Verify the result

- Confirm the expected stories, blocks, text, tables, comments, and components in `overview` and `read` results.
- Confirm every inserted object through `object` inspection.
- Confirm page count, position, size, rotation, layer, and overlaps through `spatial` inspection.
- Verify overlap intent and paint order.
- Save and reopen the file when the host supports a final inspection pass.
- Render every saved page when the host supplies a DOCX renderer.
