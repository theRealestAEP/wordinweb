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
3. Run the exact command from the invitation in the terminal. The command passes the complete URL to `wordinweb-agent connect`, starts a detached local bridge, and returns a `sessionId`.
4. Read the ready result and its room instructions.
5. Confirm that the ready result reports `wake.state` as `armed`.
6. End the current agent turn and stay idle.
7. When the bridge starts a document-agent turn for a private message, run the exact local session command from that message with `{"command":"sync"}` first.
8. Use one `context` inspection for a broad text task. It returns bounded text and edit references from all non-empty stories.
9. Use `overview`, `read`, `object`, or `spatial` only when the task needs their extra detail.
10. Run each JSON command through the local session command supplied in the turn message.
11. Sync and inspect again when an edit returns `needs_sync`.
12. Reply through the document chat command, then end the agent turn.
13. Send `{"command":"close"}` when the inviter ends the collaboration task.

The bridge places the visible agent cursor after each edit. The agent does not send a separate cursor command.

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

1. Use `context` for a broad text task across non-empty stories.
2. Use `overview` only when the task needs sections, layout, outline, or object counts.
3. Use detailed `read`, `search`, `object`, or `spatial` requests only for the relevant target.
4. Expand only the objects that matter to the task.
5. Query `word_document_capabilities` by exact operation kind.
6. Use references returned by the current inspection revision.
7. Apply one bounded transaction against that revision.
8. Re-inspect after a structural edit before use of new content.
9. Inspect affected objects and page ranges.
10. Save after semantic and spatial checks pass.

## Keep context bounded

- Start with `context` for broad text work.
- Use `search` to locate specific content.
- Use detailed `read` only when the compact context lacks a required field.
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
