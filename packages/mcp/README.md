# `@wordinweb/mcp`

A local MCP server that opens, edits, and saves Word `.docx` files. It works with
Claude Code, Claude Desktop, Codex, Cursor, and any other MCP client.

- **No account.** It runs on your computer with `npx`.
- **Your files stay local.** The server reads and writes only the `.docx` paths
  that your client names. It sends nothing over the network.
- **Word-compatible.** Edits, comments, and tracked changes use the
  [WordInWeb](https://github.com/theRealestAEP/wordinweb) engine, and Word opens
  the result.

## Install

Claude Code:

```bash
claude mcp add wordinweb -- npx -y @wordinweb/mcp
```

Codex:

```bash
codex mcp add wordinweb -- npx -y @wordinweb/mcp
```

Claude Desktop, Cursor, and other clients that use a JSON config:

```json
{
  "mcpServers": {
    "wordinweb": { "command": "npx", "args": ["-y", "@wordinweb/mcp"] }
  }
}
```

Node.js 20 or later is necessary.

## Tools

| Tool | What it does |
|---|---|
| `word_document_open` | Opens a `.docx` file and returns a `documentId` |
| `word_document_new` | Starts an empty document |
| `word_document_compose` | Fills a new document from blocks, tables, and media |
| `word_document_inspect` | Reads text, structure, search results, or page geometry |
| `word_document_project` | Shows a story as text or markdown with line anchors |
| `word_document_capabilities` | Lists the edit operations |
| `word_document_edit` | Applies edits, comments, and tracked changes |
| `word_document_patch` | Applies line edits or a diff to a projection |
| `word_document_asset` | Reads an image or embedded file |
| `word_document_save` | Writes the document to a `.docx` file |
| `word_document_close` | Forgets an open document |

## Options

Set `WORDINWEB_AUTHOR` to name the author of comments and tracked changes. The
default is `AI Agent`.

```bash
claude mcp add wordinweb -e WORDINWEB_AUTHOR="Claude" -- npx -y @wordinweb/mcp
```
