import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { describe, expect, it } from "vitest";
import { createWordServer } from "../src/server.js";

async function connect() {
  const [clientSide, serverSide] = InMemoryTransport.createLinkedPair();
  await createWordServer({ author: "Test Agent" }).connect(serverSide);
  const client = new Client({ name: "test", version: "1" });
  await client.connect(clientSide);
  const call = async (name: string, args: Record<string, unknown> = {}) => {
    const result = await client.callTool({ name, arguments: args }) as { isError?: boolean; content: { text: string }[] };
    const text = result.content[0].text;
    return { isError: result.isError === true, text, value: result.isError ? undefined : JSON.parse(text) };
  };
  return { client, call };
}

describe("wordinweb MCP server", () => {
  it("creates, saves, reopens, and reads a document on disk", async () => {
    const { client, call } = await connect();
    const names = (await client.listTools()).tools.map((tool) => tool.name);
    expect(names).toEqual(expect.arrayContaining(["word_document_open", "word_document_new", "word_document_inspect", "word_document_edit", "word_document_save"]));

    const dir = mkdtempSync(join(tmpdir(), "wordinweb-mcp-"));
    const path = join(dir, "made.docx");
    const created = await call("word_document_new");
    const composed = await call("word_document_compose", { documentId: created.value.documentId, revision: created.value.revision, body: [{ type: "paragraph", text: "Hello from the MCP server." }] });
    expect(composed.isError).toBe(false);
    expect((await call("word_document_save", { documentId: created.value.documentId, path })).value).toMatchObject({ path });
    expect(readFileSync(path).subarray(0, 2).toString()).toBe("PK");

    const opened = await call("word_document_open", { path });
    const projected = await call("word_document_project", { documentId: opened.value.documentId });
    expect(projected.text).toContain("Hello from the MCP server.");
  });

  it("refuses paths that are not .docx, unknown documents, and compose on an opened file", async () => {
    const { call } = await connect();
    const dir = mkdtempSync(join(tmpdir(), "wordinweb-mcp-"));
    writeFileSync(join(dir, "notes.txt"), "x");
    expect((await call("word_document_open", { path: join(dir, "notes.txt") })).text).toBe("The path must name a .docx file.");
    expect((await call("word_document_inspect", { documentId: "missing", kind: "overview" })).text).toMatch(/No open document/);

    const created = await call("word_document_new");
    await call("word_document_compose", { documentId: created.value.documentId, revision: created.value.revision, body: [{ type: "paragraph", text: "x" }] });
    const path = join(dir, "a.docx");
    await call("word_document_save", { documentId: created.value.documentId, path });
    const opened = await call("word_document_open", { path });
    const compose = await call("word_document_compose", { documentId: opened.value.documentId, revision: opened.value.revision, body: [] });
    expect(compose).toMatchObject({ isError: true, text: "word_document_compose works only on a document from word_document_new." });
  });
});
