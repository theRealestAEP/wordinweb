import { randomUUID } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { extname, resolve } from "node:path";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { CallToolRequestSchema, ListToolsRequestSchema, type CallToolResult, type Tool } from "@modelcontextprotocol/sdk/types.js";
import { AgentDocument } from "@wordinweb/agent";

const INSTRUCTIONS = `Open a .docx with word_document_open, or start an empty one with word_document_new.
Each call names the document by its documentId. Inspect or project the document before an edit,
and pass the revision the read returned. Call word_document_capabilities to list the edit operations.
Call word_document_save to write the result to a .docx file. Files stay on this computer.`;

const documentId = { type: "string", description: "The documentId from word_document_open or word_document_new." };

/** An MCP server over WordInWeb's agent tools. It reads and writes only the paths the client names. */
export function createWordServer(options: { author?: string } = {}): Server {
  const provenance = { author: options.author ?? "AI Agent" };
  const documents = new Map<string, { doc: AgentDocument; path?: string }>();

  // A new document lists every tool, word_document_compose included.
  // Saving is replaced by a tool that writes a file instead of returning base64.
  const documentTools: Tool[] = AgentDocument.create().tools()
    .filter((tool) => tool.name !== "word_document_save")
    .map((tool) => {
      const schema = tool.inputSchema as { properties?: Record<string, unknown>; required?: string[] };
      return {
        name: tool.name,
        description: tool.name === "word_document_compose" ? `${tool.description} Only for a document from word_document_new.` : tool.description,
        inputSchema: { ...schema, type: "object", properties: { documentId, ...schema.properties }, required: ["documentId", ...(schema.required ?? [])] },
      };
    });
  const tools: Tool[] = [
    {
      name: "word_document_open",
      description: "Open a .docx file from this computer. Returns the documentId and revision for later calls.",
      inputSchema: { type: "object", properties: { path: { type: "string", description: "Path to a .docx file." } }, required: ["path"], additionalProperties: false },
    },
    {
      name: "word_document_new",
      description: "Start an empty document. Fill it with word_document_compose, then save it.",
      inputSchema: { type: "object", additionalProperties: false },
    },
    ...documentTools,
    {
      name: "word_document_save",
      description: "Write the document to a .docx file. Without a path, overwrite the file it was opened from.",
      inputSchema: { type: "object", properties: { documentId, path: { type: "string", description: "Path of the .docx file to write." } }, required: ["documentId"], additionalProperties: false },
    },
    {
      name: "word_document_close",
      description: "Forget an open document. Unsaved edits are lost.",
      inputSchema: { type: "object", properties: { documentId }, required: ["documentId"], additionalProperties: false },
    },
  ];

  const docxPath = (path: unknown): string => {
    if (typeof path !== "string" || extname(path).toLowerCase() !== ".docx") throw new Error("The path must name a .docx file.");
    return resolve(path);
  };
  const opened = (id: unknown) => {
    const entry = documents.get(String(id));
    if (!entry) throw new Error("No open document has this documentId. Call word_document_open or word_document_new first.");
    return entry;
  };
  const add = (doc: AgentDocument, path?: string) => {
    const id = randomUUID();
    documents.set(id, { doc, path });
    return { documentId: id, revision: doc.revision, ...(path ? { path } : {}) };
  };

  async function call(name: string, input: Record<string, unknown>): Promise<unknown> {
    if (name === "word_document_open") {
      const path = docxPath(input.path);
      return add(AgentDocument.load(await readFile(path), { provenance }), path);
    }
    if (name === "word_document_new") return add(AgentDocument.create({ provenance }));
    const entry = opened(input.documentId);
    if (name === "word_document_save") {
      const path = input.path === undefined && entry.path ? entry.path : docxPath(input.path);
      await writeFile(path, entry.doc.save());
      entry.path = path;
      return { path, revision: entry.doc.revision };
    }
    if (name === "word_document_close") {
      documents.delete(String(input.documentId));
      return { closed: true };
    }
    const tool = entry.doc.tools().find((candidate) => candidate.name === name);
    if (!tool) throw new Error(name === "word_document_compose" ? "word_document_compose works only on a document from word_document_new." : `Unknown tool: ${name}`);
    const { documentId: _, ...rest } = input;
    return tool.execute(rest);
  }

  const server = new Server({ name: "wordinweb", version: "0.1.0" }, { capabilities: { tools: {} }, instructions: INSTRUCTIONS });
  server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));
  server.setRequestHandler(CallToolRequestSchema, async (request): Promise<CallToolResult> => {
    try {
      const value = await call(request.params.name, request.params.arguments ?? {});
      return { content: [{ type: "text", text: JSON.stringify(value) }] };
    } catch (error) {
      return { isError: true, content: [{ type: "text", text: error instanceof Error ? error.message : String(error) }] };
    }
  });
  return server;
}
