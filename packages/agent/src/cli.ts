import { webcrypto } from "node:crypto";
import { spawn } from "node:child_process";
import { chmodSync, mkdirSync, realpathSync, rmSync } from "node:fs";
import { createConnection, createServer, type Socket } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createInterface } from "node:readline";
import { PassThrough } from "node:stream";
import WebSocket from "ws";
import {
  CollabConnection,
  EncryptedCollabConnection,
  createWebSocketTransport,
  stretchShareCode,
  type ConnectionCallbacks,
  type ParticipantProfile,
  type RosterEntry,
} from "@wordinweb/collab/client";
import { AgentDocument } from "./document.js";

if (!globalThis.crypto) Object.defineProperty(globalThis, "crypto", { value: webcrypto });

interface InvitePayload {
  version: 1;
  room: { wsUrl: string; httpBase: string; docId: string; docKey?: string; shareCode?: string };
  invite: { inviteId: string; token: string; chatKey: string; expiresAt: number };
  agent: { name: string; instructions: string; mode?: AgentMode };
}

type AgentMode = "suggest" | "edit";

type WakeTarget = { provider: "codex" | "claude"; sessionId: string };

const AGENT_MODE_CONTROL_PREFIX = "\u0000wordinweb-agent-mode:";
interface Command {
  id?: string | number;
  command?: string;
  request?: Record<string, unknown>;
  category?: string;
  kind?: string;
  text?: string;
  timeoutMs?: number;
}

function base64UrlToBytes(value: string): Uint8Array {
  const padded = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
  return Uint8Array.from(Buffer.from(padded, "base64"));
}

function bytesToBase64Url(value: Uint8Array): string {
  return Buffer.from(value).toString("base64url");
}

async function parseInvite(url: string): Promise<InvitePayload> {
  const parsed = new URL(url);
  const encoded = parsed.hash.match(/(?:^#|&)invite=([A-Za-z0-9_-]+)/)?.[1];
  let payload: InvitePayload;
  if (encoded) {
    payload = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded))) as InvitePayload;
  } else {
    const values = new URLSearchParams(parsed.hash.slice(1));
    const inviteId = values.get("i");
    const secret = values.get("k");
    if (!inviteId || !/^[A-Za-z0-9_-]{16,64}$/.test(inviteId) || !secret) throw new Error("The URL has no AI invitation payload");
    const response = await fetch(new URL(`/agent-invites/${inviteId}`, parsed.origin));
    if (!response.ok) throw new Error(response.status === 404 ? "The AI invitation is unavailable or expired" : `The AI invitation could not be loaded (${response.status})`);
    const stored = await response.json() as { iv: string; ciphertext: string; expiresAt: number };
    if (stored.expiresAt <= Date.now()) throw new Error("The AI invitation has expired");
    const key = await crypto.subtle.importKey("raw", base64UrlToBytes(secret) as BufferSource, "AES-GCM", false, ["decrypt"]);
    const plaintext = await crypto.subtle.decrypt(
      {
        name: "AES-GCM",
        iv: base64UrlToBytes(stored.iv) as BufferSource,
        additionalData: new TextEncoder().encode(`wordinweb-agent-invite:${inviteId}`) as BufferSource,
      },
      key,
      base64UrlToBytes(stored.ciphertext) as BufferSource,
    );
    payload = JSON.parse(new TextDecoder().decode(plaintext)) as InvitePayload;
  }
  if (payload.version !== 1) throw new Error("The AI invitation version is unsupported");
  if (payload.invite.expiresAt <= Date.now()) throw new Error("The AI invitation has expired");
  return payload;
}

function parseAgentModeControl(text: string): AgentMode | null {
  if (!text.startsWith(AGENT_MODE_CONTROL_PREFIX)) return null;
  const mode = text.slice(AGENT_MODE_CONTROL_PREFIX.length);
  return mode === "suggest" || mode === "edit" ? mode : null;
}

function editRequestForMode(request: Record<string, unknown> | undefined, mode: AgentMode): Record<string, unknown> | undefined {
  if (mode === "edit" || !request) return request;
  if (!Array.isArray(request.operations)) return request;
  const operations = request.operations.map((value) => {
    if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("Each edit operation must be an object");
    const operation = value as Record<string, unknown>;
    switch (operation.kind) {
      case "insertText":
      case "splitParagraph":
        return { ...operation, suggest: true };
      case "deleteText":
        return {
          kind: "suggestRevision",
          ranges: [{ blockRef: operation.blockRef, runRef: operation.runRef, start: operation.start, end: operation.end }],
        };
      case "mergeParagraph":
        return { kind: "suggestRevision", marks: [{ blockRef: operation.blockRef, glyph: "del" }] };
      case "suggestRevision":
      case "commentRun":
      case "replyComment":
        return operation;
      default:
        throw new Error(`${String(operation.kind)} is unavailable in suggestion mode. Ask the inviter to switch this agent to editing mode`);
    }
  });
  return { ...request, operations };
}

function randomId(prefix: string): string {
  return `${prefix}_${bytesToBase64Url(crypto.getRandomValues(new Uint8Array(16)))}`;
}

export function detectWakeTarget(env: NodeJS.ProcessEnv = process.env): WakeTarget | null {
  const codexSessionId = env.CODEX_THREAD_ID?.trim();
  if (codexSessionId && /^[A-Za-z0-9_-]{8,128}$/.test(codexSessionId)) {
    return { provider: "codex", sessionId: codexSessionId };
  }
  const claudeSessionId = env.CLAUDE_CODE_SESSION_ID?.trim();
  if (claudeSessionId && /^[A-Za-z0-9_-]{8,128}$/.test(claudeSessionId)) {
    return { provider: "claude", sessionId: claudeSessionId };
  }
  return null;
}

export function agentWakePrompt(agentSessionId: string, text: string): string {
  const sessionCommand = `${JSON.stringify(process.execPath)} ${JSON.stringify(fileURLToPath(import.meta.url))} session ${JSON.stringify(agentSessionId)} '<JSON command>'`;
  return [
    "An authenticated private WordInWeb message started this document-agent turn.",
    `The inviter's message is ${JSON.stringify(text)}.`,
    `The detached document bridge session is ${agentSessionId}. Use this existing bridge for the task.`,
    `Run document commands with: ${sessionCommand}`,
    'Start with sync. For a broad text task, inspect all non-empty stories in one compact call: {"command":"inspect","request":{"kind":"context"}}. Use overview, read, object, or spatial only when the task needs their extra detail.',
    "Inspect the relevant content before every edit. If an edit returns needs_sync, sync and inspect again.",
    "Complete this task, then send the inviter a concise result with the chat command.",
    "End this agent turn after the reply and stay idle. The detached bridge will start the next turn when another message arrives.",
  ].join("\n\n");
}

function write(value: unknown): void {
  process.stdout.write(`${JSON.stringify(value)}\n`);
}

type Emit = (value: unknown) => void;

function sessionSocketPath(sessionId: string): string {
  if (!/^session_[A-Za-z0-9_-]{16,128}$/.test(sessionId)) throw new Error("The agent session ID is invalid");
  return process.platform === "win32"
    ? `\\\\.\\pipe\\wordinweb-agent-${sessionId}`
    : join("/tmp", `wordinweb-agent-${sessionId}.sock`);
}

async function importChatKey(value: string): Promise<CryptoKey> {
  return crypto.subtle.importKey("raw", base64UrlToBytes(value) as BufferSource, "AES-GCM", false, ["encrypt", "decrypt"]);
}

function chatAad(agentClientId: string, messageId: string): Uint8Array {
  return new TextEncoder().encode(`wordinweb-agent-chat:${agentClientId}:${messageId}`);
}

async function encryptChat(key: CryptoKey, agentClientId: string, messageId: string, text: string) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ciphertext = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv: iv as BufferSource, additionalData: chatAad(agentClientId, messageId) as BufferSource },
    key,
    new TextEncoder().encode(text),
  );
  return { iv: bytesToBase64Url(iv), ciphertext: bytesToBase64Url(new Uint8Array(ciphertext)) };
}

async function decryptChat(
  key: CryptoKey,
  agentClientId: string,
  messageId: string,
  iv: string,
  ciphertext: string,
): Promise<string> {
  const plaintext = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: base64UrlToBytes(iv) as BufferSource, additionalData: chatAad(agentClientId, messageId) as BufferSource },
    key,
    base64UrlToBytes(ciphertext) as BufferSource,
  );
  return new TextDecoder().decode(plaintext);
}

export async function connectAgent(
  invitationUrl: string,
  inputStream: NodeJS.ReadableStream = process.stdin,
  emit: Emit = write,
): Promise<void> {
  const payload = await parseInvite(invitationUrl);
  let mode: AgentMode = payload.agent.mode === "edit" ? "edit" : "suggest";
  const clientId = randomId("agent");
  const socket = new WebSocket(payload.room.wsUrl);
  const transport = createWebSocketTransport(socket as never);
  const chatKey = await importChatKey(payload.invite.chatKey);
  let connection: CollabConnection;
  let revision = 0;
  let ready = false;
  const connectionState: { current: "reconnecting" | "live" | "offline" } = { current: "reconnecting" };
  let roster: RosterEntry[] = [];
  let lastActivityIndex = 0;
  const events: unknown[] = [];
  const waiters: Array<(event: unknown) => void> = [];

  const publish = (event: unknown): void => {
    const waiter = waiters.shift();
    if (waiter) waiter(event);
    else events.push(event);
    emit(event);
  };
  const heartbeat = setInterval(() => {
    if (socket.readyState === WebSocket.OPEN) socket.ping();
  }, 25_000);
  socket.addEventListener("close", (event) => {
    connectionState.current = "offline";
    clearInterval(heartbeat);
    publish({ event: "connection_closed", code: event.code, reason: event.reason || undefined });
  });

  let readyResolve: (() => void) | null = null;
  let readyReject: ((error: Error) => void) | null = null;
  const readyPromise = new Promise<void>((resolve, reject) => {
    readyResolve = resolve;
    readyReject = reject;
  });
  const callbacks: ConnectionCallbacks = {
    onChange: () => {
      revision++;
      if (connection?.ready && !ready) {
        ready = true;
        connectionState.current = "live";
        readyResolve?.();
        return;
      }
      if (ready) publish({ event: "document_changed", revision: String(revision) });
    },
    onRoster: (next) => {
      roster = next;
      if (ready) publish({ event: "roster", roster });
    },
    onRefused: (reason) => {
      const error = new Error(`Room connection refused: ${reason}`);
      readyReject?.(error);
      if (ready) publish({ event: "refused", reason });
    },
    onAgentChat: (message) => {
      if (message.agentClientId !== clientId || message.sender !== "inviter") return;
      void decryptChat(chatKey, clientId, message.messageId, message.iv, message.ciphertext).then(
        (text) => {
          const nextMode = parseAgentModeControl(text);
          if (nextMode) {
            mode = nextMode;
            publish({ event: "mode_changed", mode });
          } else publish({ event: "chat", messageId: message.messageId, text });
        },
        () => publish({ event: "chat_error", messageId: message.messageId }),
      );
    },
    onError: ({ where, error }) => publish({ event: "connection_error", where, message: error instanceof Error ? error.message : String(error) }),
  };

  const stretched = payload.room.docKey && payload.room.shareCode
    ? await stretchShareCode(payload.room.shareCode, payload.room.docId)
    : undefined;
  const codeProof = stretched ? Buffer.from(stretched).toString("base64") : undefined;
  connection = payload.room.docKey
    ? new EncryptedCollabConnection(
        transport,
        clientId,
        payload.room.docKey,
        callbacks,
        stretched,
        undefined,
        payload.room.httpBase ? { httpBase: payload.room.httpBase } : undefined,
      ) as unknown as CollabConnection
    : new CollabConnection(
        transport,
        clientId,
        callbacks,
        payload.room.httpBase ? { httpBase: payload.room.httpBase } : undefined,
      );
  const profile: ParticipantProfile = { name: payload.agent.name || "AI Agent", color: "" };
  connection.join(payload.room.docId, undefined, {
    profile,
    codeProof,
    agentInvite: { inviteId: payload.invite.inviteId, token: payload.invite.token },
  });

  const timeout = setTimeout(() => readyReject?.(new Error("The document did not become ready within 30 seconds")), 30_000);
  await readyPromise.finally(() => clearTimeout(timeout));
  const document = AgentDocument.connect({
    getDocument: () => connection.doc,
    getRevision: () => revision,
    allocateIds: (count) => connection.allocIds(count),
    submit: (operation) => connection.submit(operation),
    setPresence: (position) => {
      const deadline = Date.now() + 30_000;
      const send = (): void => {
        if (connection.pendingCount === 0) connection.setPresence(position);
        else if (connectionState.current === "live" && Date.now() < deadline) setTimeout(send, 25);
      };
      send();
    },
    uploadMedia: (bytes) => connection.uploadMedia(bytes),
    getConnectionState: () => connectionState.current,
  }, { provenance: { author: profile.name } });

  emit({
    event: "ready",
    clientId,
    revision: document.revision,
    mode,
    instructions: payload.agent.instructions,
    commands: ["sync", "capabilities", "inspect", "edit", "chat", "wait", "close"],
  });

  const nextEvent = (timeoutMs: number): Promise<unknown> => {
    if (events.length) return Promise.resolve(events.shift());
    return new Promise((resolve) => {
      const timer = setTimeout(() => {
        const index = waiters.indexOf(done);
        if (index >= 0) waiters.splice(index, 1);
        resolve({ event: "wait_timeout" });
      }, timeoutMs);
      const done = (event: unknown) => {
        clearTimeout(timer);
        resolve(event);
      };
      waiters.push(done);
    });
  };

  const respond = (id: Command["id"], result?: unknown, error?: unknown): void => {
    emit(error === undefined ? { id, ok: true, result } : { id, ok: false, error: String(error) });
  };
  const input = createInterface({ input: inputStream, terminal: false });
  for await (const line of input) {
    if (!line.trim()) continue;
    let command: Command;
    try {
      command = JSON.parse(line) as Command;
    } catch {
      respond(undefined, undefined, "Each input line must be valid JSON");
      continue;
    }
    try {
      switch (command.command) {
        case "sync": {
          const activity = connection.activity.slice(lastActivityIndex);
          lastActivityIndex = connection.activity.length;
          respond(command.id, { revision: document.revision, activity, roster, mode });
          break;
        }
        case "capabilities":
          respond(command.id, document.capabilities(command.category as never, command.kind as never));
          break;
        case "inspect":
          respond(command.id, document.inspect(command.request as never));
          break;
        case "edit":
          try {
            if (connectionState.current !== "live") throw new Error("The collaboration connection is offline. Reconnect before editing");
            respond(command.id, await document.edit(editRequestForMode(command.request, mode) as never));
          } catch (error) {
            if (error instanceof Error && error.message.includes("revision is stale")) {
              respond(command.id, { status: "needs_sync", latestRevision: document.revision });
            } else throw error;
          }
          break;
        case "chat": {
          if (connectionState.current !== "live") throw new Error("The collaboration connection is offline. Reconnect before sending chat");
          if (!command.text?.trim()) throw new Error("Chat text is required");
          const messageId = randomId("message");
          const sealed = await encryptChat(chatKey, clientId, messageId, command.text.trim());
          connection.sendAgentChat(clientId, messageId, sealed.iv, sealed.ciphertext);
          respond(command.id, { messageId });
          break;
        }
        case "wait":
          respond(command.id, await nextEvent(Math.min(Math.max(command.timeoutMs ?? 30_000, 1), 55_000)));
          break;
        case "close":
          respond(command.id, { closed: true });
          clearInterval(heartbeat);
          socket.close();
          input.close();
          return;
        default:
          throw new Error("Unknown command");
      }
    } catch (error) {
      respond(command.id, undefined, error instanceof Error ? error.message : String(error));
    }
  }
  clearInterval(heartbeat);
  socket.close();
}

interface PendingCommand {
  socket: Socket;
  originalId?: string | number;
}

function sendSocketValue(socket: Socket, value: unknown): void {
  socket.end(`${JSON.stringify(value)}\n`);
}

async function listen(server: ReturnType<typeof createServer>, path: string): Promise<void> {
  if (process.platform !== "win32") rmSync(path, { force: true });
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(path, () => {
      server.off("error", reject);
      if (process.platform !== "win32") chmodSync(path, 0o600);
      resolve();
    });
  });
}

interface CodexWakeHost {
  startTurn(prompt: string): Promise<void>;
  close(): void;
}

interface AppServerRequest {
  resolve: (result: Record<string, unknown>) => void;
  reject: (error: Error) => void;
  timer: ReturnType<typeof setTimeout>;
}

async function startCodexWakeHost(sessionId: string): Promise<CodexWakeHost> {
  const cwd = join(tmpdir(), `wordinweb-agent-${sessionId}`);
  mkdirSync(cwd, { mode: 0o700 });
  const childEnv = { ...process.env };
  delete childEnv.CODEX_THREAD_ID;
  delete childEnv.CLAUDE_CODE_SESSION_ID;
  delete childEnv.CLAUDECODE;
  const child = spawn("codex", ["app-server"], {
    cwd,
    env: childEnv,
    stdio: ["pipe", "pipe", "ignore"],
  });
  const output = createInterface({ input: child.stdout, terminal: false });
  const requests = new Map<number, AppServerRequest>();
  let requestId = 0;
  let threadId = "";
  let activeTurn: { resolve: () => void; reject: (error: Error) => void } | null = null;
  let closed = false;

  const fail = (error: Error): void => {
    for (const request of requests.values()) {
      clearTimeout(request.timer);
      request.reject(error);
    }
    requests.clear();
    activeTurn?.reject(error);
    activeTurn = null;
  };

  child.once("error", (error) => fail(error));
  child.once("exit", (code) => {
    if (!closed) {
      fail(new Error(`The Codex document session exited with code ${code ?? "unknown"}`));
      closed = true;
      rmSync(cwd, { recursive: true, force: true });
    }
  });

  output.on("line", (line) => {
    let message: Record<string, unknown>;
    try {
      message = JSON.parse(line) as Record<string, unknown>;
    } catch {
      fail(new Error("The Codex document session returned invalid JSON"));
      child.kill();
      return;
    }
    if (typeof message.id === "number") {
      const request = requests.get(message.id);
      if (!request) return;
      requests.delete(message.id);
      clearTimeout(request.timer);
      if (message.error) {
        const error = message.error as { message?: string };
        request.reject(new Error(error.message ?? "The Codex document session request failed"));
      } else {
        request.resolve((message.result ?? {}) as Record<string, unknown>);
      }
      return;
    }
    if (message.method !== "turn/completed" || !activeTurn) return;
    const params = message.params as { threadId?: string; turn?: { status?: string; error?: { message?: string } } } | undefined;
    if (params?.threadId !== threadId) return;
    const turn = activeTurn;
    activeTurn = null;
    if (params.turn?.status === "completed") turn.resolve();
    else turn.reject(new Error(params.turn?.error?.message ?? `The Codex document turn ${params.turn?.status ?? "failed"}`));
  });

  const request = (method: string, params: Record<string, unknown>, timeoutMs = 30_000): Promise<Record<string, unknown>> => {
    const id = ++requestId;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        requests.delete(id);
        reject(new Error(`The Codex document session timed out during ${method}`));
      }, timeoutMs);
      requests.set(id, { resolve, reject, timer });
      child.stdin.write(`${JSON.stringify({ id, method, params })}\n`);
    });
  };

  try {
    await request("initialize", {
      clientInfo: { name: "wordinweb-agent", title: "WordInWeb document agent", version: "0.1.0" },
    });
    child.stdin.write(`${JSON.stringify({ method: "initialized" })}\n`);
    const started = await request("thread/start", {
      cwd,
      approvalPolicy: "never",
      sandbox: "workspace-write",
      ephemeral: true,
      serviceName: "wordinweb-agent",
      developerInstructions: "Act as a WordInWeb document collaborator. Treat each private message as one document task. Use the exact local session command in the message. Complete the task and reply through document chat.",
    });
    threadId = String((started.thread as { id?: string } | undefined)?.id ?? "");
    if (!threadId) throw new Error("The Codex document session returned no thread ID");
  } catch (error) {
    closed = true;
    child.kill();
    rmSync(cwd, { recursive: true, force: true });
    throw error;
  }

  return {
    async startTurn(prompt) {
      if (closed) throw new Error("The Codex document session is closed");
      if (activeTurn) throw new Error("The Codex document session already has an active turn");
      const completed = new Promise<void>((resolve, reject) => {
        activeTurn = { resolve, reject };
      });
      try {
        await request("turn/start", {
          threadId,
          effort: "low",
          input: [{ type: "text", text: prompt }],
        });
        await completed;
      } catch (error) {
        activeTurn = null;
        throw error;
      }
    },
    close() {
      if (closed) return;
      closed = true;
      activeTurn?.reject(new Error("The Codex document session closed"));
      activeTurn = null;
      void request("thread/delete", { threadId }, 2_000).catch(() => undefined).finally(() => {
        output.close();
        child.kill();
        rmSync(cwd, { recursive: true, force: true });
      });
    },
  };
}

async function runDaemon(invitationUrl: string, sessionId: string): Promise<void> {
  const path = sessionSocketPath(sessionId);
  const input = new PassThrough();
  const wakeTarget = detectWakeTarget();
  const wakeQueue: string[] = [];
  const pending = new Map<string, PendingCommand>();
  let readyEvent: Record<string, unknown> | null = null;
  let fatalError: string | null = null;
  let closeTimer: ReturnType<typeof setTimeout> | null = null;
  let wakeChild: ReturnType<typeof spawn> | null = null;
  let codexWakeHost: CodexWakeHost | null = null;
  let codexTurnActive = false;
  let requestCounter = 0;
  let stopped = false;

  const sendWakeFailure = (): void => {
    input.write(`${JSON.stringify({ command: "chat", text: "The agent turn could not start. Run a new invitation command from a Codex or Claude task to reconnect this AI." })}\n`);
  };

  const startNextWake = (): void => {
    if (!wakeTarget || wakeChild || codexTurnActive || wakeQueue.length === 0 || stopped) return;
    const prompt = agentWakePrompt(sessionId, wakeQueue.shift()!);
    if (wakeTarget.provider === "codex") {
      if (!codexWakeHost) return;
      codexTurnActive = true;
      void codexWakeHost.startTurn(prompt).then(
        () => {
          codexTurnActive = false;
          startNextWake();
        },
        () => {
          codexTurnActive = false;
          if (!stopped) sendWakeFailure();
          startNextWake();
        },
      );
      return;
    }
    const childEnv = { ...process.env };
    delete childEnv.CODEX_THREAD_ID;
    delete childEnv.CLAUDE_CODE_SESSION_ID;
    delete childEnv.CLAUDECODE;
    wakeChild = spawn("claude", ["-p", "--resume", wakeTarget.sessionId], { env: childEnv, stdio: ["pipe", "ignore", "ignore"] });
    let finished = false;
    const finish = (ok: boolean): void => {
      if (finished) return;
      finished = true;
      wakeChild = null;
      if (!ok && !stopped) sendWakeFailure();
      startNextWake();
    };
    wakeChild.once("error", () => finish(false));
    wakeChild.once("exit", (code) => finish(code === 0));
    wakeChild.stdin?.on("error", () => undefined);
    wakeChild.stdin?.end(prompt);
  };

  const shutdown = (): void => {
    if (stopped) return;
    stopped = true;
    if (closeTimer) clearTimeout(closeTimer);
    wakeChild?.kill();
    codexWakeHost?.close();
    codexWakeHost = null;
    input.end();
    server.close();
    if (process.platform !== "win32") rmSync(path, { force: true });
  };

  const emit: Emit = (value) => {
    const message = value as Record<string, unknown>;
    if (message.event === "ready") readyEvent = message;
    if (message.event === "connection_closed" && !closeTimer) {
      closeTimer = setTimeout(shutdown, 60_000);
      closeTimer.unref();
    }
    if (message.event === "chat" && typeof message.text === "string" && wakeTarget) {
      wakeQueue.push(message.text);
      startNextWake();
    }
    if (typeof message.id !== "string") return;
    const command = pending.get(message.id);
    if (!command) return;
    pending.delete(message.id);
    const response = { ...message };
    if (command.originalId === undefined) delete response.id;
    else response.id = command.originalId;
    sendSocketValue(command.socket, response);
  };

  const server = createServer((socket) => {
    socket.setEncoding("utf8");
    let buffer = "";
    socket.on("data", (chunk: string) => {
      buffer += chunk;
      const newline = buffer.indexOf("\n");
      if (newline < 0) return;
      socket.removeAllListeners("data");
      try {
        const message = JSON.parse(buffer.slice(0, newline)) as Record<string, unknown>;
        if (message.daemon === "status") {
          if (fatalError) sendSocketValue(socket, { ok: false, error: fatalError });
          else if (readyEvent) sendSocketValue(socket, { ok: true, result: { state: "ready", ready: readyEvent } });
          else sendSocketValue(socket, { ok: true, result: { state: "starting" } });
          return;
        }
        if (!readyEvent) {
          sendSocketValue(socket, { ok: false, error: fatalError ?? "The agent bridge is still starting" });
          return;
        }
        const id = `request_${++requestCounter}`;
        pending.set(id, { socket, originalId: message.id as string | number | undefined });
        input.write(`${JSON.stringify({ ...message, id })}\n`);
      } catch (error) {
        sendSocketValue(socket, { ok: false, error: error instanceof Error ? error.message : String(error) });
      }
    });
  });

  await listen(server, path);
  try {
    if (wakeTarget?.provider === "codex") codexWakeHost = await startCodexWakeHost(sessionId);
    await connectAgent(invitationUrl, input, emit);
  } catch (error) {
    fatalError = error instanceof Error ? error.message : String(error);
    for (const command of pending.values()) sendSocketValue(command.socket, { ok: false, error: fatalError });
    pending.clear();
    await new Promise<void>((resolve) => setTimeout(resolve, 1_000));
  } finally {
    shutdown();
  }
}

function daemonRequest(sessionId: string, request: unknown, timeoutMs = 60_000): Promise<Record<string, unknown>> {
  return new Promise((resolve, reject) => {
    const socket = createConnection(sessionSocketPath(sessionId));
    let buffer = "";
    const fail = (error: Error) => {
      socket.destroy();
      reject(error);
    };
    socket.setEncoding("utf8");
    socket.setTimeout(timeoutMs, () => fail(new Error("The agent bridge command timed out")));
    socket.once("error", fail);
    socket.once("connect", () => socket.write(`${JSON.stringify(request)}\n`));
    socket.on("data", (chunk: string) => {
      buffer += chunk;
      const newline = buffer.indexOf("\n");
      if (newline < 0) return;
      socket.destroy();
      try {
        resolve(JSON.parse(buffer.slice(0, newline)) as Record<string, unknown>);
      } catch (error) {
        reject(error);
      }
    });
  });
}

async function startDaemon(invitationUrl: string): Promise<void> {
  const sessionId = randomId("session");
  const wakeTarget = detectWakeTarget();
  const child = spawn(process.execPath, [fileURLToPath(import.meta.url), "daemon", invitationUrl, sessionId], {
    detached: true,
    stdio: "ignore",
  });
  child.unref();
  const deadline = Date.now() + 35_000;
  while (Date.now() < deadline) {
    try {
      const status = await daemonRequest(sessionId, { daemon: "status" }, 2_000);
      if (!status.ok) throw new Error(String(status.error));
      const result = status.result as { state?: string; ready?: Record<string, unknown> } | undefined;
      if (result?.state === "ready" && result.ready) {
        write({
          ...result.ready,
          sessionId,
          mode: "detached",
          wake: wakeTarget
            ? { state: "armed", provider: wakeTarget.provider }
            : { state: "unavailable", reason: "This command did not run inside a Codex or Claude task" },
          next: `wordinweb-agent session '${sessionId}' '<JSON command>'`,
        });
        return;
      }
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if (code !== "ENOENT" && code !== "ECONNREFUSED") throw error;
    }
    await new Promise<void>((resolve) => setTimeout(resolve, 100));
  }
  throw new Error("The detached agent bridge did not become ready within 35 seconds");
}

async function main(): Promise<void> {
  const [command, first, second] = process.argv.slice(2);
  if (command === "connect" && first) return startDaemon(first);
  if (command === "session" && first && second) {
    const request = JSON.parse(second) as Command;
    const timeoutMs = request.command === "wait"
      ? Math.min(Math.max(request.timeoutMs ?? 30_000, 1), 55_000) + 5_000
      : 60_000;
    write(await daemonRequest(first, request, timeoutMs));
    return;
  }
  if (command === "daemon" && first && second) return runDaemon(first, second);
  process.stderr.write(
    "Usage:\n" +
    "  wordinweb-agent connect '<AI invitation URL>'\n" +
    "  wordinweb-agent session '<session ID>' '<JSON command>'\n",
  );
  process.exitCode = 1;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === realpathSync(process.argv[1])) {
  void main().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
