import { webcrypto } from "node:crypto";
import { spawn } from "node:child_process";
import { chmodSync, realpathSync, rmSync } from "node:fs";
import { createConnection, createServer, type Socket } from "node:net";
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
  agent: { name: string; instructions: string };
}

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

function parseInvite(url: string): InvitePayload {
  const encoded = new URL(url).hash.match(/(?:^#|&)invite=([A-Za-z0-9_-]+)/)?.[1];
  if (!encoded) throw new Error("The URL has no AI invitation payload");
  const payload = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded))) as InvitePayload;
  if (payload.version !== 1) throw new Error("The AI invitation version is unsupported");
  if (payload.invite.expiresAt <= Date.now()) throw new Error("The AI invitation has expired");
  return payload;
}

function randomId(prefix: string): string {
  return `${prefix}_${bytesToBase64Url(crypto.getRandomValues(new Uint8Array(16)))}`;
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
  const payload = parseInvite(invitationUrl);
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
        (text) => publish({ event: "chat", messageId: message.messageId, text }),
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
    uploadMedia: (bytes) => connection.uploadMedia(bytes),
    getConnectionState: () => connectionState.current,
  }, { provenance: { author: profile.name } });

  emit({
    event: "ready",
    clientId,
    revision: document.revision,
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
          respond(command.id, { revision: document.revision, activity, roster });
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
            respond(command.id, await document.edit(command.request as never));
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

async function runDaemon(invitationUrl: string, sessionId: string): Promise<void> {
  const path = sessionSocketPath(sessionId);
  const input = new PassThrough();
  const pending = new Map<string, PendingCommand>();
  let readyEvent: Record<string, unknown> | null = null;
  let fatalError: string | null = null;
  let closeTimer: ReturnType<typeof setTimeout> | null = null;
  let requestCounter = 0;
  let stopped = false;

  const shutdown = (): void => {
    if (stopped) return;
    stopped = true;
    if (closeTimer) clearTimeout(closeTimer);
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
  parseInvite(invitationUrl);
  const sessionId = randomId("session");
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
