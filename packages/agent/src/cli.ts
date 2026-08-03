import { webcrypto } from "node:crypto";
import { realpathSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { createInterface } from "node:readline";
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

export async function connectAgent(invitationUrl: string): Promise<void> {
  const payload = parseInvite(invitationUrl);
  const clientId = randomId("agent");
  const socket = new WebSocket(payload.room.wsUrl);
  const transport = createWebSocketTransport(socket as never);
  const chatKey = await importChatKey(payload.invite.chatKey);
  let connection: CollabConnection;
  let revision = 0;
  let ready = false;
  let roster: RosterEntry[] = [];
  let lastActivityIndex = 0;
  const events: unknown[] = [];
  const waiters: Array<(event: unknown) => void> = [];

  const publish = (event: unknown): void => {
    const waiter = waiters.shift();
    if (waiter) waiter(event);
    else events.push(event);
    write(event);
  };

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
    getConnectionState: () => "live",
  }, { provenance: { author: profile.name } });

  write({
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
    write(error === undefined ? { id, ok: true, result } : { id, ok: false, error: String(error) });
  };
  const input = createInterface({ input: process.stdin, terminal: false });
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
            respond(command.id, await document.edit(command.request as never));
          } catch (error) {
            if (error instanceof Error && error.message.includes("revision is stale")) {
              respond(command.id, { status: "needs_sync", latestRevision: document.revision });
            } else throw error;
          }
          break;
        case "chat": {
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
  socket.close();
}

async function main(): Promise<void> {
  const [command, invitationUrl] = process.argv.slice(2);
  if (command !== "connect" || !invitationUrl) {
    process.stderr.write("Usage: wordinweb-agent connect '<AI invitation URL>'\n");
    process.exitCode = 1;
    return;
  }
  await connectAgent(invitationUrl);
}

if (process.argv[1] && fileURLToPath(import.meta.url) === realpathSync(process.argv[1])) {
  void main().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
