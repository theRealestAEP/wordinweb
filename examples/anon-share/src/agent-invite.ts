export interface AgentInvitePayload {
  version: 1;
  room: {
    wsUrl: string;
    httpBase: string;
    docId: string;
    docKey?: string;
    shareCode?: string;
  };
  invite: {
    inviteId: string;
    token: string;
    chatKey: string;
    expiresAt: number;
  };
  agent: {
    name: string;
    instructions: string;
    mode?: AgentMode;
  };
}

export type AgentMode = "suggest" | "edit";

export const AGENT_MODE_CONTROL_PREFIX = "\u0000wordinweb-agent-mode:";

export const AGENT_CONNECTION_INSTRUCTIONS =
  "Keep this agent task open until the inviter tells you to disconnect. " +
  "The connect command starts a detached local bridge and returns a sessionId. Use a new short session command for each sync, inspect, edit, chat, or wait request. " +
  "The agent starts in suggestion mode. Follow later mode_changed events from the inviter. " +
  "Treat new private chat messages as tasks. Sync and inspect the relevant content before every edit. " +
  "After each task, call wait again instead of sending a final response. Close the bridge when the inviter ends the collaboration.";

function bytesToBase64Url(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function base64UrlToBytes(value: string): Uint8Array {
  const padded = value.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(value.length / 4) * 4, "=");
  const binary = atob(padded);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

export function randomAgentToken(prefix: string, byteLength: number): string {
  const bytes = new Uint8Array(byteLength);
  crypto.getRandomValues(bytes);
  return `${prefix}_${bytesToBase64Url(bytes)}`;
}

export function encodeAgentInvite(payload: AgentInvitePayload): string {
  return bytesToBase64Url(new TextEncoder().encode(JSON.stringify(payload)));
}

export function decodeAgentInvite(url: string): AgentInvitePayload {
  const encoded = new URL(url).hash.match(/(?:^#|&)invite=([A-Za-z0-9_-]+)/)?.[1];
  if (!encoded) throw new Error("The URL has no AI invitation payload");
  const payload = JSON.parse(new TextDecoder().decode(base64UrlToBytes(encoded))) as AgentInvitePayload;
  if (payload.version !== 1) throw new Error("The AI invitation version is unsupported");
  return payload;
}

export function agentInviteUrl(origin: string, payload: AgentInvitePayload): string {
  const url = new URL("/agent-invite", origin);
  url.hash = `invite=${encodeAgentInvite(payload)}`;
  return url.toString();
}

async function inviteKey(value: string): Promise<CryptoKey> {
  return crypto.subtle.importKey("raw", base64UrlToBytes(value) as BufferSource, "AES-GCM", false, ["encrypt"]);
}

export async function createShortAgentInvite(origin: string, payload: AgentInvitePayload): Promise<string> {
  const inviteId = randomAgentToken("link", 12);
  const secret = randomAgentToken("key", 32).slice(4);
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const ciphertext = await crypto.subtle.encrypt(
    {
      name: "AES-GCM",
      iv: iv as BufferSource,
      additionalData: new TextEncoder().encode(`wordinweb-agent-invite:${inviteId}`) as BufferSource,
    },
    await inviteKey(secret),
    new TextEncoder().encode(JSON.stringify(payload)),
  );
  const response = await fetch(new URL(`/agent-invites/${inviteId}`, origin), {
    method: "PUT",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      iv: bytesToBase64Url(iv),
      ciphertext: bytesToBase64Url(new Uint8Array(ciphertext)),
      expiresAt: payload.invite.expiresAt,
    }),
  });
  if (!response.ok) throw new Error(`The short AI invitation could not be created (${response.status})`);
  const url = new URL("/agent-invite", origin);
  url.hash = new URLSearchParams({ i: inviteId, k: secret }).toString();
  return url.toString();
}

export function agentInviteClipboardText(invitationUrl: string): string {
  return [
    "Join my WordInWeb document as an AI collaborator.",
    "Run this exact command now. Do not open the invitation in a browser and do not search the workspace for instructions:",
    `npx -y --package='https://collab.word-in-web.com/wordinweb-agent.tgz?v=short-invite-1' wordinweb-agent connect '${invitationUrl}'`,
    "The connect command starts a detached local bridge and returns a sessionId. It does not require a persistent terminal tool call.",
    "Use the returned sessionId with a new short `wordinweb-agent session` command for every sync, inspect, edit, chat, and wait request.",
    "After you connect, wait for my private chat messages in the document.",
    "Keep this agent task open. After each request, call wait again. Do not send a final response until I tell you to disconnect.",
  ].join("\n\n");
}

export function agentModeControl(mode: AgentMode): string {
  return `${AGENT_MODE_CONTROL_PREFIX}${mode}`;
}

export function parseAgentModeControl(text: string): AgentMode | null {
  if (!text.startsWith(AGENT_MODE_CONTROL_PREFIX)) return null;
  const mode = text.slice(AGENT_MODE_CONTROL_PREFIX.length);
  return mode === "suggest" || mode === "edit" ? mode : null;
}

async function chatKey(value: string): Promise<CryptoKey> {
  return crypto.subtle.importKey("raw", base64UrlToBytes(value) as BufferSource, "AES-GCM", false, ["encrypt", "decrypt"]);
}

function chatAad(agentClientId: string, messageId: string): Uint8Array {
  return new TextEncoder().encode(`wordinweb-agent-chat:${agentClientId}:${messageId}`);
}

export async function encryptAgentChat(
  key: string,
  agentClientId: string,
  messageId: string,
  text: string,
): Promise<{ iv: string; ciphertext: string }> {
  const iv = new Uint8Array(12);
  crypto.getRandomValues(iv);
  const ciphertext = await crypto.subtle.encrypt(
    { name: "AES-GCM", iv: iv as BufferSource, additionalData: chatAad(agentClientId, messageId) as BufferSource },
    await chatKey(key),
    new TextEncoder().encode(text),
  );
  return { iv: bytesToBase64Url(iv), ciphertext: bytesToBase64Url(new Uint8Array(ciphertext)) };
}

export async function decryptAgentChat(
  key: string,
  agentClientId: string,
  messageId: string,
  iv: string,
  ciphertext: string,
): Promise<string> {
  const plaintext = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: base64UrlToBytes(iv) as BufferSource, additionalData: chatAad(agentClientId, messageId) as BufferSource },
    await chatKey(key),
    base64UrlToBytes(ciphertext) as BufferSource,
  );
  return new TextDecoder().decode(plaintext);
}
