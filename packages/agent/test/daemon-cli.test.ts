import { execFile } from "node:child_process";
import { webcrypto } from "node:crypto";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";
import { WebSocket, WebSocketServer } from "ws";
import { localName, type XmlElement } from "@wordinweb/core";
import { CollabConnection, createWebSocketTransport } from "@wordinweb/collab/client";
import { CollabHub } from "../../server/src/hub.js";
import { blankProvider } from "../../server/src/blank.js";
import { attachWebSocketServer } from "../../server/src/ws.js";

const runFile = promisify(execFile);
const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));

async function until(check: () => boolean, label: string): Promise<void> {
  const started = Date.now();
  while (!check() && Date.now() - started < 5_000) {
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  if (!check()) throw new Error(`timeout: ${label}`);
}

function documentText(root: XmlElement): string {
  let text = localName(root.name) === "t" ? root.text : "";
  for (const child of root.children) text += documentText(child);
  return text;
}

async function cliCommand(args: string[]): Promise<Record<string, unknown>> {
  const result = await runFile(process.execPath, [cli, ...args], { timeout: 45_000 });
  return JSON.parse(result.stdout.trim()) as Record<string, unknown>;
}

describe("detached agent CLI bridge", () => {
  it("keeps the collaboration socket alive between short CLI commands", async () => {
    const hub = new CollabHub(blankProvider);
    const wss = new WebSocketServer({ port: 0 });
    attachWebSocketServer(wss, hub);
    await new Promise<void>((resolve) => wss.on("listening", () => resolve()));
    const port = (wss.address() as { port: number }).port;
    const wsUrl = `ws://127.0.0.1:${port}`;
    const humanSocket = new WebSocket(wsUrl);
    let inviteRegistered = false;
    let agentClientId = "";
    const human = new CollabConnection(createWebSocketTransport(humanSocket as never), "human", {
      onAgentInviteRegistered: () => { inviteRegistered = true; },
      onAgentConnected: (agent) => { agentClientId = agent.agentClientId; },
    });
    human.join("daemon-test", undefined, { profile: { name: "Human", color: "" } });
    await until(() => human.ready, "human ready");

    const inviteId = "invite_1234567890";
    const token = "token_12345678901234567890123456789012";
    const chatKeyBytes = Buffer.alloc(32, 7);
    human.registerAgentInvite(inviteId, token, Date.now() + 120_000);
    await until(() => inviteRegistered, "invite registered");
    const payload = {
      version: 1,
      room: { wsUrl, httpBase: `http://127.0.0.1:${port}`, docId: "daemon-test" },
      invite: {
        inviteId,
        token,
        chatKey: chatKeyBytes.toString("base64url"),
        expiresAt: Date.now() + 120_000,
      },
      agent: { name: "Detached agent", instructions: "Wait for private chat tasks." },
    };
    const invitationUrl = `https://example.test/agent-invite#invite=${Buffer.from(JSON.stringify(payload)).toString("base64url")}`;
    let sessionId = "";
    try {
      const connected = await cliCommand(["connect", invitationUrl]);
      sessionId = String(connected.sessionId);
      expect(connected).toMatchObject({ event: "ready", mode: "detached" });
      expect(sessionId).toMatch(/^session_/);
      await until(() => !!agentClientId, "agent joined");

      const messageId = "message_12345678";
      const iv = Buffer.alloc(12, 3);
      const chatKey = await webcrypto.subtle.importKey("raw", chatKeyBytes, "AES-GCM", false, ["encrypt"]);
      const ciphertext = await webcrypto.subtle.encrypt(
        {
          name: "AES-GCM",
          iv,
          additionalData: new TextEncoder().encode(`wordinweb-agent-chat:${agentClientId}:${messageId}`),
        },
        chatKey,
        new TextEncoder().encode("Please make the edit."),
      );
      human.sendAgentChat(agentClientId, messageId, iv.toString("base64url"), Buffer.from(ciphertext).toString("base64url"));
      let waited: Record<string, unknown> = {};
      for (let attempt = 0; attempt < 3; attempt++) {
        waited = await cliCommand(["session", sessionId, JSON.stringify({ command: "wait", timeoutMs: 2_000 })]);
        if ((waited.result as { event?: string } | undefined)?.event === "chat") break;
      }
      expect(waited).toMatchObject({ ok: true, result: { event: "chat", text: "Please make the edit." } });

      const synced = await cliCommand(["session", sessionId, JSON.stringify({ command: "sync" })]);
      const revision = String((synced.result as { revision: string }).revision);
      const inspected = await cliCommand(["session", sessionId, JSON.stringify({
        command: "inspect",
        request: { kind: "read" },
      })]);
      const first = (inspected.result as {
        blocks: Array<{ ref: string; runs: Array<{ ref: string }> }>;
      }).blocks[0];
      const edited = await cliCommand(["session", sessionId, JSON.stringify({
        command: "edit",
        request: {
          revision,
          operations: [{
            kind: "insertText",
            at: { blockRef: first.ref, runRef: first.runs[0].ref, offset: 0 },
            text: "Detached edit",
          }],
        },
      })]);
      expect(edited).toMatchObject({ ok: true, result: { status: "submitted" } });
      await until(() => documentText(human.doc!.docRoot).includes("Detached edit"), "human receives edit");
    } finally {
      if (sessionId) {
        await cliCommand(["session", sessionId, JSON.stringify({ command: "close" })]).catch(() => undefined);
      }
      humanSocket.close();
      wss.close();
    }
  }, 20_000);
});
