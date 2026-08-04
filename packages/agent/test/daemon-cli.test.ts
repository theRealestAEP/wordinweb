import { execFile } from "node:child_process";
import { webcrypto } from "node:crypto";
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { createServer } from "node:http";
import { tmpdir } from "node:os";
import { delimiter, join } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";
import { WebSocket, WebSocketServer } from "ws";
import { collectRevisions, localName, type XmlElement } from "@wordinweb/core";
import { CollabConnection, createWebSocketTransport, type PresencePosition } from "@wordinweb/collab/client";
import { CollabHub } from "../../server/src/hub.js";
import { blankProvider } from "../../server/src/blank.js";
import { attachWebSocketServer } from "../../server/src/ws.js";

const runFile = promisify(execFile);
const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));

async function until(check: () => boolean | Promise<boolean>, label: string): Promise<void> {
  const started = Date.now();
  while (!await check() && Date.now() - started < 5_000) {
    await new Promise<void>((resolve) => setTimeout(resolve, 20));
  }
  if (!await check()) throw new Error(`timeout: ${label}`);
}

function documentText(root: XmlElement): string {
  let text = localName(root.name) === "t" ? root.text : "";
  for (const child of root.children) text += documentText(child);
  return text;
}

async function cliCommand(args: string[], env: NodeJS.ProcessEnv = {}): Promise<Record<string, unknown>> {
  const commandEnv = { ...process.env };
  delete commandEnv.CODEX_THREAD_ID;
  delete commandEnv.CLAUDE_CODE_SESSION_ID;
  Object.assign(commandEnv, env);
  const result = await runFile(process.execPath, [cli, ...args], { timeout: 45_000, env: commandEnv });
  return JSON.parse(result.stdout.trim()) as Record<string, unknown>;
}

describe("detached agent CLI bridge", () => {
  it("keeps the collaboration socket alive between short CLI commands", async () => {
    const fakeBin = await mkdtemp(join(tmpdir(), "wordinweb-wake-"));
    const wakeLog = join(fakeBin, "wake.json");
    const fakeCodex = join(fakeBin, "codex");
    await writeFile(fakeCodex, `#!/usr/bin/env node
import { appendFileSync } from "node:fs";
const log = (value) => appendFileSync(process.env.WORDINWEB_WAKE_LOG, JSON.stringify(value) + "\\n");
const reply = (value) => process.stdout.write(JSON.stringify(value) + "\\n");
let turn = 0;
log({ event: "start", args: process.argv.slice(2) });
let buffer = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
  buffer += chunk;
  let newline;
  while ((newline = buffer.indexOf("\\n")) >= 0) {
    const line = buffer.slice(0, newline);
    buffer = buffer.slice(newline + 1);
    if (!line) continue;
    const message = JSON.parse(line);
    log(message);
    if (message.method === "initialize") reply({ id: message.id, result: { userAgent: "fake-codex" } });
    if (message.method === "thread/start") reply({ id: message.id, result: { thread: { id: "thread_document_agent" } } });
    if (message.method === "turn/start") {
      const turnId = "turn_" + ++turn;
      reply({ id: message.id, result: { turn: { id: turnId, status: "inProgress", items: [] } } });
      reply({ method: "turn/completed", params: { threadId: message.params.threadId, turn: { id: turnId, status: "completed", items: [] } } });
    }
    if (message.method === "thread/delete") reply({ id: message.id, result: {} });
  }
});
`);
    await chmod(fakeCodex, 0o755);
    const wakeEvents = async (): Promise<Array<Record<string, unknown>>> => {
      if (!existsSync(wakeLog)) return [];
      return (await readFile(wakeLog, "utf8")).trim().split("\n").filter(Boolean).map((line) => JSON.parse(line) as Record<string, unknown>);
    };
    const hub = new CollabHub(blankProvider);
    let storedInvite = "";
    const httpServer = createServer((request, response) => {
      if (request.url === "/agent-invites/link_1234567890123456" && storedInvite) {
        response.writeHead(200, { "content-type": "application/json" }).end(storedInvite);
      } else response.writeHead(404).end();
    });
    const wss = new WebSocketServer({ noServer: true });
    attachWebSocketServer(wss, hub);
    httpServer.on("upgrade", (request, socket, head) => {
      wss.handleUpgrade(request, socket, head, (client) => wss.emit("connection", client, request));
    });
    await new Promise<void>((resolve) => httpServer.listen(0, resolve));
    const port = (httpServer.address() as { port: number }).port;
    const wsUrl = `ws://127.0.0.1:${port}`;
    const humanSocket = new WebSocket(wsUrl);
    let inviteRegistered = false;
    let agentClientId = "";
    let agentPresence: PresencePosition | null = null;
    const human = new CollabConnection(createWebSocketTransport(humanSocket as never), "human", {
      onAgentInviteRegistered: () => { inviteRegistered = true; },
      onAgentConnected: (agent) => { agentClientId = agent.agentClientId; },
      onPresence: (participant, position) => {
        if (participant === agentClientId) agentPresence = position;
      },
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
      agent: { name: "Detached agent", instructions: "Wait for private chat tasks.", mode: "suggest" },
    };
    const inviteSecret = Buffer.alloc(32, 9);
    const inviteIv = Buffer.alloc(12, 8);
    const inviteKey = await webcrypto.subtle.importKey("raw", inviteSecret, "AES-GCM", false, ["encrypt"]);
    const sealedInvite = await webcrypto.subtle.encrypt(
      {
        name: "AES-GCM",
        iv: inviteIv,
        additionalData: new TextEncoder().encode("wordinweb-agent-invite:link_1234567890123456"),
      },
      inviteKey,
      new TextEncoder().encode(JSON.stringify(payload)),
    );
    storedInvite = JSON.stringify({
      iv: inviteIv.toString("base64url"),
      ciphertext: Buffer.from(sealedInvite).toString("base64url"),
      expiresAt: payload.invite.expiresAt,
    });
    const invitationUrl = `http://127.0.0.1:${port}/agent-invite#i=link_1234567890123456&k=${inviteSecret.toString("base64url")}`;
    let sessionId = "";
    try {
      const connected = await cliCommand(["connect", invitationUrl], {
        CODEX_THREAD_ID: "thread_12345678",
        PATH: `${fakeBin}${delimiter}${process.env.PATH}`,
        WORDINWEB_WAKE_LOG: wakeLog,
      });
      sessionId = String(connected.sessionId);
      expect(connected).toMatchObject({ event: "ready", mode: "detached", wake: { state: "armed", provider: "codex" } });
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
      await until(async () => (await wakeEvents()).some((event) => event.method === "turn/start"), "Codex session wakes");
      const events = await wakeEvents();
      expect(events.find((event) => event.event === "start")).toMatchObject({ args: ["app-server"] });
      expect(events.filter((event) => event.method === "thread/start")).toHaveLength(1);
      const wake = events.find((event) => event.method === "turn/start") as {
        params: { threadId: string; effort: string; input: Array<{ text: string }> };
      };
      const prompt = wake.params.input[0].text;
      expect(wake.params.threadId).toBe("thread_document_agent");
      expect(wake.params.effort).toBe("low");
      expect(prompt).toContain("Please make the edit.");
      expect(prompt).toContain(sessionId);
      expect(prompt).toContain(process.execPath);
      expect(prompt).toContain(cli);
      expect(prompt).toContain("stay idle");
      expect(prompt).toContain('{"kind":"context"}');

      const nextMessageId = "message_abcdefgh";
      const nextIv = Buffer.alloc(12, 6);
      const nextCiphertext = await webcrypto.subtle.encrypt(
        {
          name: "AES-GCM",
          iv: nextIv,
          additionalData: new TextEncoder().encode(`wordinweb-agent-chat:${agentClientId}:${nextMessageId}`),
        },
        chatKey,
        new TextEncoder().encode("Please make another edit."),
      );
      human.sendAgentChat(agentClientId, nextMessageId, nextIv.toString("base64url"), Buffer.from(nextCiphertext).toString("base64url"));
      await until(async () => (await wakeEvents()).filter((event) => event.method === "turn/start").length === 2, "resident Codex session handles another message");
      const residentEvents = await wakeEvents();
      expect(residentEvents.filter((event) => event.event === "start")).toHaveLength(1);
      expect(residentEvents.filter((event) => event.method === "thread/start")).toHaveLength(1);
      const turns = residentEvents.filter((event) => event.method === "turn/start") as Array<{ params: { threadId: string } }>;
      expect(turns.map((turn) => turn.params.threadId)).toEqual(["thread_document_agent", "thread_document_agent"]);
      expect(await cliCommand(["session", sessionId, JSON.stringify({ command: "wait", timeoutMs: 2_000 })]))
        .toMatchObject({ ok: true, result: { event: "chat", text: "Please make another edit." } });

      const synced = await cliCommand(["session", sessionId, JSON.stringify({ command: "sync" })]);
      const revision = String((synced.result as { revision: string }).revision);
      const inspected = await cliCommand(["session", sessionId, JSON.stringify({
        command: "inspect",
        request: { kind: "context", includeEmpty: true },
      })]);
      const first = (inspected.result as {
        contents: Array<{ blocks: Array<{ ref: string; runs: Array<{ ref: string }> }> }>;
      }).contents[0].blocks[0];
      expect(first.runs[0]).not.toHaveProperty("formatting");
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
      await until(() => agentPresence !== null, "human receives agent cursor");
      expect(agentPresence?.anchor.offset).toBe("Detached edit".length);
      expect(collectRevisions(human.doc!)).toHaveLength(1);

      const modeMessageId = "message_mode_edit";
      const modeIv = Buffer.alloc(12, 4);
      const modeCiphertext = await webcrypto.subtle.encrypt(
        {
          name: "AES-GCM",
          iv: modeIv,
          additionalData: new TextEncoder().encode(`wordinweb-agent-chat:${agentClientId}:${modeMessageId}`),
        },
        chatKey,
        new TextEncoder().encode("\u0000wordinweb-agent-mode:edit"),
      );
      human.sendAgentChat(agentClientId, modeMessageId, modeIv.toString("base64url"), Buffer.from(modeCiphertext).toString("base64url"));
      let modeEvent: Record<string, unknown> = {};
      for (let attempt = 0; attempt < 3; attempt++) {
        modeEvent = await cliCommand(["session", sessionId, JSON.stringify({ command: "wait", timeoutMs: 2_000 })]);
        if ((modeEvent.result as { event?: string } | undefined)?.event === "mode_changed") break;
      }
      expect(modeEvent).toMatchObject({ ok: true, result: { event: "mode_changed", mode: "edit" } });

      const editSync = await cliCommand(["session", sessionId, JSON.stringify({ command: "sync" })]);
      const editRevision = String((editSync.result as { revision: string }).revision);
      expect(editSync).toMatchObject({ ok: true, result: { mode: "edit" } });
      const editInspect = await cliCommand(["session", sessionId, JSON.stringify({ command: "inspect", request: { kind: "read" } })]);
      const editFirst = (editInspect.result as { blocks: Array<{ ref: string; runs: Array<{ ref: string }> }> }).blocks[0];
      await cliCommand(["session", sessionId, JSON.stringify({
        command: "edit",
        request: {
          revision: editRevision,
          operations: [{ kind: "insertText", at: { blockRef: editFirst.ref, runRef: editFirst.runs[0].ref, offset: 0 }, text: "Direct edit" }],
        },
      })]);
      await until(() => documentText(human.doc!.docRoot).includes("Direct edit"), "human receives direct edit");
      expect(collectRevisions(human.doc!)).toHaveLength(1);
    } finally {
      if (sessionId) {
        await cliCommand(["session", sessionId, JSON.stringify({ command: "close" })]).catch(() => undefined);
        await until(async () => (await wakeEvents()).some((event) => event.method === "thread/delete"), "Codex session closes").catch(() => undefined);
      }
      humanSocket.close();
      wss.close();
      httpServer.close();
      await rm(fakeBin, { recursive: true, force: true });
    }
  }, 20_000);
});
