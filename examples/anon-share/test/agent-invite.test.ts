import { afterEach, describe, expect, it, vi } from "vitest";
import {
  agentInviteClipboardText,
  agentInviteUrl,
  createShortAgentInvite,
  decodeAgentInvite,
  decryptAgentChat,
  encryptAgentChat,
  randomAgentToken,
  type AgentInvitePayload,
} from "../src/agent-invite";

afterEach(() => vi.unstubAllGlobals());

describe("AI invitation helpers", () => {
  it("keeps the room code, document key, and instructions in the URL fragment", () => {
    const payload: AgentInvitePayload = {
      version: 1,
      room: {
        wsUrl: "wss://docs.example.test",
        httpBase: "https://docs.example.test",
        docId: "d_123",
        docKey: "document-key",
        shareCode: "482913",
      },
      invite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
        chatKey: randomAgentToken("", 32).slice(1),
        expiresAt: Date.now() + 60_000,
      },
      agent: { name: "Review agent", instructions: "Check the totals before editing." },
    };
    const url = agentInviteUrl("https://docs.example.test", payload);
    expect(new URL(url).pathname).toBe("/agent-invite");
    expect(new URL(url).search).toBe("");
    expect(decodeAgentInvite(url)).toEqual(payload);
  });

  it("copies the invitation with connection and private-chat instructions", () => {
    const url = "https://docs.example.test/agent-invite#i=short-link&k=secret";
    const copied = agentInviteClipboardText(url);

    expect(copied).toContain(url);
    expect(copied).toContain("Run this exact command in the terminal");
    expect(copied).toContain("Pass the invitation URL directly");
    expect(copied).toContain("short-invite-5");
    expect(copied).toContain("dedicated document-agent session");
    expect(copied).toContain("Stay idle after this turn");
  });

  it("stores an encrypted invite behind a short URL", async () => {
    const payload: AgentInvitePayload = {
      version: 1,
      room: { wsUrl: "wss://docs.example.test", httpBase: "https://docs.example.test", docId: "private-doc", docKey: "private-key" },
      invite: {
        inviteId: "invite_1234567890",
        token: "token_12345678901234567890123456789012",
        chatKey: randomAgentToken("", 32).slice(1),
        expiresAt: Date.now() + 60_000,
      },
      agent: { name: "Review agent", instructions: "Wait for tasks.", mode: "suggest" },
    };
    let stored = "";
    vi.stubGlobal("fetch", async (_url: URL, init: RequestInit) => {
      stored = String(init.body);
      return { ok: true, status: 201 };
    });

    const url = await createShortAgentInvite("https://docs.example.test", payload);

    expect(url.length).toBeLessThan(180);
    expect(url).toContain("#i=link_");
    expect(url).not.toContain("private-doc");
    expect(stored).not.toContain("private-key");
  });

  it("encrypts private chat for one invited agent", async () => {
    const key = randomAgentToken("", 32).slice(1);
    const sealed = await encryptAgentChat(key, "agent-one", "message-one", "Please review the new paragraph.");
    await expect(decryptAgentChat(key, "agent-one", "message-one", sealed.iv, sealed.ciphertext))
      .resolves.toBe("Please review the new paragraph.");
    await expect(decryptAgentChat(key, "agent-two", "message-one", sealed.iv, sealed.ciphertext))
      .rejects.toBeDefined();
  });
});
