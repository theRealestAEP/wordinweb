import { CollabHub } from "./hub.js";
import { attachWebSocketServer, WsServer } from "./ws.js";
import { InMemoryStorage } from "./storage.js";
import { blankProvider } from "./blank.js";

/**
 * Zero-config dev server (plan Tier 1: `npx wordinweb server`). Uses blank-doc
 * provider + in-memory (ephemeral) storage and a real `ws` WebSocketServer,
 * dynamically imported so the package carries no `ws` build dependency — the
 * operator installs it (`npm i ws`). Auth is OFF here; production embeds
 * CollabHub with a real storage/JWT/media configuration.
 */
export async function startDevServer(opts: { port?: number } = {}): Promise<{ close: () => void }> {
  const port = opts.port ?? 1234;
  // Dynamic import via an indirect specifier keeps `ws` optional at build/
  // test time (it is never statically resolved).
  const spec = "ws";
  const wsMod = (await import(spec).catch(() => {
    throw new Error("startDevServer needs the 'ws' package — run: npm i ws");
  })) as { WebSocketServer: new (o: { port: number }) => WsServer & { close(): void } };

  const hub = new CollabHub(blankProvider, new InMemoryStorage());
  const wss = new wsMod.WebSocketServer({ port });
  attachWebSocketServer(wss, hub);
  // eslint-disable-next-line no-console
  console.log(`wordinweb collab dev server on ws://localhost:${port} (ephemeral: data is not durable)`);
  return { close: () => wss.close() };
}

/**
 * The zero-custody demo server (plan doc 12): one port serving both halves —
 *
 *   HTTP  POST /docs, PUT /docs/:docId   (go-live / bring-it-back, seed-http.ts)
 *   WS    upgrade on the same listener   (the collab protocol)
 *
 * No provider (unknown docIds are `no-session`), no storage driver (nothing
 * at rest, ever), rooms evicted on a periodic sweep after the last-disconnect
 * grace. Restart-safe by design: a crash loses only live rooms, and any
 * bundle-holding browser re-seeds (doc 12 §2 "the browsers ARE the recovery
 * machinery") — which is why this process needs no volumes and no drain.
 */
export async function startZeroCustodyServer(opts: { port?: number } = {}): Promise<{ close: () => void }> {
  const port = opts.port ?? 1234;
  const spec = "ws";
  const wsMod = (await import(spec).catch(() => {
    throw new Error("startZeroCustodyServer needs the 'ws' package — run: npm i ws");
  })) as {
    WebSocketServer: new (o: { noServer: true }) => WsServer & {
      close(): void;
      handleUpgrade(req: unknown, socket: unknown, head: unknown, cb: (client: unknown) => void): void;
      emit(ev: string, ...args: unknown[]): void;
    };
  };
  const http = await import("node:http");
  const { handleSeedRequest } = await import("./seed-http.js");
  const { blankDocxBytes } = await import("./blank.js");

  const hub = new CollabHub(/*provider*/ null); // zero-custody: seed-only rooms
  const wss = new wsMod.WebSocketServer({ noServer: true });
  attachWebSocketServer(wss, hub);

  const server = http.createServer((req, res) => {
    const url = req.url ?? "";
    if (req.method === "GET" && url === "/blank") {
      // Template bytes for client-side sealing (encrypted go-live, doc 13):
      // the server CANNOT seal a blank doc itself — it has no keys — so the
      // creating browser fetches the template, seals, and POSTs the blob.
      res.writeHead(200, { "content-type": "application/octet-stream" }).end(Buffer.from(blankDocxBytes()));
      return;
    }
    // Media relay (doc 16 §3): bytes over HTTP, never the WS sequencer.
    const media = /^\/docs\/([^/]+)\/media\/([0-9a-f]{64})$/.exec(url);
    if (media) {
      const [, mDoc, mSha] = media;
      if (req.method === "GET") {
        const bytes = hub.mediaDownload(decodeURIComponent(mDoc), mSha);
        if (!bytes) res.writeHead(404, { "content-type": "application/json" }).end(`{"state":"absent"}`);
        else res.writeHead(200, { "content-type": "application/octet-stream" }).end(Buffer.from(bytes));
        return;
      }
      if (req.method === "PUT") {
        const chunks: Buffer[] = [];
        let size = 0;
        req.on("data", (c: Buffer) => {
          size += c.length;
          if (size > 12 * 1024 * 1024) req.destroy();
          else chunks.push(c);
        });
        req.on("end", () => {
          void hub.mediaUpload(decodeURIComponent(mDoc), mSha, new Uint8Array(Buffer.concat(chunks))).then((status) => {
            res.writeHead(status, { "content-type": "application/json" }).end(JSON.stringify({ status }));
          });
        });
        return;
      }
      res.writeHead(405).end();
      return;
    }
    const put = req.method === "PUT" && /^\/docs\/[^/]+$/.test(url);
    const post = req.method === "POST" && url === "/docs";
    if (!put && !post) {
      res.writeHead(404, { "content-type": "application/json" }).end(`{"error":"not-found"}`);
      return;
    }
    const chunks: Buffer[] = [];
    let size = 0;
    req.on("data", (c: Buffer) => {
      size += c.length;
      if (size > 16 * 1024 * 1024) req.destroy(); // wire-level cap above the decoded cap
      else chunks.push(c);
    });
    req.on("end", () => {
      let body: { docx?: string } = {};
      try {
        body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
      } catch {
        res.writeHead(400, { "content-type": "application/json" }).end(`{"error":"bad-json"}`);
        return;
      }
      const out = handleSeedRequest(hub, {
        method: put ? "PUT" : "POST",
        docId: put ? decodeURIComponent(url.slice("/docs/".length)) : undefined,
        body,
      });
      res.writeHead(out.status, { "content-type": "application/json" }).end(JSON.stringify(out.body));
    });
  });
  server.on("upgrade", (req, socket, head) => {
    wss.handleUpgrade(req, socket, head, (client) => wss.emit("connection", client, req));
  });
  // Eviction + token sweeps: coarse cadence is fine — grace is 60s.
  const sweeper = setInterval(() => {
    hub.sweepRooms();
    hub.sweepExpired();
    hub.sweepMedia();
  }, 10_000);
  server.listen(port);
  // eslint-disable-next-line no-console
  console.log(
    `wordinweb zero-custody server on :${port} — storage: none (sessions are not persisted; browsers hold the documents)`,
  );
  return {
    close: () => {
      clearInterval(sweeper);
      wss.close();
      server.close();
    },
  };
}

// Allow `node dist/cli.js` to run it directly. ZERO_CUSTODY=1 selects the
// seed-only server; default remains the blank-doc dev server.
if (typeof process !== "undefined" && process.argv[1] && process.argv[1].endsWith("cli.js")) {
  const port = Number(process.env.PORT ?? 1234);
  if (process.env.ZERO_CUSTODY === "1") void startZeroCustodyServer({ port });
  else void startDevServer({ port });
}
