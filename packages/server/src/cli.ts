import { CollabHub } from "./hub.js";
import { attachWebSocketServer, WsServer } from "./ws.js";
import { InMemoryStorage } from "./storage.js";
import { blankProvider } from "./blank.js";
import { createObservability, observabilityEnabled } from "./observability.js";
import { envLimits, clientIp, mediaWireCap } from "./limits.js";
import { IpGuard } from "./ip-guard.js";

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
/**
 * Where log lines go. Default STDERR — one JSON object per line, which any
 * collector (vector, fluent-bit, journald, `docker logs`) already knows how
 * to tail, with no vendor coupling anywhere in this codebase.
 *
 * WW_LOG_FILE appends to a path instead, for the zero-dependency case where
 * something ships files rather than streams. DELIBERATELY DUMB: no rotation,
 * no size cap, no reopen-on-SIGHUP. Rotation is a solved problem owned by
 * logrotate and every collector on the planet, and a half-implemented
 * rotator inside an app is how logs get silently truncated. If the file
 * cannot be opened we fall back to stderr and say so — losing logs because
 * of a bad path is exactly the silent failure this arc exists to end.
 */
async function makeLogSink(): Promise<((line: string) => void) | undefined> {
  const path = typeof process !== "undefined" ? process.env.WW_LOG_FILE : undefined;
  if (!path) return undefined; // default (stderr) is applied by observability.ts
  try {
    // Dynamic import so the default path pulls in no fs at all (and because
    // this module is ESM — `require` does not exist here).
    const fs = await import("node:fs");
    const fd = fs.openSync(path, "a");
    return (line: string) => {
      try {
        fs.writeSync(fd, line + "\n");
      } catch {
        // The file went away (unlinked, full, unmounted). Fall back rather
        // than throw inside a log call — a logger that can crash its caller
        // is worse than one that loses a line.
        process.stderr.write(line + "\n");
      }
    };
  } catch (err) {
    process.stderr.write(
      JSON.stringify({
        ts: new Date().toISOString(),
        level: "error",
        ev: "log-file-open-failed",
        where: "cli.makeLogSink",
        message: err instanceof Error ? err.message : String(err),
      }) + "\n",
    );
    return undefined;
  }
}

export async function startZeroCustodyServer(opts: { port?: number } = {}): Promise<{ close: () => void }> {
  const port = opts.port ?? 1234;
  const spec = "ws";
  const wsMod = (await import(spec).catch(() => {
    throw new Error("startZeroCustodyServer needs the 'ws' package — run: npm i ws");
  })) as {
    WebSocketServer: new (o: { noServer: true; maxPayload: number }) => WsServer & {
      close(): void;
      handleUpgrade(req: unknown, socket: unknown, head: unknown, cb: (client: unknown) => void): void;
      emit(ev: string, ...args: unknown[]): void;
    };
  };
  const http = await import("node:http");
  const { handleSeedRequest } = await import("./seed-http.js");
  const { blankDocxBytes } = await import("./blank.js");

  // Passive observability (counters + structured stderr logging). The LEVEL
  // is WW_LOG_LEVEL (default info; WW_OBS=1 still means the old firehose —
  // see envLogLevel); the /stats READ PATH stays gated on WW_OBS alone,
  // because raising a log level must never open an HTTP surface. Records
  // shapes and counts only — never content (zero-custody).
  const obs = createObservability({ out: await makeLogSink() });

  // PROCESS-LEVEL GUARDS. Without these, an async throw nobody caught prints
  // Node's default trace to stderr as unstructured text — invisible to a
  // collector parsing JSON lines — and, for unhandled rejections, KILLS the
  // process on Node ≥15 with no record in our own stream. One structured
  // error line first, then the deliberate part:
  //  - unhandledRejection: LOG AND CONTINUE. Every one seen in this codebase
  //    came from a per-connection async path (a seal, a storage write); the
  //    room's other participants are unaffected, and taking the whole server
  //    down over one socket's bad day is the worse failure.
  //  - uncaughtException: LOG AND EXIT(1). The process state is unknown after
  //    one of these, and a zero-custody server holds every live room in RAM —
  //    continuing risks serving corrupted state, and a supervisor restart
  //    loses nothing durable (clients hold the bundles).
  process.on("unhandledRejection", (reason) => {
    obs.error("cli.unhandledRejection", reason);
  });
  process.on("uncaughtException", (err) => {
    obs.error("cli.uncaughtException", err, { exiting: true });
    process.exit(1);
  });
  // Lifecycle + abuse limits, resolved from the environment ONCE (limits.ts
  // documents every knob, its default, and the 0-disables convention). The
  // guard and the hub share one clock so their deadlines cannot drift.
  const limits = envLimits();
  const ipGuard = new IpGuard(limits.ip);
  const hub = new CollabHub(/*provider*/ null, undefined, undefined, undefined, undefined, obs, limits, ipGuard); // zero-custody: seed-only rooms
  // INBOUND FRAME CAP — the first wall, and it has to be first: `ws` defaults
  // maxPayload to 100 MiB, so without this a hostile peer makes the server
  // buffer and PARSE a hundred-megabyte frame before any hub-level check —
  // including the 256 KB envelope cap — ever runs. A violating connection is
  // closed by `ws` with status 1009, which arrives at our ordinary `close`
  // handler and unwinds like any other disconnect.
  const wss = new wsMod.WebSocketServer({ noServer: true, maxPayload: limits.surge.wsMaxPayloadBytes });
  attachWebSocketServer(wss, hub, obs, {
    guard: ipGuard,
    trustProxyHops: limits.ip.trustProxyHops,
    maxBufferedBytes: limits.surge.wsMaxBufferedBytes,
    maxConns: limits.surge.maxConns,
  });

  const server = http.createServer((req, res) => {
    const url = req.url ?? "";
    // CORS (doc 06 hosting): a deployed static app and this API live on
    // different origins; the magic link (not a cookie) is the capability,
    // so a permissive ACAO is safe — no credentials ever ride these
    // requests. Preflight (OPTIONS) is answered for the seed/media routes.
    res.setHeader("Access-Control-Allow-Origin", req.headers.origin ?? "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "content-type");
    res.setHeader("Access-Control-Max-Age", "86400");
    if (req.method === "OPTIONS") {
      res.writeHead(204).end();
      return;
    }
    if (req.method === "GET" && url === "/healthz") {
      // Liveness probe — UNGATED on purpose, and the contrast with /stats
      // below is the point: a load balancer, a compose healthcheck or an
      // uptime monitor must be able to ask "is this process answering?"
      // without anyone opening the metrics surface to make it work. Gating
      // health behind WW_OBS would push operators to set WW_OBS in production
      // just to get a probe, which is exactly how a dev-only read path ends
      // up permanently exposed. Carries NO state: not a room count, not a
      // version — just proof the listener is alive.
      res.writeHead(200, { "content-type": "application/json" }).end(`{"ok":true}`);
      return;
    }
    if (req.method === "GET" && url === "/stats") {
      // Dev-only metrics read path (doc: observability). Gated on WW_OBS —
      // deliberately NOT on the log level: raising verbosity to investigate
      // something must never silently OPEN AN HTTP READ SURFACE. They were the
      // same flag before levels existed; keeping /stats on the explicit flag
      // is what stops `WW_LOG_LEVEL=debug` from also exposing an endpoint.
      // 404 when unset so it never leaks in a default run. Returns ONLY safe
      // aggregate numbers, never per-doc content or identifiers.
      if (!observabilityEnabled()) {
        res.writeHead(404, { "content-type": "application/json" }).end(`{"error":"not-found"}`);
        return;
      }
      // `ipBuckets` is a CARDINALITY, never an address — and it is the one
      // signal that diagnoses a proxy misconfiguration: a deployment behind
      // Caddy with WW_TRUST_PROXY unset attributes every socket to the proxy,
      // so this reads 1 while ip-conn-limit refusals climb. That pair means
      // "trust-proxy is wrong" and nothing else does.
      // `rooms` ATTRIBUTES the aggregates: which room is consuming what, and
      // how close each is to the limit that will end it. Same zero-custody
      // rule as everything else here — each row is keyed by an opaque,
      // freshly-random label (Room.obsLabel), never a docId, because a docId
      // is the capability that opens the document and a metrics row carrying
      // one would make this endpoint a way into other people's rooms.
      res.writeHead(200, { "content-type": "application/json" }).end(
        JSON.stringify({ ...obs.snapshot(), ipBuckets: ipGuard.distinctIps(), rooms: hub.roomsSnapshot() }),
      );
      return;
    }
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
        // ADMISSION FIRST — no open room, or no budget left, means no body is
        // ever read. Deciding this after buffering (which is what asking
        // `mediaUpload` alone would force) makes every refusal cost a full
        // body in memory, so the refusals themselves become the attack.
        const admit = hub.mediaUploadAdmission(decodeURIComponent(mDoc));
        if (admit !== 200) {
          const body =
            admit === 429
              ? { status: admit, error: "room-upload-rate", retryAfterMs: 60_000 }
              : { status: admit, error: "no-room" };
          // `connection: close` because the body is deliberately never read:
          // this socket cannot be safely reused for a following request when
          // an unconsumed body is still queued on it.
          //
          // Teardown waits for `finish` to avoid racing the response out the
          // door: destroying inline can discard writes the socket has not
          // flushed yet, which would turn a clean 404 into a dead connection
          // the caller cannot distinguish from a crash. Small bodies like
          // these usually survive an inline destroy — which is the problem
          // with relying on it, not a reason to.
          res.on("finish", () => req.destroy());
          res
            .writeHead(admit, {
              "content-type": "application/json",
              connection: "close",
              ...(admit === 429 ? { "retry-after": "60" } : {}),
            })
            .end(JSON.stringify(body));
          return;
        }
        const chunks: Buffer[] = [];
        let size = 0;
        // SOCKET-LEVEL GUARD, DERIVED from the configured blob cap rather than
        // a constant beside it. It must stay strictly ABOVE that cap: if the
        // two ever cross, a legal upload is destroyed mid-flight before the
        // hub can answer 413, and the client sees a dropped connection instead
        // of "that image is too big" — a failure invisible from both ends and
        // one that only appears when somebody raises the cap. Deriving it
        // makes the ordering structural instead of a promise two constants
        // make to each other.
        const wireCap = mediaWireCap(limits.media);
        req.on("data", (c: Buffer) => {
          size += c.length;
          if (size > wireCap) req.destroy();
          else chunks.push(c);
        });
        req.on("end", () => {
          void hub.mediaUpload(decodeURIComponent(mDoc), mSha, new Uint8Array(Buffer.concat(chunks))).then((status) => {
            // A 413 carries the LIMIT, so the UI can say "images must be under
            // 5 MB" instead of "upload failed". The number is the only thing
            // that makes the message actionable, and the client has no other
            // way to learn it.
            const body =
              status === 413
                ? { status, error: "too-large", maxBytes: limits.media.maxBlobBytes }
                : status === 507
                  ? { status, error: "room-media-full", maxBytes: limits.media.maxRoomBytes }
                  : { status };
            res.writeHead(status, { "content-type": "application/json" }).end(JSON.stringify(body));
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
      const out = handleSeedRequest(
        hub,
        {
          method: put ? "PUT" : "POST",
          docId: put ? decodeURIComponent(url.slice("/docs/".length)) : undefined,
          body,
        },
        // The seed rate limit and the live-doc cap key off this. Extraction
        // reads X-Forwarded-For only when WW_TRUST_PROXY says a proxy we
        // control put it there — see clientIp, which documents why trusting
        // the header by default would make every per-IP limit here bypassable
        // with one line of curl.
        {
          creatorIp: clientIp(req, limits.ip.trustProxyHops),
          // Both modes work by default; a deployment opts IN to refusing
          // plaintext. Defaulting to refusal would change what this server
          // does for everyone already running it, to solve a problem that
          // belongs to one deployment's config rather than to the package.
          encryptedOnly: process.env.WW_ENCRYPTED_ONLY === "1",
        },
      );
      res.writeHead(out.status, { "content-type": "application/json" }).end(JSON.stringify(out.body));
    });
  });
  server.on("upgrade", (req, socket, head) => {
    wss.handleUpgrade(req, socket, head, (client) => wss.emit("connection", client, req));
  });
  // Eviction + token sweeps: coarse cadence is fine — the shortest deadline
  // is the 60s empty-room grace, and warnings carry their MEASURED remainder
  // so a countdown stays honest even though it may be armed up to one
  // interval early (see sweepLifecycle).
  const sweeper = setInterval(() => {
    hub.sweepRooms();
    hub.sweepLifecycle();
    hub.sweepExpired();
    hub.sweepMedia();
    ipGuard.sweep();
  }, 10_000);
  server.listen(port);
  // No startup chatter on stdout by design: `scripts/dev.mjs` prints the demo's
  // own ready line, and the zero-custody server keeps stdout clean (structured
  // observability, when enabled via WW_OBS, goes to stderr — see observability.ts).
  return {
    close: () => {
      clearInterval(sweeper);
      wss.close();
      server.close();
    },
  };
}

// Allow `node dist/cli.js` to run it directly. The zero-custody server is the
// default: it is a strict superset (HTTP seed + CORS + WS) and the only one
// the demo can talk to — a bare `node cli.js` serving the ws-only dev server
// was a footgun (plain GET → 426, no CORS → the demo can't create a doc).
// Opt into the legacy blank-doc ws-only server with WW_DEV_WS=1.
// The second suffix is the npm `bin` name: on POSIX installs argv[1] is the
// `.bin/wordinweb-collab-server` symlink, which node does not realpath.
if (
  typeof process !== "undefined" &&
  process.argv[1] &&
  (process.argv[1].endsWith("cli.js") || process.argv[1].endsWith("wordinweb-collab-server"))
) {
  const port = Number(process.env.PORT ?? 1234);
  if (process.env.WW_DEV_WS === "1") void startDevServer({ port });
  else void startZeroCustodyServer({ port });
}
