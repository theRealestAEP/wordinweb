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

// Allow `node dist/cli.js` to run it directly.
if (typeof process !== "undefined" && process.argv[1] && process.argv[1].endsWith("cli.js")) {
  const port = Number(process.env.PORT ?? 1234);
  void startDevServer({ port });
}
