import type { ClientMessage, ServerMessage } from "./protocol.js";
import type { ClientTransport } from "./connection.js";

/**
 * Whether this client's edits can currently reach the sequencer.
 *
 *  - `live`         the round trip completed recently; submits will land.
 *  - `reconnecting` the socket is gone and a retry is scheduled. TRANSIENT and
 *                   usually invisible — a laptop lid closed for ten seconds
 *                   passes through here and back to `live` without the user
 *                   noticing.
 *  - `lost`         the retries are exhausted. Needs a human: reconnect or
 *                   reload.
 *
 * A NON-LIVE STATE IS NOT A WRITE GATE (doc 12 §5: offline editing is
 * first-class). What consumers MUST guarantee on anything but `live` is that
 * nothing typed can be LOST: edits apply locally and are recorded to the
 * durable offline tail (doc 15 §2) for replay on reconnect — the react
 * layer's capture path does exactly this. The write gate (`writesBlocked`)
 * stays reserved for server REFUSALS, where accepting a keystroke really is
 * accepting-then-losing it. An editable surface whose disconnected edits are
 * neither kept nor gated remains the silent-loss bug this type was created
 * to surface — keep them either captured or blocked, never dropped.
 */
export type ConnectionState = "live" | "reconnecting" | "lost";

export interface LivenessOptions {
  /** Gap between probes once the previous one was answered. */
  intervalMs?: number;
  /** How long a probe may go unanswered before the socket is declared dead. */
  timeoutMs?: number;
  setTimer?: (fn: () => void, ms: number) => unknown;
  clearTimer?: (t: unknown) => void;
}

/** Probe cadence. 15s/10s means a dead socket is noticed within ~25s at worst
 * and the wake-up probe (below) makes the common phone case immediate, while
 * costing two tiny frames a minute on an idle tab. */
const DEFAULT_INTERVAL_MS = 15_000;
const DEFAULT_TIMEOUT_MS = 10_000;

/**
 * Application-level heartbeat over an established transport.
 *
 * WHY THE TRANSPORT CANNOT ANSWER THIS ITSELF. `onclose` and `onerror` catch a
 * socket that dies politely — a server restart, a closed lid that the OS
 * notices, a dropped Wi-Fi association. They do NOT catch the case this whole
 * arc exists for: a phone browser suspended mid-session leaves a HALF-OPEN
 * connection where the peer is gone but nothing local knows it. No event
 * fires, `readyState` still reads OPEN, and `send()` keeps succeeding into a
 * void. Asking for a reply and noticing its absence is the only way to learn
 * the truth, so the heartbeat is not a belt-and-braces addition to the close
 * handlers — it is the only detector that covers the motivating scenario.
 *
 * Timers are injected so tests drive the clock deterministically, matching
 * BundlePersister rather than inventing a second convention.
 */
export class LivenessMonitor {
  private probeTimer: unknown = null;
  private deadlineTimer: unknown = null;
  private nonce = 0;
  /** The nonce of the probe currently awaiting an answer; null when idle. */
  private awaiting: number | null = null;
  private stopped = false;

  constructor(
    private readonly send: (msg: ClientMessage) => void,
    /** Called ONCE per monitor when a probe goes unanswered. */
    private readonly onDead: () => void,
    private readonly opts: LivenessOptions = {},
  ) {}

  private get intervalMs(): number {
    return this.opts.intervalMs ?? DEFAULT_INTERVAL_MS;
  }
  private get timeoutMs(): number {
    return this.opts.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  }
  private setTimer(fn: () => void, ms: number): unknown {
    return (this.opts.setTimer ?? ((f, m) => setTimeout(f, m)))(fn, ms);
  }
  private clearTimer(t: unknown): void {
    if (t !== null) (this.opts.clearTimer ?? clearTimeout)(t as never);
  }

  /** Begin probing. Call when the socket opens. */
  start(): void {
    if (this.stopped) return;
    this.arm();
  }

  /** Stop every timer and answer nothing further. Idempotent — the reconnect
   * path calls it from both the close handler and the effect teardown. */
  stop(): void {
    this.stopped = true;
    this.clearTimer(this.probeTimer);
    this.clearTimer(this.deadlineTimer);
    this.probeTimer = null;
    this.deadlineTimer = null;
    this.awaiting = null;
  }

  /** Schedule the next probe one interval out. */
  private arm(): void {
    if (this.stopped) return;
    this.clearTimer(this.probeTimer);
    this.probeTimer = this.setTimer(() => {
      this.probeTimer = null;
      this.probe();
    }, this.intervalMs);
  }

  /**
   * Probe NOW rather than waiting for the interval — wired to
   * `visibilitychange`, because a tab waking from sleep is the exact moment
   * the connection is most likely already dead and the worst moment to make
   * the user wait a full interval to find out. Without this the motivating
   * scenario (unlock phone, start typing) spends up to 25s in the silent-loss
   * state the arc exists to close.
   *
   * A probe already in flight is left alone: re-arming its deadline on every
   * focus event would let a user who keeps switching tabs postpone the verdict
   * indefinitely.
   */
  probe(): void {
    if (this.stopped || this.awaiting !== null) return;
    const nonce = ++this.nonce;
    this.awaiting = nonce;
    this.deadlineTimer = this.setTimer(() => {
      this.deadlineTimer = null;
      // Still waiting on the same probe ⇒ the round trip never completed.
      if (this.stopped || this.awaiting !== nonce) return;
      this.stop(); // one verdict per monitor; the reconnect owns what follows
      this.onDead();
    }, this.timeoutMs);
    // Sent AFTER the deadline is armed: a transport that throws synchronously
    // must still be caught by the timeout rather than leaving a monitor that
    // silently never probes again.
    try {
      this.send({ t: "ping", nonce });
    } catch {
      // Ignored on purpose — the armed deadline above is the report. Throwing
      // here would propagate into a timer callback with no handler.
    }
  }

  /**
   * Feed an inbound frame to the monitor.
   *
   * ONLY A MATCHING PONG COUNTS AS PROOF OF LIFE — deliberately, and this is
   * the subtle decision in the file. The obvious optimisation is to treat any
   * inbound traffic as evidence the connection works, which is what most
   * heartbeats do. It is wrong here: it assumes both directions fail together.
   * A connection whose OUTBOUND half is broken while broadcasts still arrive
   * would look permanently healthy under that rule, and that is precisely the
   * user-visible bug being fixed — edits going nowhere while the screen keeps
   * updating with everyone else's. A pong proves the full round trip an edit
   * actually needs, so nothing weaker is accepted.
   */
  noteInbound(msg: ServerMessage): void {
    if (this.stopped) return;
    if (msg.t !== "pong") return;
    if (this.awaiting === null || msg.nonce !== this.awaiting) return; // stale
    this.awaiting = null;
    this.clearTimer(this.deadlineTimer);
    this.deadlineTimer = null;
    this.arm();
  }
}

/**
 * Wrap a transport with a heartbeat, returning both halves.
 *
 * A WRAPPER RATHER THAN A HOOK INSIDE THE CONNECTION, because
 * `ClientTransport.onMessage` holds exactly ONE handler — registering a second
 * would overwrite the connection's own and silently stop the document
 * updating. Teeing here keeps that single-handler contract intact and leaves
 * both CollabConnection and EncryptedCollabConnection completely unaware that
 * a heartbeat exists.
 *
 * Pongs are consumed rather than forwarded: neither connection class has a
 * `default:` in its dispatch, so forwarding would be harmless — but passing a
 * message no consumer handles invites exactly the "why is this ignored?"
 * confusion the rest of this protocol avoids.
 */
export function monitorTransport(
  inner: ClientTransport,
  onDead: () => void,
  opts: LivenessOptions = {},
): { transport: ClientTransport; monitor: LivenessMonitor } {
  const monitor = new LivenessMonitor((msg) => inner.send(msg), onDead, opts);
  const transport: ClientTransport = {
    send: (msg) => inner.send(msg),
    onMessage: (cb) => {
      inner.onMessage((msg) => {
        monitor.noteInbound(msg);
        if (msg.t === "pong") return;
        cb(msg);
      });
    },
  };
  return { transport, monitor };
}
