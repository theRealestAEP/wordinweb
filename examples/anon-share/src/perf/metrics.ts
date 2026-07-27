import { attachTap, detachTap, type InEntry, type OutFrame } from "./ws-tap";

/**
 * The perf HUD's measurement engine (demo layer only).
 *
 * Everything here is OBSERVATION: the collab numbers come off the wire tap
 * (see ws-tap.ts) and the rendering numbers come from rAF + PerformanceObserver.
 * No collab package is patched, no document text is read — the only document
 * facts collected are a paragraph COUNT and a serialized BYTE LENGTH, both on
 * a 2s idle throttle, never on the hot path.
 *
 * Nothing runs until `start()`: the rAF loop, the long-task observer and the
 * socket listeners are all created there and torn down in `stop()`, so a
 * hidden HUD costs the demo exactly nothing.
 */

/** Rolling windows (ms). */
const FPS_WINDOW = 1000;
const WORST_WINDOW = 5000;
const RATE_WINDOW = 5000;
const DOC_STATS_EVERY = 2000;
/** How many round-trip / repaint samples the rolling averages keep. */
const SAMPLE_CAP = 40;

export interface DocStats {
  paragraphs: number;
  bytes: number;
}

export interface PerfSnapshot {
  /** Frames per second over the last second (rAF ticks). */
  fps: number;
  /** Worst single frame interval in the last 5s (ms). */
  worstFrameMs: number;
  /** 'longtask' entries in the last 5s, and since the HUD opened. */
  longTasks: number;
  longTasksTotal: number;
  /** Highest server-assigned seq this client has seen on the wire. */
  seq: number;
  /** Local intents sent but not yet echoed back by the sequencer. */
  pending: number;
  /** submit → own-echo, rolling average and last sample (ms). */
  rttAvgMs: number | null;
  rttLastMs: number | null;
  rttMaxMs: number | null;
  /** Remote (other clients') sequenced entries per second, 5s window. */
  opsInPerSec: number;
  /** Totals since the HUD opened. */
  opsIn: number;
  opsOut: number;
  /** Bytes moved on the wire since the HUD opened (frame sizes, not content). */
  wireInBytes: number;
  wireOutBytes: number;
  /** Remote entry arrival → the paint that shows it (ms, rolling average). */
  repaintAvgMs: number | null;
  repaintLastMs: number | null;
  /** Throttled document stats (null until the first sample lands). */
  doc: (DocStats & { ageMs: number }) | null;
  /** Wall time the HUD has been measuring (ms). */
  upMs: number;
}

interface Stamped {
  t: number;
  v: number;
}

function prune(xs: Stamped[], now: number, window: number): void {
  let i = 0;
  while (i < xs.length && now - xs[i].t > window) i++;
  if (i) xs.splice(0, i);
}

function avg(xs: number[]): number | null {
  return xs.length ? xs.reduce((a, b) => a + b, 0) / xs.length : null;
}

function push(xs: number[], v: number): void {
  xs.push(v);
  if (xs.length > SAMPLE_CAP) xs.shift();
}

export class PerfMonitor {
  private started = 0;
  private raf = 0;
  private observer: PerformanceObserver | null = null;
  private docTimer: ReturnType<typeof setTimeout> | null = null;
  private clientId = "";
  private docStatsFn: (() => DocStats | null) | null = null;

  // Rendering
  private frames: Stamped[] = []; // v = frame interval (ms)
  private lastFrameAt = 0;
  private longTaskTimes: Stamped[] = [];
  private longTaskTotal = 0;

  // Collab (wire tap)
  private outstanding = new Map<number, number>(); // clientSeq → send time
  private seq = 0;
  private rtt: number[] = [];
  private rttMax = 0;
  private remoteArrivals: Stamped[] = [];
  private opsIn = 0;
  private opsOut = 0;
  private wireIn = 0;
  private wireOut = 0;

  // Repaint cost: armed by a remote entry, sampled two frames later (the
  // frame after the one that commits the repaint) — an approximation of
  // "arrival → the user can see it", deliberately erring high.
  private repaintArmedAt: number | null = null;
  private repaintStage = 0;
  private repaint: number[] = [];

  private doc: (DocStats & { at: number }) | null = null;

  get running(): boolean {
    return this.started !== 0;
  }

  start(opts: { clientId: string; docStats: () => DocStats | null }): void {
    if (this.running) return;
    this.clientId = opts.clientId;
    this.docStatsFn = opts.docStats;
    this.started = performance.now();
    this.lastFrameAt = this.started;

    const tick = (now: number) => {
      const dt = now - this.lastFrameAt;
      this.lastFrameAt = now;
      this.frames.push({ t: now, v: dt });
      prune(this.frames, now, WORST_WINDOW);
      if (this.repaintArmedAt !== null) {
        if (this.repaintStage === 0) this.repaintStage = 1;
        else {
          push(this.repaint, now - this.repaintArmedAt);
          this.repaintArmedAt = null;
          this.repaintStage = 0;
        }
      }
      this.raf = requestAnimationFrame(tick);
    };
    this.raf = requestAnimationFrame(tick);

    // Long tasks: not implemented everywhere (Safari/Firefox) — absence just
    // leaves the counter at 0 rather than breaking the HUD.
    try {
      if (typeof PerformanceObserver !== "undefined" && PerformanceObserver.supportedEntryTypes?.includes("longtask")) {
        this.observer = new PerformanceObserver((list) => {
          for (const e of list.getEntries()) {
            this.longTaskTotal++;
            this.longTaskTimes.push({ t: performance.now(), v: e.duration });
          }
        });
        this.observer.observe({ entryTypes: ["longtask"] });
      }
    } catch {
      this.observer = null;
    }

    attachTap({
      onSubmit: (f: OutFrame) => {
        this.opsOut++;
        this.wireOut += f.bytes;
        this.outstanding.set(f.clientSeq, f.at);
      },
      onEntries: (entries: InEntry[], bytes: number, at: number) => {
        this.wireIn += bytes;
        let sawRemote = false;
        for (const e of entries) {
          if (e.seq > this.seq) this.seq = e.seq;
          if (e.clientId === this.clientId) {
            const sentAt = this.outstanding.get(e.clientSeq);
            if (sentAt !== undefined) {
              this.outstanding.delete(e.clientSeq);
              const rtt = at - sentAt;
              push(this.rtt, rtt);
              if (rtt > this.rttMax) this.rttMax = rtt;
            }
          } else {
            this.opsIn++;
            sawRemote = true;
            this.remoteArrivals.push({ t: at, v: 1 });
          }
        }
        // One repaint measurement per burst — arm on the FIRST unmeasured
        // remote arrival so a flood doesn't reset the clock forever.
        if (sawRemote && this.repaintArmedAt === null) {
          this.repaintArmedAt = at;
          this.repaintStage = 0;
        }
      },
    });

    const sampleDoc = () => {
      const run = () => {
        try {
          const s = this.docStatsFn?.() ?? null;
          if (s) this.doc = { ...s, at: performance.now() };
        } catch {
          /* a mid-reload session: skip this sample */
        }
        this.docTimer = setTimeout(sampleDoc, DOC_STATS_EVERY);
      };
      // Serializing the document is the one expensive thing the HUD does, so
      // it waits for idle time and never competes with typing or a repaint.
      const ric = (window as unknown as { requestIdleCallback?: (cb: () => void, o?: { timeout: number }) => number }).requestIdleCallback;
      if (ric) ric(run, { timeout: 1000 });
      else run();
    };
    this.docTimer = setTimeout(sampleDoc, 200);
  }

  stop(): void {
    if (!this.running) return;
    cancelAnimationFrame(this.raf);
    this.observer?.disconnect();
    this.observer = null;
    if (this.docTimer) clearTimeout(this.docTimer);
    this.docTimer = null;
    detachTap();
    this.started = 0;
    this.frames = [];
    this.longTaskTimes = [];
    this.longTaskTotal = 0;
    this.outstanding.clear();
    this.rtt = [];
    this.rttMax = 0;
    this.repaint = [];
    this.repaintArmedAt = null;
    this.remoteArrivals = [];
    this.opsIn = 0;
    this.opsOut = 0;
    this.wireIn = 0;
    this.wireOut = 0;
    this.seq = 0;
    this.doc = null;
  }

  snapshot(): PerfSnapshot | null {
    if (!this.running) return null;
    const now = performance.now();
    prune(this.frames, now, WORST_WINDOW);
    prune(this.longTaskTimes, now, WORST_WINDOW);
    prune(this.remoteArrivals, now, RATE_WINDOW);
    const recent = this.frames.filter((f) => now - f.t <= FPS_WINDOW);
    const fpsSpan = recent.length ? Math.min(FPS_WINDOW, now - this.started) : 0;
    const rttAvg = avg(this.rtt);
    const repaintAvg = avg(this.repaint);
    // The rate window is short at startup — divide by the time actually
    // measured so the first seconds don't read artificially low.
    const rateSpan = Math.min(RATE_WINDOW, Math.max(now - this.started, 1));
    return {
      fps: fpsSpan > 0 ? (recent.length * 1000) / fpsSpan : 0,
      worstFrameMs: this.frames.reduce((m, f) => Math.max(m, f.v), 0),
      longTasks: this.longTaskTimes.length,
      longTasksTotal: this.longTaskTotal,
      seq: this.seq,
      pending: this.outstanding.size,
      rttAvgMs: rttAvg,
      rttLastMs: this.rtt.length ? this.rtt[this.rtt.length - 1] : null,
      rttMaxMs: this.rtt.length ? this.rttMax : null,
      opsInPerSec: (this.remoteArrivals.length * 1000) / rateSpan,
      opsIn: this.opsIn,
      opsOut: this.opsOut,
      wireInBytes: this.wireIn,
      wireOutBytes: this.wireOut,
      repaintAvgMs: repaintAvg,
      repaintLastMs: this.repaint.length ? this.repaint[this.repaint.length - 1] : null,
      doc: this.doc ? { paragraphs: this.doc.paragraphs, bytes: this.doc.bytes, ageMs: now - this.doc.at } : null,
      upMs: now - this.started,
    };
  }
}

/** One monitor per page — the HUD component drives it, `__ww.perf()` reads it. */
export const perfMonitor = new PerfMonitor();
