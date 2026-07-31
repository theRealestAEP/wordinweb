import { useEffect, useRef, useState } from "react";
import { perfMonitor, type DocStats, type PerfSnapshot } from "./metrics";

/**
 * The perf debug menu (demo layer, dev harness only).
 *
 * Open with `Ctrl+Shift+P` or `?perf=1`. Closed, it is a single keydown
 * listener and nothing else — the rAF loop, the long-task observer, the socket
 * listeners and the throttled doc-stat sampler all live inside PerfMonitor and
 * are created on open / destroyed on close.
 *
 * It reports how the collaboration FEELS (frames, round-trips, repaint lag)
 * and never what it CONTAINS: paragraph count and serialized size are the only
 * document facts on the panel, both from a 2s idle sample.
 */

const num = (v: number | null | undefined, digits = 0, unit = ""): string =>
  v === null || v === undefined || Number.isNaN(v) ? "—" : `${v.toFixed(digits)}${unit}`;

const bytes = (b: number | null | undefined): string => {
  if (b === null || b === undefined) return "—";
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / (1024 * 1024)).toFixed(2)} MB`;
};

/** Green / amber / red against a "good below" and "bad above" pair. */
const grade = (v: number | null, good: number, bad: number, invert = false): string => {
  if (v === null) return "#9aa0a6";
  const hot = invert ? v < good : v > bad;
  const warm = invert ? v < bad : v > good;
  return hot ? "#f28b82" : warm ? "#fdd663" : "#81c995";
};

function Row({ label, value, color, testid }: { label: string; value: string; color?: string; testid?: string }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", gap: 12, lineHeight: "17px" }}>
      <span style={{ color: "#9aa0a6" }}>{label}</span>
      <span data-testid={testid} style={{ color: color ?? "#e8eaed", fontVariantNumeric: "tabular-nums" }}>{value}</span>
    </div>
  );
}

export function PerfHud({ clientId, docStats }: { clientId: string; docStats: () => DocStats | null }) {
  const [open, setOpen] = useState(() => {
    try {
      return new URLSearchParams(location.search).get("perf") === "1";
    } catch {
      return false;
    }
  });
  const [snap, setSnap] = useState<PerfSnapshot | null>(null);

  // The shortcut is the only thing alive while the HUD is closed.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.ctrlKey && e.shiftKey && (e.code === "KeyP" || e.key === "P" || e.key === "p")) {
        e.preventDefault();
        setOpen((v) => !v);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // The app's `docStats` closes over the live session, so its identity changes
  // on every reconciled edit. Reading it through a ref keeps the effect below
  // keyed on the HUD's OWN state — otherwise every keystroke would restart the
  // monitor (detaching the wire tap and zeroing every counter mid-measurement).
  const docStatsRef = useRef(docStats);
  docStatsRef.current = docStats;

  useEffect(() => {
    if (!open) {
      perfMonitor.stop();
      setSnap(null);
      return;
    }
    perfMonitor.start({ clientId, docStats: () => docStatsRef.current() });
    // 2 Hz: fast enough to watch a burst, slow enough that the HUD's own
    // re-render never shows up in the numbers it is reporting.
    const id = setInterval(() => setSnap(perfMonitor.snapshot()), 500);
    setSnap(perfMonitor.snapshot());
    return () => {
      clearInterval(id);
      perfMonitor.stop();
    };
  }, [open, clientId]);

  if (!open) return null;

  const s = snap;
  return (
    <div
      data-testid="perf-hud"
      style={{
        position: "fixed", right: 12, bottom: 12, zIndex: 2147483000, width: 244,
        background: "rgba(20,22,26,0.88)", color: "#e8eaed", borderRadius: 8,
        border: "1px solid rgba(255,255,255,0.14)", boxShadow: "0 6px 24px rgba(0,0,0,0.35)",
        font: "11px/1.45 ui-monospace, SFMono-Regular, Menlo, monospace",
        padding: "8px 10px 9px", backdropFilter: "blur(6px)", userSelect: "none",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 5 }}>
        <b style={{ letterSpacing: 0.4 }}>perf</b>
        <span style={{ color: "#5f6368" }}>
          ⌃⇧P{" "}
          <button
            data-testid="perf-close"
            onClick={() => setOpen(false)}
            title="Close (Ctrl+Shift+P)"
            style={{ background: "none", border: "none", color: "#9aa0a6", cursor: "pointer", font: "inherit", padding: 0 }}
          >
            ✕
          </button>
        </span>
      </div>

      <div style={{ color: "#5f6368", margin: "4px 0 2px" }}>render</div>
      <Row label="fps" value={num(s?.fps ?? null, 0)} color={grade(s?.fps ?? null, 50, 30, true)} testid="perf-fps" />
      <Row label="worst frame 5s" value={num(s?.worstFrameMs ?? null, 1, "ms")} color={grade(s?.worstFrameMs ?? null, 34, 100)} testid="perf-worst" />
      <Row label="long tasks 5s" value={s ? `${s.longTasks} (${s.longTasksTotal})` : "—"} color={grade(s?.longTasks ?? null, 0, 3)} testid="perf-longtasks" />

      <div style={{ color: "#5f6368", margin: "6px 0 2px" }}>collab</div>
      <Row label="seq" value={s ? String(s.seq) : "—"} testid="perf-seq" />
      <Row label="pending" value={s ? String(s.pending) : "—"} color={grade(s?.pending ?? null, 1, 5)} testid="perf-pending" />
      <Row
        label="rtt avg/last"
        value={s ? `${num(s.rttAvgMs, 0)}/${num(s.rttLastMs, 0)}ms` : "—"}
        color={grade(s?.rttAvgMs ?? null, 60, 250)}
        testid="perf-rtt"
      />
      <Row label="rtt max" value={num(s?.rttMaxMs ?? null, 0, "ms")} color={grade(s?.rttMaxMs ?? null, 150, 500)} testid="perf-rtt-max" />
      <Row label="ops in/s" value={num(s?.opsInPerSec ?? null, 1)} testid="perf-opsin" />
      <Row label="ops in/out" value={s ? `${s.opsIn}/${s.opsOut}` : "—"} testid="perf-ops" />
      <Row
        label="repaint avg/last"
        value={s ? `${num(s.repaintAvgMs, 0)}/${num(s.repaintLastMs, 0)}ms` : "—"}
        color={grade(s?.repaintAvgMs ?? null, 50, 200)}
        testid="perf-repaint"
      />
      <Row label="wire in/out" value={s ? `${bytes(s.wireInBytes)}/${bytes(s.wireOutBytes)}` : "—"} testid="perf-wire" />

      <div style={{ color: "#5f6368", margin: "6px 0 2px" }}>doc (2s sample)</div>
      <Row label="paragraphs" value={s?.doc ? String(s.doc.paragraphs) : "—"} testid="perf-paras" />
      <Row label="serialized" value={bytes(s?.doc?.bytes ?? null)} testid="perf-size" />
      <div style={{ color: "#5f6368", marginTop: 5, fontSize: 10 }}>metrics only — no document text is read</div>
    </div>
  );
}
