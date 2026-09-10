/**
 * Word-parity warning. The document asks for faces this browser cannot
 * render, so substitutes are painted and line breaks may drift from Word.
 * Both screens (local and live) show it above the page; the host clears the
 * list to dismiss, and the next document load reports afresh.
 */
export function MissingFontsBanner({ fonts, onDismiss }: {
  fonts: { family: string }[];
  onDismiss: () => void;
}) {
  if (fonts.length === 0) return null;
  return (
    <div
      data-testid="missing-fonts-banner"
      role="status"
      style={{
        display: "flex", alignItems: "center", gap: 8, padding: "6px 14px",
        background: "#fff7e0", borderBottom: "1px solid #e8d9a0",
        font: "12.5px system-ui, sans-serif", color: "#6b5518",
      }}
    >
      <span>
        Some fonts this document asks for aren't available here, so substitutes are shown —
        layout may differ from Word: <b>{fonts.map((f) => f.family).join(", ")}</b>
      </span>
      <button
        data-testid="missing-fonts-dismiss"
        onClick={onDismiss}
        style={{ marginLeft: "auto", border: "none", background: "none", cursor: "pointer", color: "inherit" }}
        title="Dismiss"
      >
        ✕
      </button>
    </div>
  );
}
