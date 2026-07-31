/**
 * URL scheme allowlist for user-authored links (hyperlinks, web-video URLs,
 * embedded-object sources). Document content is authored by other users — in
 * the collaborative/anonymous-demo setting by strangers — so a link target
 * reaching `anchor.href` or `window.open` unchecked is stored XSS
 * (`javascript:`, `data:text/html`, `vbscript:`). Only these schemes may
 * navigate; everything else renders as inert text.
 *
 * Threat-model launch gate #1 (internal/collab-plan/11-threat-model.md).
 */

const SAFE_SCHEMES = new Set(["http:", "https:", "mailto:"]);

/** True if `url` is safe to navigate to. Internal bookmark refs ("#name") are
 * safe (handled by in-app scrolling, never navigation). Relative/scheme-less
 * URLs are treated as safe http(s)-style references. */
export function isSafeUrl(url: string): boolean {
  // Browsers strip C0 control chars and tab/newline anywhere in the URL
  // before resolving the scheme, so "java\tscript:" runs. Normalize the same
  // way, then evaluate — matching the parser we are defending.
  const normalized = url.replace(/[\x00-\x20]/g, "");
  if (normalized === "") return false;
  if (normalized.startsWith("#")) return true;
  const colon = normalized.indexOf(":");
  const slash = normalized.indexOf("/");
  // No scheme (colon absent or appearing after the first path separator):
  // a relative or protocol-relative reference — safe.
  if (colon === -1 || (slash !== -1 && slash < colon)) return true;
  const scheme = normalized.slice(0, colon + 1).toLowerCase();
  return SAFE_SCHEMES.has(scheme);
}

/** The URL if safe to navigate to, otherwise `about:blank` — use when a value
 * must still be assigned (e.g. anchor.href) but must never be dangerous. */
export function safeUrlOrBlank(url: string): string {
  return isSafeUrl(url) ? url : "about:blank";
}
