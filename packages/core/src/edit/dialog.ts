export interface TextInputDialogOptions {
  title: string;
  label?: string;
  value?: string;
  placeholder?: string;
  submitLabel?: string;
  multiline?: boolean;
}

let dialogSequence = 0;

/** Opens the editor's CSS-token-driven text dialog. This replaces browser
 * prompts while keeping the same promise result: submitted text or null. */
export function requestTextInputDialog(
  anchor: HTMLElement,
  options: TextInputDialogOptions,
): Promise<string | null> {
  const doc = anchor.ownerDocument;
  const win = doc.defaultView;
  return new Promise((resolve) => {
    const id = `dxw-input-dialog-${++dialogSequence}`;
    const previousFocus = doc.activeElement instanceof HTMLElement ? doc.activeElement : null;
    const backdrop = doc.createElement("div");
    backdrop.className = "dxw-input-dialog-backdrop";
    backdrop.dataset.dxwInputDialog = "";
    backdrop.style.cssText =
      "position:fixed;inset:0;z-index:var(--dxw-dialog-z-index,1000);" +
      "display:grid;place-items:center;padding:16px;box-sizing:border-box;" +
      "background:var(--dxw-dialog-backdrop,rgba(32,33,36,.38));";

    if (win) {
      const computed = win.getComputedStyle(anchor);
      for (const property of [
        "--dxw-toolbar-fg", "--dxw-toolbar-border", "--dxw-toolbar-muted",
        "--dxw-accent", "--dxw-accent-fg", "--dxw-popover-bg",
        "--dxw-popover-shadow", "--dxw-dialog-backdrop", "--dxw-dialog-z-index",
        "--dxw-dialog-width", "--dxw-dialog-radius", "--dxw-dialog-field-bg",
      ]) {
        const value = computed.getPropertyValue(property);
        if (value) backdrop.style.setProperty(property, value);
      }
    }

    const form = doc.createElement("form");
    form.className = "dxw-input-dialog";
    form.setAttribute("role", "dialog");
    form.setAttribute("aria-modal", "true");
    form.setAttribute("aria-labelledby", id);
    form.style.cssText =
      "width:min(var(--dxw-dialog-width,420px),calc(100vw - 32px));box-sizing:border-box;" +
      "display:grid;gap:10px;padding:16px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:var(--dxw-dialog-radius,10px);color:var(--dxw-toolbar-fg,#3c4043);" +
      "background:var(--dxw-popover-bg,#fff);box-shadow:var(--dxw-popover-shadow,0 8px 28px rgba(0,0,0,.24));" +
      "font:13px system-ui,sans-serif;";

    const title = doc.createElement("strong");
    title.id = id;
    title.className = "dxw-input-dialog-title";
    title.textContent = options.title;
    title.style.cssText = "font-size:15px;line-height:1.35;";
    form.appendChild(title);

    const label = doc.createElement("label");
    label.className = "dxw-input-dialog-label";
    label.textContent = options.label ?? options.title;
    label.style.cssText = "display:grid;gap:5px;color:var(--dxw-toolbar-muted,#5f6368);font-size:12px;";
    const field = options.multiline ? doc.createElement("textarea") : doc.createElement("input");
    field.className = "dxw-input-dialog-field";
    field.value = options.value ?? "";
    field.placeholder = options.placeholder ?? "";
    field.setAttribute("aria-label", options.label ?? options.title);
    if (field instanceof HTMLTextAreaElement) field.rows = 4;
    field.style.cssText =
      "width:100%;min-height:34px;box-sizing:border-box;padding:7px 9px;" +
      "border:1px solid var(--dxw-toolbar-border,#dadce0);border-radius:6px;outline:none;" +
      "color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-dialog-field-bg,var(--dxw-popover-bg,#fff));" +
      "font:13px system-ui,sans-serif;resize:vertical;";
    label.appendChild(field);
    form.appendChild(label);

    const actions = doc.createElement("div");
    actions.className = "dxw-input-dialog-actions";
    actions.style.cssText = "display:flex;justify-content:flex-end;gap:7px;";
    const cancel = doc.createElement("button");
    cancel.type = "button";
    cancel.className = "dxw-input-dialog-cancel";
    cancel.textContent = "Cancel";
    cancel.style.cssText =
      "min-height:32px;padding:0 12px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:7px;color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-popover-bg,#fff);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    const submit = doc.createElement("button");
    submit.type = "submit";
    submit.className = "dxw-input-dialog-submit";
    submit.textContent = options.submitLabel ?? "Apply";
    submit.style.cssText =
      "min-height:32px;padding:0 14px;border:1px solid transparent;border-radius:7px;" +
      "color:var(--dxw-accent-fg,#fff);background:var(--dxw-accent,#1a73e8);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    actions.append(cancel, submit);
    form.appendChild(actions);
    backdrop.appendChild(form);
    doc.body.appendChild(backdrop);

    let finished = false;
    const finish = (value: string | null) => {
      if (finished) return;
      finished = true;
      doc.removeEventListener("keydown", onKeyDown, true);
      backdrop.remove();
      previousFocus?.focus({ preventScroll: true });
      resolve(value);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        finish(null);
      }
    };
    doc.addEventListener("keydown", onKeyDown, true);
    backdrop.addEventListener("mousedown", (event) => {
      if (event.target === backdrop) finish(null);
    });
    cancel.addEventListener("click", () => finish(null));
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      finish(field.value);
    });
    requestAnimationFrame(() => {
      field.focus({ preventScroll: true });
      field.select();
    });
  });
}
