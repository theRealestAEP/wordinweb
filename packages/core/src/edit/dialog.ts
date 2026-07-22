export interface TextInputDialogOptions {
  title: string;
  label?: string;
  value?: string;
  placeholder?: string;
  submitLabel?: string;
  multiline?: boolean;
  inputType?: "text" | "url" | "number";
  min?: number;
  max?: number;
  step?: number;
}

export interface NumberPairDialogValue {
  first: number;
  second: number;
}

export interface NumberPairDialogOptions {
  title: string;
  firstLabel: string;
  secondLabel: string;
  value: NumberPairDialogValue;
  min?: number;
  step?: number;
}

export interface ColorDialogOptions {
  title: string;
  value: string | null;
  allowNone?: boolean;
}

export interface LineStyleDialogValue {
  color: string;
  width: number;
  style: "solid" | "dashed" | "dotted";
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
    if (field instanceof HTMLInputElement) {
      field.type = options.inputType ?? "text";
      if (options.min !== undefined) field.min = String(options.min);
      if (options.max !== undefined) field.max = String(options.max);
      if (options.step !== undefined) field.step = String(options.step);
    }
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

/** Opens two numeric fields for values such as width/height and X/Y. */
export function requestNumberPairDialog(
  anchor: HTMLElement,
  options: NumberPairDialogOptions,
): Promise<NumberPairDialogValue | null> {
  const doc = anchor.ownerDocument;
  const win = doc.defaultView;
  return new Promise((resolve) => {
    const id = `dxw-number-pair-dialog-${++dialogSequence}`;
    const previousFocus = doc.activeElement instanceof HTMLElement ? doc.activeElement : null;
    const backdrop = doc.createElement("div");
    backdrop.className = "dxw-input-dialog-backdrop";
    backdrop.dataset.dxwNumberPairDialog = "";
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
        const token = computed.getPropertyValue(property);
        if (token) backdrop.style.setProperty(property, token);
      }
    }

    const form = doc.createElement("form");
    form.className = "dxw-input-dialog dxw-number-pair-dialog";
    form.setAttribute("role", "dialog");
    form.setAttribute("aria-modal", "true");
    form.setAttribute("aria-labelledby", id);
    form.style.cssText =
      "width:min(var(--dxw-dialog-width,420px),calc(100vw - 32px));box-sizing:border-box;" +
      "display:grid;gap:12px;padding:16px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:var(--dxw-dialog-radius,10px);color:var(--dxw-toolbar-fg,#3c4043);" +
      "background:var(--dxw-popover-bg,#fff);box-shadow:var(--dxw-popover-shadow,0 8px 28px rgba(0,0,0,.24));" +
      "font:13px system-ui,sans-serif;";

    const title = doc.createElement("strong");
    title.id = id;
    title.textContent = options.title;
    title.style.cssText = "font-size:15px;line-height:1.35;";
    form.appendChild(title);

    const fields = doc.createElement("div");
    fields.style.cssText = "display:grid;grid-template-columns:1fr 1fr;gap:10px;";
    const makeField = (labelText: string, value: number) => {
      const label = doc.createElement("label");
      label.textContent = labelText;
      label.style.cssText = "display:grid;gap:5px;color:var(--dxw-toolbar-muted,#5f6368);font-size:12px;";
      const input = doc.createElement("input");
      input.type = "number";
      input.required = true;
      input.value = String(value);
      if (options.min !== undefined) input.min = String(options.min);
      if (options.step !== undefined) input.step = String(options.step);
      input.setAttribute("aria-label", labelText);
      input.style.cssText =
        "width:100%;min-height:34px;box-sizing:border-box;padding:7px 9px;" +
        "border:1px solid var(--dxw-toolbar-border,#dadce0);border-radius:6px;outline:none;" +
        "color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-dialog-field-bg,var(--dxw-popover-bg,#fff));" +
        "font:13px system-ui,sans-serif;";
      label.appendChild(input);
      fields.appendChild(label);
      return input;
    };
    const first = makeField(options.firstLabel, options.value.first);
    const second = makeField(options.secondLabel, options.value.second);
    form.appendChild(fields);

    const actions = doc.createElement("div");
    actions.style.cssText = "display:flex;justify-content:flex-end;gap:7px;";
    const cancel = doc.createElement("button");
    cancel.type = "button";
    cancel.textContent = "Cancel";
    cancel.style.cssText =
      "min-height:32px;padding:0 12px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:7px;color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-popover-bg,#fff);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    const submit = doc.createElement("button");
    submit.type = "submit";
    submit.textContent = "Apply";
    submit.style.cssText =
      "min-height:32px;padding:0 14px;border:1px solid transparent;border-radius:7px;" +
      "color:var(--dxw-accent-fg,#fff);background:var(--dxw-accent,#1a73e8);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    actions.append(cancel, submit);
    form.appendChild(actions);
    backdrop.appendChild(form);
    doc.body.appendChild(backdrop);

    let finished = false;
    const finish = (value: NumberPairDialogValue | null) => {
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
      const firstValue = Number(first.value);
      const secondValue = Number(second.value);
      if (!Number.isFinite(firstValue) || !Number.isFinite(secondValue)) {
        form.reportValidity();
        return;
      }
      finish({ first: firstValue, second: secondValue });
    });
    requestAnimationFrame(() => {
      first.focus({ preventScroll: true });
      first.select();
    });
  });
}

/** Opens a native color picker with an exact hex field and optional no-fill choice. */
export function requestColorDialog(
  anchor: HTMLElement,
  options: ColorDialogOptions,
): Promise<string | null | undefined> {
  const doc = anchor.ownerDocument;
  const win = doc.defaultView;
  return new Promise((resolve) => {
    const id = `dxw-color-dialog-${++dialogSequence}`;
    const previousFocus = doc.activeElement instanceof HTMLElement ? doc.activeElement : null;
    const backdrop = doc.createElement("div");
    backdrop.className = "dxw-input-dialog-backdrop";
    backdrop.dataset.dxwColorDialog = "";
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
        const token = computed.getPropertyValue(property);
        if (token) backdrop.style.setProperty(property, token);
      }
    }

    const form = doc.createElement("form");
    form.className = "dxw-input-dialog dxw-color-dialog";
    form.setAttribute("role", "dialog");
    form.setAttribute("aria-modal", "true");
    form.setAttribute("aria-labelledby", id);
    form.style.cssText =
      "width:min(var(--dxw-dialog-width,420px),calc(100vw - 32px));box-sizing:border-box;" +
      "display:grid;gap:12px;padding:16px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:var(--dxw-dialog-radius,10px);color:var(--dxw-toolbar-fg,#3c4043);" +
      "background:var(--dxw-popover-bg,#fff);box-shadow:var(--dxw-popover-shadow,0 8px 28px rgba(0,0,0,.24));" +
      "font:13px system-ui,sans-serif;";

    const title = doc.createElement("strong");
    title.id = id;
    title.textContent = options.title;
    title.style.cssText = "font-size:15px;line-height:1.35;";
    form.appendChild(title);

    const initial = /^#?([0-9a-f]{6})$/i.exec(options.value ?? "")?.[1] ?? "4472C4";
    const colorLabel = doc.createElement("label");
    colorLabel.textContent = "Color";
    colorLabel.style.cssText = "display:grid;gap:5px;color:var(--dxw-toolbar-muted,#5f6368);font-size:12px;";
    const colorRow = doc.createElement("span");
    colorRow.style.cssText = "display:grid;grid-template-columns:42px 1fr;gap:7px;";
    const picker = doc.createElement("input");
    picker.type = "color";
    picker.value = `#${initial}`;
    picker.setAttribute("aria-label", `${options.title} color picker`);
    picker.style.cssText =
      "width:42px;height:34px;box-sizing:border-box;padding:2px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:6px;background:var(--dxw-dialog-field-bg,var(--dxw-popover-bg,#fff));";
    const field = doc.createElement("input");
    field.type = "text";
    field.value = `#${initial.toUpperCase()}`;
    field.required = true;
    field.pattern = "#?[0-9a-fA-F]{6}";
    field.setAttribute("aria-label", `${options.title} color`);
    field.style.cssText =
      "width:100%;min-height:34px;box-sizing:border-box;padding:7px 9px;" +
      "border:1px solid var(--dxw-toolbar-border,#dadce0);border-radius:6px;outline:none;" +
      "color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-dialog-field-bg,var(--dxw-popover-bg,#fff));" +
      "font:13px system-ui,sans-serif;";
    let noFill: HTMLInputElement | null = null;
    picker.addEventListener("input", () => {
      field.value = picker.value.toUpperCase();
      if (noFill) noFill.checked = false;
    });
    field.addEventListener("input", () => {
      if (noFill) noFill.checked = false;
      const match = /^#?([0-9a-f]{6})$/i.exec(field.value.trim());
      if (match) picker.value = `#${match[1]}`;
    });
    colorRow.append(picker, field);
    colorLabel.appendChild(colorRow);
    form.appendChild(colorLabel);

    if (options.allowNone) {
      const noneLabel = doc.createElement("label");
      noneLabel.style.cssText = "display:flex;align-items:center;gap:7px;color:var(--dxw-toolbar-fg,#3c4043);font-size:12px;";
      noFill = doc.createElement("input");
      noFill.type = "checkbox";
      noFill.checked = options.value === null;
      noFill.setAttribute("aria-label", "No fill");
      noneLabel.append(noFill, doc.createTextNode("No fill"));
      form.appendChild(noneLabel);
    }

    const actions = doc.createElement("div");
    actions.style.cssText = "display:flex;justify-content:flex-end;gap:7px;";
    const cancel = doc.createElement("button");
    cancel.type = "button";
    cancel.textContent = "Cancel";
    cancel.style.cssText =
      "min-height:32px;padding:0 12px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:7px;color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-popover-bg,#fff);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    const submit = doc.createElement("button");
    submit.type = "submit";
    submit.textContent = "Apply";
    submit.style.cssText =
      "min-height:32px;padding:0 14px;border:1px solid transparent;border-radius:7px;" +
      "color:var(--dxw-accent-fg,#fff);background:var(--dxw-accent,#1a73e8);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    actions.append(cancel, submit);
    form.appendChild(actions);
    backdrop.appendChild(form);
    doc.body.appendChild(backdrop);

    let finished = false;
    const finish = (value: string | null | undefined) => {
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
        finish(undefined);
      }
    };
    doc.addEventListener("keydown", onKeyDown, true);
    backdrop.addEventListener("mousedown", (event) => {
      if (event.target === backdrop) finish(undefined);
    });
    cancel.addEventListener("click", () => finish(undefined));
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      if (noFill?.checked) {
        finish(null);
        return;
      }
      const color = /^#?([0-9a-f]{6})$/i.exec(field.value.trim());
      if (!color) {
        form.reportValidity();
        return;
      }
      finish(`#${color[1].toUpperCase()}`);
    });
    requestAnimationFrame(() => {
      (noFill?.checked ? noFill : field).focus({ preventScroll: true });
      if (!noFill?.checked) field.select();
    });
  });
}

/** Opens structured controls for a DrawingML line or shape outline. */
export function requestLineStyleDialog(
  anchor: HTMLElement,
  value: LineStyleDialogValue,
  titleText = "Line style",
): Promise<LineStyleDialogValue | null> {
  const doc = anchor.ownerDocument;
  const win = doc.defaultView;
  return new Promise((resolve) => {
    const fieldName = titleText === "Outline" ? "Outline" : "Line";
    const id = `dxw-line-style-dialog-${++dialogSequence}`;
    const previousFocus = doc.activeElement instanceof HTMLElement ? doc.activeElement : null;
    const backdrop = doc.createElement("div");
    backdrop.className = "dxw-input-dialog-backdrop";
    backdrop.dataset.dxwLineStyleDialog = "";
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
        const token = computed.getPropertyValue(property);
        if (token) backdrop.style.setProperty(property, token);
      }
    }

    const form = doc.createElement("form");
    form.className = "dxw-input-dialog dxw-line-style-dialog";
    form.setAttribute("role", "dialog");
    form.setAttribute("aria-modal", "true");
    form.setAttribute("aria-labelledby", id);
    form.style.cssText =
      "width:min(var(--dxw-dialog-width,420px),calc(100vw - 32px));box-sizing:border-box;" +
      "display:grid;gap:12px;padding:16px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:var(--dxw-dialog-radius,10px);color:var(--dxw-toolbar-fg,#3c4043);" +
      "background:var(--dxw-popover-bg,#fff);box-shadow:var(--dxw-popover-shadow,0 8px 28px rgba(0,0,0,.24));" +
      "font:13px system-ui,sans-serif;";

    const title = doc.createElement("strong");
    title.id = id;
    title.className = "dxw-input-dialog-title";
    title.textContent = titleText;
    title.style.cssText = "font-size:15px;line-height:1.35;";
    form.appendChild(title);

    const fieldStyle =
      "width:100%;min-height:34px;box-sizing:border-box;padding:7px 9px;" +
      "border:1px solid var(--dxw-toolbar-border,#dadce0);border-radius:6px;outline:none;" +
      "color:var(--dxw-toolbar-fg,#3c4043);background:var(--dxw-dialog-field-bg,var(--dxw-popover-bg,#fff));" +
      "font:13px system-ui,sans-serif;";
    const labelStyle = "display:grid;gap:5px;color:var(--dxw-toolbar-muted,#5f6368);font-size:12px;";

    const colorLabel = doc.createElement("label");
    colorLabel.textContent = "Color";
    colorLabel.style.cssText = labelStyle;
    const colorRow = doc.createElement("span");
    colorRow.style.cssText = "display:grid;grid-template-columns:42px 1fr;gap:7px;";
    const colorPicker = doc.createElement("input");
    colorPicker.type = "color";
    colorPicker.value = value.color;
    colorPicker.setAttribute("aria-label", `${fieldName} color picker`);
    colorPicker.style.cssText =
      "width:42px;height:34px;box-sizing:border-box;padding:2px;border:1px solid var(--dxw-toolbar-border,#dadce0);" +
      "border-radius:6px;background:var(--dxw-dialog-field-bg,var(--dxw-popover-bg,#fff));";
    const colorField = doc.createElement("input");
    colorField.type = "text";
    colorField.value = value.color.toUpperCase();
    colorField.required = true;
    colorField.pattern = "#?[0-9a-fA-F]{6}";
    colorField.setAttribute("aria-label", `${fieldName} color`);
    colorField.style.cssText = fieldStyle;
    colorPicker.addEventListener("input", () => { colorField.value = colorPicker.value.toUpperCase(); });
    colorField.addEventListener("input", () => {
      const match = /^#?([0-9a-f]{6})$/i.exec(colorField.value.trim());
      if (match) colorPicker.value = `#${match[1]}`;
    });
    colorRow.append(colorPicker, colorField);
    colorLabel.appendChild(colorRow);
    form.appendChild(colorLabel);

    const widthLabel = doc.createElement("label");
    widthLabel.textContent = "Width (pixels)";
    widthLabel.style.cssText = labelStyle;
    const widthField = doc.createElement("input");
    widthField.type = "number";
    widthField.min = "0.25";
    widthField.step = "any";
    widthField.required = true;
    widthField.value = String(Number(value.width.toFixed(2)));
    widthField.setAttribute("aria-label", `${fieldName} width in pixels`);
    widthField.style.cssText = fieldStyle;
    widthLabel.appendChild(widthField);
    form.appendChild(widthLabel);

    const styleLabel = doc.createElement("label");
    styleLabel.textContent = "Style";
    styleLabel.style.cssText = labelStyle;
    const styleField = doc.createElement("select");
    styleField.setAttribute("aria-label", `${fieldName} style`);
    styleField.style.cssText = fieldStyle;
    for (const [style, label] of [["solid", "Solid"], ["dashed", "Dashed"], ["dotted", "Dotted"]] as const) {
      const option = doc.createElement("option");
      option.value = style;
      option.textContent = label;
      option.selected = value.style === style;
      styleField.appendChild(option);
    }
    styleLabel.appendChild(styleField);
    form.appendChild(styleLabel);

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
    submit.textContent = "Apply";
    submit.style.cssText =
      "min-height:32px;padding:0 14px;border:1px solid transparent;border-radius:7px;" +
      "color:var(--dxw-accent-fg,#fff);background:var(--dxw-accent,#1a73e8);" +
      "cursor:pointer;font:600 12px system-ui,sans-serif;";
    actions.append(cancel, submit);
    form.appendChild(actions);
    backdrop.appendChild(form);
    doc.body.appendChild(backdrop);

    let finished = false;
    const finish = (result: LineStyleDialogValue | null) => {
      if (finished) return;
      finished = true;
      doc.removeEventListener("keydown", onKeyDown, true);
      backdrop.remove();
      previousFocus?.focus({ preventScroll: true });
      resolve(result);
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
      const color = /^#?([0-9a-f]{6})$/i.exec(colorField.value.trim());
      const width = Number(widthField.value);
      if (!color || !Number.isFinite(width) || width <= 0) {
        form.reportValidity();
        return;
      }
      finish({
        color: `#${color[1].toUpperCase()}`,
        width,
        style: styleField.value as LineStyleDialogValue["style"],
      });
    });
    requestAnimationFrame(() => {
      colorField.focus({ preventScroll: true });
      colorField.select();
    });
  });
}
