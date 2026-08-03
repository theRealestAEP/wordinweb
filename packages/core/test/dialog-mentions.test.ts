// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { requestTextInputDialog } from "../src/edit/dialog.js";

describe("comment mention shortcuts", () => {
  it("inserts a selected collaborator name into the comment field", async () => {
    const anchor = document.createElement("div");
    document.body.appendChild(anchor);
    const result = requestTextInputDialog(anchor, {
      title: "New comment",
      multiline: true,
      mentions: ["Alex", "Review agent"],
    });
    const button = [...document.querySelectorAll("button")].find((item) => item.textContent === "@Review agent");
    button?.click();
    const field = document.querySelector("textarea") as HTMLTextAreaElement;
    expect(field.value).toBe("@Review agent ");
    field.value += "please review this section";
    (document.querySelector("form") as HTMLFormElement).dispatchEvent(new SubmitEvent("submit", { bubbles: true, cancelable: true }));
    await expect(result).resolves.toBe("@Review agent please review this section");
  });
});
