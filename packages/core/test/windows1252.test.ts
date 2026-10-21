import { expect, it } from "vitest";
import { windows1252 } from "../src/windows1252.js";

it("preserves Windows punctuation, Latin characters and unassigned controls", () => {
  const bytes = Uint8Array.of(0x41, 0x80, 0x91, 0x92, 0x93, 0x94, 0x96, 0x97, 0x85, 0x8c, 0x9c, 0x9f, 0xe9, 0x81, 0x8d, 0x90, 0x9d);
  expect(windows1252.decode(bytes)).toBe("A€‘’“”–—…ŒœŸé\u0081\u008d\u0090\u009d");
});
