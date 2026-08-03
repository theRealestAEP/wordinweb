import { localName, type StableIds, type XmlElement } from "@wordinweb/core";

export function blockRef(id: number): string {
  return `block:${id}`;
}

export function runRef(id: number): string {
  return `run:${id}`;
}

export function objectRef(runId: number, index: number): string {
  return `object:${runId}:${index}`;
}

export function blockObjectRef(blockId: number, runIndex: number, contentIndex: number): string {
  return `object:block:${blockId}:${runIndex}:${contentIndex}`;
}

export function viewRef(revision: string, kind: "block" | "run" | "object", index: number): string {
  return `view:${encodeURIComponent(revision)}:${kind}:${index}`;
}

export function assetRef(index: number): string {
  return `asset:${index}`;
}

function idFromRef(ref: unknown, prefix: "block" | "run"): number {
  if (typeof ref !== "string") throw new Error(`${prefix} reference must be a string`);
  const match = new RegExp(`^${prefix}:(\\d+)$`).exec(ref);
  if (!match) throw new Error(`Invalid ${prefix} reference`);
  return Number(match[1]);
}

function assertElement(ids: StableIds, id: number, names: readonly string[], label: string): XmlElement {
  const element = ids.elOf(id);
  if (!element || !names.includes(localName(element.name))) throw new Error(`Unknown ${label} reference`);
  return element;
}

export function resolveBlockRef(ids: StableIds, ref: unknown): number {
  const id = idFromRef(ref, "block");
  assertElement(ids, id, ["p", "tbl"], "block");
  return id;
}

export function resolveParagraphRef(ids: StableIds, ref: unknown): number {
  const id = idFromRef(ref, "block");
  assertElement(ids, id, ["p"], "paragraph");
  return id;
}

export function resolveRunRef(ids: StableIds, ref: unknown): number {
  const id = idFromRef(ref, "run");
  assertElement(ids, id, ["r"], "run");
  return id;
}

export function parseObjectRef(ref: string): { runId?: number; objectIndex?: number } {
  const run = /^object:(\d+):(\d+)$/.exec(ref);
  if (run) return { runId: Number(run[1]), objectIndex: Number(run[2]) };
  if (/^object:block:\d+:\d+:\d+$/.test(ref) || /^view:[^:]+:object:\d+$/.test(ref)) return {};
  throw new Error("Invalid object reference");
}
