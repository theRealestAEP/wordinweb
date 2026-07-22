/**
 * Client entry surface for @wordinweb/collab. Currently exposes the shared
 * wire types and the transform (used by reconciliation replay); the
 * optimistic client engine (pending queue, confirmed snapshot,
 * rollback-replay — plan doc 03 Stage B) lands in a later phase.
 */
export * from "./intents.js";
export { transformIntent, transformPosition } from "./transform.js";
