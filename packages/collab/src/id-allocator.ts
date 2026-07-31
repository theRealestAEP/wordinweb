/**
 * Allocates carried node ids from a client-specific range.
 *
 * Keep one allocator for the lifetime of a client session. A reconnect creates
 * a new connection, but it must not restart this sequence and reuse ids that
 * already exist in the document.
 */
export class CarriedIdAllocator {
  private counter = 0;
  private readonly base: number;

  constructor(clientId: string) {
    let hash = 0;
    for (let i = 0; i < clientId.length; i++) hash = (hash * 31 + clientId.charCodeAt(i)) >>> 0;
    this.base = 1_000_000_000 + (hash % 100_000) * 10_000_000;
  }

  next(): number {
    return this.base + this.counter++;
  }
}
