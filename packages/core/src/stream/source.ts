/** Random access to a .docx's bytes: a URL read with HTTP range requests, or
 * bytes already in memory. */
export interface ByteSource {
  size(): Promise<number>;
  /** The bytes [start, end). */
  read(start: number, end: number): Promise<Uint8Array>;
}

/**
 * A .docx at a URL whose server honours Range requests. A caller with the
 * file on disk serves it under a URL of its own: over the DevTools protocol,
 * Fetch.enable on the page and fulfill each paused request from the file
 * (render/print.ts has the shape of that loop).
 */
export function fetchSource(url: string): ByteSource {
  let size: number | undefined;
  const range = async (start: number, end: number): Promise<Response> => {
    const response = await fetch(url, { headers: { Range: `bytes=${start}-${end - 1}` } });
    if (response.status !== 206) throw new Error(`${url}: range request answered ${response.status}; the server must honour Range`);
    return response;
  };
  return {
    async size() {
      if (size === undefined) {
        const response = await range(0, 1);
        const total = /\/(\d+)$/.exec(response.headers.get("Content-Range") ?? "");
        if (!total) throw new Error(`${url}: no Content-Range on a range response`);
        size = Number(total[1]);
      }
      return size;
    },
    async read(start, end) {
      if (end <= start) return new Uint8Array(0);
      return new Uint8Array(await (await range(start, end)).arrayBuffer());
    },
  };
}

export function bytesSource(bytes: Uint8Array): ByteSource {
  return {
    size: async () => bytes.length,
    read: async (start, end) => bytes.subarray(start, end),
  };
}
