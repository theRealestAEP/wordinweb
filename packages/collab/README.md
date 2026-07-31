# @wordinweb/collab

The transport-independent collaboration engine for WordInWeb.

React applications should usually use the `wordinweb/collab` entry. It
includes the React binding and the browser collaboration client:

```tsx
import { CollabEditor, IndexedDbBundleStore } from "wordinweb/collab";
```

Install this lower-level package when you need its client or authoritative
session APIs directly:

```sh
npm install @wordinweb/collab
```

```ts
import { CollabConnection } from "@wordinweb/collab/client";
import { DocumentSession } from "@wordinweb/collab/server";
```

The client entry contains browser collaboration logic. The server entry
contains the transport-independent document state machine. This package does
not contain an HTTP or WebSocket host.

Install `@wordinweb/server` separately when you need the provided Node server.
See the [anonymous sharing example][example] for a complete encrypted
deployment.

[example]: https://github.com/theRealestAEP/wordinweb/tree/main/examples/anon-share

## License

MIT. See `LICENSE`.
