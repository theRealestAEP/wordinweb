# @wordinweb/server

The Node collaboration host for WordInWeb.

This package is separate from `wordinweb` and `wordinweb/collab`. Browser
applications do not install or bundle it.

## Run the included server

```sh
npm install @wordinweb/server
npx wordinweb-collab-server
```

The command starts the HTTP and WebSocket server on port `1234`. Set `PORT` to
change the port.

The server accepts encrypted and plaintext rooms by default. Set
`WW_ENCRYPTED_ONLY=1` when the deployment must reject plaintext documents:

```sh
WW_ENCRYPTED_ONLY=1 npx wordinweb-collab-server
```

## Embed the server

```ts
import { CollabHub, startZeroCustodyServer } from "@wordinweb/server";
```

See the [anonymous sharing example][example] for the encrypted deployment,
configuration variables, and operational limits.

[example]: https://github.com/theRealestAEP/wordinweb/tree/main/examples/anon-share

## License

MIT. See `LICENSE`.
