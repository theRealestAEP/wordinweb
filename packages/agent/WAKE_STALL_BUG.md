# Bug: Codex document-agent wake blocked by macOS Seatbelt

## Status

- Reproduced on macOS on August 4, 2026.
- Affected flow: the resident Codex wake host in `@wordinweb/agent`.
- User impact: private chat shows activity, while the agent sends no edit or reply.
- Root-cause confidence: confirmed.

## Summary

The WordInWeb bridge connects and receives the private message. The message starts a resident Codex turn.

The bridge socket and Codex workspace now share the same private directory. macOS Seatbelt still treats a Unix-socket connection as network access. The resident turn used the `workspaceWrite` sandbox with its default network policy, so Seatbelt returned `EPERM` when the agent ran `sync`.

The wake completion guard then reported `turn_completed_without_chat`, because the Codex turn ended before a WordInWeb chat command succeeded.

## Evidence

The failed session returned this error:

```text
connect EPERM .../wiw-FFTvlHCoCiBH2FEFcND-rA/bridge.sock
```

Each private message followed this sequence:

1. The bridge started a Codex turn.
2. The agent ran `sync`.
3. macOS Seatbelt rejected the Unix-socket connection.
4. The Codex turn completed without a WordInWeb chat command.
5. The bridge reported `turn_completed_without_chat`.

The bridge remained healthy. Manual commands from the parent Codex task reached the same socket.

## Root cause

`thread/start` selects the `workspaceWrite` sandbox for the resident Codex thread. Each `turn/start` request inherited the default network policy.

For the Codex app-server protocol, `workspaceWrite` uses a boolean `networkAccess` field. A Unix-socket connection on macOS requires this permission even when the socket sits inside the writable workspace.

## Fix

Set the resident turn policy explicitly:

```ts
sandboxPolicy: {
  type: "workspaceWrite",
  networkAccess: true,
},
```

The same `sync` command then returned exit code `0` with the document revision and roster.

## Regression coverage

The daemon test checks that every resident `turn/start` request includes the network-enabled `workspaceWrite` policy. The fake Codex process continues to test the wake ID, session command, reply, timeout, and resident-session recovery paths.

The confirmed macOS reproduction supplies the Seatbelt integration check. The unit test locks the app-server request shape without requiring Codex authentication in CI.

## Acceptance criteria

- A resident Codex turn can connect to the bridge socket on macOS.
- The first `sync` command returns the document revision and roster.
- Every resident `turn/start` request enables network access for `workspaceWrite`.
- The task can inspect, edit, and reply through WordInWeb chat.
- The wake completion guard records a successful chat for the current wake ID.
