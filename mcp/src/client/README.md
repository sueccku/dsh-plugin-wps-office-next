# client

Windows-only WPS communication client.

- wps-client.ts: spawns the COM bridge and normalizes responses.

The macOS/Linux polling transport (mac-poll-server.ts, wps-keepalive.ts) was removed in the
Windows/COM-only fork; see THIRD_PARTY_NOTICES.md.
