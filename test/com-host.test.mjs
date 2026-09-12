// Regression test for the resident COM host: recovery, serialization, protocol framing.
// Run: node test/com-host.test.mjs
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { comHost } = require("../mcp/dist/client/com-host.js");

const results = [];
function check(name, ok, detail) {
  results.push({ name, ok });
  console.log((ok ? "PASS " : "FAIL ") + name + (detail ? "  " + detail : ""));
}
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const first = await comHost.invoke("ping", {});
check("cold ping succeeds", first.success === true, JSON.stringify(first).slice(0, 80));

const warmStart = Date.now();
await comHost.invoke("ping", {});
const warmMs = Date.now() - warmStart;
check("warm ping under 100ms", warmMs < 100, "warm=" + warmMs + "ms");

// Concurrent calls must be serialized by the client and all resolve correctly.
const concurrent = await Promise.all([
  comHost.invoke("ping", {}),
  comHost.invoke("wireCheck", {}),
  comHost.invoke("ping", {}),
  comHost.invoke("wireCheck", {}),
  comHost.invoke("ping", {})
]);
check("5 concurrent calls all succeed", concurrent.every((r) => r.success === true), "ok=" + concurrent.filter((r) => r.success).length + "/5");

// Kill the host process and confirm the next call transparently respawns it.
const child = comHost.child;
check("host process is attached", !!child, child ? "pid=" + child.pid : "no child");
try { child.kill(); } catch { /* already gone */ }
await sleep(400);
let recovered = null;
try {
  recovered = await comHost.invoke("ping", {});
} catch (error) {
  recovered = { success: false, error: String(error.message) };
}
check("host recovers after being killed", recovered.success === true, JSON.stringify(recovered).slice(0, 100));

// An unknown action must come back as a structured failure, never as a silent success.
const unknown = await comHost.invoke("__definitely_not_an_action", {});
check("unknown action fails loudly", unknown.success === false && typeof unknown.error === "string", JSON.stringify(unknown).slice(0, 100));

await comHost.stop();
const failed = results.filter((r) => !r.ok).length;
console.log(failed === 0 ? "COM HOST TESTS OK (" + results.length + ")" : "COM HOST TESTS FAILED (" + failed + "/" + results.length + ")");
process.exit(failed === 0 ? 0 : 1);
