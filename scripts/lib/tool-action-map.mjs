// Shared parser: map each tool to the bridge action it drives and to the argument keys it sends.
//
// The handler source is the only place that says what a tool really sends: most tools pass
// their schema parameters through, but some rename them in code (createChart sends camelCase
// dataRange while its schema says data_range). Both the parameter-contract sweep and the
// operation-spec bootstrap read this module, so the two can never drift apart.
//
// 一旦我被修改，请更新我的头部注释。
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

/**
 * Keys of the object literal handed to executeMethod, or null when it cannot be read statically.
 * `from` must be the offset just past the action string, so that a result-type generic such as
 * executeMethod<{ success: boolean }>(...) is never mistaken for the argument object.
 */
export function sentKeys(source, from) {
  const after = source.slice(from);
  // The argument object may be preceded by a comment, and may be a variable holding it.
  const m = /^\s*,\s*(?:\/\/[^\n]*\n\s*)*\{/.exec(after);
  if (!m) {
    const idm = /^\s*,\s*([A-Za-z_$][\w$]*)\s*(?:,|\))/.exec(after);
    if (!idm) return null;
    const decl = new RegExp("(?:const|let|var)\\s+" + idm[1] + "\\s*(?::[^=]*)?=\\s*\\{").exec(source);
    if (!decl) return null;
    return collectKeys(source, decl.index + decl[0].length - 1);
  }
  return collectKeys(after, m[0].length - 1);
}

export function collectKeys(text, braceAt) {
  const after = text;
  let depth = 0;
  let end = -1;
  for (let i = braceAt; i < after.length; i++) {
    const ch = after[i];
    if (ch === "{") depth++;
    else if (ch === "}") { depth--; if (depth === 0) { end = i; break; } }
  }
  if (end < 0) return null;
  const body = after.slice(braceAt + 1, end);
  const keys = new Set();
  let d = 0;
  let segment = "";
  const segments = [];
  for (const ch of body) {
    if ("{[(".includes(ch)) d++;
    else if ("}])".includes(ch)) d--;
    if (ch === "," && d === 0) { segments.push(segment); segment = ""; continue; }
    segment += ch;
  }
  segments.push(segment);
  for (const raw of segments) {
    const s = raw.replace(/\/\/[^\n]*/g, "").trim();
    if (!s) continue;
    if (s.startsWith("...")) return null; // a spread hides the real key set
    const m = /^([A-Za-z_$][\w$]*)\s*(?::([\s\S]*))?$/.exec(s);
    if (!m) return null;
    keys.add(m[1]);
  }
  return keys;
}

// Some tools go through a wpsClient helper (getRangeData/setRangeData) rather than executeMethod.
// Their sent keys are fixed inside wps-client.ts, so read them from there instead of guessing.
export const CLIENT_HELPERS = { getRangeData: 'getRangeData', setRangeData: 'setRangeData' };
export function clientHelperKeys() {
  const map = new Map();
  let src = '';
  try { src = readFileSync('mcp/src/client/wps-client.ts', 'utf8'); } catch { return map; }
  for (const m of src.matchAll(/invokeAction(?:<[^>]*>)?\(\s*'([A-Za-z0-9]+)'\s*,\s*\{([^}]*)\}/g)) {
    const keys = [...m[2].matchAll(/([A-Za-z_$][\w$]*)\s*:/g)].map((k) => k[1]);
    map.set(m[1], keys);
  }
  return map;
}
export const HELPER_KEYS = clientHelperKeys();
export function analyseToolSource() {
  const map = new Map();
  const unparsed = [];
  const walk = (dir) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) { walk(p); continue; }
      if (!entry.name.endsWith(".ts")) continue;
      const src = readFileSync(p, "utf8");
      const names = [...src.matchAll(/name:\s*'(wps_[a-z0-9_]+)'/g)];
      for (let i = 0; i < names.length; i++) {
        const from = names[i].index;
        const to = i + 1 < names.length ? names[i + 1].index : src.length;
        const block = src.slice(from, to);
        // Only a call with an explicitly named action can be traced without executing it.
        const calls = [...block.matchAll(/executeMethod[\s\S]{0,600}?'([A-Za-z][A-Za-z0-9]*)'/g)];
        if (calls.length !== 1) {
          const helper = Object.keys(CLIENT_HELPERS).find((h) => block.includes('wpsClient.' + h + '('));
          if (helper) {
            const action = CLIENT_HELPERS[helper];
            map.set(names[i][1], { action, keys: HELPER_KEYS.get(action) || ['sheet', 'range', 'data'], mentioned: () => true });
            continue;
          }
          unparsed.push({ tool: names[i][1], reason: calls.length + " executeMethod calls", action: calls.length ? calls[0][1] : null });
          continue;
        }
        const keys = sentKeys(block, calls[0].index + calls[0][0].length);
        if (!keys) { unparsed.push({ tool: names[i][1], reason: "argument object not statically readable", action: calls[0][1] }); continue; }
        // A schema key is dropped only if the handler never mentions it at all; renaming it (for
        // example filePath -> path) is a legitimate adaptation, not a defect.
        const mentioned = (key) => new RegExp("\\b" + key + "\\b").test(block);
        map.set(names[i][1], { action: calls[0][1], keys: [...keys], mentioned });
      }
    }
  };
  walk("mcp/src/tools");
  return { map, unparsed };
}
