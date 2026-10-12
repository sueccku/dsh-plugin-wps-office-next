/**
 * Input: none (pure Node built-ins)
 * Output: the Node executable an MCP stdio child must be spawned with
 * Pos: shared by plugin.js (publishes WPS_OFFICE_MCP_NODE) and cordis.patch.yml (reads it).
 *      为什么需要它：DSH Desktop 给 Host 的 env 里 DSH_DESKTOP_NODE_EXECUTABLE 是空的
 *      （宿主以 bin=void 0 启动 Host），而这个变量名又会被 dsh-mcp-client 的
 *      scrubbedParentEnv() 整体剥掉，所以裸 `node` 在不装 Node.js 的机器上 spawn ENOENT，
 *      整行 MCP 条目 fiberPhase=failed、95 个工具全不注册。这里改用随 DSH 打包的真实
 *      node.exe（真 Node，不是 Electron-as-Node），找不到才退到 process.execPath。
 *      一旦我被修改，请更新我的头部注释。
 */
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

/** Windows ships node.exe; every other platform ships a bare `node`. */
const EXECUTABLE_NAME = process.platform === 'win32' ? 'node.exe' : 'node';

/** Path below the DSH runtimes root down to the pinned Node executable. */
const NODE_BELOW_RUNTIMES = ['dsh-primary-runtime', 'dependencies', 'node', 'bin', EXECUTABLE_NAME];

/**
 * The DSH-managed runtimes root: `$DSH_HOME/dsh-runtimes`, defaulting to `~/.dsh/dsh-runtimes`.
 * DSH sets `DSH_HOME` for the Host; the homedir default is what the installer uses when unset.
 */
function runtimesRoot() {
  const home = process.env.DSH_HOME;
  return join(home && home.trim() !== '' ? home : join(homedir(), '.dsh'), 'dsh-runtimes');
}

/**
 * Every candidate Node executable in preference order. Candidates are not existence-checked here.
 * @returns {string[]} absolute paths, ending with `process.execPath`.
 */
export function nodeExecutableCandidates() {
  const candidates = [];
  // 1. Explicit override: lets a user pin any Node build without editing the bundle patch.
  const override = process.env.WPS_OFFICE_MCP_NODE;
  if (override && override.trim() !== '') candidates.push(override.trim());
  // 2. The Node shipped inside the DSH-managed runtime — a real Node, present on Node-less machines.
  candidates.push(join(runtimesRoot(), ...NODE_BELOW_RUNTIMES));
  // 3. Whatever is running the Host: Electron's binary in Node mode, or Node itself for the CLI.
  candidates.push(process.execPath);
  return candidates.filter((candidate, index) => candidate && candidates.indexOf(candidate) === index);
}

/**
 * Resolve the Node executable for the MCP stdio child.
 * @returns {string} the first candidate that exists on disk, else the ambient `node` name.
 */
export function resolveNodeExecutable() {
  return nodeExecutableCandidates().find((candidate) => existsSync(candidate)) ?? 'node';
}
