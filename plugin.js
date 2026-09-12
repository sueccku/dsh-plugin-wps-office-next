/**
 * Input: DSH cordis loader
 * Output: skill registrations, plus the resolved MCP entry path in the process env
 * Pos: DSH entry for the Windows/COM-only WPS bundle.
 *      这个模块必须排在 cordis.patch.yml 的 MCP 条目之前：它在顶层把包内绝对路径
 *      写进 WPS_OFFICE_MCP_ENTRY，因此用户不需要再配置 WPS_SKILLS_ROOT。
 *      一旦我被修改，请更新我的头部注释。
 */
import { existsSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const name = 'wps-office-next';
export const inject = ['skills'];

const here = dirname(fileURLToPath(import.meta.url));

/** Absolute path to the bundled MCP server entry. */
export const mcpEntry = join(here, 'mcp', 'dist', 'index.js');

/** Absolute path to the resident COM host script. */
export const hostScript = join(here, 'host', 'wps-com-host.ps1');

// Publish the resolved paths for the MCP client entry and the spawned server.
// cordis.patch.yml reads WPS_OFFICE_MCP_ENTRY when the mcp entry activates.
process.env.WPS_OFFICE_MCP_ENTRY = process.env.WPS_OFFICE_MCP_ENTRY || mcpEntry;
process.env.WPS_OFFICE_HOST_SCRIPT = process.env.WPS_OFFICE_HOST_SCRIPT || hostScript;

/** Skills shipped by this bundle, in registration order. */
const SKILLS = [
  { name: 'wps-office-next', dir: 'wps-office-next' },
  { name: 'wps-excel', dir: 'wps-excel' },
  { name: 'wps-word', dir: 'wps-word' },
  { name: 'wps-ppt', dir: 'wps-ppt' },
];

/** Split YAML frontmatter from a skill body. Only the scalar keys this bundle uses. */
function readFrontmatter(raw) {
  const match = /^---\r?\n([\s\S]*?)\r?\n---\r?\n?/.exec(raw);
  if (!match) return { meta: {}, body: raw };
  const meta = {};
  for (const line of match[1].split(/\r?\n/)) {
    const pair = /^([A-Za-z][A-Za-z0-9_-]*):\s*(.*)$/.exec(line);
    if (pair) meta[pair[1]] = pair[2].trim();
  }
  return { meta, body: raw.slice(match[0].length) };
}

export function apply(ctx) {
  const disposers = [];

  if (!existsSync(mcpEntry)) {
    throw new Error('dsh-plugin-wps-office-next: MCP entry not found at ' + mcpEntry + '; the package is missing its mcp/dist build output');
  }

  for (const entry of SKILLS) {
    const skillDir = join(here, 'skills', entry.dir);
    const skillFile = join(skillDir, 'SKILL.md');
    if (!existsSync(skillFile)) continue;
    const { meta, body } = readFrontmatter(readFileSync(skillFile, 'utf8'));
    disposers.push(
      ctx.skills.register({
        name: meta.name || entry.name,
        description: meta.description || entry.name,
        whenToUse: meta.whenToUse,
        source: 'runtime',
        resourceBase: { kind: 'directory', path: skillDir },
        content: body.trim(),
      })
    );
  }

  return () => {
    for (const dispose of disposers) {
      try { dispose(); } catch { /* already disposed */ }
    }
  };
}
