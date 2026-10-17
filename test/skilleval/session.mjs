// 会话轨迹（session log）读取：给 P1 的 §3.2 验证与 D1 的指标统计用。
//
// DSH 把会话写成 $DSH_HOME/sessions/<encoded-cwd>/<session-id>/session.v<N>.jsonl.zstd，
// 每次 flush 追加**一个 zstd 帧**，而 node 的 zstdDecompressSync 只解第一帧 —— 所以必须按帧边界切。
// 这段逻辑与 scripts/e2e.mjs 同源（那边是脚本内联的），这里抽出来给评测复用。
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { zstdDecompressSync } from 'node:zlib';

const ZSTD_MAGIC = Buffer.from([0x28, 0xb5, 0x2f, 0xfd]);
const SESSION_FILE = /^session\.v\d+\.jsonl\.zstd$/;

/** 解出文件里的所有 zstd 帧并拼起来。 */
export function decodeZstdFrames(buf) {
  const starts = [];
  for (let i = 0; i + 4 <= buf.length; i++) if (buf.compare(ZSTD_MAGIC, 0, 4, i, i + 4) === 0) starts.push(i);
  if (!starts.length) throw new Error('no zstd frame found');
  const bounds = starts.concat([buf.length]);
  const out = [];
  let i = 0;
  while (i < bounds.length - 1) {
    let done = false;
    for (let j = i + 1; j < bounds.length && !done; j++) {
      try {
        out.push(zstdDecompressSync(buf.subarray(bounds[i], bounds[j])));
        i = j;
        done = true;
      } catch {
        // 这个边界落在帧内部，试下一个
      }
    }
    if (!done) throw new Error('undecodable frame at offset ' + bounds[i]);
  }
  return Buffer.concat(out);
}

/** 读一个会话文件为记录数组。 */
export function loadRecords(file) {
  const text = decodeZstdFrames(readFileSync(file)).toString('utf8');
  return text
    .split(/\r?\n/)
    .filter((line) => line.trim())
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        return { __unparsed: line };
      }
    });
}

/** 列出 $DSH_HOME/sessions 下的所有会话文件（两种目录深度都接受）。 */
export function listSessionFiles(dshHome) {
  const root = join(dshHome, 'sessions');
  const files = [];
  if (!existsSync(root)) return files;
  for (const workspace of readdirSync(root)) {
    const workspaceDir = join(root, workspace);
    let entries;
    try {
      entries = readdirSync(workspaceDir);
    } catch {
      continue;
    }
    for (const entry of entries) {
      if (SESSION_FILE.test(entry)) files.push(join(workspaceDir, entry));
    }
    for (const entry of entries) {
      let nested;
      try {
        nested = readdirSync(join(workspaceDir, entry));
      } catch {
        continue;
      }
      for (const inner of nested) if (SESSION_FILE.test(inner)) files.push(join(workspaceDir, entry, inner));
    }
  }
  return files;
}

export function flatten(content) {
  if (content == null) return '';
  if (typeof content === 'string') return content;
  if (Array.isArray(content)) return content.map(flatten).join(' | ');
  if (typeof content === 'object') {
    if (typeof content.text === 'string') return content.text;
    if (content.content) return flatten(content.content);
  }
  return '';
}

/**
 * 找一次运行对应的会话文件。
 * 用"第一条 user/message 的正文 == 我们发给 dsh 的任务"来认，比按时间猜可靠；
 * sinceMs 只是第一道筛，避免去解码几百个老会话。
 */
export function findSession({ dshHome, task, sinceMs, slackMs = 15000 }) {
  const candidates = [];
  for (const file of listSessionFiles(dshHome)) {
    let stat;
    try {
      stat = statSync(file);
    } catch {
      continue;
    }
    if (sinceMs && stat.mtimeMs < sinceMs - slackMs) continue;
    let records;
    try {
      records = loadRecords(file);
    } catch {
      continue;
    }
    candidates.push({ file, mtimeMs: stat.mtimeMs, records });
  }
  candidates.sort((a, b) => b.mtimeMs - a.mtimeMs);
  const wanted = String(task).trim();
  for (const candidate of candidates) {
    const firstUser = candidate.records.find((r) => r.type === 'user/message');
    const text = firstUser ? flatten(firstUser.data && firstUser.data.content).trim() : '';
    if (text === wanted || (wanted.length > 40 && text.startsWith(wanted.slice(0, 40)))) {
      return candidate;
    }
  }
  return null;
}

/** 把记录压成"这次运行做了什么"：工具调用、工具结果、助手文本、用户消息、真实 token 用量。 */
export function summarize(records) {
  const toolCalls = [];
  const toolResults = [];
  const assistantTexts = [];
  const userMessages = [];
  const pwshCommands = [];
  const nameByCall = new Map();
  // 改进指南 §10.3 要求"实际计费 tokens，不能只按字符估算" —— 会话日志里每个 assistant/message
  // 都带 usage，直接累加就是这一轮的真实开销。
  const usage = { inputTokens: 0, outputTokens: 0, cacheReadTokens: 0, cacheWriteTokens: 0, totalTokens: 0, turns: 0 };
  const loadedChars = { skill: 0, read: 0 };
  for (const record of records) {
    const data = record.data || {};
    if (record.type === 'user/message') {
      userMessages.push(flatten(data.content));
    } else if (record.type === 'assistant/message') {
      for (const part of (data.message && data.message.content) || []) {
        if (part.type === 'text' && part.text) assistantTexts.push(part.text);
      }
      if (data.usage && typeof data.usage === 'object') {
        usage.turns++;
        for (const key of ['inputTokens', 'outputTokens', 'cacheReadTokens', 'cacheWriteTokens', 'totalTokens']) {
          const value = Number(data.usage[key]);
          if (Number.isFinite(value)) usage[key] += value;
        }
      }
    } else if (record.type === 'tool/call') {
      const name = String(data.name || '');
      const args = String(data.arguments || '');
      nameByCall.set(data.callId, name);
      toolCalls.push({ name, args, callId: data.callId });
      if (name === 'pwsh') {
        try {
          pwshCommands.push(String(JSON.parse(args).command || ''));
        } catch {
          // 参数不是 JSON 就只保留原文
        }
      }
    } else if (record.type === 'tool/result') {
      const name = nameByCall.get(data.message && data.message.source && data.message.source.callId) || '?';
      const text = flatten(data.message && data.message.content);
      toolResults.push({ name, text });
      // 模型实际"读进来"的技能/参考文件体量（字符，不是 token —— token 看 usage）。
      if (name === 'skill' || name === 'read') loadedChars[name] += text.length;
    }
  }
  const byTool = {};
  for (const call of toolCalls) byTool[call.name] = (byTool[call.name] || 0) + 1;
  return {
    toolCalls,
    toolResults,
    assistantTexts,
    userMessages,
    pwshCommands,
    usage,
    loadedChars,
    skillLoads: byTool.skill || 0,
    helpQueries: byTool.wps_help || 0,
    toolCounts: byTool,
    finalAnswer: assistantTexts.length ? assistantTexts[assistantTexts.length - 1] : '',
  };
}
