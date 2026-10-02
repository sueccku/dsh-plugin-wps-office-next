/**
 * Input: 桥侧结果里的 warnings（宿主把它们同时挂在结果根与 data.warnings 上）
 * Output: 一次工具调用期间收集到的注意事项，交给注册表附回结果
 * Pos: 「尽力而为的失败」的唯一出口。第一方 handler 各自取字段、自建 data，会把 warnings 整个丢掉；
 *      所以收集放在客户端（所有调用都经过它），回传放在注册表（所有结果都从它出去）。作用域用
 *      AsyncLocalStorage 绑定到单次 tools/call，避免并发调用互相串味（FIXES 73）。
 */
import { AsyncLocalStorage } from 'node:async_hooks';

const warningStore = new AsyncLocalStorage<string[]>();

/** 在本次工具调用的作用域里执行 handler，并把收集到的注意事项一并返回。 */
export async function runWithWarningCollector<T>(
  fn: () => Promise<T>
): Promise<{ value: T; warnings: string[] }> {
  return warningStore.run([], async () => {
    const value = await fn();
    return { value, warnings: [...(warningStore.getStore() ?? [])] };
  });
}

/** 把桥侧结果里的 warnings 收进当前调用；不在工具调用作用域里（例如启动自检）就静默忽略。 */
export function collectToolWarnings(value: unknown): void {
  const sink = warningStore.getStore();
  if (!sink) return;
  for (const message of extractMessages(value)) {
    if (!sink.includes(message)) sink.push(message);
  }
}

function extractMessages(value: unknown): string[] {
  if (value === null || value === undefined) return [];
  if (typeof value === 'string') return value.trim() ? [value] : [];
  if (Array.isArray(value)) return value.flatMap((item) => extractMessages(item));
  return [];
}
