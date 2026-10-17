// 受控故障注入（WPS_OFFICE_FAULT_INJECT）的单元测试：不需要 WPS，也不启动 MCP server。
//
// 只测三件事：
//   ① 规格解析：合法片段被接受、写坏的片段进 invalid（不静默吞掉）；
//   ② 命中与消耗：只按 tool 名 / wps_call 的内层 tool 名匹配，count 用满即失效；
//   ③ **惰性**：不设环境变量时 active=false、take() 恒为 null —— 正常路径行为不受影响。
// Run: node test/fault-injection.test.mjs
import { parseFaultSpec, FaultInjector, injectedTimeoutText, injectedShortfallText } from '../mcp/dist/utils/fault-inject.js';

let pass = 0;
let fail = 0;
function check(name, ok, detail) {
  if (ok) { pass += 1; console.log('PASS ' + name); }
  else { fail += 1; console.log('FAIL ' + name + (detail ? '  ' + detail : '')); }
}

// ① 解析
{
  const parsed = parseFaultSpec('wps_word_insert_text:timeout:2;wps_word_set_font:shortfall');
  check('两条合法片段都被解析', parsed.rules.length === 2 && parsed.invalid.length === 0, JSON.stringify(parsed.invalid));
  check('默认 count=1', parsed.rules[1].remaining === 1, String(parsed.rules[1].remaining));
  check('显式 count 生效', parsed.rules[0].remaining === 2, String(parsed.rules[0].remaining));
}
{
  const parsed = parseFaultSpec('good_tool:timeout;bad-mode:explode;bad-count:timeout:0;bad-regex:":timeout;empty::');
  check('写坏的片段进 invalid，不抛异常', parsed.invalid.length === 4, JSON.stringify(parsed.invalid));
  check('非法片段不产生规则', parsed.rules.every((rule) => rule.raw === 'good_tool:timeout'), parsed.rules.map((r) => r.raw).join(','));
}
check('空规格 = 没有规则', parseFaultSpec(undefined).rules.length === 0 && parseFaultSpec('').rules.length === 0);

// ② 命中与消耗
{
  const injector = new FaultInjector('wps_word_insert_text:timeout:2');
  check('按 tool 名命中', injector.take('wps_word_insert_text') === 'timeout');
  check('count=2 时第二次仍命中', injector.take('wps_word_insert_text') === 'timeout');
  check('用满后不再命中', injector.take('wps_word_insert_text') === null);
}
{
  const injector = new FaultInjector('wps_word_set_font:shortfall');
  check('wps_call 用内层 tool 名命中', injector.take('wps_call', { tool: 'wps_word_set_font' }) === 'shortfall');
  check('无关 tool 不命中', injector.take('wps_word_get_document_text') === null);
  check('wps_call 内层是别的 tool 也不命中', injector.take('wps_call', { tool: 'wps_word_get_paragraphs' }) === null);
}
{
  const injector = new FaultInjector('^wps_word_(insert_text|set_paragraph)$:timeout');
  check('正则片段按正则匹配', injector.take('wps_word_set_paragraph') === 'timeout');
  check('正则片段不匹配前缀相同的别的 tool', injector.take('wps_word_set_paragraph_format') === null);
}
{
  // 模型可以绕过直达工具，用 wps_call → wps_execute_method 走同一个桥动作；
  // 只看 tool 名会让注入"看起来生效、实际没挡住"（实测踩过）。
  const injector = new FaultInjector('^(setFont|applyStyle)$:shortfall');
  check(
    'wps_call 深处的桥动作名也参与匹配',
    injector.take('wps_call', { tool: 'wps_execute_method', args: { method: 'setFont' } }) === 'shortfall',
  );
  check(
    'wps_call 深处的无关动作不匹配',
    injector.take('wps_call', { tool: 'wps_execute_method', args: { method: 'getDocumentText' } }) === null,
  );
}

// ③ 惰性（最重要：默认关闭时零行为）
{
  const injector = new FaultInjector(undefined);
  check('不设变量时 active=false', injector.active === false);
  check('不设变量时 take() 恒为 null', injector.take('wps_word_insert_text', { tool: 'wps_call' }) === null);
}
{
  const injector = new FaultInjector('   ');
  check('空白变量也不启用', injector.active === false && injector.invalid.length === 0);
}

// 注入文案：必须带「状态未知」，否则模型识别不出超时的语义
check('注入超时文案含「状态未知」', /状态未知/.test(injectedTimeoutText('wps_word_insert_text')));
check('注入短欠文案说明"这一步没有生效"', /没有生效/.test(injectedShortfallText('wps_word_set_font')));

console.log(fail === 0 ? 'FAULT INJECTION TESTS OK (' + pass + ')' : 'FAULT INJECTION TESTS FAILED (' + fail + ' failed)');
process.exit(fail === 0 ? 0 : 1);
