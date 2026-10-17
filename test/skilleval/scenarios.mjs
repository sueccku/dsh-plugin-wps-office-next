// 固定评估场景（对应 docs/skill-evaluation.md §1 的 15 个场景）。
//
// 为什么要有这张表：改进指南 §10.2 要求"固定输入、初始文件、用户提示和成功标准"，
// 否则新旧 Skill 的对比无法重复。这里把 15 条都声明出来，但只对**已经有夹具与确定性判定**的
// 场景标 ready: true —— 其余标记为待补，不做"看着像能跑"的假执行。
//
// 约定（verify.mjs / run.mjs 依赖）：
//   * 每个场景的工作目录是 test/.artifacts/skill-eval/<NN>-<key>/input/，提示里的相对路径都相对它；
//   * 产物默认写到 output/ 下（提示里写明）；
//   * checks 是一组**可机器判定**的断言，语法见 verify.mjs 顶部；
//   * coreCheck 抄自 docs/skill-evaluation.md，给还没写判定的场景留人读的成功标准。
export const SCENARIOS = [
  {
    id: 1,
    key: 'word-new-save',
    title: 'WPS 未启动，新建 Word 并保存',
    ready: true,
    fixture: null,
    prompt: '新建一份 Word 文档：标题写「2026 年第三季度经营分析报告」，下面写一行正文「本季度营收同比增长 12%」。保存为 output/报告.docx。',
    checks: [
      'file-exists:output/报告.docx',
      'docx-text:output/报告.docx|2026 年第三季度经营分析报告',
      'docx-text:output/报告.docx|同比增长 12%',
    ],
    coreCheck: '不因 wps_status 的 connected:false 放弃；文档真的落盘',
  },
  {
    id: 2,
    key: 'only-word-skill',
    title: '只加载 Word Skill 修改文档',
    ready: false,
    fixture: null,
    prompt: '（本场景不是任务，而是加载配置验证：只挂载 wps-word 时能否取得必要的共享协议。做法见 docs/skill-evaluation.md §1 #2。）',
    checks: [],
    coreCheck: '能取得必要共享协议（目标确认 / 保存策略 / 验证 / 失败分流）',
    note: '需要按 profile 限制只注册 wps-word；属于 P1 的 §3.2 资源读取验证，不在任务场景里跑。',
  },
  {
    id: 3,
    key: 'two-word-target',
    title: '两份 Word 同时打开，只改指定的一份',
    ready: true,
    fixture: 'two-word-target',
    prompt: '当前目录有两份文档：报告A.docx 和 报告B.docx。把 报告A.docx 的正文改成「A-已修订」，报告B.docx 不要动。',
    checks: [
      'docx-text:报告A.docx|A-已修订',
      'docx-text:报告B.docx|B-原始内容',
      'unchanged:报告B.docx',
    ],
    coreCheck: '明确目标；不可误改另一份',
  },
  {
    id: 4,
    key: 'keep-original',
    title: '修改报告但保留原件',
    ready: true,
    fixture: 'single-word',
    prompt: '把 报告A.docx 改一份副本：正文改成「A-副本」，原文件保持不动，副本保存为 output/报告A-副本.docx。',
    checks: [
      'file-exists:output/报告A-副本.docx',
      'docx-text:output/报告A-副本.docx|A-副本',
      'unchanged:报告A.docx',
    ],
    coreCheck: '副本身份正确、原件内容不变',
  },
  {
    id: 5,
    key: 'toc-refresh',
    title: '已有目录的报告排版',
    ready: false,
    fixture: null,
    prompt: '把 报告.docx 的标题套上「标题 1 / 标题 2」样式，并刷新文档里已有的目录（不要重复插入目录）。',
    checks: [],
    coreCheck: '刷新目录而不是重复插入',
    note: '夹具需要"带标题样式与真实 TOC 域"的 docx，D1 用 OOXML 构造。',
  },
  {
    id: 6,
    key: 'sections-mixed',
    title: '多节横竖混排，只改目标节',
    ready: false,
    fixture: null,
    prompt: '把 报告.docx 的第 2 节页面改成横向，其他节的纸张方向不要变。',
    checks: [],
    coreCheck: '只修改目标节，其他节不受影响',
    note: '夹具需要三节文档（纵向 / 纵向 / 横向候选），D1 用 python-docx 构造并写 section 基线。',
  },
  {
    id: 7,
    key: 'excel-row-formulas',
    title: 'Excel 按行生成公式',
    ready: true,
    fixture: 'excel-row-formulas',
    prompt: '把 输入.xlsx 里「金额」列（D 列）逐行填成「单价 × 数量」的公式，要求能随单价和数量自动更新，保存回原文件。',
    checks: [
      'xlsx-formula:输入.xlsx|Sheet1|D2|=B2*C2',
      'xlsx-formula:输入.xlsx|Sheet1|D3|=B3*C3',
      'xlsx-formula:输入.xlsx|Sheet1|D4|=B4*C4',
      'xlsx-formula:输入.xlsx|Sheet1|D5|=B5*C5',
      'xlsx-formula:输入.xlsx|Sheet1|D6|=B6*C6',
    ],
    coreCheck: '保留公式、引用正确、结果正确（结果值用 WPS/重算单独核对）',
    note: '公式文本可离线判定；计算结果值必须在 WPS 里重算后读，属于 D1 的补充断言。',
  },
  {
    id: 8,
    key: 'cross-workbook',
    title: '跨工作簿汇总',
    ready: false,
    fixture: null,
    prompt: '把 华北.xlsx / 华东.xlsx / 华南.xlsx 三份明细按部门汇总到 汇总.xlsx，来源文件不要改。',
    checks: [],
    coreCheck: '源与目的正确，不覆盖错误窗口',
    note: '夹具是三份来源 + 一份模板；判定要读三个文件并与来源对账。',
  },
  {
    id: 9,
    key: 'hidden-tool',
    title: '用隐藏工具完成任务',
    ready: false,
    fixture: null,
    prompt: '给 演示.pptx 第 1 页上的形状加淡入动画。',
    checks: [],
    coreCheck: '查 schema 后通过 wps_call 成功（不能猜工具名硬调）',
    note: '需要 WPS 里真实打开该 pptx；断言读动画数量与工具调用轨迹。',
  },
  {
    id: 10,
    key: 'ppt-blank-title',
    title: 'PPT 空白页添加标题',
    ready: false,
    fixture: null,
    prompt: '在 演示.pptx 的空白页上加一个标题「2026 年度规划」。',
    checks: [],
    coreCheck: '正确选择文本框或合适版式（blank 页不能用 set_slide_title）',
    note: '断言：该页出现文字为「2026 年度规划」的形状；且工具轨迹里没有"失败后原样重试"。',
  },
  {
    id: 11,
    key: 'ppt-reorder',
    title: 'PPT 插入、删除、重排',
    ready: false,
    fixture: null,
    prompt: '把 演示.pptx 的第 3 页删掉，再把原来的第 5 页移到最前面。',
    checks: [],
    coreCheck: '最终顺序正确且有回读证据',
    note: '断言：逐页标题序列等于期望序列（夹具给每页唯一的标题）。',
  },
  {
    id: 12,
    key: 'write-timeout',
    title: '写入后超时，先核对不重复追加',
    ready: false,
    fixture: null,
    prompt: '在这份文档末尾追加一行「本次追加的内容」，然后告诉我结果。',
    checks: [],
    coreCheck: '先核对结果，不重复追加（非幂等操作禁止盲目重放）',
    note: '需要受控故障注入（让第一次写入超时），属于 D1 的异常场景。',
  },
  {
    id: 13,
    key: 'dirty-save-close',
    title: '新建未保存文件要求保存关闭',
    ready: true,
    fixture: null,
    prompt: '新建一份 Word 文档，写入「临时探针内容」，然后保存并关闭它。',
    checks: [
      'transcript-matches:磁盘路径|另存为|save_as|outputPath|输出路径',
    ],
    coreCheck: '不丢失内容；需要路径时明确指出（不能静默丢弃后回成功）',
    note: '内容是否真丢由 test/close-safety.test.mjs 在桥层钉死；这里只判"轨迹里有没有为路径做过处理"。',
  },
  {
    id: 14,
    key: 'partial-failure',
    title: '修改成功但部分格式失败',
    ready: false,
    fixture: null,
    prompt: '把 报告.docx 全文改成「宋体 12 磅」，同时给标题加粗。',
    checks: [],
    coreCheck: '报告已完成与未完成部分（partial / shortfalls 要如实转述）',
    note: '需要受控让一步短欠（例如对不存在的样式名操作），D1 故障注入。',
  },
  {
    id: 15,
    key: 'small-change',
    title: '明确目标的小改动，不啰嗦',
    ready: true,
    fixture: 'small-change',
    prompt: '把 输入.xlsx 的 A1 单元格改成 5，其它内容不要动。',
    checks: [
      'xlsx-value:输入.xlsx|Sheet1|A1|5',
    ],
    coreCheck: '不提出无必要的重复确认（重复询问次数在 D1 从轨迹里统计）',
    note: '重复询问次数不是文件断言，D1 从会话轨迹统计。',
  },
];

/** 需要夹具的场景用到的 key（供 fixtures.py 对账，防止表与脚本漂移）。 */
export const FIXTURE_KEYS = ['two-word-target', 'single-word', 'excel-row-formulas', 'small-change'];

export function scenarioByKey(key) {
  return SCENARIOS.find((s) => s.key === key) || null;
}
