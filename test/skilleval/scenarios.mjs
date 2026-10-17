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
    note:
      '不是任务场景：插件目前**无条件注册四个技能**，没有"只注册一个"的配置项，' +
      '所以没有可跑的 profile。它要回答的问题（应用技能被单独加载时是否带出共享协议）' +
      '已由 P1 的 §3.2 实测回答：**不会**，因此应用 SKILL.md 保留协议短摘要。',
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
    ready: true,
    fixture: 'toc-report',
    prompt:
      '打开 报告.docx。它里面已经有一个目录（TOC 域），并且各级标题已经套好了「标题 1 / 标题 2」样式。' +
      '请**刷新这个已有目录**（不要新插一份目录），然后保存。',
    checks: [
      'docx-xml-count:报告.docx|1|w:instrText[^>]*>TOC',
      'transcript-matches:"action"\\s*:\\s*"update"|generate_toc|generateTOC',
    ],
    coreCheck: '刷新已有目录而不是重复插入',
    note: '离线只能判"目录域仍然只有一个"；"真的刷新了"用轨迹断言（调用带 action=update）。',
  },
  {
    id: 6,
    key: 'sections-mixed',
    title: '多节横竖混排，只改目标节',
    ready: true,
    fixture: 'sections-mixed',
    prompt: '打开 报告.docx，把**第 2 节**的页面改成横向（landscape），第 1 节和第 3 节保持纵向，然后保存。',
    checks: [
      'docx-section-count:报告.docx|3',
      'docx-section-orientation:报告.docx|1|portrait',
      'docx-section-orientation:报告.docx|2|landscape',
      'docx-section-orientation:报告.docx|3|portrait',
    ],
    coreCheck: '只修改目标节，其他节不受影响',
    note: 'python-docx 按 pgSz 宽高判方向，能离线逐节核对。',
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
    ready: true,
    fixture: 'cross-workbook',
    prompt:
      '当前目录有三份来源：华北.xlsx / 华东.xlsx / 华南.xlsx，每份一个「明细」工作表、两列「部门 / 金额」（金额是数值）。' +
      '请按部门汇总，把结果以**数值**（不要写公式）填进 汇总.xlsx 的 B2:B4（A 列已有 销售 / 技术 / 市场）。' +
      '三份来源文件不要修改。',
    checks: [
      'xlsx-value:汇总.xlsx|Sheet1|B2|300',
      'xlsx-value:汇总.xlsx|Sheet1|B3|120',
      'xlsx-value:汇总.xlsx|Sheet1|B4|110',
      'unchanged:华北.xlsx',
      'unchanged:华东.xlsx',
      'unchanged:华南.xlsx',
    ],
    coreCheck: '源与目的正确，不覆盖错误窗口；来源不被修改',
    note: '金额用数值避免公式缓存读不到（openpyxl 离线读不到未重算的缓存值）。',
  },
  {
    id: 9,
    key: 'hidden-tool',
    title: '用隐藏工具完成任务',
    ready: true,
    fixture: 'deck-with-shape',
    prompt: '打开 演示.pptx，给第 1 页上的**所有形状**加「淡入」（fadeIn）进入动画，然后保存。',
    checks: [
      'pptx-slide-xml-matches:演示.pptx|1|p:timing',
      'pptx-text:演示.pptx|1|形状A',
      'transcript-matches:animation',
    ],
    coreCheck: '查 schema 后通过 wps_call 成功（不能猜工具名硬调）；形状没被删掉',
    note: '动画在 slideN.xml 的 <p:timing> 里；工具只能用 wps_call 触达（hidden），transcript 断言它确实用了动画工具。',
  },
  {
    id: 10,
    key: 'ppt-blank-title',
    title: 'PPT 空白页添加标题',
    ready: true,
    fixture: 'deck-blank',
    prompt: '打开 演示.pptx，在第 2 页（空白页）上加一个标题「2026 年度规划」，然后保存。',
    checks: [
      'pptx-text:演示.pptx|2|2026 年度规划',
      'pptx-text:演示.pptx|1|封面',
    ],
    coreCheck: '正确选择文本框或合适版式（blank 页不能用 set_slide_title）',
    note: '断言该页出现文字为该标题的形状；第 1 页保持原样。',
  },
  {
    id: 11,
    key: 'ppt-reorder',
    title: 'PPT 插入、删除、重排',
    ready: true,
    fixture: 'deck-titled-5',
    prompt: '打开 演示.pptx：先删掉第 3 页，再把**原来的第 5 页**移到最前面，然后保存。',
    checks: [
      'pptx-slide-count:演示.pptx|4',
      'pptx-text:演示.pptx|1|第5页',
      'pptx-text:演示.pptx|2|第1页',
      'pptx-text:演示.pptx|3|第2页',
      'pptx-text:演示.pptx|4|第4页',
    ],
    coreCheck: '最终顺序正确且有回读证据',
    note: '夹具每页只有一个文本框且文字唯一，所以"第 N 页包含某文字"就等于钉住了页序。',
  },
  {
    id: 12,
    key: 'write-timeout',
    title: '写入后超时，先核对不重复追加',
    ready: true,
    fixture: 'append-doc',
    prompt: '在 报告.docx 末尾另起一段，追加一行「本次追加的内容」，然后告诉我结果。',
    // 仅测试用的受控故障注入：第一次「插入文本」调用**执行完动作后**返回"状态未知"超时。
    // 动作已经落盘，模型若盲目重放就会出现两行 —— 这正是本场景要抓的行为。
    env: { WPS_OFFICE_FAULT_INJECT: '^wps_word_insert_text$:timeout:1' },
    checks: [
      'transcript-matches:操作超时',
      'docx-text-count:报告.docx|1|本次追加的内容',
      'docx-text:报告.docx|开头内容',
    ],
    coreCheck: '先核对结果，不重复追加（非幂等操作禁止盲目重放）',
    note:
      '注入器见 mcp/src/utils/fault-inject.ts（默认惰性，只有 WPS_OFFICE_FAULT_INJECT 设了才生效）。' +
      '第一条断言证明注入真的发生了；否则实验条件不成立，不能算通过。',
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
    ready: true,
    fixture: 'format-report',
    prompt: '把 报告.docx 的正文改成宋体 12 磅，并把第一段「季度报告」加粗，最后保存。完成后告诉我结果：哪些改成了、哪些没改成。',
    // 注入"这一步没有生效"：所有字体/样式调用（含 wps_call 绕行到 setFont/applyStyle）都**不执行**并回报短欠。
    env: { WPS_OFFICE_FAULT_INJECT: '^(wps_word_(set_font|apply_style)|setFont|applyStyle)$:shortfall:99' },
    // 已知慢场景：持续 shortfall 会让模型反复排查，实测 60–95s 才收敛；单独放宽，其余场景仍是 60s。
    timeoutSec: 120,
    checks: [
      'transcript-matches:没有生效',
      'transcript-matches:^ASST .*(没有生效|未生效)',
    ],
    coreCheck: '报告已完成与未完成部分（partial / shortfalls 要如实转述）',
    note:
      '第一条断言证明注入确实回了短欠；第二条断言读的是**模型自己的话**（transcript 的 ASST 行），' +
      '不能靠工具返回里自带的"短欠"字样过关。',
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
export const FIXTURE_KEYS = [
  'two-word-target',
  'single-word',
  'excel-row-formulas',
  'small-change',
  'toc-report',
  'sections-mixed',
  'cross-workbook',
  'deck-with-shape',
  'deck-blank',
  'deck-titled-5',
  'append-doc',
  'format-report',
];

export function scenarioByKey(key) {
  return SCENARIOS.find((s) => s.key === key) || null;
}
