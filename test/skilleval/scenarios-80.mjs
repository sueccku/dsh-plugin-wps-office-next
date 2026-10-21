// 80 场景基准（P2-2）。从 80 场景评估的两份报告抽取；不预跑模型。
// 每个案例给 id/标题/应用/提示/确定性轨迹断言/coreCheck；ready:false 表示夹具或轨迹尚未落盘。
export const BENCHMARK80 = [
  {
    "id": "E1",
    "key": "p2-e1",
    "title": "6 个部门月报合成一张总表，以后加数据能自动更新",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：6 个部门月报合成一张总表，以后加数据能自动更新。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E1 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E2",
    "key": "p2-e2",
    "title": "签到表按 3–4 条公式规则给整行上色，空行不变",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：签到表按 3–4 条公式规则给整行上色，空行不变。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E2 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E3",
    "key": "p2-e3",
    "title": "两表按工号匹配，未匹配显示\"未匹配\"",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：两表按工号匹配，未匹配显示\"未匹配\"。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E3 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E4",
    "key": "p2-e4",
    "title": "一格多个姓名拆成一列一行 + 去重 + 删空行",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：一格多个姓名拆成一列一行 + 去重 + 删空行。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E4 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E5",
    "key": "p2-e5",
    "title": "按品类 × 月份做透视表并刷新",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：按品类 × 月份做透视表并刷新。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E5 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E6",
    "key": "p2-e6",
    "title": "每页重复表头、缩到一页宽、Excel 从第 3 页接页码",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：每页重复表头、缩到一页宽、Excel 从第 3 页接页码。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E6 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E7",
    "key": "p2-e7",
    "title": "折线图改日期轴、空值断开、别改公式",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：折线图改日期轴、空值断开、别改公式。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E7 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "E8",
    "key": "p2-e8",
    "title": "多人同表只改自己的行、保护密码、共享冲突",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：多人同表只改自己的行、保护密码、共享冲突。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 E8 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "W1",
    "key": "p2-w1",
    "title": "论文/公文按模板排样式、页边距、参考文献、引号",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：论文/公文按模板排样式、页边距、参考文献、引号。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 W1 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "W2",
    "key": "p2-w2",
    "title": "目录页码全 1 / 个别不对 / 正文误入目录",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：目录页码全 1 / 个别不对 / 正文误入目录。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 W2 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "W3",
    "key": "p2-w3",
    "title": "模板 + 200 行名单批量出合同，一人一份",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：模板 + 200 行名单批量出合同，一人一份。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 W3 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "W4",
    "key": "p2-w4",
    "title": "接受修订、清批注、关修订、交干净版",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：接受修订、清批注、关修订、交干净版。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 W4 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "W5",
    "key": "p2-w5",
    "title": "表格跨页每页有表头、简历不被挤到下一页",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：表格跨页每页有表头、简历不被挤到下一页。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 W5 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "W6",
    "key": "p2-w6",
    "title": "同类文字格式批量统一/格式刷",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：同类文字格式批量统一/格式刷。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 W6 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "P1",
    "key": "p2-p1",
    "title": "有内容的 PPT 套公司模板，字体/LOGO/页脚不跑偏",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：有内容的 PPT 套公司模板，字体/LOGO/页脚不跑偏。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 P1 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "P2",
    "key": "p2-p2",
    "title": "6 张图排一行、等距、整排相对页面居中",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：6 张图排一行、等距、整排相对页面居中。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 P2 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "P3",
    "key": "p2-p3",
    "title": "全篇统一动画与\"平滑\"切换，播放不闪",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：全篇统一动画与\"平滑\"切换，播放不闪。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 P3 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "P4",
    "key": "p2-p4",
    "title": "只改文本框字色不动字号/占位符/表格；全文替换",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：只改文本框字色不动字号/占位符/表格；全文替换。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 P4 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "X1",
    "key": "p2-x1",
    "title": "Excel 数据/图表进 PPT，保留可刷新链接",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：Excel 数据/图表进 PPT，保留可刷新链接。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 X1 原判定 A；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL1",
    "key": "p2-xl1",
    "title": "多条件 SUMIFS 全返 0",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：多条件 SUMIFS 全返 0。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL1 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL2",
    "key": "p2-xl2",
    "title": "产品×月份二维取值",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：产品×月份二维取值。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL2 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL3",
    "key": "p2-xl3",
    "title": "全表错误值混排",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：全表错误值混排。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL3 原判定 C+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL4",
    "key": "p2-xl4",
    "title": "地址拆省市",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：地址拆省市。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL4 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL5",
    "key": "p2-xl5",
    "title": "图标集对公式格不生效",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：图标集对公式格不生效。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL5 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL6",
    "key": "p2-xl6",
    "title": "固定宽度分列+Ctrl+E",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：固定宽度分列+Ctrl+E。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL6 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL7",
    "key": "p2-xl7",
    "title": "CSV 乱码/001/日期",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：CSV 乱码/001/日期。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL7 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL8",
    "key": "p2-xl8",
    "title": "去重保留最近修改",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：去重保留最近修改。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL8 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL9",
    "key": "p2-xl9",
    "title": "透视计算字段/季度/切片器",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：透视计算字段/季度/切片器。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL9 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL10",
    "key": "p2-xl10",
    "title": "柱+折线次轴+趋势线",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：柱+折线次轴+趋势线。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL10 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL11",
    "key": "p2-xl11",
    "title": "冻结/分级/隐藏列",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：冻结/分级/隐藏列。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL11 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL12",
    "key": "p2-xl12",
    "title": "允许编辑区域/隐藏公式",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：允许编辑区域/隐藏公式。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL12 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL13",
    "key": "p2-xl13",
    "title": "保护密码忘了",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：保护密码忘了。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL13 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL14",
    "key": "p2-xl14",
    "title": "12 万行大表卡顿",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：12 万行大表卡顿。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL14 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL15",
    "key": "p2-xl15",
    "title": "外链更新/源丢失/断开",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：外链更新/源丢失/断开。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL15 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL16",
    "key": "p2-xl16",
    "title": "12 个月度文件静态合并",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：12 个月度文件静态合并。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL16 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL17",
    "key": "p2-xl17",
    "title": "自定义顺序+多键排序+高级筛选",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：自定义顺序+多键排序+高级筛选。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL17 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL18",
    "key": "p2-xl18",
    "title": "批注/超链接/名称管理器清理",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：批注/超链接/名称管理器清理。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL18 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL19",
    "key": "p2-xl19",
    "title": "目标倒推/敏感性分析",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：目标倒推/敏感性分析。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL19 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XL20",
    "key": "p2-xl20",
    "title": "打印区域/分页符/页脚",
    "app": "excel",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：打印区域/分页符/页脚。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_excel_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XL20 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD1",
    "key": "p2-wd1",
    "title": "水印（三连）",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：水印（三连）。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD1 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD2",
    "key": "p2-wd2",
    "title": "分节/横纵混排",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：分节/横纵混排。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD2 原判定 C+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD3",
    "key": "p2-wd3",
    "title": "封面无页眉/奇偶页不同",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：封面无页眉/奇偶页不同。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD3 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD4",
    "key": "p2-wd4",
    "title": "页眉横线删不掉",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：页眉横线删不掉。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD4 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD5",
    "key": "p2-wd5",
    "title": "题注/图表编号",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：题注/图表编号。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD5 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD6",
    "key": "p2-wd6",
    "title": "交叉引用/PDF 跳转",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：交叉引用/PDF 跳转。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD6 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD7",
    "key": "p2-wd7",
    "title": "图目录/表目录",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：图目录/表目录。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD7 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD8",
    "key": "p2-wd8",
    "title": "两版文档比较",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：两版文档比较。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD8 原判定 D+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD9",
    "key": "p2-wd9",
    "title": "邮件合并标签/信封/席卡",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：邮件合并标签/信封/席卡。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD9 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD10",
    "key": "p2-wd10",
    "title": "带照片证书合并",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：带照片证书合并。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD10 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD11",
    "key": "p2-wd11",
    "title": "内容控件/可填写表单",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：内容控件/可填写表单。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD11 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD12",
    "key": "p2-wd12",
    "title": "按作者筛选修订",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：按作者筛选修订。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD12 原判定 C+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD13",
    "key": "p2-wd13",
    "title": "Word 表格公式/排序/斜线",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：Word 表格公式/排序/斜线。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD13 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD14",
    "key": "p2-wd14",
    "title": "文档属性/PDF 元数据",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：文档属性/PDF 元数据。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD14 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "WD15",
    "key": "p2-wd15",
    "title": "限制编辑/只读",
    "app": "word",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：限制编辑/只读。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_word_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 WD15 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT1",
    "key": "p2-pt1",
    "title": "母版新版式/LOGO/占位符",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：母版新版式/LOGO/占位符。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT1 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT2",
    "key": "p2-pt2",
    "title": "主题色/主题字体",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：主题色/主题字体。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT2 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT3",
    "key": "p2-pt3",
    "title": "4:3→16:9",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：4:3→16:9。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT3 原判定 C+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT4",
    "key": "p2-pt4",
    "title": "只换字体不动字号",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：只换字体不动字号。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT4 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT5",
    "key": "p2-pt5",
    "title": "字体嵌入/缺失",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：字体嵌入/缺失。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT5 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT6",
    "key": "p2-pt6",
    "title": "图片裁剪/压缩/替换",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：图片裁剪/压缩/替换。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT6 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT7",
    "key": "p2-pt7",
    "title": "SmartArt",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：SmartArt。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT7 原判定 D（D3 边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT8",
    "key": "p2-pt8",
    "title": "音视频",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：音视频。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT8 原判定 D（D3 边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT9",
    "key": "p2-pt9",
    "title": "节/隐藏页/导出",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：节/隐藏页/导出。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT9 原判定 D；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "PT10",
    "key": "p2-pt10",
    "title": "备注/提词/导出带备注 PDF",
    "app": "ppt",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：备注/提词/导出带备注 PDF。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_ppt_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 PT10 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA1",
    "key": "p2-xa1",
    "title": "表格→Word 报告",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：表格→Word 报告。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA1 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA2",
    "key": "p2-xa2",
    "title": "表格→PPT 静态页",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：表格→PPT 静态页。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA2 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA3",
    "key": "p2-xa3",
    "title": "Word 大纲→PPT",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：Word 大纲→PPT。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA3 原判定 C；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA4",
    "key": "p2-xa4",
    "title": "PPT 备注→Word 讲义",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：PPT 备注→Word 讲义。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA4 原判定 B；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA5",
    "key": "p2-xa5",
    "title": "PDF 合并",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：PDF 合并。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA5 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA6",
    "key": "p2-xa6",
    "title": "PDF 拆分",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：PDF 拆分。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA6 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA7",
    "key": "p2-xa7",
    "title": "PDF→Word/Excel",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：PDF→Word/Excel。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA7 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA8",
    "key": "p2-xa8",
    "title": "PDF 水印/页码",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：PDF 水印/页码。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA8 原判定 D（边界/会员）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA9",
    "key": "p2-xa9",
    "title": "批量转换",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：批量转换。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA9 原判定 D+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA10",
    "key": "p2-xa10",
    "title": "批量改名/整理",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：批量改名/整理。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA10 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA11",
    "key": "p2-xa11",
    "title": "批量套模板/统一格式",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：批量套模板/统一格式。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA11 原判定 C+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA12",
    "key": "p2-xa12",
    "title": "OLE/链接对象",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：OLE/链接对象。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA12 原判定 D+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA13",
    "key": "p2-xa13",
    "title": "手机端批量操作",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：手机端批量操作。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA13 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA14",
    "key": "p2-xa14",
    "title": "WPS AI 生成/改写",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：WPS AI 生成/改写。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA14 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "XA15",
    "key": "p2-xa15",
    "title": "十万行/长文档卡顿",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：十万行/长文档卡顿。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 XA15 原判定 C+E；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  },
  {
    "id": "F1",
    "key": "p2-f1",
    "title": "批量打印/导出 PDF",
    "app": "cross",
    "ready": false,
    "fixture": null,
    "prompt": "请完成这个真实 WPS 场景：批量打印/导出 PDF。先确认目标与输入；不确定工具时用 wps_help 查询；完成后报告已执行/已核对/未验证。",
    "checks": [
      "transcript-matches:\"wps_(common|word|excel|ppt|convert)_\"",
      "transcript-not-matches:\"unknown tool\""
    ],
    "coreCheck": "场景 F1 原判定 D（边界）；不得无依据成功，源文件改动要如实报告",
    "note": "P2 benchmark 案例：提示/确定性断言已定义；夹具与真机轨迹按需补（不预跑模型）。"
  }
];
