"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.operations = void 0;
const types_1 = require("./types");
// GENERATED DRAFT: bootstrap of the operation spec from the pre-P1 surface (scripts/extract-spec.mjs).
// It is a starting point to curate, not the finished spec. Aliases and containers are carried over
// from the old translation tables so the generator can reproduce the pre-P1 artifacts byte for byte
// (that reproduction is the P1 acceptance test). Curation targets: drop aliases, add effect/enum/units,
// split the union-parameter interfaces, and delete the opaque entries once nothing is unreadable.
//
// 一旦我被修改，请更新 docs/tool-roadmap.md 的 P1 状态。
exports.operations = [
    (0, types_1.op)({
        "tool": "wps_batch",
        "action": null,
        "app": "common",
        "summary": "按顺序批量执行多个 WPS 工具调用，用于跨应用的连续操作；单次最多 50 项",
        "params": {
            "calls": {
                "type": "array",
                "description": "调用列表，每项为 {tool, args}",
                "items": {
                    "type": "object"
                },
                "schema": {
                    "type": "array",
                    "description": "调用列表，每项为 {tool, args}",
                    "items": {
                        "type": "object",
                        "properties": {
                            "tool": {
                                "type": "string"
                            },
                            "args": {
                                "type": "object"
                            }
                        },
                        "required": [
                            "tool"
                        ]
                    }
                },
                "required": true,
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "calls"
        ],
        "engine": "opaque"
    }),
    (0, types_1.op)({
        "tool": "wps_call",
        "action": null,
        "app": "common",
        "summary": "执行任意已注册但未直接广告的 WPS 工具；先用 wps_help 取到工具名与参数",
        "params": {
            "tool": {
                "type": "string",
                "description": "完整工具名，如 wps_ppt_set_animation",
                "required": true,
                "kind": "local"
            },
            "args": {
                "type": "object",
                "description": "该工具的参数对象",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "tool"
        ],
        "engine": "opaque"
    }),
    (0, types_1.op)({
        "tool": "wps_common_get_app_info",
        "action": "getAppInfo",
        "app": "common",
        "summary": "获取WPS应用的基本信息。\n\n使用场景：\n- \"WPS是什么版本\"\n- \"查看WPS信息\"\n- \"获取应用状态\"\n\n特点：\n- 返回WPS版本号、构建信息\n- 返回当前打开的文档信息\n- 返回运行平台信息",
        "params": {},
        "effect": "read",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_common_get_selected_text",
        "action": "getSelectedText",
        "app": "common",
        "summary": "获取当前文档中选中的文本内容。\n\n使用场景：\n- \"读取我选中的内容\"\n- \"获取当前选区的文字\"\n- \"看看我选了什么\"\n\n特点：\n- 返回当前选区的纯文本内容\n- 适用于Word、Excel、PPT",
        "params": {},
        "effect": "read",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_common_ping",
        "action": "ping",
        "app": "common",
        "summary": "检测WPS应用连接状态。\n\n使用场景：\n- \"WPS连上了吗\"\n- \"检查一下WPS是否在线\"\n- \"测试WPS连接\"\n\n特点：\n- 快速检测WPS加载项是否可达\n- 返回连接状态信息",
        "params": {},
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_common_save",
        "action": "save",
        "app": "common",
        "summary": "保存当前文档。\n\n使用场景：\n- \"保存一下\"\n- \"Ctrl+S\"\n- \"把修改存起来\"\n\n特点：\n- 自动保存当前活动文档\n- 如果文档从未保存过，可能会提示选择保存路径",
        "params": {},
        "effect": "lifecycle",
        "advertised": true,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_common_save_as",
        "action": "saveAs",
        "app": "common",
        "summary": "将当前文档另存为指定路径和格式。\n\n使用场景：\n- \"另存为到桌面\"\n- \"换个名字保存\"\n- \"保存一份副本到指定位置\"\n\n特点：\n- 支持指定完整文件路径\n- 可选指定保存格式",
        "params": {
            "filePath": {
                "type": "string",
                "description": "目标文件完整路径，包含文件名和扩展名",
                "required": true
            },
            "format": {
                "type": "string",
                "description": "保存格式（可选），如 docx, xlsx, pptx 等"
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "filePath"
        ],
        "engine": "bridge",
        "aliases": {
            "filePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_common_set_selected_text",
        "action": "setSelectedText",
        "app": "common",
        "summary": "替换当前文档中选中的文本内容。\n\n使用场景：\n- \"把选中的内容替换成xxx\"\n- \"修改选区的文字\"\n- \"用新内容替换选中部分\"\n\n特点：\n- 将当前选区的文本替换为指定内容\n- 适用于Word、Excel、PPT",
        "params": {
            "text": {
                "type": "string",
                "description": "要替换为的新文本内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_common_wire_check",
        "action": "wireCheck",
        "app": "common",
        "summary": "检查与WPS加载项之间的通信线路状态。\n\n使用场景：\n- \"检查通信状态\"\n- \"诊断连接问题\"\n- \"通信线路是否正常\"\n\n特点：\n- 比ping更详细的通信诊断\n- 返回线路延迟、协议状态等信息",
        "params": {},
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_convert_format",
        "action": "convertFormat",
        "app": "common",
        "summary": "将当前文档转换为其他格式。\n\n使用场景：\n- \"把docx转成doc格式\"\n- \"保存为旧版Excel格式兼容老系统\"\n- \"把PPT转成pptx\"\n- \"导出为RTF格式\"\n- \"转换成HTML网页格式\"\n\n支持的格式：\n- Word: doc, docx, rtf, txt, html, xml\n- Excel: xls, xlsx, xlsm, xlsb, csv, html\n- PPT: ppt, pptx, pptm, html, png, jpg, gif, bmp\n\n注意：\n- 转换时会保留原文档，另存为新格式\n- 某些格式转换可能会丢失部分效果（如doc不支持的新特性）",
        "params": {
            "targetFormat": {
                "type": "string",
                "description": "目标格式扩展名，如 doc, xlsx, ppt, rtf, csv, html 等",
                "required": true
            },
            "outputPath": {
                "type": "string",
                "description": "输出路径（包含文件名），如不指定则使用原文件名改为新扩展名"
            },
            "app_type": {
                "type": "string",
                "description": "要转换的应用；不填则按 Excel→Word→PPT 选第一个正在运行的文档",
                "enum": [
                    "excel",
                    "word",
                    "ppt"
                ]
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "targetFormat"
        ],
        "engine": "bridge",
        "aliases": {
            "app_type": "appType"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_convert_to_pdf",
        "action": "convertToPDF",
        "app": "common",
        "summary": "将当前文档转换为PDF格式。支持Word、Excel、PPT文档。\n\n使用场景：\n- \"把这个Word转成PDF\"\n- \"导出PDF给客户看\"\n- \"把表格保存成PDF格式\"\n- \"PPT转PDF方便打印\"\n\n特点：\n- 自动检测当前打开的文档类型\n- 可以指定输出路径，不指定则使用原文件名.pdf\n- 保持原文档格式和排版",
        "params": {
            "outputPath": {
                "type": "string",
                "description": "PDF输出路径（包含文件名），如不指定则使用原文件路径，把扩展名改为.pdf"
            },
            "openAfterExport": {
                "type": "boolean",
                "description": "导出后是否自动打开PDF，默认false"
            },
            "app_type": {
                "type": "string",
                "description": "要导出的应用；不填则按 Excel→Word→PPT 选第一个正在运行的文档",
                "enum": [
                    "excel",
                    "word",
                    "ppt"
                ]
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "app_type": "appType"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_execute_method",
        "action": null,
        "app": "common",
        "summary": "执行自定义WPS API方法",
        "params": {
            "method": {
                "type": "string",
                "description": "API方法名",
                "required": true,
                "kind": "local"
            },
            "params": {
                "type": "object",
                "description": "方法参数",
                "kind": "local"
            },
            "appType": {
                "type": "string",
                "description": "应用类型：wps（文字）、et（表格）、wpp（演示）",
                "enum": [
                    "wps",
                    "et",
                    "wpp"
                ],
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "method"
        ],
        "engine": "opaque"
    }),
    (0, types_1.op)({
        "tool": "wps_help",
        "action": null,
        "app": "common",
        "summary": "查询未直接广告的工具：无参看分组概览，传 app 或 query 查目录，传 tool 取完整参数 schema",
        "params": {
            "app": {
                "type": "string",
                "description": "应用：excel / word / ppt / common",
                "kind": "local"
            },
            "query": {
                "type": "string",
                "description": "按名称或描述搜索关键字",
                "kind": "local"
            },
            "tool": {
                "type": "string",
                "description": "工具名，返回完整 inputSchema",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": true,
        "engine": "opaque"
    }),
    (0, types_1.op)({
        "tool": "wps_status",
        "action": null,
        "app": "common",
        "summary": "查看 WPS 连接状态、当前活动应用与所选工具面；编辑前先调用它",
        "params": {},
        "effect": "read",
        "advertised": true,
        "engine": "opaque"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_add_comment",
        "action": "addCellComment",
        "app": "excel",
        "summary": "给单元格添加批注。",
        "params": {
            "cell": {
                "type": "string",
                "description": "单元格地址，如 A1、B2",
                "required": true
            },
            "comment": {
                "type": "string",
                "description": "批注内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "cell",
            "comment"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_add_list_row",
        "action": "addListRow",
        "app": "excel",
        "summary": "给表末尾追加一行，可同时写入这一行的值（按列顺序）。列的格式与公式会自动带上。使用场景：往结构化表格里持续追加记录。",
        "params": {
            "table": {
                "type": "string",
                "description": "表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）",
                "required": true
            },
            "values": {
                "type": "array",
                "description": "按列顺序写这一行的值，如 [\"华东\", \"A\", 10]；数字请写数字（写成字符串会当文本写进单元格）。不填只加一个空行"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "table"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_add_sparkline",
        "action": "addSparkline",
        "app": "excel",
        "summary": "在单元格区域里加迷你图（单元格内的微型图表）：dataRange 是数据，location 是放图的位置，两者形状要一致（如 B2:B5 → C2:C5）。使用场景：在表格旁边一行一个小趋势图，不占地方。",
        "params": {
            "dataRange": {
                "type": "string",
                "description": "数据区域，如 B2:B5（每个单元格一条迷你图时按列给）",
                "required": true
            },
            "location": {
                "type": "string",
                "description": "放置位置，如 C2:C5；形状要与 dataRange 一致",
                "required": true
            },
            "sparklineType": {
                "type": "string",
                "description": "迷你图类型，默认 line",
                "enum": [
                    "line",
                    "column",
                    "winloss"
                ]
            },
            "markers": {
                "type": "boolean",
                "description": "是否标出数据点（折线图）"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "dataRange",
            "location"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_auto_filter",
        "action": "autoFilter",
        "app": "excel",
        "summary": "对Excel指定范围应用自动筛选。",
        "params": {
            "range": {
                "type": "string",
                "description": "筛选范围，如 A1:D100",
                "required": true
            },
            "column": {
                "type": "string",
                "description": "筛选列标识"
            },
            "criteria": {
                "type": "string",
                "description": "筛选条件"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_auto_fit",
        "action": "autoFitAll",
        "app": "excel",
        "summary": "按内容自动调整列宽与行高。使用场景：\"列宽太窄看不清\"、\"让表格自适应内容\"。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            },
            "range": {
                "type": "string",
                "description": "要调整的区域，如 A1:D20；不填则用整张表的已用范围"
            }
        },
        "effect": "write",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_auto_fit_columns",
        "action": "autoFitColumn",
        "app": "excel",
        "summary": "按内容自动调整列宽（不动行高）。使用 column 可只调整某一列。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            },
            "range": {
                "type": "string",
                "description": "要调整的区域，如 A1:D20；不填则用整张表的已用范围"
            },
            "column": {
                "type": "string",
                "description": "只调整这一列（列名如 B，或列号）"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_auto_fit_rows",
        "action": "autoFitRow",
        "app": "excel",
        "summary": "按内容自动调整行高（不动列宽）。使用 row 可只调整某一行。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            },
            "range": {
                "type": "string",
                "description": "要调整的区域，如 A1:D20；不填则用整张表的已用范围"
            },
            "row": {
                "type": "number",
                "description": "只调整这一行（从 1 开始）"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_auto_sum",
        "action": "autoSum",
        "app": "excel",
        "summary": "对指定范围的列或行自动求和，并将结果写入目标单元格。",
        "params": {
            "range": {
                "type": "string",
                "description": "要求和的数据范围，如 \"A1:A10\" 或 \"B2:F2\"",
                "required": true
            },
            "targetCell": {
                "type": "string",
                "description": "求和结果写入的目标单元格，如 \"A11\" 或 \"G2\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "targetCell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_calculate",
        "action": "calculateSheet",
        "app": "excel",
        "summary": "强制重算公式。all 为 true 时重算整个工作簿，否则只重算指定工作表。使用场景：刚写入公式要立刻拿结果；表格显示的是过期值（手动计算模式）。",
        "params": {
            "all": {
                "type": "boolean",
                "description": "true 重算整个工作簿；默认 false 只重算一张表"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表（all 为 true 时忽略）"
            }
        },
        "effect": "write",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_clean_data",
        "action": "cleanData",
        "app": "excel",
        "summary": "数据清洗工具，支持多种清洗操作的组合。\n\n使用场景：\n- \"把A列的前后空格去掉\" -> 使用 trim 操作\n- \"把日期格式统一成年-月-日\" -> 使用 unify_date 操作\n- \"删除空行\" -> 使用 remove_empty_rows 操作\n\n支持的操作：\n- trim: 去除单元格前后空格\n- remove_duplicates: 删除重复行\n- unify_date: 统一日期格式为 yyyy-mm-dd\n- remove_empty_rows: 删除空行",
        "params": {
            "range": {
                "type": "string",
                "description": "要清洗的数据范围，如 A1:D100",
                "required": true
            },
            "operations": {
                "type": "array",
                "description": "要执行的清洗操作列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要执行的清洗操作列表",
                    "items": {
                        "type": "string",
                        "enum": [
                            "trim",
                            "remove_duplicates",
                            "unify_date",
                            "remove_empty_rows"
                        ]
                    }
                },
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "operations"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_clear_formats",
        "action": "clearFormats",
        "app": "excel",
        "summary": "清除区域的格式（字体、颜色、边框、数字格式），单元格内容保留。使用场景：格式被弄乱了，恢复成默认样子。",
        "params": {
            "range": {
                "type": "string",
                "description": "目标区域，如 A1:D20",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": true,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_clear_pivot_table",
        "action": "clearPivotTable",
        "app": "excel",
        "summary": "清除透视表在表上的报表区域（数据源不动）。注意：WPS 清掉报表后透视表对象会留到保存/重开，期间它仍出现在透视表列表里——工具会如实报告剩余数量，不谎称已删除。使用场景：把临时透视表从工作表上拿掉。",
        "params": {
            "pivotTable": {
                "type": "string",
                "description": "要清除的透视表名",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": true,
        "required": [
            "pivotTable"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_clear_range",
        "action": "clearRange",
        "app": "excel",
        "summary": "清除指定范围的内容、格式或全部。",
        "params": {
            "range": {
                "type": "string",
                "description": "范围地址，如 A1:C10",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前工作表"
            },
            "type": {
                "type": "string",
                "description": "清除类型：all（全部）、contents（仅内容）、formats（仅格式）",
                "enum": [
                    "all",
                    "contents",
                    "formats"
                ]
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_clear_sparkline",
        "action": "clearSparkline",
        "app": "excel",
        "summary": "清除指定区域上的迷你图（数据不动）。使用场景：不想要这些微型图了。",
        "params": {
            "location": {
                "type": "string",
                "description": "迷你图所在区域，如 C2:C5",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "location"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_close_workbook",
        "action": "closeWorkbook",
        "app": "excel",
        "summary": "关闭指定的Excel工作簿，可选是否保存。",
        "params": {
            "name": {
                "type": "string",
                "description": "工作簿名称，不填则关闭当前工作簿"
            },
            "save": {
                "type": "boolean",
                "description": "是否保存，默认true"
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_consolidate",
        "action": "consolidate",
        "app": "excel",
        "summary": "把多块来源区域按指定函数汇总写入目标区域（Excel 的合并计算）。sources 形如 [Sheet1!A1:B4, Sheet2!A1:B4]。使用场景：多张同结构表加总到一张。",
        "params": {
            "destination": {
                "type": "string",
                "description": "汇总结果写入的区域左上角，如 E1",
                "required": true
            },
            "sources": {
                "type": "array",
                "description": "来源区域列表，如 [Sheet1!A1:B4, Sheet2!A1:B4]",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "items": {
                        "type": "string"
                    },
                    "description": "来源区域列表，如 [Sheet1!A1:B4, Sheet2!A1:B4]"
                },
                "required": true
            },
            "function": {
                "type": "string",
                "description": "汇总函数，默认 sum",
                "enum": [
                    "sum",
                    "count",
                    "average",
                    "max",
                    "min"
                ]
            },
            "topRow": {
                "type": "boolean",
                "description": "按标签合并：来源首行是标题。默认 false（按位置逐格相加，纯数字表用这个）"
            },
            "leftColumn": {
                "type": "boolean",
                "description": "按标签合并：来源首列是标题。默认 false"
            },
            "createLinks": {
                "type": "boolean",
                "description": "与来源建立链接，默认 false"
            },
            "sheet": {
                "type": "string",
                "description": "目标工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "destination",
            "sources"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_copy_format",
        "action": "copyFormat",
        "app": "excel",
        "summary": "把一块区域的格式复制到另一块区域（只复制格式，不改数值与公式）。使用场景：把 A1 的样式刷到整个 A 列、统一表头外观。要连值一起搬请用 wps_excel_write_range。",
        "params": {
            "source": {
                "type": "string",
                "description": "格式来源区域，如 A1",
                "required": true
            },
            "target": {
                "type": "string",
                "description": "格式目标区域，如 A2:A100",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "source",
            "target"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_copy_range",
        "action": "copyRange",
        "app": "excel",
        "summary": "复制Excel指定范围到目标位置。",
        "params": {
            "source": {
                "type": "string",
                "description": "源范围，如 A1:C10",
                "required": true
            },
            "destination": {
                "type": "string",
                "description": "目标位置，如 E1",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "source",
            "destination"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_copy_sheet",
        "action": "copySheet",
        "app": "excel",
        "summary": "复制当前工作簿中的指定工作表。可指定新名称和插入位置。",
        "params": {
            "name": {
                "type": "string",
                "description": "要复制的工作表名称",
                "required": true
            },
            "newName": {
                "type": "string",
                "description": "复制后的工作表名称，不填则自动生成"
            },
            "position": {
                "type": "number",
                "description": "插入位置索引，从0开始（0=最前）；不填则追加到末尾"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_create_chart",
        "action": "createChart",
        "app": "excel",
        "summary": "在Excel中创建图表。支持柱状图、折线图、饼图、散点图等多种类型。\n\n使用场景：\n- \"帮我用A1:B10的数据画个柱状图\" -> 创建 column_clustered\n- \"把这些数据做成折线图看趋势\" -> 创建 line\n- \"显示各部门占比\" -> 创建 pie 饼图\n- \"分析两个变量的相关性\" -> 创建 scatter 散点图\n\n支持的图表类型：\n- column_clustered: 簇状柱形图（默认，最常用）\n- column_stacked: 堆积柱形图\n- bar_clustered: 簇状条形图\n- line: 折线图\n- line_markers: 带标记的折线图\n- pie: 饼图\n- doughnut: 环形图\n- scatter: 散点图\n- area: 面积图\n- radar: 雷达图",
        "params": {
            "data_range": {
                "type": "string",
                "description": "数据范围，如 A1:C10，图表数据的来源",
                "required": true
            },
            "chart_type": {
                "type": "string",
                "description": "图表类型，默认 column_clustered（簇状柱形图）",
                "enum": [
                    "column_clustered",
                    "column_stacked",
                    "bar_clustered",
                    "line",
                    "line_markers",
                    "pie",
                    "doughnut",
                    "scatter",
                    "area",
                    "radar"
                ]
            },
            "title": {
                "type": "string",
                "description": "图表标题，如 \"销售趋势图\""
            },
            "position": {
                "type": "object",
                "description": "图表位置，不填则自动放在数据右侧",
                "schema": {
                    "type": "object",
                    "description": "图表位置，不填则自动放在数据右侧",
                    "properties": {
                        "left": {
                            "type": "number",
                            "description": "左边距（像素）"
                        },
                        "top": {
                            "type": "number",
                            "description": "上边距（像素）"
                        },
                        "width": {
                            "type": "number",
                            "description": "图表宽度（像素），默认480"
                        },
                        "height": {
                            "type": "number",
                            "description": "图表高度（像素），默认300"
                        }
                    }
                }
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            },
            "show_legend": {
                "type": "boolean",
                "description": "是否显示图例，默认true"
            },
            "show_data_labels": {
                "type": "boolean",
                "description": "是否显示数据标签，默认false"
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "data_range"
        ],
        "engine": "bridge",
        "aliases": {
            "data_range": "dataRange",
            "chart_type": "chartType",
            "show_legend": "showLegend",
            "show_data_labels": "showDataLabels"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_excel_create_list_object",
        "action": "createListObject",
        "app": "excel",
        "summary": "把一块区域变成「表」（ListObject）：自动带表头与筛选按钮，并可用结构化引用（表名[列名]）写公式，之后能按表加行/删行/加总计行。使用场景：数据要反复增删、要按列筛选、要用结构化引用。默认认为首行是标题。",
        "params": {
            "range": {
                "type": "string",
                "description": "建表的数据区域，含表头，如 A1:C10",
                "required": true
            },
            "name": {
                "type": "string",
                "description": "表名（不填由 WPS 自动命名，如 表1）；这个名字要能写进公式"
            },
            "hasHeaders": {
                "type": "boolean",
                "description": "首行是否为标题，默认 true；纯数据请传 false"
            },
            "tableStyle": {
                "type": "string",
                "description": "表格样式名，如 TableStyleMedium2；不填用 WPS 默认"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_create_pivot_table",
        "action": "createPivotTable",
        "app": "excel",
        "summary": "创建Excel透视表，用于数据汇总和分析。\n\n使用场景：\n- \"帮我按部门统计销售额\" -> 行字段=部门，值字段=销售额(SUM)\n- \"按月份和产品分析数量\" -> 行字段=月份，列字段=产品，值字段=数量(SUM)\n- \"统计各地区的订单数\" -> 行字段=地区，值字段=订单号(COUNT)\n\n注意事项：\n- sourceRange必须包含表头行\n- rowFields、columnFields、valueFields中的field必须是表头中的列名\n- 透视表会从destinationCell开始向右下方扩展",
        "params": {
            "sourceRange": {
                "type": "string",
                "description": "数据源范围，如\"A1:E100\"。必须包含表头行，数据要连续无空行",
                "required": true
            },
            "destinationCell": {
                "type": "string",
                "description": "透视表放置位置（左上角单元格），如\"G1\"、\"H5\"",
                "required": true
            },
            "rowFields": {
                "type": "array",
                "description": "行字段列名列表，如[\"部门\", \"员工\"]。这些字段会作为透视表的行标签",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "行字段列名列表，如[\"部门\", \"员工\"]。这些字段会作为透视表的行标签",
                    "items": {
                        "type": "string"
                    }
                },
                "required": true
            },
            "columnFields": {
                "type": "array",
                "description": "列字段列名列表（可选），如[\"月份\"]。这些字段会作为透视表的列标签",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "列字段列名列表（可选），如[\"月份\"]。这些字段会作为透视表的列标签",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "valueFields": {
                "type": "array",
                "description": "值字段配置列表。每项包含field(字段名)和aggregation(聚合方式:SUM/COUNT/AVERAGE/MAX/MIN)",
                "items": {
                    "type": "object"
                },
                "schema": {
                    "type": "array",
                    "description": "值字段配置列表。每项包含field(字段名)和aggregation(聚合方式:SUM/COUNT/AVERAGE/MAX/MIN)",
                    "items": {
                        "type": "object",
                        "properties": {
                            "field": {
                                "type": "string",
                                "description": "字段名/列名"
                            },
                            "aggregation": {
                                "type": "string",
                                "description": "聚合方式：SUM(求和)、COUNT(计数)、AVERAGE(平均)、MAX(最大)、MIN(最小)",
                                "enum": [
                                    "SUM",
                                    "COUNT",
                                    "AVERAGE",
                                    "MAX",
                                    "MIN"
                                ]
                            }
                        },
                        "required": [
                            "field",
                            "aggregation"
                        ]
                    }
                },
                "required": true
            },
            "filterFields": {
                "type": "array",
                "description": "筛选字段列名列表（可选），这些字段会出现在透视表上方作为筛选器",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "筛选字段列名列表（可选），这些字段会出现在透视表上方作为筛选器",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "tableName": {
                "type": "string",
                "description": "透视表名称（可选），不填则自动生成"
            },
            "destinationSheet": {
                "type": "string",
                "description": "目标工作表名称（可选），透视表将创建在此工作表，不填则使用当前工作表"
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "sourceRange",
            "destinationCell",
            "rowFields",
            "valueFields"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_create_sheet",
        "action": "createSheet",
        "app": "excel",
        "summary": "在当前工作簿中创建新的工作表。可指定名称和插入位置。",
        "params": {
            "name": {
                "type": "string",
                "description": "新工作表的名称",
                "required": true
            },
            "position": {
                "type": "number",
                "description": "插入位置索引，从0开始（0=最前）；不填则追加到末尾"
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_create_workbook",
        "action": "createWorkbook",
        "app": "excel",
        "summary": "新建一个空白Excel工作簿。",
        "params": {},
        "effect": "lifecycle",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_cell_comment",
        "action": "deleteCellComment",
        "app": "excel",
        "summary": "删除Excel单元格上的批注。",
        "params": {
            "cell": {
                "type": "string",
                "description": "单元格地址，如 A1、B2",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "cell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_chart",
        "action": "deleteChart",
        "app": "excel",
        "summary": "删除工作表上的图表（不删它引用的数据）。给 chart 名字或序号；表上只有一张图时可以省略。使用场景：清掉临时图表。",
        "params": {
            "chart": {
                "type": "string",
                "description": "图表名（如 Chart 1）或该表上的序号；只有一张图时可省略"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_columns",
        "action": "deleteColumns",
        "app": "excel",
        "summary": "删除Excel中指定位置的一列或多列。",
        "params": {
            "column": {
                "type": "string",
                "description": "起始列字母，如 A、B、C",
                "required": true
            },
            "count": {
                "type": "number",
                "description": "删除列数，默认1"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "column"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_list_row",
        "action": "deleteListRow",
        "app": "excel",
        "summary": "删除表里的第几行（表体行，从 1 开始，不含表头）。使用场景：剔除一条记录。要按条件删请先 read_range 找到行号，或者直接重写整块数据。",
        "params": {
            "table": {
                "type": "string",
                "description": "表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）",
                "required": true
            },
            "rowIndex": {
                "type": "number",
                "description": "删第几行（表体行，从 1 开始，不含表头）",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "table",
            "rowIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_named_range",
        "action": "deleteNamedRange",
        "app": "excel",
        "summary": "删除指定的命名范围（只删名字，不动单元格内容）。使用场景：\"把这个没用的名字去掉\"。",
        "params": {
            "name": {
                "type": "string",
                "description": "要删除的命名范围名称",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_rows",
        "action": "deleteRows",
        "app": "excel",
        "summary": "删除Excel中指定位置的一行或多行。",
        "params": {
            "startRow": {
                "type": "number",
                "description": "起始行号（从1开始）"
            },
            "row": {
                "type": "number",
                "description": "要删除的行号（与 startRow 等价，便于合并旧工具）"
            },
            "count": {
                "type": "number",
                "description": "删除行数，默认1"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_delete_sheet",
        "action": "deleteSheet",
        "app": "excel",
        "summary": "删除当前工作簿中的指定工作表。注意：此操作不可撤销。",
        "params": {
            "name": {
                "type": "string",
                "description": "要删除的工作表名称",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_diagnose_formula",
        "action": "diagnoseFormula",
        "app": "excel",
        "summary": "诊断公式错误，分析原因并提供修复建议。\n\n使用场景：\n- 用户说\"这个公式报错了\"\n- 用户说\"#REF! 是什么意思\"\n- 用户说\"帮我看看公式哪里有问题\"\n\n支持诊断的错误类型：\n- #REF! - 引用了不存在的单元格\n- #N/A - 查找函数未找到匹配值\n- #VALUE! - 参数类型错误\n- #NAME? - 函数名称错误\n- #DIV/0! - 除数为零\n- #NUM! - 数值问题\n- #NULL! - 交集为空",
        "params": {
            "cell": {
                "type": "string",
                "description": "包含错误公式的单元格地址，如 A1、B2",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "cell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_evaluate_formula",
        "action": "evaluateFormula",
        "app": "excel",
        "summary": "计算并返回公式结果（由 Excel 求值，不写入任何单元格）。\n\n需要在某个单元格里求值并保留公式时，用 wps_excel_set_formula 写入再从该格读值。",
        "params": {
            "formula": {
                "type": "string",
                "description": "要计算的公式，如 =SUM(A1:A10)",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "formula"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_export_chart_as_image",
        "action": "exportChartAsImage",
        "app": "excel",
        "summary": "将工作表中指定的图表导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。\n\n调用底层 WPS Excel 原生接口 Chart.Export(FileName, FilterName)，最简单可靠，\n实现 1:1 像素级还原，避免通过 PDF + pdf2image 中转造成的版式/字体/坐标轴失真。\n\n支持的 format（FilterName）取值：\n- PNG（默认，推荐用于截图与无损展示）\n- JPG / JPEG（自动按 JPG 滤镜处理，体积更小）\n- GIF（限 256 色，适合简单图形）\n- BMP（无压缩位图，文件最大）\n\n使用场景：\n- \"把 Sheet1 上的 Chart 1 导出成 PNG\"\n- \"导出销售分析柱状图给我\"\n- \"把图表保存为高清图片用于报告\"\n\n注意：\n- outputPath 必须是绝对路径\n- chartName 通常为 \"Chart 1\"、\"图表 1\" 等，可通过 wps_excel_create_chart 返回值或界面查看\n- macOS 上建议输出到 ~/Downloads 等用户可写目录，避免沙箱权限拒绝",
        "params": {
            "chartName": {
                "type": "string",
                "description": "图表名称（如 \"Chart 1\" 或 \"图表 1\"）",
                "required": true
            },
            "outputPath": {
                "type": "string",
                "description": "输出图片文件的绝对路径（如 /Users/xxx/Downloads/chart1.png）",
                "required": true
            },
            "format": {
                "type": "string",
                "description": "图片格式，默认 PNG",
                "enum": [
                    "PNG",
                    "JPG",
                    "JPEG",
                    "GIF",
                    "BMP"
                ]
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "export",
        "advertised": false,
        "required": [
            "chartName",
            "outputPath"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_export_range_as_image",
        "action": "exportRangeAsImage",
        "app": "excel",
        "summary": "将工作表中指定区域导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。\n\n实现原理（经典临时图表法）：\n1. Range.CopyPicture 把区域复制为位图到剪贴板\n2. 临时插入一个等尺寸 ChartObject\n3. Chart.Paste 把剪贴板位图贴入图表\n4. Chart.Export 导出图表为图片\n5. 删除临时图表\n\n效果：1:1 还原单元格内容、字体、边框、合并单元格、单元格背景色等所有视觉元素，\n避免通过 PDF + pdf2image 中转造成的版式失真。\n\n支持的 format 取值：\n- PNG（默认）/ JPG / JPEG / GIF / BMP（JPEG 在底层归一化为 JPG）\n\n使用场景：\n- \"把 A1:F20 表格导出成 PNG\"\n- \"导出当前月份的销售统计表为图片用于报告\"\n- \"把这块数据区域保存为图片粘贴到 PPT\"\n\n注意：\n- outputPath 必须是绝对路径\n- 此方法依赖剪贴板，并发或 headless 模式下可能失败，handler 内部已包含临时图表清理与异常回滚\n- range 必须为合法 A1 范围格式（如 \"A1:F20\"），不支持命名范围",
        "params": {
            "range": {
                "type": "string",
                "description": "区域地址（如 \"A1:F20\"）",
                "required": true
            },
            "outputPath": {
                "type": "string",
                "description": "输出图片文件的绝对路径（如 /Users/xxx/Downloads/range.png）",
                "required": true
            },
            "format": {
                "type": "string",
                "description": "图片格式，默认 PNG",
                "enum": [
                    "PNG",
                    "JPG",
                    "JPEG",
                    "GIF",
                    "BMP"
                ]
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "export",
        "advertised": false,
        "required": [
            "range",
            "outputPath"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_fill_series",
        "action": "fillSeries",
        "app": "excel",
        "summary": "自动填充序列数据。",
        "params": {
            "range": {
                "type": "string",
                "description": "填充范围，如 A1:A10"
            },
            "direction": {
                "type": "string",
                "description": "填充方向",
                "enum": [
                    "down",
                    "right",
                    "up",
                    "left"
                ]
            },
            "type": {
                "type": "string",
                "description": "填充类型",
                "enum": [
                    "linear",
                    "growth",
                    "date",
                    "auto"
                ]
            },
            "step": {
                "type": "number",
                "description": "步长值"
            },
            "startValue": {
                "type": "number",
                "description": "序列起始值，默认1"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            },
            "sourceRange": {
                "type": "string",
                "description": "自动填充的源区域（与 targetRange 配对使用）"
            },
            "targetRange": {
                "type": "string",
                "description": "自动填充的目标区域"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_find_in_sheet",
        "action": "findInSheet",
        "app": "excel",
        "summary": "在工作表里查找文本并返回每一个命中的单元格地址（不改动任何内容）。\n\n使用场景：\n- \"帮我找一下'华东'出现在哪些格子里\"\n- 先定位再决定怎么改，比直接替换安全\n\n只统计、不修改。要替换请用 wps_excel_find_replace。",
        "params": {
            "searchText": {
                "type": "string",
                "description": "要查找的文本",
                "required": true
            },
            "range": {
                "type": "string",
                "description": "查找范围；不填则用已用范围"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            },
            "matchCase": {
                "type": "boolean",
                "description": "是否区分大小写，默认 false"
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "searchText"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_find_replace",
        "action": "findReplaceExcel",
        "app": "excel",
        "summary": "在Excel中查找并替换内容。",
        "params": {
            "find": {
                "type": "string",
                "description": "要查找的文本",
                "required": true
            },
            "replace": {
                "type": "string",
                "description": "替换为的文本",
                "required": true
            },
            "matchCase": {
                "type": "boolean",
                "description": "是否区分大小写，默认false"
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "find",
            "replace"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_freeze_panes",
        "action": "freezePanes",
        "app": "excel",
        "summary": "冻结/取消冻结窗格。可指定冻结的行和列位置。",
        "params": {
            "row": {
                "type": "number",
                "description": "冻结到第几行（从1开始），不填则不冻结行"
            },
            "column": {
                "type": "number",
                "description": "冻结到第几列（从1开始），不填则不冻结列"
            },
            "freeze": {
                "type": "boolean",
                "description": "是否冻结，默认true。设为false则取消冻结"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_generate_formula",
        "action": "getContext",
        "app": "excel",
        "summary": "根据自然语言描述生成Excel公式。这是解决用户\"公式不会写\"痛点的核心工具。\n\n使用场景：\n- 用户说\"帮我写个公式查价格\" -> 生成 VLOOKUP/XLOOKUP\n- 用户说\"如果大于100就显示达标\" -> 生成 IF 公式\n- 用户说\"统计每个部门的人数\" -> 生成 COUNTIF\n- 用户说\"求这列的平均值\" -> 生成 AVERAGE\n\n调用此工具会返回当前工作表的上下文信息，包括表头、选中区域等，便于生成准确的公式。",
        "params": {
            "description": {
                "type": "string",
                "description": "用户对公式需求的自然语言描述，如\"查找产品名对应的价格\"、\"计算A列的总和\"",
                "required": true,
                "kind": "local"
            },
            "target_cell": {
                "type": "string",
                "description": "目标单元格地址，如 B2。不填则返回上下文让用户确认",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "description"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_cell_comments",
        "action": "getCellComments",
        "app": "excel",
        "summary": "获取Excel指定范围内的所有批注。不指定范围则获取当前工作表所有批注。",
        "params": {
            "range": {
                "type": "string",
                "description": "要查询的范围，如 A1:D10。不填则获取所有批注"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_cell_info",
        "action": "getCellInfo",
        "app": "excel",
        "summary": "获取单元格的详细信息（值、公式、格式等）。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名称",
                "required": true
            },
            "cell": {
                "type": "string",
                "description": "单元格地址，如 A1",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "sheet",
            "cell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_cell_value",
        "action": "getCellValue",
        "app": "excel",
        "summary": "获取Excel指定单元格的值。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名称",
                "required": true
            },
            "row": {
                "type": "number",
                "description": "行号（从1开始）",
                "required": true
            },
            "col": {
                "type": "number",
                "description": "列号（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "sheet",
            "row",
            "col"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_conditional_formats",
        "action": "getConditionalFormats",
        "app": "excel",
        "summary": "列出区域上生效的条件格式规则（序号 + 类型），用于先看清楚再改。使用场景：这个表为什么某些格子会变红；删规则之前先确认删哪一条。",
        "params": {
            "range": {
                "type": "string",
                "description": "要查看的区域，如 A1:A100",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_data_validations",
        "action": "getDataValidations",
        "app": "excel",
        "summary": "读取区域上的数据验证规则（类型、来源公式、提示语）。使用场景：这个下拉框的选项是从哪来的；删规则之前先确认规则内容。",
        "params": {
            "range": {
                "type": "string",
                "description": "要查看的区域，如 B2:B100",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_formula",
        "action": "getFormula",
        "app": "excel",
        "summary": "获取Excel指定单元格的公式。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名称",
                "required": true
            },
            "cell": {
                "type": "string",
                "description": "单元格地址，如 A1、B2",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "sheet",
            "cell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_formula_audit",
        "action": "getFormulaAudit",
        "app": "excel",
        "summary": "审计一个单元格的公式依赖：它引用了谁（precedents）、谁引用了它（dependents）、直接引用几处，并可选在界面上画出追踪箭头。使用场景：\"这个数是怎么算出来的\"、\"改这个格子会影响哪些单元格\"。",
        "params": {
            "cell": {
                "type": "string",
                "description": "要审计的单元格，如 C5",
                "required": true
            },
            "showPrecedents": {
                "type": "boolean",
                "description": "在界面上画出引用来源箭头"
            },
            "showDependents": {
                "type": "boolean",
                "description": "在界面上画出被引用箭头"
            },
            "clearArrows": {
                "type": "boolean",
                "description": "先清除工作表上的所有追踪箭头"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "cell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_list_objects",
        "action": "getListObjects",
        "app": "excel",
        "summary": "列出工作簿（或指定工作表）里的全部「表」及其结构：名字、范围、行列数、列名、表格样式、总计行开关，以及每列可直接写进公式的结构化引用。使用场景：先看清有哪几张表、列名叫什么，再决定怎么改。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "只看这张工作表；不填则列出整个工作簿"
            }
        },
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_named_ranges",
        "action": "getNamedRanges",
        "app": "excel",
        "summary": "列出工作簿里的全部命名范围及其引用位置。使用场景：\"这个工作簿里定义了哪些名字\"。",
        "params": {},
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_open_workbooks",
        "action": "getOpenWorkbooks",
        "app": "excel",
        "summary": "获取当前所有已打开的Excel工作簿列表。",
        "params": {},
        "effect": "lifecycle",
        "advertised": true,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_pivot_tables",
        "action": "getPivotTables",
        "app": "excel",
        "summary": "列出工作簿（或指定工作表）上的透视表：名字、所在区域，以及（能读到时）行字段与数据字段。使用场景：先看清有哪几张透视表、它们叫什么，再刷新或清除。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "只看这张工作表；不填则列出整个工作簿"
            }
        },
        "effect": "read",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_selection",
        "action": "getSelection",
        "app": "excel",
        "summary": "获取当前Excel中选中区域的信息，包括范围地址、行列数等。",
        "params": {},
        "effect": "read",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_sheet_info",
        "action": "getExcelContext",
        "app": "excel",
        "summary": "获取工作表的结构信息：工作簿与工作表名、已用范围地址、表头、当前单元格。\n\n使用场景：\n- \"这张表有多大\" / \"数据到哪一行\"\n- 读数据之前先确定范围，而不是猜一个很大的区域\n\n先调用它拿到 usedRange，再用 read_range 精确读取，避免把整片空白也读回来。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_sheet_list",
        "action": "getSheetList",
        "app": "excel",
        "summary": "获取当前工作簿的所有工作表列表，包含名称、索引和是否为活动工作表。",
        "params": {},
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_get_sheet_settings",
        "action": "getSheetSettings",
        "app": "excel",
        "summary": "读一张工作表的页面设置、打印设置、页眉页脚与外观：方向、纸张、页边距（磅）、缩放或按页适配、是否居中、打印区域与打印标题、页眉页脚、可见性、标签色、手动分页符数量。使用场景：打印前先看清现状；改完再回读确认。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_goal_seek",
        "action": "goalSeek",
        "app": "excel",
        "summary": "单变量求解：反复调整 changingCell，直到 cell 的公式结果等于 goal。cell 必须是带公式的单元格。使用场景：\"要利润到 100 万，销量得多少\"。结果是近似解，工具会回读调整后的取值。",
        "params": {
            "cell": {
                "type": "string",
                "description": "带公式的目标单元格，如 B4",
                "required": true
            },
            "goal": {
                "type": "number",
                "description": "希望目标单元格达到的值",
                "required": true
            },
            "changingCell": {
                "type": "string",
                "description": "被反复调整的单元格（不能有公式），如 B2",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "cell",
            "goal",
            "changingCell"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_group_columns",
        "action": "groupColumns",
        "app": "excel",
        "summary": "把一段列折叠分组（分级显示）。列名如 B、E 或列号 2、5。使用场景：把中间的计算列收起来，只留结果列。",
        "params": {
            "startColumn": {
                "type": "string",
                "description": "起始列（列名如 B，或列号）",
                "required": true
            },
            "endColumn": {
                "type": "string",
                "description": "结束列（列名如 E，或列号）",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "startColumn",
            "endColumn"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_group_rows",
        "action": "groupRows",
        "app": "excel",
        "summary": "对Excel中指定范围的行进行分组，便于折叠/展开管理。",
        "params": {
            "startRow": {
                "type": "number",
                "description": "起始行号（从1开始）",
                "required": true
            },
            "endRow": {
                "type": "number",
                "description": "结束行号（从1开始）",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "startRow",
            "endRow"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_hide_column",
        "action": "hideColumns",
        "app": "excel",
        "summary": "隐藏或显示指定列。可指定起始列号、列数和隐藏/显示状态。",
        "params": {
            "column": {
                "type": "number",
                "description": "起始列号（从1开始）",
                "required": true
            },
            "count": {
                "type": "number",
                "description": "列数，默认1",
                "required": true
            },
            "hide": {
                "type": "boolean",
                "description": "是否隐藏，true为隐藏，false为显示",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "column",
            "count",
            "hide"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_hide_rows",
        "action": "hideRows",
        "app": "excel",
        "summary": "隐藏Excel中指定范围的行。",
        "params": {
            "startRow": {
                "type": "number",
                "description": "起始行号（从1开始）"
            },
            "endRow": {
                "type": "number",
                "description": "结束行号（从1开始）"
            },
            "row": {
                "type": "number",
                "description": "单行行号（与 startRow 等价）"
            },
            "rows": {
                "type": "array",
                "description": "行号数组"
            },
            "count": {
                "type": "number",
                "description": "从 row 起的连续行数"
            },
            "hide": {
                "type": "boolean",
                "description": "true 隐藏（默认），false 显示"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_insert_columns",
        "action": "insertColumns",
        "app": "excel",
        "summary": "在Excel中指定位置插入一列或多列。",
        "params": {
            "column": {
                "type": "string",
                "description": "在哪一列前插入，如 A、B、C",
                "required": true
            },
            "count": {
                "type": "number",
                "description": "插入列数，默认1"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "column"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_insert_excel_image",
        "action": "insertExcelImage",
        "app": "excel",
        "summary": "在Excel中插入图片到指定位置。",
        "params": {
            "filePath": {
                "type": "string",
                "description": "图片文件路径",
                "required": true
            },
            "cell": {
                "type": "string",
                "description": "插入位置的单元格地址，如 A1。不填则插入到当前选中位置"
            },
            "width": {
                "type": "number",
                "description": "图片宽度（像素），不填则使用原始宽度"
            },
            "height": {
                "type": "number",
                "description": "图片高度（像素），不填则使用原始高度"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "filePath"
        ],
        "engine": "bridge",
        "aliases": {
            "filePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_excel_insert_rows",
        "action": "insertRows",
        "app": "excel",
        "summary": "在Excel中指定位置插入一行或多行。",
        "params": {
            "row": {
                "type": "number",
                "description": "在第几行前插入（从1开始）",
                "required": true
            },
            "count": {
                "type": "number",
                "description": "插入行数，默认1"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "row"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_lock_cells",
        "action": "lockCells",
        "app": "excel",
        "summary": "锁定或解锁Excel指定范围的单元格（需配合工作表保护使用）。",
        "params": {
            "range": {
                "type": "string",
                "description": "要锁定/解锁的范围，如 A1:D10",
                "required": true
            },
            "locked": {
                "type": "boolean",
                "description": "true为锁定，false为解锁",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "locked"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_merge_cells",
        "action": "mergeCells",
        "app": "excel",
        "summary": "合并Excel指定范围的单元格。合并后保留左上角单元格的值。",
        "params": {
            "range": {
                "type": "string",
                "description": "要合并的范围，如 A1:C3",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_move_sheet",
        "action": "moveSheet",
        "app": "excel",
        "summary": "移动指定工作表到新的位置。",
        "params": {
            "name": {
                "type": "string",
                "description": "要移动的工作表名称",
                "required": true
            },
            "position": {
                "type": "number",
                "description": "目标位置索引，从0开始（0=最前）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "name",
            "position"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_open_workbook",
        "action": "openWorkbook",
        "app": "excel",
        "summary": "打开指定路径的Excel工作簿文件。",
        "params": {
            "filePath": {
                "type": "string",
                "description": "工作簿文件路径",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "filePath"
        ],
        "engine": "bridge",
        "aliases": {
            "filePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_excel_paste_range",
        "action": "pasteRange",
        "app": "excel",
        "summary": "粘贴已复制的内容到指定位置。",
        "params": {
            "destination": {
                "type": "string",
                "description": "目标位置，如 E1",
                "required": true
            },
            "pasteType": {
                "type": "string",
                "description": "粘贴类型",
                "enum": [
                    "all",
                    "values",
                    "formats",
                    "formulas"
                ]
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "destination"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_protect_sheet",
        "action": "protectSheet",
        "app": "excel",
        "summary": "保护或取消保护工作表。",
        "params": {
            "password": {
                "type": "string",
                "description": "保护密码（可选）"
            },
            "protect": {
                "type": "boolean",
                "description": "是否保护，默认true"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_protect_workbook",
        "action": "protectWorkbook",
        "app": "excel",
        "summary": "保护或取消保护工作簿，防止结构被修改（如添加/删除工作表）。",
        "params": {
            "password": {
                "type": "string",
                "description": "保护密码",
                "required": true
            },
            "protect": {
                "type": "boolean",
                "description": "是否保护，true为保护，false为取消保护",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "password",
            "protect"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_read_range",
        "action": "getRangeData",
        "app": "excel",
        "summary": "读取Excel指定范围的单元格数据，返回二维数组格式的数据。",
        "params": {
            "range": {
                "type": "string",
                "description": "要读取的范围，如 A1:C10、B2:D5",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            },
            "include_header": {
                "type": "boolean",
                "description": "是否将第一行作为表头返回，默认false",
                "kind": "local"
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_refresh_all_data",
        "action": "refreshAllData",
        "app": "excel",
        "summary": "刷新整个工作簿的外部数据连接与透视表（相当于 Excel 的「全部刷新」）。使用场景：多个数据源都要更新一次。",
        "params": {},
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_refresh_links",
        "action": "refreshLinks",
        "app": "excel",
        "summary": "刷新工作簿引用的全部外部链接并报告条数。使用场景：数据源文件更新了，把引用拉一遍。没有链接时如实报告 0 条。",
        "params": {},
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_refresh_pivot_tables",
        "action": "refreshPivotTables",
        "app": "excel",
        "summary": "刷新透视表：给 pivotTable 只刷新那一张，不给就刷新目标工作表上的全部。使用场景：源数据改了，透视表还是旧数字。",
        "params": {
            "pivotTable": {
                "type": "string",
                "description": "透视表名；不填则刷新目标工作表上的全部透视表"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_remove_conditional_format",
        "action": "removeConditionalFormat",
        "app": "excel",
        "summary": "删除区域上的条件格式规则。给 index 删指定的一条（序号见 wps_excel_get_conditional_formats），不填则删该区域的全部规则。只删规则，不动内容与普通格式。",
        "params": {
            "range": {
                "type": "string",
                "description": "目标区域，如 A1:A100",
                "required": true
            },
            "index": {
                "type": "number",
                "description": "只删第几条规则（从 1 开始）；不填则删全部"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_remove_data_validation",
        "action": "removeDataValidation",
        "app": "excel",
        "summary": "删除区域上的数据验证规则（下拉框、输入限制）。使用场景：去掉这列的下拉限制。只删规则，不动单元格内容。",
        "params": {
            "range": {
                "type": "string",
                "description": "目标区域，如 B2:B100",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_remove_duplicates",
        "action": "removeDuplicates",
        "app": "excel",
        "summary": "删除指定范围内的重复行。可以指定根据哪些列判断重复。",
        "params": {
            "range": {
                "type": "string",
                "description": "要处理的数据范围，如 A1:D100",
                "required": true
            },
            "columns": {
                "type": "array",
                "description": "用于判断重复的列，如 [\"A\", \"B\"]。不填则根据所有列判断",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "用于判断重复的列，如 [\"A\", \"B\"]。不填则根据所有列判断",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "has_header": {
                "type": "boolean",
                "description": "第一行是否为表头，默认true"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge",
        "aliases": {
            "has_header": "hasHeader"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_excel_rename_sheet",
        "action": "renameSheet",
        "app": "excel",
        "summary": "重命名当前工作簿中的指定工作表。",
        "params": {
            "oldName": {
                "type": "string",
                "description": "当前工作表名称",
                "required": true
            },
            "newName": {
                "type": "string",
                "description": "新的工作表名称",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "oldName",
            "newName"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_reset_page_breaks",
        "action": "resetPageBreaks",
        "app": "excel",
        "summary": "清除工作表上的手动分页符，恢复按内容自动分页。使用场景：手工插过分页符之后想回到自动分页。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_resize_list_object",
        "action": "resizeListObject",
        "app": "excel",
        "summary": "调整表覆盖的区域（长短变化），表名、样式与结构化引用都保留。使用场景：数据变长了，把表扩到新范围。",
        "params": {
            "table": {
                "type": "string",
                "description": "表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）",
                "required": true
            },
            "range": {
                "type": "string",
                "description": "表的新范围，含表头，如 A1:C20",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "table",
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_array_formula",
        "action": "setArrayFormula",
        "app": "excel",
        "summary": "为Excel指定范围设置数组公式（CSE数组公式）。",
        "params": {
            "range": {
                "type": "string",
                "description": "数组公式应用的范围，如 A1:A10",
                "required": true
            },
            "formula": {
                "type": "string",
                "description": "数组公式，如 =A1:A10*B1:B10",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "formula"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_border",
        "action": "setBorder",
        "app": "excel",
        "summary": "设置Excel单元格边框样式，支持不同粗细、位置和颜色。",
        "params": {
            "range": {
                "type": "string",
                "description": "要设置边框的范围，如 A1:C10",
                "required": true
            },
            "borderStyle": {
                "type": "string",
                "description": "边框线条样式",
                "enum": [
                    "thin",
                    "medium",
                    "thick",
                    "double",
                    "none"
                ],
                "required": true
            },
            "position": {
                "type": "string",
                "description": "边框位置，默认 all（全部边框）",
                "enum": [
                    "all",
                    "top",
                    "bottom",
                    "left",
                    "right",
                    "outline"
                ]
            },
            "color": {
                "type": "string",
                "description": "边框颜色，十六进制如 #000000，默认黑色"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "borderStyle"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_cell_format",
        "action": "setCellFormat",
        "app": "excel",
        "summary": "设置Excel单元格格式，包括字体、颜色、背景色、粗体、斜体、字号等。",
        "params": {
            "range": {
                "type": "string",
                "description": "要设置格式的范围，如 A1:C10、B2:D5",
                "required": true
            },
            "format": {
                "type": "object",
                "description": "格式设置对象，可包含 bold(粗体)、italic(斜体)、fontSize(字号)、fontName(字体名)、fontColor(字体颜色，如#FF0000)、bgColor(背景颜色)、underline(下划线)、strikethrough(删除线)、horizontalAlignment(水平对齐: left/center/right)、verticalAlignment(垂直对齐: top/center/bottom)、wrapText(自动换行)",
                "schema": {
                    "type": "object",
                    "description": "格式设置对象，可包含 bold(粗体)、italic(斜体)、fontSize(字号)、fontName(字体名)、fontColor(字体颜色，如#FF0000)、bgColor(背景颜色)、underline(下划线)、strikethrough(删除线)、horizontalAlignment(水平对齐: left/center/right)、verticalAlignment(垂直对齐: top/center/bottom)、wrapText(自动换行)",
                    "properties": {
                        "bold": {
                            "type": "boolean",
                            "description": "是否粗体"
                        },
                        "italic": {
                            "type": "boolean",
                            "description": "是否斜体"
                        },
                        "fontSize": {
                            "type": "number",
                            "description": "字号大小"
                        },
                        "fontName": {
                            "type": "string",
                            "description": "字体名称，如 微软雅黑、Arial"
                        },
                        "fontColor": {
                            "type": "string",
                            "description": "字体颜色，十六进制如 #FF0000"
                        },
                        "bgColor": {
                            "type": "string",
                            "description": "背景颜色，十六进制如 #FFFF00"
                        },
                        "underline": {
                            "type": "boolean",
                            "description": "是否下划线"
                        },
                        "strikethrough": {
                            "type": "boolean",
                            "description": "是否删除线"
                        },
                        "horizontalAlignment": {
                            "type": "string",
                            "description": "水平对齐方式",
                            "enum": [
                                "left",
                                "center",
                                "right"
                            ]
                        },
                        "verticalAlignment": {
                            "type": "string",
                            "description": "垂直对齐方式",
                            "enum": [
                                "top",
                                "center",
                                "bottom"
                            ]
                        },
                        "wrapText": {
                            "type": "boolean",
                            "description": "是否自动换行"
                        }
                    }
                },
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "range",
            "format"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_cell_style",
        "action": "setCellStyle",
        "app": "excel",
        "summary": "应用预定义样式到Excel单元格，如标题、强调、输入、输出等内置样式。",
        "params": {
            "range": {
                "type": "string",
                "description": "要应用样式的范围，如 A1:C10",
                "required": true
            },
            "style": {
                "type": "string",
                "description": "预定义样式名称，如 标题、强调、好、差、适中、输入、输出、计算、检查单元格、解释性文本、汇总 等",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "style"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_cell_value",
        "action": "setCellValue",
        "app": "excel",
        "summary": "设置Excel指定单元格的值。",
        "params": {
            "sheet": {
                "type": "string",
                "description": "工作表名称",
                "required": true
            },
            "row": {
                "type": "number",
                "description": "行号（从1开始）",
                "required": true
            },
            "col": {
                "type": "number",
                "description": "列号（从1开始）",
                "required": true
            },
            "value": {
                "type": "string",
                "description": "要设置的值",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "sheet",
            "row",
            "col",
            "value"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_chart_labels",
        "action": "setChartLabels",
        "app": "excel",
        "summary": "给图表加标题与坐标轴标题（分类轴 = 横轴，数值轴 = 纵轴）。使用场景：裸图没人看得懂，补上「月度销售」「月份」「金额」。",
        "params": {
            "title": {
                "type": "string",
                "description": "图表标题"
            },
            "categoryAxisTitle": {
                "type": "string",
                "description": "分类轴（横轴）标题"
            },
            "valueAxisTitle": {
                "type": "string",
                "description": "数值轴（纵轴）标题"
            },
            "chart": {
                "type": "string",
                "description": "图表名（如 Chart 1）或该表上的序号；只有一张图时可省略"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_column_width",
        "action": "setColumnWidth",
        "app": "excel",
        "summary": "设置Excel指定列的列宽。支持单列或连续多列。",
        "params": {
            "column": {
                "type": "string",
                "description": "列标识，如 A（单列）或 A:C（连续多列）",
                "required": true
            },
            "width": {
                "type": "number",
                "description": "列宽数值（字符宽度单位），如 15、20",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "column",
            "width"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_conditional_format",
        "action": "addConditionalFormat",
        "app": "excel",
        "summary": "设置条件格式。",
        "params": {
            "range": {
                "type": "string",
                "description": "要设置条件格式的范围，如 A1:D100",
                "required": true
            },
            "condition": {
                "type": "string",
                "description": "条件表达式，如 \">100\"、\"=0\"、\"between(1,10)\"",
                "required": true
            },
            "format": {
                "type": "string",
                "description": "格式描述，如 \"red_fill\"、\"bold\"、\"green_font\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "condition",
            "format"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_data_validation",
        "action": "addDataValidation",
        "app": "excel",
        "summary": "设置Excel单元格的数据验证规则，如下拉列表、数值范围、日期范围等。",
        "params": {
            "range": {
                "type": "string",
                "description": "要设置验证的范围，如 A1:A100",
                "required": true
            },
            "type": {
                "type": "string",
                "description": "验证类型：list(下拉列表)、whole(整数)、decimal(小数)、date(日期)、textLength(文本长度)、custom(自定义)",
                "required": true
            },
            "formula": {
                "type": "string",
                "description": "验证公式。list类型用逗号分隔值如\"选项1,选项2,选项3\"；数值类型如\"1,100\"表示范围；custom类型为Excel公式",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "type",
            "formula"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_formula",
        "action": "setFormula",
        "app": "excel",
        "summary": "在指定单元格设置Excel公式。公式必须以=开头，支持所有Excel内置函数。",
        "params": {
            "range": {
                "type": "string",
                "description": "目标单元格地址，如 A1、B2:B10",
                "required": true
            },
            "formula": {
                "type": "string",
                "description": "Excel公式，必须以=开头，如 =SUM(A1:A10)、=VLOOKUP(A1,B:C,2,0)",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "range",
            "formula"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_hyperlink",
        "action": "setHyperlink",
        "app": "excel",
        "summary": "为Excel单元格设置超链接。",
        "params": {
            "cell": {
                "type": "string",
                "description": "单元格地址，如 A1",
                "required": true
            },
            "url": {
                "type": "string",
                "description": "超链接URL地址",
                "required": true
            },
            "text": {
                "type": "string",
                "description": "显示文本，不填则显示URL"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "cell",
            "url"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_list_object_totals",
        "action": "setListObjectTotals",
        "app": "excel",
        "summary": "开/关总计行，并可指定某一列的汇总方式。总计行写的是 SUBTOTAL 公式，会随筛选结果变化，这也正是它和普通求和公式的区别。使用场景：给金额列加合计。",
        "params": {
            "table": {
                "type": "string",
                "description": "表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）",
                "required": true
            },
            "show": {
                "type": "boolean",
                "description": "是否显示总计行；不填则保持现状（指定 column 时自动打开）"
            },
            "column": {
                "type": "string",
                "description": "要设置汇总方式的列（列名或序号）；不填只开关总计行"
            },
            "function": {
                "type": "string",
                "description": "汇总方式，默认 sum",
                "enum": [
                    "sum",
                    "average",
                    "count",
                    "countNums",
                    "max",
                    "min",
                    "stdDev",
                    "var",
                    "none"
                ]
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "table"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_named_range",
        "action": "createNamedRange",
        "app": "excel",
        "summary": "设置命名范围。为指定单元格区域创建或更新命名范围。",
        "params": {
            "name": {
                "type": "string",
                "description": "命名范围的名称",
                "required": true
            },
            "range": {
                "type": "string",
                "description": "单元格区域，如 \"A1:D10\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "name",
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_number_format",
        "action": "setNumberFormat",
        "app": "excel",
        "summary": "设置Excel单元格的数字格式。\n\n常用格式代码：\n- #,##0.00 - 千分位+2位小数\n- 0.00% - 百分比\n- yyyy-mm-dd - 日期\n- ¥#,##0.00 - 人民币\n- $#,##0.00 - 美元\n- 0.00E+00 - 科学计数法\n- @ - 文本格式",
        "params": {
            "range": {
                "type": "string",
                "description": "要设置格式的范围，如 A1:C10",
                "required": true
            },
            "format": {
                "type": "string",
                "description": "数字格式代码，如 #,##0.00、0.00%、yyyy-mm-dd",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "range",
            "format"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_outline_levels",
        "action": "setOutlineLevels",
        "app": "excel",
        "summary": "控制分级显示的展开层级与汇总位置：rowLevels/columnLevels 指定行/列显示到第几级（1 表示全部折叠），summaryRow/summaryColumn 指定汇总行在上还是下、汇总列在左还是右。使用场景：分组之后把明细收起来只留汇总。",
        "params": {
            "rowLevels": {
                "type": "number",
                "description": "行显示到第几级（1 表示只显示第 1 级，深层的折叠）"
            },
            "columnLevels": {
                "type": "number",
                "description": "列显示到第几级"
            },
            "summaryRow": {
                "type": "string",
                "description": "汇总行在明细的上方还是下方",
                "enum": [
                    "above",
                    "below"
                ]
            },
            "summaryColumn": {
                "type": "string",
                "description": "汇总列在明细的左侧还是右侧",
                "enum": [
                    "left",
                    "right"
                ]
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_print_area",
        "action": "setPrintArea",
        "app": "excel",
        "summary": "设置打印区域",
        "params": {
            "range": {
                "type": "string",
                "description": "打印区域，如 A1:D20",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_row_height",
        "action": "setRowHeight",
        "app": "excel",
        "summary": "设置Excel指定行的行高。",
        "params": {
            "row": {
                "type": "number",
                "description": "行号，从1开始",
                "required": true
            },
            "height": {
                "type": "number",
                "description": "行高数值（磅为单位），如 20、30",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "row",
            "height"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_sheet_appearance",
        "action": "setSheetAppearance",
        "app": "excel",
        "summary": "设置工作表的可见性与标签色：可见 / 隐藏 / 深度隐藏（veryHidden，用户界面上无法取消隐藏），以及标签颜色（十六进制如 #FF9900）。使用场景：把中间计算表藏起来、给关键工作表标个颜色。",
        "params": {
            "visible": {
                "type": "string",
                "description": "可见性，默认不变；veryHidden 在界面上无法恢复，脚本可恢复",
                "enum": [
                    "visible",
                    "hidden",
                    "veryHidden"
                ]
            },
            "tabColor": {
                "type": "string",
                "description": "标签颜色，如 #FF9900；传空字符串恢复默认"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_sheet_header_footer",
        "action": "setSheetHeaderFooter",
        "app": "excel",
        "summary": "设置打印页眉页脚。文本里可以用 Excel 的域代码：&P 页码、&N 总页数、&D 日期、&T 时间、&F 文件名、&A 工作表名。使用场景：页脚写「第 &P 页 / 共 &N 页」。",
        "params": {
            "leftHeader": {
                "type": "string",
                "description": "页眉左侧文本；传空字符串清除"
            },
            "centerHeader": {
                "type": "string",
                "description": "页眉中间文本"
            },
            "rightHeader": {
                "type": "string",
                "description": "页眉右侧文本"
            },
            "leftFooter": {
                "type": "string",
                "description": "页脚左侧文本"
            },
            "centerFooter": {
                "type": "string",
                "description": "页脚中间文本，如 第 &P 页 / 共 &N 页"
            },
            "rightFooter": {
                "type": "string",
                "description": "页脚右侧文本"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_sheet_page_setup",
        "action": "setSheetPageSetup",
        "app": "excel",
        "summary": "设置工作表的页面：方向、纸张、页边距、缩放、是否居中、是否打印网格线与行列标题。页边距单位是磅（1 厘米 ≈ 28.35 磅），与 Word 侧一致。缩放比例与按页适配互斥，同时给以后者为准。使用场景：把表调成横向 A4、一页宽、水平居中再打印。",
        "params": {
            "orientation": {
                "type": "string",
                "description": "纸张方向，默认纵向",
                "enum": [
                    "portrait",
                    "landscape"
                ]
            },
            "paperSize": {
                "type": "string",
                "description": "纸张大小，默认不变（当前多为 A4）",
                "enum": [
                    "A4",
                    "A3",
                    "A5",
                    "B5",
                    "letter",
                    "legal",
                    "tabloid"
                ]
            },
            "topMargin": {
                "type": "number",
                "description": "上边距（磅）"
            },
            "bottomMargin": {
                "type": "number",
                "description": "下边距（磅）"
            },
            "leftMargin": {
                "type": "number",
                "description": "左边距（磅）"
            },
            "rightMargin": {
                "type": "number",
                "description": "右边距（磅）"
            },
            "headerMargin": {
                "type": "number",
                "description": "页眉距顶边（磅）"
            },
            "footerMargin": {
                "type": "number",
                "description": "页脚距底边（磅）"
            },
            "zoom": {
                "type": "number",
                "description": "缩放百分比（如 90）；与 fitToPages* 互斥"
            },
            "fitToPagesWide": {
                "type": "number",
                "description": "按页适配：横向压到几页宽（1 表示一页宽）"
            },
            "fitToPagesTall": {
                "type": "number",
                "description": "按页适配：纵向压到几页高"
            },
            "centerHorizontally": {
                "type": "boolean",
                "description": "水平居中打印"
            },
            "centerVertically": {
                "type": "boolean",
                "description": "垂直居中打印"
            },
            "printGridlines": {
                "type": "boolean",
                "description": "打印网格线"
            },
            "printHeadings": {
                "type": "boolean",
                "description": "打印行号列标"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_sheet_print_titles",
        "action": "setSheetPrintTitles",
        "app": "excel",
        "summary": "设置打印时每页重复的行/列（打印标题）：如行 $1:$1 让表头每页都出现，列 $A:$A 让第一列每页都出现。使用场景：多页表格打印出来每一页都有表头。",
        "params": {
            "printTitleRows": {
                "type": "string",
                "description": "每页重复的行，如 $1:$1；传空字符串清除"
            },
            "printTitleColumns": {
                "type": "string",
                "description": "每页重复的列，如 $A:$A；传空字符串清除"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_wrap_text",
        "action": "wrapText",
        "app": "excel",
        "summary": "设置或取消单元格的自动换行。使用场景：\"让长文本在单元格里换行显示\"。",
        "params": {
            "range": {
                "type": "string",
                "description": "目标区域，如 A1:C10",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            },
            "wrap": {
                "type": "boolean",
                "description": "true 打开自动换行，false 关闭，默认 true"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_set_zoom",
        "action": "setZoom",
        "app": "excel",
        "summary": "设置当前工作表的缩放比例（10-400%）。",
        "params": {
            "percent": {
                "type": "number",
                "description": "缩放百分比，范围10-400",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "percent"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_show_columns",
        "action": "showColumns",
        "app": "excel",
        "summary": "显示Excel中已隐藏的列。",
        "params": {
            "startColumn": {
                "type": "string",
                "description": "起始列字母，如 A",
                "required": true
            },
            "endColumn": {
                "type": "string",
                "description": "结束列字母，如 D",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "startColumn",
            "endColumn"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_show_rows",
        "action": "showRows",
        "app": "excel",
        "summary": "显示Excel中已隐藏的行。",
        "params": {
            "startRow": {
                "type": "number",
                "description": "起始行号（从1开始）"
            },
            "endRow": {
                "type": "number",
                "description": "结束行号（从1开始）"
            },
            "row": {
                "type": "number",
                "description": "单行行号（与 startRow 等价）"
            },
            "rows": {
                "type": "array",
                "description": "行号数组"
            },
            "count": {
                "type": "number",
                "description": "从 row 起的连续行数"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_sort_range",
        "action": "sortRange",
        "app": "excel",
        "summary": "对Excel选定区域按指定列排序。",
        "params": {
            "range": {
                "type": "string",
                "description": "要排序的范围，如 A1:D100",
                "required": true
            },
            "column": {
                "type": "number",
                "description": "排序依据的列号（从1开始）",
                "required": true
            },
            "ascending": {
                "type": "boolean",
                "description": "是否升序，默认true"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "column"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_subtotal",
        "action": "subtotal",
        "app": "excel",
        "summary": "创建分类汇总。",
        "params": {
            "range": {
                "type": "string",
                "description": "数据范围，如 A1:D100",
                "required": true
            },
            "groupBy": {
                "type": "string",
                "description": "分组依据列在 range 内的序号（从 1 开始），如 range=A1:C4 时用 \"1\" 按第一列分组",
                "required": true
            },
            "function": {
                "type": "string",
                "description": "汇总函数",
                "enum": [
                    "sum",
                    "count",
                    "average",
                    "max",
                    "min"
                ],
                "required": true
            },
            "columns": {
                "type": "array",
                "description": "要汇总的列在 range 内的序号列表（从 1 开始），如 [3] 表示对第三列求和",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "items": {
                        "type": "string"
                    },
                    "description": "要汇总的列在 range 内的序号列表（从 1 开始），如 [3] 表示对第三列求和"
                },
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range",
            "groupBy",
            "function",
            "columns"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_switch_sheet",
        "action": "switchSheet",
        "app": "excel",
        "summary": "切换到指定的工作表，使其成为活动工作表。",
        "params": {
            "name": {
                "type": "string",
                "description": "要切换到的工作表名称",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_switch_workbook",
        "action": "switchWorkbook",
        "app": "excel",
        "summary": "切换到指定名称的Excel工作簿。",
        "params": {
            "name": {
                "type": "string",
                "description": "工作簿名称",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_text_to_columns",
        "action": "textToColumns",
        "app": "excel",
        "summary": "将文本按分隔符拆分到多列。",
        "params": {
            "range": {
                "type": "string",
                "description": "要拆分的范围，如 A1:A100",
                "required": true
            },
            "delimiter": {
                "type": "string",
                "description": "分隔符，默认逗号"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_transpose",
        "action": "transpose",
        "app": "excel",
        "summary": "转置数据（行列互换）。",
        "params": {
            "source": {
                "type": "string",
                "description": "源范围，如 A1:C3",
                "required": true
            },
            "destination": {
                "type": "string",
                "description": "目标位置，如 E1",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "source",
            "destination"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_unlist_list_object",
        "action": "unlistListObject",
        "app": "excel",
        "summary": "把表转回普通区域（数据与格式保留）：结构化引用、筛选按钮与表对象都会消失。使用场景：交付前清掉表对象，避免对方打开时出现意料之外的引用。",
        "params": {
            "table": {
                "type": "string",
                "description": "表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "table"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_unmerge_cells",
        "action": "unmergeCells",
        "app": "excel",
        "summary": "拆分Excel中已合并的单元格，恢复为独立的单元格。",
        "params": {
            "range": {
                "type": "string",
                "description": "要拆分的合并单元格范围，如 A1:C3",
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "range"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_unprotect_sheet",
        "action": "unprotectSheet",
        "app": "excel",
        "summary": "取消保护当前工作表。",
        "params": {
            "password": {
                "type": "string",
                "description": "保护密码（如果设置了密码保护则需要提供）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_update_chart",
        "action": "updateChart",
        "app": "excel",
        "summary": "更新Excel图表的属性，包括标题、颜色、图例、数据标签等。\n\n使用场景：\n- \"把图表标题改成销售报表\" -> 更新 title\n- \"隐藏图例\" -> 设置 show_legend: false\n- \"显示数据标签\" -> 设置 show_data_labels: true\n- \"改变图表类型为折线图\" -> 设置 chart_type: line\n\n注意：需要先通过 wps_excel_create_chart 创建图表，或者指定已存在图表的名称/索引",
        "params": {
            "chart_index": {
                "type": "number",
                "description": "图表索引（从1开始），与chart_name二选一"
            },
            "chart_name": {
                "type": "string",
                "description": "图表名称，与chart_index二选一"
            },
            "title": {
                "type": "string",
                "description": "新的图表标题"
            },
            "chart_type": {
                "type": "string",
                "description": "更改图表类型",
                "enum": [
                    "column_clustered",
                    "column_stacked",
                    "bar_clustered",
                    "line",
                    "line_markers",
                    "pie",
                    "doughnut",
                    "scatter",
                    "area",
                    "radar"
                ]
            },
            "show_legend": {
                "type": "boolean",
                "description": "是否显示图例"
            },
            "legend_position": {
                "type": "string",
                "description": "图例位置：bottom（下）、top（上）、left（左）、right（右）",
                "enum": [
                    "bottom",
                    "top",
                    "left",
                    "right"
                ]
            },
            "show_data_labels": {
                "type": "boolean",
                "description": "是否显示数据标签"
            },
            "data_range": {
                "type": "string",
                "description": "更改数据源范围"
            },
            "colors": {
                "type": "array",
                "description": "系列颜色数组，如 [\"#FF0000\", \"#00FF00\", \"#0000FF\"]",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "系列颜色数组，如 [\"#FF0000\", \"#00FF00\", \"#0000FF\"]",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "sheet": {
                "type": "string",
                "description": "图表所在的工作表名称"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "chart_index": "chartIndex",
            "chart_name": "chartName",
            "chart_type": "chartType",
            "show_legend": "showLegend",
            "legend_position": "legendPosition",
            "show_data_labels": "showDataLabels",
            "data_range": "dataRange"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_excel_update_list_object",
        "action": "updateListObject",
        "app": "excel",
        "summary": "改表本身的设置：改名、换表格样式、显示或隐藏表头行、显示或隐藏筛选按钮。使用场景：表名要能写进公式、筛选按钮碍事要关掉、换一个配色。",
        "params": {
            "table": {
                "type": "string",
                "description": "表名（如 表1、Sales）或该表在工作表上的序号（从 1 开始）",
                "required": true
            },
            "name": {
                "type": "string",
                "description": "新的表名（要能写进公式，不能与已有表名或命名范围重复）"
            },
            "tableStyle": {
                "type": "string",
                "description": "新的表格样式名，如 TableStyleMedium2"
            },
            "showHeaders": {
                "type": "boolean",
                "description": "是否显示表头行"
            },
            "showAutoFilter": {
                "type": "boolean",
                "description": "是否显示表头里的筛选按钮"
            },
            "sheet": {
                "type": "string",
                "description": "工作表名或序号；不填则用当前活动工作表"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "table"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_update_pivot_table",
        "action": "updatePivotTable",
        "app": "excel",
        "summary": "更新已有透视表的配置，包括添加/移除字段、修改聚合方式等。\n\n使用场景：\n- \"给透视表加个筛选器\" -> 使用addFilterFields\n- \"把销售额改成求平均\" -> 使用updateValueFields\n- \"去掉月份这个列字段\" -> 使用removeColumnFields\n- \"刷新透视表数据\" -> 使用refresh=true\n\n支持的操作：\n- 添加/移除行字段\n- 添加/移除列字段\n- 添加/移除/更新值字段\n- 添加/移除筛选字段\n- 刷新数据",
        "params": {
            "pivotTableName": {
                "type": "string",
                "description": "透视表名称。如果不确定名称，可以先用其他工具查看工作表内容"
            },
            "pivotTableCell": {
                "type": "string",
                "description": "透视表所在的任意单元格地址（可选）。如果提供了pivotTableName则忽略此参数"
            },
            "addRowFields": {
                "type": "array",
                "description": "要添加的行字段列名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要添加的行字段列名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "removeRowFields": {
                "type": "array",
                "description": "要移除的行字段列名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要移除的行字段列名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "addColumnFields": {
                "type": "array",
                "description": "要添加的列字段列名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要添加的列字段列名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "removeColumnFields": {
                "type": "array",
                "description": "要移除的列字段列名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要移除的列字段列名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "addValueFields": {
                "type": "array",
                "description": "要添加的值字段配置列表",
                "items": {
                    "type": "object"
                },
                "schema": {
                    "type": "array",
                    "description": "要添加的值字段配置列表",
                    "items": {
                        "type": "object",
                        "properties": {
                            "field": {
                                "type": "string",
                                "description": "字段名"
                            },
                            "aggregation": {
                                "type": "string",
                                "description": "聚合方式",
                                "enum": [
                                    "SUM",
                                    "COUNT",
                                    "AVERAGE",
                                    "MAX",
                                    "MIN"
                                ]
                            }
                        },
                        "required": [
                            "field",
                            "aggregation"
                        ]
                    }
                }
            },
            "removeValueFields": {
                "type": "array",
                "description": "要移除的值字段名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要移除的值字段名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "updateValueFields": {
                "type": "array",
                "description": "要更新的值字段配置（修改聚合方式）",
                "items": {
                    "type": "object"
                },
                "schema": {
                    "type": "array",
                    "description": "要更新的值字段配置（修改聚合方式）",
                    "items": {
                        "type": "object",
                        "properties": {
                            "field": {
                                "type": "string",
                                "description": "字段名"
                            },
                            "aggregation": {
                                "type": "string",
                                "description": "新的聚合方式",
                                "enum": [
                                    "SUM",
                                    "COUNT",
                                    "AVERAGE",
                                    "MAX",
                                    "MIN"
                                ]
                            }
                        },
                        "required": [
                            "field",
                            "aggregation"
                        ]
                    }
                }
            },
            "addFilterFields": {
                "type": "array",
                "description": "要添加的筛选字段列名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要添加的筛选字段列名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "removeFilterFields": {
                "type": "array",
                "description": "要移除的筛选字段列名列表",
                "items": {
                    "type": "string"
                },
                "schema": {
                    "type": "array",
                    "description": "要移除的筛选字段列名列表",
                    "items": {
                        "type": "string"
                    }
                }
            },
            "refresh": {
                "type": "boolean",
                "description": "是否刷新透视表数据（当源数据变化时使用）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_excel_write_range",
        "action": "setRangeData",
        "app": "excel",
        "summary": "向Excel指定范围写入数据。数据格式为二维数组，从指定单元格开始向右下方填充。",
        "params": {
            "range": {
                "type": "string",
                "description": "起始单元格地址，如 A1、B2",
                "required": true
            },
            "data": {
                "type": "array2d",
                "description": "二维数组数据，如 [[\"姓名\",\"年龄\"],[\"张三\",25],[\"李四\",30]]",
                "schema": {
                    "type": "array",
                    "description": "二维数组数据，如 [[\"姓名\",\"年龄\"],[\"张三\",25],[\"李四\",30]]",
                    "items": {
                        "type": "array",
                        "items": {
                            "type": "string"
                        }
                    }
                },
                "required": true
            },
            "sheet": {
                "type": "string",
                "description": "工作表名称，不填则使用当前活动工作表"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "range",
            "data"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_animation",
        "action": "addAnimation",
        "app": "ppt",
        "summary": "为幻灯片中的形状添加动画效果。\n\n使用场景：\n- \"给第1页第2个形状加个淡入动画\"\n- \"为标题添加飞入效果\"\n- \"添加动画让元素依次出现\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "effect": {
                "type": "string",
                "description": "动画效果名称，如 \"fadeIn\"、\"flyIn\"、\"wipe\" 等",
                "required": true
            },
            "trigger": {
                "type": "string",
                "description": "触发方式",
                "enum": [
                    "onClick",
                    "withPrevious",
                    "afterPrevious"
                ]
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "effect"
        ],
        "engine": "bridge",
        "aliases": {
            "shapeIndex": "shapeName"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_animation_preset",
        "action": "addAnimationPreset",
        "app": "ppt",
        "summary": "为形状添加预设入场动画效果。\n\n支持的预设：\n- fadeIn: 淡入\n- flyIn: 飞入\n- wipe: 擦除\n- zoom: 缩放\n- bounce: 弹跳\n- spin: 旋转\n\n使用场景：\n- \"给标题加个淡入效果\"\n- \"让这个形状飞入\"\n- \"添加弹跳入场动画\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "preset": {
                "type": "string",
                "description": "预设动画类型",
                "enum": [
                    "fadeIn",
                    "flyIn",
                    "wipe",
                    "zoom",
                    "bounce",
                    "spin"
                ],
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "preset"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_emphasis_animation",
        "action": "addEmphasisAnimation",
        "app": "ppt",
        "summary": "为形状添加强调动画效果，用于在演示时突出显示元素。\n\n支持的效果：\n- pulse: 脉冲\n- spin: 陀螺旋\n- grow: 放大/缩小\n- teeter: 跷跷板\n- colorPulse: 颜色脉冲\n\n使用场景：\n- \"让这个元素闪烁突出\"\n- \"添加脉冲强调效果\"\n- \"让图片旋转强调\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "effect": {
                "type": "string",
                "description": "强调动画效果",
                "enum": [
                    "pulse",
                    "spin",
                    "grow",
                    "teeter",
                    "colorPulse"
                ],
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "effect"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_master_element",
        "action": "addMasterElement",
        "app": "ppt",
        "summary": "向母版中添加新元素。\n\n支持添加文本框、形状、图片、Logo等母版级元素。\n\n使用场景：\n- \"在母版上添加公司Logo\"\n- \"给母版加个页脚\"\n- \"在母版添加水印\"",
        "params": {
            "element": {
                "type": "object",
                "description": "元素配置对象，如 {type:\"textbox\",text:\"...\",left:0,top:0,width:100,height:50} 或 {type:\"image\",path:\"...\",left:0,top:0}",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "element"
        ],
        "engine": "bridge",
        "containers": [
            "element"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_ppt_hyperlink",
        "action": "addPptHyperlink",
        "app": "ppt",
        "summary": "为幻灯片中的形状添加超链接。\n\n支持网页链接、邮件链接、幻灯片内部跳转等。\n\n使用场景：\n- \"给按钮添加链接\"\n- \"设置点击跳转到网页\"\n- \"添加邮件链接\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "url": {
                "type": "string",
                "description": "超链接地址，如 \"https://example.com\" 或 \"mailto:test@example.com\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "url"
        ],
        "engine": "bridge",
        "aliases": {
            "url": "address"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_shape",
        "action": "addShape",
        "app": "ppt",
        "summary": "在幻灯片中添加形状。\n\n支持的形状类型：\n- rectangle: 矩形\n- oval: 椭圆\n- triangle: 三角形\n- diamond: 菱形\n- pentagon: 五边形\n- hexagon: 六边形\n- arrow: 箭头\n- star: 星形\n- heart: 心形\n- cloud: 云形\n\n使用场景：\n- \"在第1页添加一个矩形\"\n- \"插入一个蓝色的圆形\"\n- \"添加一个带文字的箭头\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始），默认1"
            },
            "type": {
                "type": "string",
                "description": "形状类型",
                "enum": [
                    "rectangle",
                    "oval",
                    "triangle",
                    "diamond",
                    "pentagon",
                    "hexagon",
                    "arrow",
                    "star",
                    "heart",
                    "cloud"
                ]
            },
            "left": {
                "type": "number",
                "description": "左边距（像素），默认100"
            },
            "top": {
                "type": "number",
                "description": "上边距（像素），默认100"
            },
            "width": {
                "type": "number",
                "description": "宽度（像素），默认100"
            },
            "height": {
                "type": "number",
                "description": "高度（像素），默认100"
            },
            "text": {
                "type": "string",
                "description": "形状内的文本"
            },
            "fillColor": {
                "type": "string",
                "description": "填充颜色，十六进制如 #FF0000"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_slide",
        "action": "addSlide",
        "app": "ppt",
        "summary": "添加新幻灯片到演示文稿。\n\n支持的布局类型：\n- title: 标题页\n- title_content: 标题+内容（最常用）\n- blank: 空白页\n- two_column: 两栏内容\n- comparison: 对比布局\n\n使用场景：\n- \"新建一页PPT\"\n- \"添加一个标题页\"\n- \"在第3页后面插入一页\"",
        "params": {
            "layout": {
                "type": "string",
                "description": "幻灯片布局类型",
                "enum": [
                    "title",
                    "title_content",
                    "blank",
                    "two_column",
                    "comparison"
                ]
            },
            "position": {
                "type": "number",
                "description": "插入位置（页码），不填则在末尾添加"
            },
            "title": {
                "type": "string",
                "description": "幻灯片标题"
            },
            "content": {
                "type": "string",
                "description": "幻灯片内容（针对有内容区域的布局）"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_add_textbox",
        "action": "addTextBox",
        "app": "ppt",
        "summary": "在幻灯片中添加文本框。\n\n使用场景：\n- \"在第1页添加一个文本框\"\n- \"插入一个写着标题的文本框\"\n- \"添加文本框并设置字号\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始），默认1"
            },
            "left": {
                "type": "number",
                "description": "左边距（像素），默认100"
            },
            "top": {
                "type": "number",
                "description": "上边距（像素），默认100"
            },
            "width": {
                "type": "number",
                "description": "宽度（像素），默认200"
            },
            "height": {
                "type": "number",
                "description": "高度（像素），默认50"
            },
            "text": {
                "type": "string",
                "description": "文本框内容"
            },
            "fontSize": {
                "type": "number",
                "description": "字号大小"
            },
            "fontName": {
                "type": "string",
                "description": "字体名称，如 \"微软雅黑\""
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_align_shapes",
        "action": "alignShapes",
        "app": "ppt",
        "summary": "对齐幻灯片中的多个形状。\n\n支持的对齐方式：\n- left: 左对齐\n- center: 水平居中\n- right: 右对齐\n- top: 顶部对齐\n- middle: 垂直居中\n- bottom: 底部对齐\n\n使用场景：\n- \"把这几个形状左对齐\"\n- \"让第1、3、5个形状垂直居中\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndices": {
                "type": "array",
                "description": "要对齐的形状索引数组（从1开始）",
                "items": {
                    "type": "number"
                },
                "schema": {
                    "type": "array",
                    "items": {
                        "type": "number"
                    },
                    "description": "要对齐的形状索引数组（从1开始）"
                }
            },
            "alignment": {
                "type": "string",
                "description": "对齐方式",
                "enum": [
                    "left",
                    "center",
                    "right",
                    "top",
                    "middle",
                    "bottom"
                ],
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "alignment"
        ],
        "engine": "bridge",
        "aliases": {
            "shapeIndices": "names"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_apply_transition_to_all",
        "action": "applyTransitionToAll",
        "app": "ppt",
        "summary": "为所有幻灯片应用统一的切换效果。\n\n使用场景：\n- \"给所有页面加上淡出切换\"\n- \"统一设置切换效果\"\n- \"所有幻灯片用推入切换\"",
        "params": {
            "effect": {
                "type": "string",
                "description": "切换效果名称，如 \"fade\"、\"push\"、\"wipe\" 等",
                "required": true
            },
            "duration": {
                "type": "number",
                "description": "切换持续时间（秒），默认1秒"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "effect"
        ],
        "engine": "bridge",
        "aliases": {
            "effect": "transition"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_beautify",
        "action": "beautifySlide",
        "app": "ppt",
        "summary": "一键美化幻灯片，优化排版、配色、字体和间距。\n\n支持的配色方案：\n- business: 商务风（深蓝+灰色）\n- tech: 科技风（蓝色+绿色）\n- creative: 创意风（珊瑚红+金色）\n- minimal: 简约风（黑白灰）\n\n美化包含的操作：\n- 统一字体\n- 应用配色方案\n- 对齐元素\n- 优化间距\n\n使用场景：\n- \"美化这页PPT\"\n- \"用商务风格优化一下\"\n- \"把PPT弄好看点\"",
        "params": {
            "slide_index": {
                "type": "number",
                "description": "要美化的幻灯片页码，不填则美化当前页"
            },
            "color_scheme": {
                "type": "string",
                "description": "配色方案",
                "enum": [
                    "business",
                    "tech",
                    "creative",
                    "minimal"
                ],
                "kind": "local"
            },
            "font": {
                "type": "string",
                "description": "统一使用的字体，如 \"微软雅黑\"、\"思源黑体\"",
                "kind": "local"
            },
            "beautify_all": {
                "type": "boolean",
                "description": "是否美化所有幻灯片，默认false只美化指定页",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "slide_index": "slideIndex"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_close_presentation",
        "action": "closePresentation",
        "app": "ppt",
        "summary": "关闭演示文稿。可指定文稿名称，不指定则关闭当前活动文稿。\n\n使用场景：\n- \"关闭这个PPT\"\n- \"关闭演示文稿不保存\"\n- \"关闭指定的PPT\"",
        "params": {
            "name": {
                "type": "string",
                "description": "要关闭的演示文稿名称，不填则关闭当前活动文稿"
            },
            "save": {
                "type": "boolean",
                "description": "关闭前是否保存，默认true"
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_copy_slide",
        "action": "duplicateSlide",
        "app": "ppt",
        "summary": "复制幻灯片到指定位置。\n\n使用场景：\n- \"复制第2页幻灯片\"\n- \"把这页复制到第5页后面\"\n- \"克隆当前幻灯片\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "要复制的幻灯片索引（从1开始）",
                "required": true
            },
            "targetIndex": {
                "type": "number",
                "description": "目标位置索引（从1开始），不填则在原位置后插入"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_create_presentation",
        "action": "createPresentation",
        "app": "ppt",
        "summary": "新建空白演示文稿。\n\n使用场景：\n- \"新建一个PPT\"\n- \"创建一个演示文稿\"\n- \"打开一个新的PPT\"",
        "params": {},
        "effect": "lifecycle",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_delete_ppt_image",
        "action": "deletePptImage",
        "app": "ppt",
        "summary": "删除幻灯片中指定的图片。\n\n使用场景：\n- \"删除第2页的第1张图片\"\n- \"移除这张图片\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "imageIndex": {
                "type": "number",
                "description": "图片索引（从1开始）",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex",
            "imageIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_delete_shape",
        "action": "deleteShape",
        "app": "ppt",
        "summary": "删除幻灯片中指定的形状。\n\n使用场景：\n- \"删除第2页的第3个形状\"\n- \"移除这个形状\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_delete_slide",
        "action": "deleteSlide",
        "app": "ppt",
        "summary": "删除指定的幻灯片。\n\n使用场景：\n- \"删除第3页幻灯片\"\n- \"把最后一页删掉\"\n- \"移除多余的页面\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "要删除的幻灯片索引（从1开始）",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_delete_textbox",
        "action": "deleteTextBox",
        "app": "ppt",
        "summary": "删除幻灯片上指定的文本框。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "textboxIndex": {
                "type": "number",
                "description": "文本框索引",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex",
            "textboxIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_distribute_shapes",
        "action": "distributeShapes",
        "app": "ppt",
        "summary": "等距分布幻灯片中的多个形状。\n\n支持的分布方向：\n- horizontal: 水平等距分布\n- vertical: 垂直等距分布\n\n使用场景：\n- \"让这些形状水平等距排列\"\n- \"垂直均匀分布这些元素\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndices": {
                "type": "array",
                "description": "要分布的形状索引数组（从1开始）",
                "items": {
                    "type": "number"
                },
                "schema": {
                    "type": "array",
                    "items": {
                        "type": "number"
                    },
                    "description": "要分布的形状索引数组（从1开始）"
                },
                "required": true
            },
            "direction": {
                "type": "string",
                "description": "分布方向",
                "enum": [
                    "horizontal",
                    "vertical"
                ],
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndices",
            "direction"
        ],
        "engine": "bridge",
        "aliases": {
            "shapeIndices": "names"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_duplicate_shape",
        "action": "duplicateShape",
        "app": "ppt",
        "summary": "复制幻灯片中的指定形状。\n\n使用场景：\n- \"复制这个形状\"\n- \"把第1页的第2个形状复制一份\"\n- \"克隆这个元素\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "要复制的形状索引（从1开始）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_export_slide_as_image",
        "action": "exportSlideAsImage",
        "app": "ppt",
        "summary": "将指定幻灯片导出为位图图片（PNG/JPG/JPEG/GIF/BMP）。\n\n调用底层 WPS PowerPoint 原生接口 Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)，\n实现 1:1 像素级还原，避免通过 PDF 中转再转图片造成的版式/字体/形状失真问题。\n\n支持的 format（FilterName）取值：\n- PNG（默认，推荐用于截图与无损展示）\n- JPG / JPEG（自动按 JPG 滤镜处理，体积更小）\n- GIF（限 256 色，适合简单图形）\n- BMP（无压缩位图，文件最大）\n\n使用场景：\n- \"把第3页 PPT 导出成 PNG 给我\"\n- \"导出整个演示文稿每一页为 1920x1080 的 JPG\"\n- \"把封面页保存为高清图片用于网页\"\n\n注意：\n- outputPath 必须是绝对路径\n- macOS 上建议输出到 ~/Downloads 或用户可写目录，避免沙箱权限拒绝\n- 不指定 width/height 时使用 1280x720（16:9 默认尺寸）",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片序号（从1开始）",
                "required": true
            },
            "outputPath": {
                "type": "string",
                "description": "输出图片文件的绝对路径（如 /Users/xxx/Downloads/slide1.png）",
                "required": true
            },
            "format": {
                "type": "string",
                "description": "图片格式，默认 PNG",
                "enum": [
                    "PNG",
                    "JPG",
                    "JPEG",
                    "GIF",
                    "BMP"
                ]
            },
            "width": {
                "type": "number",
                "description": "输出图片宽度（像素），默认 1280"
            },
            "height": {
                "type": "number",
                "description": "输出图片高度（像素），默认 720"
            }
        },
        "effect": "export",
        "advertised": true,
        "required": [
            "slideIndex",
            "outputPath"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_find_ppt_text",
        "action": "findPptText",
        "app": "ppt",
        "summary": "在演示文稿中搜索指定文本。\n\n返回包含目标文本的幻灯片页码和形状信息。\n\n使用场景：\n- \"搜索PPT中的某段文字\"\n- \"查找包含关键词的幻灯片\"\n- \"找到所有提到xxx的位置\"",
        "params": {
            "text": {
                "type": "string",
                "description": "要搜索的文本内容",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_animations",
        "action": "getAnimations",
        "app": "ppt",
        "summary": "获取幻灯片上所有动画效果的列表。\n\n使用场景：\n- \"查看第1页有哪些动画\"\n- \"列出这页的所有动画\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_open_presentations",
        "action": "getOpenPresentations",
        "app": "ppt",
        "summary": "获取当前所有已打开的演示文稿列表。\n\n使用场景：\n- \"有哪些PPT打开着\"\n- \"列出所有演示文稿\"\n- \"查看打开的PPT\"",
        "params": {},
        "effect": "lifecycle",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_shapes",
        "action": "getShapes",
        "app": "ppt",
        "summary": "获取幻灯片中所有形状的列表信息。\n\n返回每个形状的类型、位置、大小等属性。\n\n使用场景：\n- \"这页PPT有哪些形状\"\n- \"列出第3页的所有元素\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_slide_count",
        "action": "getSlideCount",
        "app": "ppt",
        "summary": "获取演示文稿中的幻灯片总数。\n\n使用场景：\n- \"一共有多少页幻灯片\"\n- \"PPT有几页\"\n- \"查看幻灯片数量\"",
        "params": {},
        "effect": "read",
        "advertised": true,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_slide_info",
        "action": "getSlideInfo",
        "app": "ppt",
        "summary": "获取指定幻灯片的详细信息，包括布局、元素列表等。\n\n使用场景：\n- \"查看第3页的信息\"\n- \"这页有什么内容\"\n- \"获取幻灯片详情\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_slide_master",
        "action": "getSlideMaster",
        "app": "ppt",
        "summary": "获取当前演示文稿的母版信息。\n\n返回母版的布局、背景、元素等详细信息。\n\n使用场景：\n- \"查看母版信息\"\n- \"获取母版布局\"\n- \"看看母版有什么元素\"",
        "params": {},
        "effect": "read",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_slide_notes",
        "action": "getSlideNotes",
        "app": "ppt",
        "summary": "获取指定幻灯片的备注内容。\n\n使用场景：\n- \"查看第3页的备注\"\n- \"读取演讲备注\"\n- \"这页有什么备注\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_slide_title",
        "action": "getSlideTitle",
        "app": "ppt",
        "summary": "获取指定幻灯片的标题。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_table_cell",
        "action": "getPptTableCell",
        "app": "ppt",
        "summary": "获取PPT表格中指定单元格的文本内容。\n\n使用场景：\n- \"读取表格第1行第1列的内容\"\n- \"获取单元格文本\"\n- \"查看表格数据\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "tableIndex": {
                "type": "number",
                "description": "表格索引（从1开始）",
                "required": true
            },
            "row": {
                "type": "number",
                "description": "行号（从1开始）",
                "required": true
            },
            "col": {
                "type": "number",
                "description": "列号（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "slideIndex",
            "tableIndex",
            "row",
            "col"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_get_textboxes",
        "action": "getTextBoxes",
        "app": "ppt",
        "summary": "获取幻灯片上所有文本框的列表。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_group_shapes",
        "action": "groupShapes",
        "app": "ppt",
        "summary": "将幻灯片中的多个形状组合为一个组。\n\n使用场景：\n- \"把这几个形状组合在一起\"\n- \"将第1、2、3个形状编组\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndices": {
                "type": "array",
                "description": "要组合的形状索引数组（从1开始，至少2个）",
                "items": {
                    "type": "number"
                },
                "schema": {
                    "type": "array",
                    "items": {
                        "type": "number"
                    },
                    "description": "要组合的形状索引数组（从1开始，至少2个）"
                },
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndices"
        ],
        "engine": "bridge",
        "aliases": {
            "shapeIndices": "names"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_insert_ppt_chart",
        "action": "insertPptChart",
        "app": "ppt",
        "summary": "在幻灯片中插入数据图表。\n\n支持的图表类型：\n- bar: 柱状图\n- line: 折线图\n- pie: 饼图\n- area: 面积图\n- scatter: 散点图\n- doughnut: 环形图\n- radar: 雷达图\n\n使用场景：\n- \"在第2页插入一个柱状图\"\n- \"添加销售数据的饼图\"\n- \"插入折线图展示趋势\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "chartType": {
                "type": "string",
                "description": "图表类型",
                "required": true
            },
            "title": {
                "type": "string",
                "description": "图表标题"
            },
            "left": {
                "type": "number",
                "description": "图表左边距（磅），默认自动居中"
            },
            "top": {
                "type": "number",
                "description": "图表上边距（磅），默认自动居中"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "chartType"
        ],
        "engine": "bridge",
        "aliases": {
            "chartType": "type"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_insert_ppt_image",
        "action": "insertPptImage",
        "app": "ppt",
        "summary": "插入图片到幻灯片中。\n\n支持常见图片格式（PNG、JPG、BMP、GIF等）。\n可以指定图片的位置和大小，不指定则使用默认值。\n\n使用场景：\n- \"在第2页插入一张图片\"\n- \"把这张图片放到PPT里\"\n- \"在幻灯片右下角插入logo\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "filePath": {
                "type": "string",
                "description": "图片文件路径",
                "required": true
            },
            "left": {
                "type": "number",
                "description": "左边距（磅），可选"
            },
            "top": {
                "type": "number",
                "description": "上边距（磅），可选"
            },
            "width": {
                "type": "number",
                "description": "宽度（磅），可选，不指定则按原始比例"
            },
            "height": {
                "type": "number",
                "description": "高度（磅），可选，不指定则按原始比例"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "slideIndex",
            "filePath"
        ],
        "engine": "bridge",
        "aliases": {
            "filePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_insert_slides_from_file",
        "action": "insertSlidesFromFile",
        "app": "ppt",
        "summary": "从另一个 PPT 文件把整页幻灯片插入到【当前活动演示文稿】，并保留来源幻灯片的原始格式（字体/配色/版式/图片）。用于把多个 PPT 整合成一个。\n\n使用场景：\n- \"把可行性报告.pptx 的第3到5页插到当前PPT第10页后面\"\n- \"整合多个PPT：把另一个演示文稿的所有幻灯片合并进来\"\n- \"从某个PPT复制整页过来，保持原样式\"\n\n说明：\n- 先用 wps_ppt_switch_presentation 切换到【目标/接收页】演示文稿，再调用本工具\n- afterIndex 表示插入到第几页之后（0=插到最前，不填=追加到末尾）\n- slideStart/slideEnd 指定只导入来源文件的某段页码范围，不填则导入全部\n- 底层调用 WPS COM Slides.InsertFromFile，原样保留来源格式，避免AI重排导致的版式失真",
        "params": {
            "filePath": {
                "type": "string",
                "description": "来源 PPT 文件的完整路径",
                "required": true
            },
            "afterIndex": {
                "type": "number",
                "description": "插入到当前演示文稿第几页之后（0=最前，不填=末尾追加）"
            },
            "slideStart": {
                "type": "number",
                "description": "来源文件起始页码（从1开始，可选）"
            },
            "slideEnd": {
                "type": "number",
                "description": "来源文件结束页码（可选，与slideStart配合）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "filePath"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_insert_table",
        "action": "insertPptTable",
        "app": "ppt",
        "summary": "在幻灯片中插入表格。\n\n支持指定：\n- 表格行列数\n- 表格位置（左上角坐标）\n\n使用场景：\n- \"在第2页插入一个3行4列的表格\"\n- \"添加一个表格到当前页\"\n- \"插入数据表格\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "rows": {
                "type": "number",
                "description": "表格行数",
                "required": true
            },
            "cols": {
                "type": "number",
                "description": "表格列数",
                "required": true
            },
            "left": {
                "type": "number",
                "description": "表格左上角X坐标（磅），可选"
            },
            "top": {
                "type": "number",
                "description": "表格左上角Y坐标（磅），可选"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "slideIndex",
            "rows",
            "cols"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_move_slide",
        "action": "moveSlide",
        "app": "ppt",
        "summary": "移动幻灯片到指定位置。\n\n使用场景：\n- \"把第5页移到第2页\"\n- \"把最后一页移到开头\"\n- \"调整幻灯片顺序\"",
        "params": {
            "fromIndex": {
                "type": "number",
                "description": "原位置索引（从1开始）",
                "required": true
            },
            "toIndex": {
                "type": "number",
                "description": "目标位置索引（从1开始）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "fromIndex",
            "toIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_open_presentation",
        "action": "openPresentation",
        "app": "ppt",
        "summary": "打开指定路径的演示文稿文件。\n\n支持的文件格式：\n- .pptx: PowerPoint 演示文稿\n- .ppt: 旧版 PowerPoint 格式\n- .dps: WPS 演示格式\n\n使用场景：\n- \"打开桌面上的演示文稿\"\n- \"打开这个PPT文件\"",
        "params": {
            "filePath": {
                "type": "string",
                "description": "演示文稿文件的完整路径",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "filePath"
        ],
        "engine": "bridge",
        "aliases": {
            "filePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_remove_animation",
        "action": "removeAnimation",
        "app": "ppt",
        "summary": "移除幻灯片中指定的动画效果。\n\n使用场景：\n- \"删除第1页的第2个动画\"\n- \"移除这个动画效果\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "animationIndex": {
                "type": "number",
                "description": "动画索引（从1开始）",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex",
            "animationIndex"
        ],
        "engine": "bridge",
        "aliases": {
            "animationIndex": "index"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_remove_ppt_hyperlink",
        "action": "removePptHyperlink",
        "app": "ppt",
        "summary": "移除幻灯片中形状的超链接。\n\n使用场景：\n- \"删除按钮的链接\"\n- \"移除超链接\"\n- \"取消链接\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_remove_slide_transition",
        "action": "removeSlideTransition",
        "app": "ppt",
        "summary": "移除幻灯片的切换效果。\n\n使用场景：\n- \"取消第1页的切换效果\"\n- \"移除页面切换动画\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            }
        },
        "effect": "delete",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_replace_ppt_image",
        "action": "replacePptImage",
        "app": "ppt",
        "summary": "原位替换幻灯片中的某张图片：保留原图的位置、尺寸、旋转角度，删除旧图后在同一矩形内插入新图。用于在保持版式不变的前提下把模板里的旧图换成自己的图。\n\n使用场景：\n- \"把第5页的第2个形状（图片）换成 D:/figs/图15.png，位置大小不变\"\n- \"替换封面主视觉图，但保持原来的排版\"\n- \"把这页的配图换掉，别动版式\"\n\n说明：\n- 先用 wps_ppt_get_shapes 查到目标图片的形状索引(shapeIndex)\n- 新图会被拉伸/适配到旧图原有的位置与尺寸，从而不破坏版式",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "要替换的图片形状索引（从1开始，通过 get_shapes 获取）"
            },
            "name": {
                "type": "string",
                "description": "要替换的图片形状名称（与 shapeIndex 二选一）"
            },
            "filePath": {
                "type": "string",
                "description": "新图片文件的完整路径",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "filePath"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_replace_ppt_text",
        "action": "replacePptText",
        "app": "ppt",
        "summary": "在演示文稿中查找并替换文本。\n\n批量替换所有匹配的文本内容。\n\n使用场景：\n- \"把PPT中所有的A替换成B\"\n- \"批量替换公司名称\"\n- \"修改所有页面的标题\"",
        "params": {
            "find": {
                "type": "string",
                "description": "要查找的文本",
                "required": true
            },
            "replace": {
                "type": "string",
                "description": "替换为的文本",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "find",
            "replace"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_3d_depth",
        "action": "set3DDepth",
        "app": "ppt",
        "summary": "设置幻灯片中形状的3D挤出深度。\n\n通过调整深度值使形状产生立体挤出效果。\n\n使用场景：\n- \"给形状添加3D深度\"\n- \"设置立体厚度\"\n- \"增加形状深度效果\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "depth": {
                "type": "number",
                "description": "挤出深度值（磅），如 20、50、100",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "depth"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_3d_material",
        "action": "set3DMaterial",
        "app": "ppt",
        "summary": "设置幻灯片中形状的3D材质效果。\n\n支持的材质类型：\n- matte: 哑光\n- plastic: 塑料\n- metal: 金属\n- wireframe: 线框\n- soft_edge: 柔化边缘\n- flat: 平面\n- dark_edge: 暗边\n\n使用场景：\n- \"给形状设置金属材质\"\n- \"改成塑料质感\"\n- \"用哑光效果\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "material": {
                "type": "string",
                "description": "材质类型：matte(哑光)、plastic(塑料)、metal(金属)、wireframe(线框)、soft_edge(柔化)、flat(平面)、dark_edge(暗边)",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "material"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_3d_rotation",
        "action": "set3DRotation",
        "app": "ppt",
        "summary": "设置幻灯片中形状的3D旋转效果。\n\n通过调整X/Y/Z轴旋转角度实现3D透视效果。\n\n使用场景：\n- \"给形状添加3D旋转效果\"\n- \"设置3D透视角度\"\n- \"让形状有立体感\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "rotation": {
                "type": "object",
                "description": "旋转参数对象，如 {rotX:30,rotY:45,rotZ:0,perspective:50}",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "rotation"
        ],
        "engine": "bridge",
        "containers": [
            "rotation"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_active_target",
        "action": "getOpenPresentations",
        "app": "ppt",
        "summary": "锁定后续所有 PPT 操作的【目标演示文稿】（按文件名）。锁定后，本服务器会给所有演示类调用自动注入该文稿名，底层精确定位，**彻底避免同时打开多个 PPT 时\"活动文稿漂移\"导致改错文件**。\n\n使用场景：\n- 批量改某个 PPT 前先锁定：\"把目标锁定为 关节模组申报PPT-生成版.pptx\"\n- 传 clear=true 或留空 name 取消锁定，恢复使用当前活动文稿\n\n强烈建议：在对一个特定 PPT 做成批增改前，先调用本工具锁定它。",
        "params": {
            "name": {
                "type": "string",
                "description": "目标演示文稿文件名（如 关节模组申报PPT-生成版.pptx）；留空或 clear=true 则取消锁定",
                "kind": "local"
            },
            "clear": {
                "type": "boolean",
                "description": "为 true 时清除锁定，恢复使用当前活动文稿",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_animation_order",
        "action": "setAnimationOrder",
        "app": "ppt",
        "summary": "调整动画在时间线上的播放顺序。\n\n使用场景：\n- \"把第3个动画移到第1个播放\"\n- \"调整动画顺序\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "animationIndex": {
                "type": "number",
                "description": "当前动画索引（从1开始）",
                "required": true
            },
            "newOrder": {
                "type": "number",
                "description": "新的播放顺序位置（从1开始）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "animationIndex",
            "newOrder"
        ],
        "engine": "bridge",
        "aliases": {
            "animationIndex": "from",
            "newOrder": "to"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_background_color",
        "action": "setBackgroundColor",
        "app": "ppt",
        "summary": "设置幻灯片背景为指定颜色。\n\n使用场景：\n- \"把第1页背景改成蓝色\"\n- \"设置背景颜色为 #FF5733\"\n- \"背景换成白色\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "color": {
                "type": "string",
                "description": "十六进制颜色值，如 \"#FF0000\"、\"#FFFFFF\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "color"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_background_gradient",
        "action": "setBackgroundGradient",
        "app": "ppt",
        "summary": "为幻灯片设置渐变色背景。\n\n渐变配置对象（gradient）属性：\n- type: 渐变类型（linear/radial），默认linear\n- angle: 渐变角度（0-360度），仅linear有效，默认180\n- colors: 渐变颜色数组（如 [\"#1a1a2e\", \"#16213e\", \"#0f3460\"]）\n- stops: 颜色停靠点数组（如 [0, 0.5, 1]），与colors对应\n\n使用场景：\n- \"设置蓝色渐变背景\"\n- \"给第2页加个从深到浅的渐变\"\n- \"设置PPT背景为渐变色\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "gradient": {
                "type": "object",
                "description": "渐变配置对象，包含type（linear/radial）、angle（角度）、colors（颜色数组）、stops（停靠点数组）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "gradient"
        ],
        "engine": "bridge",
        "containers": [
            "gradient"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_background_image",
        "action": "setBackgroundImage",
        "app": "ppt",
        "summary": "设置幻灯片背景为指定图片。\n\n使用场景：\n- \"用这张图片做背景\"\n- \"设置第2页的背景图片\"\n- \"把图片设为幻灯片背景\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "imagePath": {
                "type": "string",
                "description": "图片文件的完整路径",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "imagePath"
        ],
        "engine": "bridge",
        "aliases": {
            "imagePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_font_color",
        "action": "setFontColor",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的文字颜色。\n\n使用场景：\n- \"把标题改成红色\"\n- \"设置第2页第1个文本框的文字颜色为蓝色\"\n- \"修改文字颜色\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "color": {
                "type": "string",
                "description": "颜色值，支持十六进制如 \"#FF0000\" 或颜色名如 \"red\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "color"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_image_style",
        "action": "setImageStyle",
        "app": "ppt",
        "summary": "设置幻灯片中指定图片的样式。\n\nstyle对象属性：\n- border: 边框设置 {enabled: boolean, color: string, weight: number}\n- shadow: 阴影设置 {enabled: boolean, color: string, blur: number, offsetX: number, offsetY: number}\n- opacity: 透明度 (0-100)\n- rotation: 旋转角度 (0-360)\n- cropTop/cropBottom/cropLeft/cropRight: 裁剪比例 (0-1)\n\n使用场景：\n- \"给图片加边框\"\n- \"设置图片阴影效果\"\n- \"旋转图片45度\"\n- \"裁剪图片\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "imageIndex": {
                "type": "number",
                "description": "图片索引（从1开始）",
                "required": true
            },
            "style": {
                "type": "object",
                "description": "图片样式配置对象",
                "schema": {
                    "type": "object",
                    "description": "图片样式配置对象",
                    "properties": {
                        "border": {
                            "type": "object",
                            "description": "边框设置",
                            "properties": {
                                "enabled": {
                                    "type": "boolean",
                                    "description": "是否启用边框"
                                },
                                "color": {
                                    "type": "string",
                                    "description": "边框颜色"
                                },
                                "weight": {
                                    "type": "number",
                                    "description": "边框粗细（磅）"
                                }
                            }
                        },
                        "shadow": {
                            "type": "object",
                            "description": "阴影设置",
                            "properties": {
                                "enabled": {
                                    "type": "boolean",
                                    "description": "是否启用阴影"
                                },
                                "color": {
                                    "type": "string",
                                    "description": "阴影颜色"
                                },
                                "blur": {
                                    "type": "number",
                                    "description": "模糊半径"
                                },
                                "offsetX": {
                                    "type": "number",
                                    "description": "水平偏移"
                                },
                                "offsetY": {
                                    "type": "number",
                                    "description": "垂直偏移"
                                }
                            }
                        },
                        "opacity": {
                            "type": "number",
                            "description": "透明度 (0-100)"
                        },
                        "rotation": {
                            "type": "number",
                            "description": "旋转角度 (0-360)"
                        },
                        "cropTop": {
                            "type": "number",
                            "description": "顶部裁剪比例 (0-1)"
                        },
                        "cropBottom": {
                            "type": "number",
                            "description": "底部裁剪比例 (0-1)"
                        },
                        "cropLeft": {
                            "type": "number",
                            "description": "左侧裁剪比例 (0-1)"
                        },
                        "cropRight": {
                            "type": "number",
                            "description": "右侧裁剪比例 (0-1)"
                        }
                    }
                },
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "imageIndex",
            "style"
        ],
        "engine": "bridge",
        "containers": [
            "style"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_master_background",
        "action": "setMasterBackground",
        "app": "ppt",
        "summary": "设置母版背景样式。\n\n支持纯色、渐变、图片等背景类型。\n\n使用场景：\n- \"修改母版背景为蓝色\"\n- \"设置母版背景渐变\"\n- \"给母版换个背景图片\"",
        "params": {
            "background": {
                "type": "object",
                "description": "背景配置对象，支持 {type:\"solid\",color:\"#xxx\"}, {type:\"gradient\",colors:[...]}, {type:\"image\",path:\"...\"}",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "background"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_ppt_chart_data",
        "action": "setPptChartData",
        "app": "ppt",
        "summary": "更新幻灯片中已有图表的数据。\n\n使用场景：\n- \"更新第2页图表的数据\"\n- \"修改图表数据\"\n- \"替换图表中的数值\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "chartIndex": {
                "type": "number",
                "description": "图表索引（从1开始）",
                "required": true
            },
            "data": {
                "type": "object",
                "description": "新的图表数据，包含 categories（类别数组）和 series（系列数组）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "chartIndex",
            "data"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_ppt_chart_style",
        "action": "setPptChartStyle",
        "app": "ppt",
        "summary": "设置幻灯片中图表的样式属性。\n\n可设置的样式属性（通过 style 对象传入）：\n- colorScheme: 配色方案名称\n- showLegend: 是否显示图例\n- legendPosition: 图例位置（top/bottom/left/right）\n- showDataLabels: 是否显示数据标签\n- fontSize: 字体大小\n- title: 图表标题\n\n使用场景：\n- \"修改图表配色为蓝色系\"\n- \"显示图表的数据标签\"\n- \"把图例移到底部\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "chartIndex": {
                "type": "number",
                "description": "图表索引（从1开始）",
                "required": true
            },
            "style": {
                "type": "object",
                "description": "图表样式配置对象",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "chartIndex",
            "style"
        ],
        "engine": "bridge",
        "containers": [
            "style"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_ppt_date_time",
        "action": "setPptDateTime",
        "app": "ppt",
        "summary": "设置演示文稿日期时间显示。\n\n支持的日期格式（format）：\n- \"YYYY-MM-DD\": 如 2026-03-21\n- \"YYYY/MM/DD\": 如 2026/03/21\n- \"MM/DD/YYYY\": 如 03/21/2026\n- \"DD/MM/YYYY\": 如 21/03/2026\n\n使用场景：\n- \"显示日期时间\"\n- \"设置自动更新日期\"\n- \"隐藏日期\"",
        "params": {
            "show": {
                "type": "boolean",
                "description": "是否显示日期时间",
                "required": true
            },
            "autoUpdate": {
                "type": "boolean",
                "description": "是否自动更新日期时间，默认为true"
            },
            "format": {
                "type": "string",
                "description": "日期时间格式，如 \"YYYY-MM-DD\""
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "show"
        ],
        "engine": "bridge",
        "aliases": {
            "show": "visible"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_ppt_footer",
        "action": "setPptFooter",
        "app": "ppt",
        "summary": "设置演示文稿页脚文本。\n\n使用场景：\n- \"添加页脚'公司名称'\"\n- \"设置页脚为'机密文件'\"\n- \"隐藏页脚\"",
        "params": {
            "text": {
                "type": "string",
                "description": "页脚文本内容",
                "required": true
            },
            "show": {
                "type": "boolean",
                "description": "是否显示页脚，默认为true"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_border",
        "action": "setShapeBorder",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的边框样式。\n\nborder对象属性：\n- enabled: 是否启用边框 (boolean)\n- color: 边框颜色，如 \"#000000\"\n- weight: 边框粗细（磅）\n- style: 边框样式，如 \"solid\"(实线)、\"dash\"(虚线)、\"dot\"(点线)\n\n使用场景：\n- \"给形状加边框\"\n- \"设置红色虚线边框\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "border": {
                "type": "object",
                "description": "边框配置对象",
                "schema": {
                    "type": "object",
                    "description": "边框配置对象",
                    "properties": {
                        "enabled": {
                            "type": "boolean",
                            "description": "是否启用边框"
                        },
                        "color": {
                            "type": "string",
                            "description": "边框颜色"
                        },
                        "weight": {
                            "type": "number",
                            "description": "边框粗细（磅）"
                        },
                        "style": {
                            "type": "string",
                            "description": "边框样式",
                            "enum": [
                                "solid",
                                "dash",
                                "dot",
                                "dash_dot",
                                "dash_dot_dot"
                            ]
                        }
                    }
                },
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "border"
        ],
        "engine": "bridge",
        "containers": [
            "border"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_fill",
        "action": "setShapeFill",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的填充颜色。\n\n使用场景：\n- \"把第1页的第2个形状填充为红色\"\n- \"修改形状背景色为#00FF00\"\n- \"设置形状的填充颜色\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "color": {
                "type": "string",
                "description": "填充颜色，十六进制如 #FF0000",
                "required": true
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "slideIndex",
            "shapeIndex",
            "color"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_gradient",
        "action": "setShapeGradient",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的渐变填充效果。\n\ngradient对象属性：\ngradient对象属性：\n- stops: 渐变色标数组 [{color: \"#FF0000\", position: 0}, {color: \"#0000FF\", position: 1}]（当前仅支持两个色标）\n注意：渐变角度与类型在 WPS 上不可设置（会挂起 COM 调用），因此不再提供 angle/type 参数。\n\n使用场景：\n- \"给形状加渐变色\"\n- \"设置从红到蓝的渐变\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "gradient": {
                "type": "object",
                "description": "渐变配置对象",
                "schema": {
                    "type": "object",
                    "description": "渐变配置对象",
                    "properties": {
                        "stops": {
                            "type": "array",
                            "description": "渐变色标数组",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "color": {
                                        "type": "string",
                                        "description": "颜色值"
                                    },
                                    "position": {
                                        "type": "number",
                                        "description": "位置 (0-1)"
                                    }
                                }
                            }
                        }
                    }
                },
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "gradient"
        ],
        "engine": "bridge",
        "containers": [
            "gradient"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_position",
        "action": "setShapePosition",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的位置和大小。\n\n使用场景：\n- \"把这个形状移到左上角\"\n- \"调整形状大小\"\n- \"设置形状位置为(100, 200)\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "left": {
                "type": "number",
                "description": "左边距（磅）",
                "required": true
            },
            "top": {
                "type": "number",
                "description": "上边距（磅）",
                "required": true
            },
            "width": {
                "type": "number",
                "description": "宽度（磅），可选"
            },
            "height": {
                "type": "number",
                "description": "高度（磅），可选"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "left",
            "top"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_shadow",
        "action": "setShapeShadow",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的阴影效果。\n\nshadow对象属性：\n- enabled: 是否启用阴影 (boolean)\n- color: 阴影颜色，如 \"#000000\"\n- blur: 模糊半径（磅）\n- offsetX: 水平偏移（磅）\n- offsetY: 垂直偏移（磅）\n- opacity: 透明度 (0-1)\n\n使用场景：\n- \"给这个形状加阴影\"\n- \"设置阴影效果\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "shadow": {
                "type": "object",
                "description": "阴影配置对象",
                "schema": {
                    "type": "object",
                    "description": "阴影配置对象",
                    "properties": {
                        "enabled": {
                            "type": "boolean",
                            "description": "是否启用阴影"
                        },
                        "color": {
                            "type": "string",
                            "description": "阴影颜色"
                        },
                        "blur": {
                            "type": "number",
                            "description": "模糊半径（磅）"
                        },
                        "offsetX": {
                            "type": "number",
                            "description": "水平偏移（磅）"
                        },
                        "offsetY": {
                            "type": "number",
                            "description": "垂直偏移（磅）"
                        },
                        "opacity": {
                            "type": "number",
                            "description": "透明度 (0-1)"
                        }
                    }
                },
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "shadow"
        ],
        "engine": "bridge",
        "containers": [
            "shadow"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_style",
        "action": "setShapeStyle",
        "app": "ppt",
        "summary": "设置幻灯片中形状的样式，包括填充颜色、边框颜色和边框粗细。\n\n使用场景：\n- \"把矩形改成红色\"\n- \"设置形状的边框为蓝色\"\n- \"修改形状样式\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始），默认1"
            },
            "name": {
                "type": "string",
                "description": "形状名称（通过getSlideInfo获取）"
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（与name二选一）"
            },
            "fillColor": {
                "type": "string",
                "description": "填充颜色，十六进制如 #FF0000"
            },
            "lineColor": {
                "type": "string",
                "description": "边框颜色，十六进制如 #000000"
            },
            "lineWidth": {
                "type": "number",
                "description": "边框粗细（磅）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_text",
        "action": "setShapeText",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的文字内容。\n\n使用场景：\n- \"把第1页的第2个形状文字改成'销售额'\"\n- \"修改形状里的文字\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "text": {
                "type": "string",
                "description": "要设置的文字内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "slideIndex",
            "shapeIndex",
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_transparency",
        "action": "setShapeTransparency",
        "app": "ppt",
        "summary": "设置幻灯片中指定形状的透明度。\n\n使用场景：\n- \"把这个形状设为半透明\"\n- \"设置形状透明度为50%\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "transparency": {
                "type": "number",
                "description": "透明度值 (0-100)，0为完全不透明，100为完全透明",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "transparency"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_shape_z_order",
        "action": "setShapeZOrder",
        "app": "ppt",
        "summary": "设置形状在幻灯片中的层级顺序（Z轴排列）。\n\n支持的层级操作（order）：\n- front: 置于顶层\n- back: 置于底层\n- forward: 上移一层\n- backward: 下移一层\n\n使用场景：\n- \"把这个形状移到最前面\"\n- \"将形状置于底层\"\n- \"上移一层\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "shapeIndex": {
                "type": "number",
                "description": "形状索引（从1开始）",
                "required": true
            },
            "order": {
                "type": "string",
                "description": "层级操作类型",
                "enum": [
                    "front",
                    "back",
                    "forward",
                    "backward"
                ],
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "shapeIndex",
            "order"
        ],
        "engine": "bridge",
        "aliases": {
            "order": "zOrder"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_background",
        "action": "setSlideBackground",
        "app": "ppt",
        "summary": "设置幻灯片背景，支持多种背景类型。\n\n支持的背景类型（background.type）：\n- solid: 纯色背景，需提供 color 字段\n- gradient: 渐变背景，需提供 colors 数组\n- image: 图片背景，需提供 imagePath 字段\n- pattern: 图案背景，需提供 pattern 和 color 字段\n\n使用场景：\n- \"给第1页设置蓝色背景\"\n- \"设置渐变背景\"\n- \"用图片做背景\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "background": {
                "type": "object",
                "description": "背景配置对象，包含 type/color/colors/imagePath/pattern 等字段"
            },
            "color": {
                "type": "string",
                "description": "纯色背景的颜色（等价于 background.color）"
            },
            "imagePath": {
                "type": "string",
                "description": "图片背景的路径（等价于 background.imagePath）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_content",
        "action": "setSlideContent",
        "app": "ppt",
        "summary": "设置幻灯片的正文内容区域。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "content": {
                "type": "string",
                "description": "正文内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "slideIndex",
            "content"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_layout",
        "action": "setSlideLayout",
        "app": "ppt",
        "summary": "设置幻灯片的版式布局。\n\n支持的布局类型：\n- title: 标题页\n- title_content: 标题+内容\n- blank: 空白页\n- two_column: 两栏内容\n- comparison: 对比布局\n- section_header: 节标题\n- title_only: 仅标题\n\n使用场景：\n- \"把这页改成空白布局\"\n- \"设置为两栏内容\"\n- \"改成标题页版式\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "layout": {
                "type": "string",
                "description": "布局名称",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "layout"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_notes",
        "action": "setSlideNotes",
        "app": "ppt",
        "summary": "设置幻灯片的备注内容，用于演讲提示。\n\n使用场景：\n- \"给第1页添加备注\"\n- \"写一些演讲提示\"\n- \"修改备注内容\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "notes": {
                "type": "string",
                "description": "备注内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "notes"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_number",
        "action": "setSlideNumber",
        "app": "ppt",
        "summary": "设置幻灯片页码的显示状态和起始编号。\n\n使用场景：\n- \"显示页码\"\n- \"隐藏幻灯片编号\"\n- \"页码从第2页开始编号\"",
        "params": {
            "show": {
                "type": "boolean",
                "description": "是否显示页码",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "show"
        ],
        "engine": "bridge",
        "aliases": {
            "show": "visible"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_size",
        "action": "setSlideSize",
        "app": "ppt",
        "summary": "设置演示文稿的幻灯片尺寸。\n\n常用尺寸：\n- 标准(4:3): 宽960, 高720\n- 宽屏(16:9): 宽960, 高540\n- 宽屏(16:10): 宽960, 高600\n- A4横版: 宽1123, 高794\n- A4竖版: 宽794, 高1123\n\n使用场景：\n- \"把PPT改成16:9宽屏\"\n- \"设置幻灯片为A4尺寸\"\n- \"修改幻灯片大小\"",
        "params": {
            "width": {
                "type": "number",
                "description": "幻灯片宽度（像素）",
                "required": true
            },
            "height": {
                "type": "number",
                "description": "幻灯片高度（像素）",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "width",
            "height"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_subtitle",
        "action": "setSlideSubtitle",
        "app": "ppt",
        "summary": "设置幻灯片的副标题。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "subtitle": {
                "type": "string",
                "description": "副标题内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "subtitle"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_theme",
        "action": "setSlideTheme",
        "app": "ppt",
        "summary": "把演示文稿套用为指定的主题模板（.thmx / .potx / .pptx）。\n\n使用场景：\n- \"换成这套主题\" / \"套用这个模板\"\n- 按公司模板统一整套演示的字体与配色\n\n注意：theme 必须是**磁盘上存在的模板文件路径**，不是\"商务/简约\"这类名称；\n按名称换肤请改用配色、母版与形状样式类工具逐个调整。",
        "params": {
            "theme": {
                "type": "string",
                "description": "模板文件路径（.thmx / .potx / .pptx），必须存在",
                "required": true
            },
            "presentationName": {
                "type": "string",
                "description": "目标演示文稿名称；不填则用当前活动文稿"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "theme"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_title",
        "action": "setSlideTitle",
        "app": "ppt",
        "summary": "设置幻灯片的标题文本。\n\n使用场景：\n- \"把第1页标题改成'年度总结'\"\n- \"设置标题为'项目进展'\"\n- \"修改幻灯片标题\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始），默认1"
            },
            "title": {
                "type": "string",
                "description": "标题文本",
                "required": true
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "title"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_slide_transition",
        "action": "setSlideTransition",
        "app": "ppt",
        "summary": "设置幻灯片的页面切换效果。\n\n常用切换效果：\n- fade: 淡出\n- push: 推入\n- wipe: 擦除\n- split: 分割\n- reveal: 显露\n- cover: 覆盖\n- curtains: 帷幕\n- blinds: 百叶窗\n\n使用场景：\n- \"给第1页加个淡出切换\"\n- \"设置页面切换为推入效果\"\n- \"添加切换动画\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "effect": {
                "type": "string",
                "description": "切换效果名称",
                "required": true
            },
            "duration": {
                "type": "number",
                "description": "切换持续时间（秒），默认1秒"
            },
            "sound": {
                "type": "string",
                "description": "切换时播放的声音文件路径（可选）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "effect"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_table_cell",
        "action": "setPptTableCell",
        "app": "ppt",
        "summary": "设置PPT表格中指定单元格的文本内容。\n\n使用场景：\n- \"把表格第1行第2列的内容改成'销售额'\"\n- \"设置单元格文本\"\n- \"填写表格内容\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "tableIndex": {
                "type": "number",
                "description": "表格索引（从1开始）",
                "required": true
            },
            "row": {
                "type": "number",
                "description": "行号（从1开始）",
                "required": true
            },
            "col": {
                "type": "number",
                "description": "列号（从1开始）",
                "required": true
            },
            "text": {
                "type": "string",
                "description": "要设置的文本内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "tableIndex",
            "row",
            "col",
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_table_cell_style",
        "action": "setPptTableCellStyle",
        "app": "ppt",
        "summary": "设置PPT表格中指定单元格的样式。\n\n支持的样式属性（通过style对象传入）：\n- backgroundColor: 单元格背景色\n- fontName: 字体名称\n- fontSize: 字体大小（磅）\n- fontColor: 字体颜色\n- bold: 是否加粗（boolean）\n- italic: 是否斜体（boolean）\n- alignment: 文本对齐方式（left/center/right）\n- verticalAlignment: 垂直对齐方式（top/middle/bottom）\n\n使用场景：\n- \"把表格第1行第1列的背景设为蓝色\"\n- \"加粗标题单元格\"\n- \"设置单元格居中对齐\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "tableIndex": {
                "type": "number",
                "description": "表格索引（从1开始）",
                "required": true
            },
            "row": {
                "type": "number",
                "description": "行号（从1开始）",
                "required": true
            },
            "col": {
                "type": "number",
                "description": "列号（从1开始）",
                "required": true
            },
            "style": {
                "type": "object",
                "description": "样式配置对象，包含backgroundColor、fontName、fontSize、fontColor、bold、italic、alignment、verticalAlignment等属性",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "tableIndex",
            "row",
            "col",
            "style"
        ],
        "engine": "bridge",
        "containers": [
            "style"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_table_row_style",
        "action": "setPptTableRowStyle",
        "app": "ppt",
        "summary": "设置PPT表格中指定行的样式。\n\n支持的样式属性（通过style对象传入）：\n- backgroundColor: 行背景色\n- fontName: 字体名称\n- fontSize: 字体大小（磅）\n- fontColor: 字体颜色\n- bold: 是否加粗（boolean）\n- italic: 是否斜体（boolean）\n- alignment: 文本对齐方式（left/center/right）\n- height: 行高（磅）\n\n使用场景：\n- \"把表头行设为蓝色背景白色字\"\n- \"加粗第一行\"\n- \"设置表格行高\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "tableIndex": {
                "type": "number",
                "description": "表格索引（从1开始）",
                "required": true
            },
            "row": {
                "type": "number",
                "description": "行号（从1开始）",
                "required": true
            },
            "style": {
                "type": "object",
                "description": "样式配置对象，包含backgroundColor、fontName、fontSize、fontColor、bold、italic、alignment、height等属性",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "tableIndex",
            "row",
            "style"
        ],
        "engine": "bridge",
        "containers": [
            "style"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_table_style",
        "action": "setPptTableStyle",
        "app": "ppt",
        "summary": "设置PPT表格的整体样式。\n\n支持的样式属性（通过style对象传入）：\n- borderColor: 边框颜色（如 \"#000000\"）\n- borderWidth: 边框宽度（磅）\n- backgroundColor: 背景色（如 \"#FFFFFF\"）\n- fontName: 字体名称（如 \"微软雅黑\"）\n- fontSize: 字体大小（磅）\n- fontColor: 字体颜色（如 \"#333333\"）\n- headerBackground: 表头行背景色\n- alternateRowColor: 隔行变色颜色\n\n使用场景：\n- \"设置表格为蓝色主题\"\n- \"修改表格边框和背景\"\n- \"美化表格样式\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片页码（从1开始）",
                "required": true
            },
            "tableIndex": {
                "type": "number",
                "description": "表格索引（从1开始）",
                "required": true
            },
            "style": {
                "type": "object",
                "description": "样式配置对象，包含borderColor、borderWidth、backgroundColor、fontName、fontSize、fontColor、headerBackground、alternateRowColor等属性",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "tableIndex",
            "style"
        ],
        "engine": "bridge",
        "containers": [
            "style"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_textbox_style",
        "action": "setTextBoxStyle",
        "app": "ppt",
        "summary": "设置文本框样式（字体大小、颜色、粗体等）。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "textboxIndex": {
                "type": "number",
                "description": "文本框索引",
                "required": true
            },
            "style": {
                "type": "object",
                "description": "样式对象，包含 fontSize/fontColor/bold/italic/align 等",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "textboxIndex",
            "style"
        ],
        "engine": "bridge",
        "containers": [
            "style"
        ]
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_set_textbox_text",
        "action": "setTextBoxText",
        "app": "ppt",
        "summary": "设置指定文本框的文本内容。",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "幻灯片索引（从1开始）",
                "required": true
            },
            "textboxIndex": {
                "type": "number",
                "description": "文本框索引",
                "required": true
            },
            "text": {
                "type": "string",
                "description": "文本内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "slideIndex",
            "textboxIndex",
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_start_slide_show",
        "action": "startSlideShow",
        "app": "ppt",
        "summary": "开始幻灯片放映。\n\n可以从指定页面开始放映，默认从第1页开始。\n\n使用场景：\n- \"放映幻灯片\"\n- \"从第3页开始演示\"\n- \"开始PPT放映\"",
        "params": {
            "fromSlide": {
                "type": "number",
                "description": "从第几页开始放映（从1开始），默认从第1页开始"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_switch_presentation",
        "action": "switchPresentation",
        "app": "ppt",
        "summary": "切换到指定名称的演示文稿。\n\n使用场景：\n- \"切换到另一个PPT\"\n- \"打开那个叫xxx的演示文稿\"\n- \"切换演示文稿\"",
        "params": {
            "name": {
                "type": "string",
                "description": "要切换到的演示文稿名称",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_switch_slide",
        "action": "switchSlide",
        "app": "ppt",
        "summary": "切换到指定的幻灯片页面。\n\n使用场景：\n- \"切换到第5页\"\n- \"跳到最后一页\"\n- \"显示第1页\"",
        "params": {
            "slideIndex": {
                "type": "number",
                "description": "目标幻灯片索引（从1开始）",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [
            "slideIndex"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_ppt_unify_font",
        "action": "unifyFont",
        "app": "ppt",
        "summary": "统一演示文稿中所有幻灯片的字体。\n\n使用场景：\n- \"把所有页面的字体都改成微软雅黑\"\n- \"统一字体\"\n- \"换个字体\"\n\n常用字体推荐：\n- 微软雅黑：现代简洁，适合商务\n- 思源黑体：开源免费，适合各种场合\n- 黑体：传统正式\n- 宋体：适合正式文档",
        "params": {
            "font_name": {
                "type": "string",
                "description": "要统一使用的字体名称，如 \"微软雅黑\"、\"思源黑体\"",
                "required": true
            },
            "slide_index": {
                "type": "number",
                "description": "只处理指定页，不填则处理所有页"
            },
            "include_title": {
                "type": "boolean",
                "description": "是否包含标题，默认true"
            },
            "include_body": {
                "type": "boolean",
                "description": "是否包含正文，默认true"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "font_name"
        ],
        "engine": "bridge",
        "aliases": {
            "font_name": "fontName",
            "slide_index": "slideIndex",
            "include_title": "includeTitle",
            "include_body": "includeBody"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_apply_style",
        "action": "applyStyle",
        "app": "word",
        "summary": "应用Word样式到当前选中区域或指定范围。\n\n支持的常用样式：\n- 标题1、标题2、标题3...（或 Heading 1, Heading 2...）\n- 正文、正文首行缩进\n- 引用、强调\n- 列表段落\n\n使用场景：\n- \"把这段设成标题1\"\n- \"应用正文样式\"",
        "params": {
            "style_name": {
                "type": "string",
                "description": "样式名称，如 \"标题 1\"、\"正文\"、\"Heading 1\"",
                "required": true
            },
            "range": {
                "type": "object",
                "description": "指定范围，不填则应用到当前选中区域",
                "schema": {
                    "type": "object",
                    "description": "指定范围，不填则应用到当前选中区域",
                    "properties": {
                        "start": {
                            "type": "number",
                            "description": "起始位置（字符索引）"
                        },
                        "end": {
                            "type": "number",
                            "description": "结束位置（字符索引）"
                        }
                    }
                }
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "style_name"
        ],
        "engine": "bridge",
        "aliases": {
            "style_name": "styleName"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_close_document",
        "action": "closeDocument",
        "app": "word",
        "summary": "关闭 Word 文档，可选是否保存。\n\n使用场景：\n- \"关掉这个文档，别留着\"\n- 一批任务收尾时清理打开的文档\n\n从未保存到磁盘的文档不会被强制保存（不会弹出保存对话框），此时结果里会带 warning 说明。",
        "params": {
            "name": {
                "type": "string",
                "description": "要关闭的文档名称；不填则关闭当前活动文档"
            },
            "save": {
                "type": "boolean",
                "description": "是否保存后关闭，默认 true"
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_create_document",
        "action": "createDocument",
        "app": "word",
        "summary": "新建一个空白 Word 文档（不是打开已有文件）。\n\n使用场景：\n- \"新建一个 Word 文档\"\n- \"把结论写进一个新文档里\"\n- 需要从零起草，而不是编辑现有文档时\n\n新建后用 wps_word_insert_text 写内容，用 wps_common_save_as 保存到磁盘。",
        "params": {},
        "effect": "lifecycle",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_enable_track_changes",
        "action": "enableTrackChanges",
        "app": "word",
        "summary": "开启或关闭Word文档的修订模式（Track Changes）。\n在执行校对修改前必须先开启修订模式，确保所有修改可追溯。\n\n使用场景：\n- \"开始校对，开启修订模式\"\n- \"关闭修订模式\"\n- \"开启修订，我要开始改文档了\"",
        "params": {
            "enable": {
                "type": "boolean",
                "description": "true=开启修订模式，false=关闭修订模式。默认true"
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_find_in_document",
        "action": "findInDocument",
        "app": "word",
        "summary": "在Word文档中查找文本并返回位置信息，不执行替换操作。\n\n使用场景：\n- \"找一下'项目名称'在文档的哪个位置\"\n- \"看看文档里有哪些地方需要填写\"\n- 在使用smart_fill_field之前，先用此工具定位关键字\n\n返回信息包括：匹配文本、字符起止位置、所在段落索引、上下文（前后50字符）。\n与find_replace不同，此工具仅查找不替换，返回位置信息供后续操作使用。",
        "params": {
            "find_text": {
                "type": "string",
                "description": "要查找的文本",
                "required": true
            },
            "match_case": {
                "type": "boolean",
                "description": "是否区分大小写，默认false"
            },
            "match_whole_word": {
                "type": "boolean",
                "description": "是否全字匹配，默认false"
            },
            "max_results": {
                "type": "number",
                "description": "最大返回结果数，默认20"
            }
        },
        "effect": "read",
        "advertised": false,
        "required": [
            "find_text"
        ],
        "engine": "bridge",
        "aliases": {
            "find_text": "findText",
            "match_case": "matchCase",
            "match_whole_word": "matchWholeWord",
            "max_results": "maxResults"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_find_replace",
        "action": "findReplace",
        "app": "word",
        "summary": "在Word文档中查找并替换文本。\n\n使用场景：\n- \"把所有的'公司'替换成'集团'\"\n- \"把文档里的错别字改过来\"\n- \"批量替换某个词\"\n\n支持选项：\n- 区分大小写\n- 全字匹配\n- 全部替换或仅替换一处",
        "params": {
            "find_text": {
                "type": "string",
                "description": "要查找的文本",
                "required": true
            },
            "replace_text": {
                "type": "string",
                "description": "替换为的文本，如果只是查找不替换，可以不填"
            },
            "replace_all": {
                "type": "boolean",
                "description": "是否全部替换，默认true"
            },
            "match_case": {
                "type": "boolean",
                "description": "是否区分大小写，默认false"
            },
            "match_whole_word": {
                "type": "boolean",
                "description": "是否全字匹配，默认false"
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [
            "find_text"
        ],
        "engine": "bridge",
        "aliases": {
            "find_text": "findText",
            "replace_text": "replaceText",
            "replace_all": "replaceAll",
            "match_case": "matchCase",
            "match_whole_word": "matchWholeWord"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_generate_toc",
        "action": "generateTOC",
        "app": "word",
        "summary": "根据文档中的标题样式自动生成目录。\n\n前提条件：文档中的标题必须使用\"标题1\"、\"标题2\"等样式。\n\n使用场景：\n- \"帮我生成目录\"\n- \"在文档开头插入目录\"",
        "params": {
            "position": {
                "type": "string",
                "description": "插入位置，可选值: \"start\"(文档开头), \"cursor\"(当前光标位置)。默认start",
                "enum": [
                    "start",
                    "cursor"
                ]
            },
            "levels": {
                "type": "number",
                "description": "目录包含的标题级别数，如 3 表示包含标题1-3。默认3"
            },
            "include_page_numbers": {
                "type": "boolean",
                "description": "是否包含页码，默认true"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "include_page_numbers": "includePageNumbers"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_get_active_document",
        "action": "getActiveDocument",
        "app": "word",
        "summary": "获取当前WPS Writer活动文档的基本信息",
        "params": {},
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_get_document_text",
        "action": "getDocumentText",
        "app": "word",
        "summary": "获取当前Word文档的文本内容。\n\n使用场景：\n- \"读取文档内容\"\n- \"获取文档的前100个字符\"\n- \"查看文档从第50到第200个字符的内容\"\n\n可指定起始和结束位置来获取部分文本。",
        "params": {
            "start": {
                "type": "number",
                "description": "起始位置（字符索引），默认从头开始",
                "kind": "local"
            },
            "end": {
                "type": "number",
                "description": "结束位置（字符索引），默认到文档末尾",
                "kind": "local"
            }
        },
        "effect": "read",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_get_open_documents",
        "action": "getOpenDocuments",
        "app": "word",
        "summary": "获取当前WPS Writer中所有已打开的文档列表。\n\n使用场景：\n- \"看看现在打开了哪些文档\"\n- \"列出所有打开的Word文件\"\n- \"查看当前文档列表\"",
        "params": {},
        "effect": "lifecycle",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_get_paragraphs",
        "action": "getDocumentParagraphs",
        "app": "word",
        "summary": "获取Word文档的段落结构信息，返回每段的文本、样式和字符位置。\n\n使用场景：\n- \"了解文档的结构\"\n- \"查看文档有哪些段落\"\n- \"帮我看看模板里有哪些需要填写的位置\"\n- 在填写模板前，先读取文档结构以识别填写位置\n\n返回信息包括：段落索引、文本内容、样式名称、字符起止位置。\n支持分页获取（startParagraph/endParagraph），默认返回前50段。",
        "params": {
            "start_paragraph": {
                "type": "number",
                "description": "起始段落索引（从1开始），默认1"
            },
            "end_paragraph": {
                "type": "number",
                "description": "结束段落索引，默认为起始+49"
            }
        },
        "effect": "read",
        "advertised": true,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "start_paragraph": "startParagraph",
            "end_paragraph": "endParagraph"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_get_track_changes_status",
        "action": "getTrackChangesStatus",
        "app": "word",
        "summary": "获取当前文档的修订模式状态。\n\n使用场景：\n- \"看看修订模式开了没\"\n- \"当前文档有多少处修订\"\n- \"检查修订状态\"",
        "params": {},
        "effect": "read",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_bookmark",
        "action": "insertBookmark",
        "app": "word",
        "summary": "在当前光标位置或选中区域插入书签。\n\n书签可用于：\n- 交叉引用\n- 超链接跳转目标\n- 文档内快速导航\n\n使用场景：\n- \"在这里插入一个书签\"\n- \"标记这个位置为'章节开头'\"",
        "params": {
            "name": {
                "type": "string",
                "description": "书签名称，不能包含空格，建议使用英文或下划线连接",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_comment",
        "action": "addComment",
        "app": "word",
        "summary": "在Word文档选中内容处插入批注",
        "params": {
            "text": {
                "type": "string",
                "description": "批注内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_footer",
        "action": "insertFooter",
        "app": "word",
        "summary": "设置页脚内容。\n\n使用场景：\n- \"给文档加个页脚\"\n- \"设置页脚为页码\"\n- \"修改第1节的页脚\"",
        "params": {
            "text": {
                "type": "string",
                "description": "页脚文本内容",
                "required": true
            },
            "section": {
                "type": "number",
                "description": "节编号（从1开始），默认1",
                "default": 1,
                "schema": {
                    "type": "number",
                    "description": "节编号（从1开始），默认1",
                    "default": 1
                }
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_header",
        "action": "insertHeader",
        "app": "word",
        "summary": "设置页眉内容。\n\n使用场景：\n- \"给文档加个页眉\"\n- \"设置页眉为公司名称\"\n- \"修改第2节的页眉\"",
        "params": {
            "text": {
                "type": "string",
                "description": "页眉文本内容",
                "required": true
            },
            "section": {
                "type": "number",
                "description": "节编号（从1开始），默认1",
                "default": 1,
                "schema": {
                    "type": "number",
                    "description": "节编号（从1开始），默认1",
                    "default": 1
                }
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_image",
        "action": "insertImage",
        "app": "word",
        "summary": "在Word文档中插入图片。\n\n使用场景：\n- \"在文档中插入一张图片\"\n- \"把这个截图放到文档里\"\n- \"在光标位置插入logo\"",
        "params": {
            "imagePath": {
                "type": "string",
                "description": "图片文件路径",
                "required": true
            },
            "width": {
                "type": "number",
                "description": "图片宽度（磅），可选"
            },
            "height": {
                "type": "number",
                "description": "图片高度（磅），可选"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "imagePath"
        ],
        "engine": "bridge",
        "aliases": {
            "imagePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_page_break",
        "action": "insertPageBreak",
        "app": "word",
        "summary": "在文档光标位置插入分页符",
        "params": {},
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_section_break",
        "action": "insertSectionBreak",
        "app": "word",
        "summary": "插入分节符，用于将文档分为不同的节，以便对各节应用不同的页面设置。\n\n使用场景：\n- \"插入一个分节符\"\n- \"从下一页开始新的一节\"\n- \"在这里分节\"",
        "params": {
            "breakType": {
                "type": "string",
                "description": "分节符类型：nextPage(下一页)、continuous(连续)、evenPage(偶数页)、oddPage(奇数页)，默认nextPage",
                "default": "nextPage",
                "schema": {
                    "type": "string",
                    "description": "分节符类型：nextPage(下一页)、continuous(连续)、evenPage(偶数页)、oddPage(奇数页)，默认nextPage",
                    "default": "nextPage"
                }
            }
        },
        "effect": "write",
        "advertised": false,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_table",
        "action": "insertTable",
        "app": "word",
        "summary": "在Word文档光标位置插入表格",
        "params": {
            "rows": {
                "type": "number",
                "description": "表格行数",
                "required": true
            },
            "cols": {
                "type": "number",
                "description": "表格列数",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "rows",
            "cols"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_insert_text",
        "action": "insertText",
        "app": "word",
        "summary": "在Word文档中插入文本。\n\n使用场景：\n- \"在文档开头加个标题\"\n- \"在光标位置插入这段话\"\n- \"在文档末尾添加备注\"",
        "params": {
            "text": {
                "type": "string",
                "description": "要插入的文本内容",
                "required": true
            },
            "position": {
                "type": "string",
                "description": "插入位置，可选值: \"cursor\"(光标位置), \"start\"(文档开头), \"end\"(文档结尾)。默认cursor",
                "enum": [
                    "cursor",
                    "start",
                    "end"
                ]
            },
            "style": {
                "type": "string",
                "description": "插入后应用的样式，如 \"标题 1\"、\"正文\""
            },
            "new_paragraph": {
                "type": "boolean",
                "description": "插入后是否新起一段，默认false",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_open_document",
        "action": "openDocument",
        "app": "word",
        "summary": "打开指定路径的Word文档。\n\n使用场景：\n- \"打开桌面上的报告.docx\"\n- \"帮我打开这个文件路径的文档\"\n- \"加载指定位置的Word文件\"",
        "params": {
            "filePath": {
                "type": "string",
                "description": "要打开的文档文件路径",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": true,
        "required": [
            "filePath"
        ],
        "engine": "bridge",
        "aliases": {
            "filePath": "path"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_proofread_basic",
        "action": null,
        "app": "word",
        "summary": "对中文文本进行基础校对，检测常见问题。\n使用正则规则快速检测，纯本地校验，无需 API 调用。\n\n检测范围：\n- 常见易混淆字（的/得/地、在/再 等）\n- 重复字符（如\"了了\"、\"的的\"）\n- 重复标点（如。。、，，、！！、？？）\n- 中英文标点混用\n- 常见错误搭配\n- 数字前后异常空格\n- 常见网络用语/拼写错误\n\n返回每个问题的位置、原文、建议修改和问题类型。\n\n使用场景：\n- 校对前先做基础检查\n- 批量处理段落文本\n- 快速定位文档中的常见错误\n注意：单次最多处理 50000 字符，超出请分批调用。建议每批约 20 段。",
        "params": {
            "text": {
                "type": "string",
                "description": "要校对的文本内容",
                "required": true,
                "kind": "local"
            },
            "start_offset": {
                "type": "number",
                "description": "文本在文档中的起始偏移位置，用于定位问题在文档中的准确位置",
                "kind": "local"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "text"
        ],
        "engine": "local"
    }),
    (0, types_1.op)({
        "tool": "wps_word_replace_bookmark_content",
        "action": "replaceBookmarkContent",
        "app": "word",
        "summary": "替换Word文档中书签的内容。通过书签名定位，替换书签范围内的文本，保持原有格式。\n\n使用场景：\n- \"把书签'project_name'的内容改为'XX项目'\"\n- 模板文档使用书签标记填写位置时使用\n\n注意：替换后书签会自动重建，不会丢失。",
        "params": {
            "name": {
                "type": "string",
                "description": "书签名称",
                "required": true
            },
            "text": {
                "type": "string",
                "description": "要替换为的文本内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "name",
            "text"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_replace_range",
        "action": "replaceRange",
        "app": "word",
        "summary": "按字符范围精确替换Word文档中的文本。\n在修订模式下，此操作会自动产生修订标记。\n\n使用场景：\n- 校对时替换指定位置的错别字\n- 精确替换某一段落中的文本\n- 在已知字符起止位置时替换内容\n\n注意：请先调用 wps_word_enable_track_changes 开启修订模式。",
        "params": {
            "start_pos": {
                "type": "number",
                "description": "起始字符位置（从0开始）",
                "required": true
            },
            "end_pos": {
                "type": "number",
                "description": "结束字符位置",
                "required": true
            },
            "text": {
                "type": "string",
                "description": "替换后的文本内容",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "start_pos",
            "end_pos",
            "text"
        ],
        "engine": "bridge",
        "aliases": {
            "start_pos": "startPos",
            "end_pos": "endPos"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_set_font",
        "action": "setFont",
        "app": "word",
        "summary": "设置字体格式，包括字体名称、字号、加粗、斜体、颜色等。\n\n使用场景：\n- \"把标题改成微软雅黑24号加粗\"\n- \"把这段文字改成红色\"\n- \"全文字体改成宋体小四\"",
        "params": {
            "font_name": {
                "type": "string",
                "description": "字体名称，如 \"微软雅黑\"、\"宋体\"、\"Arial\""
            },
            "font_size": {
                "type": "number",
                "description": "字号，如 12、14、24"
            },
            "bold": {
                "type": "boolean",
                "description": "是否加粗"
            },
            "italic": {
                "type": "boolean",
                "description": "是否斜体"
            },
            "underline": {
                "type": "boolean",
                "description": "是否下划线"
            },
            "color": {
                "type": "string",
                "description": "字体颜色，支持颜色名称(red/blue/green)或十六进制(#FF0000)"
            },
            "range": {
                "type": "string",
                "description": "应用范围，可选值: \"selection\"(当前选中), \"all\"(全文)。默认selection",
                "enum": [
                    "selection",
                    "all"
                ]
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "font_name": "fontName",
            "font_size": "fontSize"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_set_line_spacing",
        "action": "setLineSpacing",
        "app": "word",
        "summary": "设置段落行距。\n\n使用场景：\n- \"把行距设为1.5倍\"\n- \"设置第3段的行距为2倍\"\n- \"调整行距\"",
        "params": {
            "lineSpacing": {
                "type": "number",
                "description": "行距值（如1.0、1.5、2.0等）",
                "required": true
            },
            "paragraphIndex": {
                "type": "number",
                "description": "段落索引（从0开始），不指定则应用于当前段落或全文"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "lineSpacing"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_set_page_setup",
        "action": "setPageSetup",
        "app": "word",
        "summary": "设置文档页面布局，包括页面方向和边距。\n\n使用场景：\n- \"把页面改成横向\"\n- \"设置上下边距为2厘米\"\n- \"调整页面为A4横向，边距2cm\"",
        "params": {
            "orientation": {
                "type": "string",
                "description": "页面方向: \"portrait\"(纵向) 或 \"landscape\"(横向)",
                "enum": [
                    "portrait",
                    "landscape"
                ]
            },
            "marginTop": {
                "type": "number",
                "description": "上边距（磅值）"
            },
            "marginBottom": {
                "type": "number",
                "description": "下边距（磅值）"
            },
            "marginLeft": {
                "type": "number",
                "description": "左边距（磅值）"
            },
            "marginRight": {
                "type": "number",
                "description": "右边距（磅值）"
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [],
        "engine": "bridge",
        "aliases": {
            "marginTop": "topMargin",
            "marginBottom": "bottomMargin",
            "marginLeft": "leftMargin",
            "marginRight": "rightMargin"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_set_paragraph",
        "action": "setParagraph",
        "app": "word",
        "summary": "设置当前段落格式（对齐方式、行间距等）",
        "params": {
            "alignment": {
                "type": "string",
                "description": "对齐方式: left/center/right/justify"
            },
            "lineSpacing": {
                "type": "number",
                "description": "行间距倍数，如1.5、2"
            }
        },
        "effect": "write",
        "advertised": true,
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_set_text_color",
        "action": "setTextColor",
        "app": "word",
        "summary": "设置Word文档中选中文字的颜色",
        "params": {
            "color": {
                "type": "string",
                "description": "颜色值，如 \"#FF0000\"、\"red\"",
                "required": true
            }
        },
        "effect": "write",
        "advertised": false,
        "required": [
            "color"
        ],
        "engine": "bridge"
    }),
    (0, types_1.op)({
        "tool": "wps_word_smart_fill_field",
        "action": "smartFillField",
        "app": "word",
        "summary": "智能填写Word模板中的字段。自动识别关键字附近的填写模式，在正确位置插入内容并保持原有格式。\n\n使用场景：\n- \"把项目名称填写为'XX信息化项目'\"\n- \"填写建设单位为'XX公司'\"\n- \"模板里的项目编号填上'2026-001'\"\n- 任何需要在模板文档中\"填写\"而非\"替换\"的场景\n\n支持的填写模式（fill_mode，默认auto自动判断）：\n- auto: 自动判断（推荐）\n- underline: 关键字后有下划线___，替换下划线为填写内容\n- afterColon: 关键字后有冒号（：或:），在冒号后插入\n- afterLabel: 关键字是标签，直接在关键字后插入\n- placeholder: 关键字被{}或【】包裹，替换整个占位符\n\n重要：模板填写场景应优先使用此工具，而非find_replace。find_replace会删除关键字本身并可能破坏格式。",
        "params": {
            "keyword": {
                "type": "string",
                "description": "要填写的关键字（如\"项目名称\"、\"建设单位\"）",
                "required": true
            },
            "value": {
                "type": "string",
                "description": "要填写的值（如\"XX信息化项目\"）",
                "required": true
            },
            "fill_mode": {
                "type": "string",
                "description": "填写模式: auto(自动判断), underline(下划线), afterColon(冒号后), afterLabel(标签后), placeholder(占位符)。默认auto",
                "enum": [
                    "auto",
                    "underline",
                    "afterColon",
                    "afterLabel",
                    "placeholder"
                ]
            }
        },
        "effect": "write",
        "advertised": true,
        "required": [
            "keyword",
            "value"
        ],
        "engine": "bridge",
        "aliases": {
            "fill_mode": "fillMode"
        }
    }),
    (0, types_1.op)({
        "tool": "wps_word_switch_document",
        "action": "switchDocument",
        "app": "word",
        "summary": "切换到指定名称的文档。\n\n使用场景：\n- \"切换到报告.docx\"\n- \"打开另一个文档窗口\"\n- \"切换到那个合同文档\"",
        "params": {
            "name": {
                "type": "string",
                "description": "要切换到的文档名称",
                "required": true
            }
        },
        "effect": "lifecycle",
        "advertised": false,
        "required": [
            "name"
        ],
        "engine": "bridge"
    }),
];
//# sourceMappingURL=operations.js.map