/**
 * Input: 无
 * Output: 桥侧的参数兼容声明：公开参数名 -> 该 action 实际读取的键；以及嵌套容器名
 * Pos: P1 契约真源的一部分。生成器从这里产出 spec/param-aliases.json 与 spec/param-containers.json，
 *      宿主脚本再据此在派发前改名与展平——取代了 build-host-actions.ps1 里那两张手写表。
 *      一旦我被修改，请更新我的头部注释，以及 docs/tool-roadmap.md 的 P1 状态。
 */

/** 公开参数名 -> 桥读取的键。含工具 schema 没有声明、但历史上被接受的兼容拼写。 */
export const paramAliases: Record<string, Record<string, string>> = {
  "removeAnimation": {
    "animationIndex": "index"
  },
  "setAnimationOrder": {
    "animationIndex": "from",
    "newOrder": "to"
  },
  "setShapeZOrder": {
    "order": "zOrder"
  },
  "addAnimation": {
    "animationType": "effect",
    "shapeIndex": "shapeName"
  },
  "setBackgroundImage": {
    "imagePath": "path",
    "filePath": "path"
  },
  "insertPptImage": {
    "imagePath": "path",
    "filePath": "path"
  },
  "replacePptImage": {
    "imagePath": "path"
  },
  "openPresentation": {
    "filePath": "path"
  },
  "alignShapes": {
    "shapeIndices": "names"
  },
  "distributeShapes": {
    "shapeIndices": "names"
  },
  "groupShapes": {
    "shapeIndices": "names"
  },
  "setSlideNumber": {
    "show": "visible"
  },
  "setPptDateTime": {
    "show": "visible"
  },
  "applyTransitionToAll": {
    "effect": "transition"
  },
  "setSlideTransition": {
    "transition": "effect"
  },
  "addPptHyperlink": {
    "url": "address"
  },
  "insertPptChart": {
    "chartType": "type"
  }
};

/** 需要展平的嵌套容器名（其属性会被并到该 action 的扁平键集合上）。 */
export const paramContainers: Record<string, string[]> = {
  "setShapeShadow": [
    "shadow"
  ],
  "setShapeBorder": [
    "border"
  ],
  "setShapeGradient": [
    "gradient"
  ],
  "setBackgroundGradient": [
    "gradient"
  ],
  "addMasterElement": [
    "element"
  ],
  "set3DRotation": [
    "rotation"
  ],
  "setImageStyle": [
    "style"
  ],
  "setTextBoxStyle": [
    "style"
  ],
  "setPptChartStyle": [
    "style"
  ],
  "setPptTableStyle": [
    "style"
  ],
  "setPptTableCellStyle": [
    "style"
  ],
  "setPptTableRowStyle": [
    "style"
  ]
};
