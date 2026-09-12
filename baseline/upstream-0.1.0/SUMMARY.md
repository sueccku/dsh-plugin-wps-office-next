# Baseline: upstream wps-skills @ a825336 (2026-06-30)

Captured by scripts/snapshot-tools.mjs plus source extraction, for regression comparison.
Do not edit by hand; regenerate when the baseline is intentionally re-cut.

## Live MCP surface (measured over stdio)

| Metric | Value |
|---|---|
| Advertised tools | 250 |
| tools/list schema bytes | 141872 |
| Approx tokens per request | 40535 |

Pro tools (238): excel=82, word=32, ppt=115, common=7, convert=2. Builtins (12): names not prefixed wps_excel/word/ppt/common/convert.

## Inventories

| Artifact | Count | Source |
|---|---|---|
| tools-catalog.json | 250 | live tools/list |
| tools-from-source.json | 250 | ToolDefinition name fields in src |
| com-actions.json | 245 | switch cases in wps-com.ps1 |

Duplicate tool names in source: none.

## Why this baseline exists

- Every model request currently pays the full tool-schema cost; the MVP must shrink it and
  must be able to prove it did not grow again.
- MVP budget: standard toolset at most 40 tools and 22,000 schema bytes (the shipped list is 31 / 17,797).