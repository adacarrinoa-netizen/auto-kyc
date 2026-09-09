---
name: auto-kyc
description: >-
  按《KYC工作指引（2026年修订版）》附件4结构自动生成茅台租赁KYC报告（结论置前）：
  启信慧眼MCP、财汇/预警通MCP、同花顺iFinD、万得Wind交叉核验。
  用于立项导入KYC、客户身份尽职调查、附件4版KYC报告撰写。
  Use when user mentions 自动KYC、KYC报告、KYC工作指引、附件4、立项KYC、茅台租赁KYC.
---

# 自动 KYC（2026 指引 · 附件4 默认版）

茅台租赁公开信息 KYC。**默认输出《指引》附件4结构**（结论置前、模块编号正文），**不要**默认使用多章「三源核验版 / Codex 样式」。

| 数据源 | Skill / 入口 |
|--------|----------------|
| 启信慧眼 | MCP `user-qixin`（工商、实控人、失信/被执行、处罚、诉讼概览） |
| 财汇企业预警通 | `caihui-mcp` / MCP `user-caihui_mcp` |
| 同花顺 iFinD | `ifind-finance-data` |
| 万得 Wind | `wind-mcp-skill` |

金样：`~/Downloads/KYC/上海城投控股股份有限公司_KYC报告.docx`  
生成器：`scripts/att4_kyc_docx.py`（默认）· `scripts/sanyuan_kyc_docx.py`（仅用户明确要多章版时）

## 何时使用

- 债权平台「立项导入」执行自动 KYC
- 用户要求按 2026 指引 / 附件4 出具 KYC
- 工商、实控人、评级、司法、发债交叉核对

## 报告结构（必须遵守）

按附件4顺序撰写，**1、KYC综合结论必须在最前**：

1. KYC综合结论（风险级别 / 负面事项汇总 / 风险影响研判 / 管控建议 / 是否建议准入）
2. 客户相关证照信息与行业认定
3. 人行征信信息（公开数据通常待补）
4. 基本信息（工商、实控人、股东、工商变更）
5. 企业融资（财务、租赁/应收中登窗、发债与存量债、对外担保）
6. 经营预警
7. 监管处罚
8. 司法诉讼
9. 公司负面舆情
10. 国家企业信用信息公示系统、信用中国等
11. 发债企业重大信息公告
12. 上市公司重大信息公告
13. 对客户有影响力的自然人

负面事项与风险影响分析**合并写在结论区**，不要每个模块单独拉一条「风险影响分析」。

**结论区「（2）负面事项汇总」硬性要求**：除司法/处罚、启信风险、评级、征信中登外，**必须单列「公司负面舆情」分析**（可写“近窗未见重大负面/命中摘要+研判”）。不得只把舆情留在第9节正文。Agent 填写 `negative_items` 时须含舆情条目；同时填 `news_rows` / `news_note`（或 `news_summary`），生成器会在缺失时自动补一条。

## Agent 工作流程

1. 确认客户全称（承租人/担保人可并行）
2. **启信慧眼**（若已接入）：首跳 `qixin_insight_lookup(path="/instructions")` + `enterprise_resolve`；再查照面、实控人、失信/被执行、异常/严重违法/处罚、诉讼概览
3. **财汇**：`get_company_basic_info`、司法、评级、存量债、工商变更、**负面舆情（search_news）** 等 → `execute_tool`
4. **iFinD / Wind**：股东、财务、存续债交叉（**历史发行只数 ≠ 余额>0 存量只数**；Wind 宽口径可能含子公司 ABS，须双口径披露）；可交叉公告/新闻
5. 对照《指引》附件1判断是否直接高风险；国有/上市等可作低风险身份认定，信用仍须实质分析
6. 用附件4结构输出 `{主体}_KYC报告.docx`；口径不一致必须写明，不得静默取一侧；**结论负面事项汇总须含舆情分析**
7. 脚本入口（债权平台 / 批量）：

```bash
python3 ~/.cursor/skills/auto-kyc/scripts/run_auto_kyc.py \
  --companies '["上海城投控股股份有限公司"]' \
  --outdir ~/Downloads/KYC/out \
  --style att4
```

`--style`：`att4`（默认）| `sanyuan`（多章版）| `both`

stdout 为 JSON 摘要；`--outdir` 写入附件4 docx、`kyc_summary.json` / `kyc_summary.md`。

## Agent 富化（推荐）

脚本侧主要拉财汇基础字段与司法/处罚摘要。完整「城投样例」质量报告时，Agent 应把启信/iFinD/Wind 结果填入 `att4_kyc_docx.build_att4_report(..., extras={...})`，关键 extras 键：

- `negative_items`（**必须含公司负面舆情一条**）/ `fact_summary` / `relevance` / `business_impact` / `controls` / `admission` / `judgment` / `risk_level`
- `shareholder_rows` / `judicial_rows` / `news_rows` / `news_note` / `news_summary` / `finance_rows`
- `bond_note` / `control_path` / `penalty_note` / `qixin_risk_note` / `executed_note`

## 铁律

- 禁止编造工商、司法、评级、债券只数、舆情条数
- 三源不一致必须写明口径，不得静默取一侧
- **1、KYC综合结论 → 负面事项汇总必须包含公司负面舆情分析**（可与第9节呼应，不得省略）
- 公开 KYC 不能替代征信、中登、合同与决议核验
- **默认禁止**输出多章「2026三源核验版」样式，除非用户明确要求
