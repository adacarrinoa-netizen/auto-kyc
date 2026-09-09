# auto-kyc

按《KYC工作指引（2026年修订版）》**附件4**结构，自动生成茅台租赁公开信息 KYC 报告（结论置前）。

## 安装（Cursor）

将本目录复制到 Cursor skills 路径：

```bash
cp -R . ~/.cursor/skills/auto-kyc
```

依赖：

```bash
pip3 install python-docx
```

数据源需在 Cursor 中配置对应 MCP / Skill：

| 数据源 | 入口 |
|--------|------|
| 启信慧眼 | MCP `user-qixin` |
| 财汇企业预警通 | MCP `user-caihui_mcp` |
| 同花顺 iFinD | `ifind-finance-data` skill |
| 万得 Wind | `wind-mcp-skill` |

> 密钥只放在本机 `~/.cursor/mcp.json`，**不要**提交到仓库。

## 快速使用

Agent 对话中提及「自动KYC / KYC报告 / 附件4」即可触发本 skill。

批量脚本入口：

```bash
python3 ~/.cursor/skills/auto-kyc/scripts/run_auto_kyc.py \
  --companies '["上海城投控股股份有限公司"]' \
  --outdir ~/Downloads/KYC/out \
  --style att4
```

`--style`：`att4`（默认）| `sanyuan` | `both`

## 目录

- `SKILL.md` — Agent 工作流与铁律
- `reference.md` — 参考说明
- `scripts/att4_kyc_docx.py` — 附件4 Word 生成器
- `scripts/run_auto_kyc.py` — 财汇拉取 + 报告落盘入口
- `scripts/sanyuan_kyc_docx.py` — 多章三源版（仅明确要求时使用）

## 注意

- **不包含**任何 KYC 报告 Word/PDF 成品；仓库仅含 skill 与生成脚本（`*.docx` 已在 `.gitignore`）
- 公开 KYC 不能替代人行征信、中登、合同与决议核验
- 禁止编造工商/司法/评级/债券只数；多源不一致须写明口径
