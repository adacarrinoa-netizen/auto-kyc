# 附件4 KYC extras 字段说明

供 Agent 调用 `scripts/att4_kyc_docx.py` 的 `build_att4_report(company, profile, extras, out_path)`。

## profile（财汇/启信合并）

| 键 | 含义 |
|----|------|
| name | 企业全称 |
| uscc | 统一社会信用代码 |
| legal_representative | 法定代表人 |
| registered_capital | 注册资本（可写双口径） |
| establishment_date | 成立日期 |
| registered_address | 注册地址 |
| enterprise_nature | 企业性质 |
| controller | 实际控制人 |
| latest_subject_rating / _organization / _date | 主体评级 |

## extras（研判与模块正文）

| 键 | 含义 |
|----|------|
| risk_level | 结论中的「低 / 一般关注 / 高」等 |
| judgment | 「结论：…」 |
| admission | 「是否建议准入」全文 |
| risk_basis | 认定依据段落 |
| negative_items | list[str]，负面事项①②③…；**必须含「公司负面舆情」分析**（生成器缺则自动补） |
| news_summary | 可选；结论区舆情一句研判，优先于自动拼接 |
| fact_summary / relevance / business_impact / controls | 研判四段 |
| control_path | 控制路径一句话 |
| shareholder_rows | [[股东, 比例, 数量, 性质], ...] |
| judicial_rows | [[日期, 案由, 身份, 案号, 进展], ...] |
| news_rows | [[日期, 标题, 要点, 来源], ...] |
| news_note | 第9节舆情正文；亦用于拼结论区舆情条目 |
| finance_rows | 年报指标行 |
| bond_note | 存量债与口径说明 |
| penalty_note / qixin_risk_note / executed_note / limit_consume_note / bankruptcy_note | 风险模块 |
| stock_code / is_listed / industry / scope / status | 身份补充 |

金样报告：`~/Downloads/KYC/上海城投控股股份有限公司_KYC报告.docx`
