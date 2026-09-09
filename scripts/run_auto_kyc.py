#!/usr/bin/env python3
"""自动 KYC 入口：财汇 MCP 公开查询 + 关联既有 KYC 报告 + 输出摘要 JSON。

供债权全流程平台「立项导入」调用，对应 ~/.cursor/skills/auto-kyc/SKILL.md
"""
from __future__ import annotations

import argparse
import json
import ssl
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any
import shutil

MCP_JSON = Path.home() / ".cursor" / "mcp.json"
KNOWN_KYC_DIRS = [
    Path.home() / "Downloads" / "KYC",
    Path.home() / "Downloads" / "0702自动化报告",
    Path.home() / "Downloads" / "0702自动化报告" / "credit-platform" / "data" / "uploads",
]

# Skill 成品目录：「2026三源核验版」落盘位置
SKILL_KYC_DIRS = [
    Path.home() / "Downloads" / "KYC",
    Path.home() / "Downloads" / "0702自动化报告",
]
SKILL_SANYUAN_MARKER = "三源核验"


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def load_caihui() -> tuple[str, str]:
    cfg = json.loads(MCP_JSON.read_text(encoding="utf-8"))
    srv = cfg["mcpServers"]["caihui_mcp"]
    return srv["url"], srv["headers"]["x-api-key"]


def caihui_rpc(method: str, params: dict, req_id: int = 1) -> dict:
    url, key = load_caihui()
    payload = {"jsonrpc": "2.0", "id": req_id, "method": method, "params": params}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    req.add_header("x-api-key", key)
    ctx = ssl._create_unverified_context()
    with urllib.request.urlopen(req, context=ctx, timeout=90) as resp:
        return json.loads(resp.read().decode("utf-8"))


def parse_execute_tool_result(rpc: dict) -> dict:
    content = (((rpc or {}).get("result") or {}).get("content") or [])
    if not content:
        return {"raw": rpc}
    text = content[0].get("text") or ""
    try:
        return json.loads(text)
    except Exception:
        return {"raw_text": text[:5000]}


def query_basic(companies: list[str]) -> dict:
    rpc = caihui_rpc(
        "tools/call",
        {
            "name": "execute_tool",
            "arguments": {
                "tool_name": "get_company_basic_info",
                "arguments": {
                    "target_company": companies,
                    "indicator_name": [
                        "企业名称",
                        "法定代表人",
                        "注册资本",
                        "成立日期",
                        "注册地址",
                        "企业性质",
                        "实际控制人",
                        "主体评级",
                    ],
                },
            },
        },
        req_id=2,
    )
    return parse_execute_tool_result(rpc)


_INDEX_ALIASES = {
    "企业名称": "name",
    "企业全称": "name",
    "法定代表人": "legal_representative",
    "注册资本(万元)": "registered_capital",
    "注册资本": "registered_capital",
    "成立日期": "establishment_date",
    "注册地址": "registered_address",
    "企业性质": "enterprise_nature",
    "实际控制人": "controller",
    "最新主体评级": "latest_subject_rating",
    "最新主体评级日期": "latest_subject_rating_date",
    "最新主体评级机构": "latest_subject_rating_organization",
    "统一社会信用代码": "uscc",
    "主体评级": "latest_subject_rating",
}


def records_to_profiles(payload: dict) -> list[dict]:
    """兼容财汇宽表与「企业名+指标名+指标值」长表两种返回。"""
    data = ((payload or {}).get("data") or {}).get("records") or {}
    heads = data.get("headInfo") or []
    rows = data.get("data") or []
    fields = [h.get("field") or h.get("name") for h in heads]
    if not rows:
        return []

    # 长表：company_name / index_name / index_value
    if set(fields) >= {"company_name", "index_name", "index_value"} or (
        len(fields) >= 3 and "指标" in str(heads)
    ):
        by_name: dict[str, dict] = {}
        for row in rows:
            item = {fields[i]: row[i] for i in range(min(len(fields), len(row)))}
            cname = item.get("company_name") or item.get("企业名称") or ""
            idx = item.get("index_name") or item.get("指标名") or ""
            val = item.get("index_value") or item.get("指标值")
            if not cname:
                continue
            prof = by_name.setdefault(
                cname,
                {
                    "name": cname,
                    "legal_representative": None,
                    "registered_capital": None,
                    "establishment_date": None,
                    "registered_address": None,
                    "enterprise_nature": None,
                    "controller": None,
                    "latest_subject_rating": None,
                    "latest_subject_rating_date": None,
                    "latest_subject_rating_organization": None,
                    "uscc": None,
                    "raw": {},
                },
            )
            key = _INDEX_ALIASES.get(str(idx).strip())
            if key == "registered_capital" and val not in (None, ""):
                # 保留万元数值，生成器侧可再格式化
                prof[key] = f"{val}万元" if "万" not in str(val) else val
            elif key and val not in (None, ""):
                # name 已有全称时不覆盖
                if key == "name" and prof.get("name"):
                    pass
                else:
                    prof[key] = val
            prof["raw"][str(idx)] = val
        return list(by_name.values())

    out = []
    for row in rows:
        item = {}
        for i, f in enumerate(fields):
            if i < len(row):
                item[f] = row[i]
        out.append(
            {
                "name": item.get("company_name") or item.get("企业名称"),
                "legal_representative": item.get("legal_representative"),
                "registered_capital": item.get("registered_capital")
                or item.get("expansion_registered_capital"),
                "establishment_date": item.get("establishment_date"),
                "registered_address": item.get("registered_address"),
                "enterprise_nature": item.get("enterprise_nature"),
                "controller": item.get("enterprise_ultimate_controller")
                or item.get("controller"),
                "latest_subject_rating": item.get("latest_subject_rating"),
                "latest_subject_rating_date": item.get("latest_subject_rating_date"),
                "latest_subject_rating_organization": item.get(
                    "latest_subject_rating_organization"
                ),
                "uscc": item.get("unified_social_credit_code") or item.get("uscc"),
                "raw": item,
            }
        )
    return out


def find_skill_sanyuan_kyc(company: str | None = None) -> list[Path]:
    """查找本主体已落盘的 Skill 三源核验版 KYC（不挂无关样例）。"""
    hits: list[Path] = []
    if not company:
        return hits
    key = company.replace("有限公司", "").replace("股份有限公司", "").replace("股份", "")[:8]
    for d in SKILL_KYC_DIRS:
        if not d.exists():
            continue
        for f in list(d.glob(f"*{SKILL_SANYUAN_MARKER}*.docx")):
            if not f.is_file():
                continue
            name = f.name
            if any(x in name for x in ("指引", "议案", "培训", "整改", "原始文件", "拟自动化实现")):
                continue
            # 仅匹配当前主体，禁止挂入上海城投等无关样例
            if company in name or (key and key in name):
                hits.append(f)
    by: dict[str, Path] = {}
    for h in hits:
        try:
            s = str(h.resolve())
        except OSError:
            s = str(h)
        by[s] = h
    return sorted(by.values(), key=lambda p: -p.stat().st_mtime)


def find_existing_kyc(company: str) -> list[Path]:
    hits = []
    key = company.replace("有限公司", "").replace("股份", "")[:8]
    for d in KNOWN_KYC_DIRS:
        if not d.exists():
            continue
        patterns = ["*KYC*.docx", "*kyc*.docx", "*KYC*.pdf", "*kyc*.pdf"]
        globs = []
        for pat in patterns:
            globs.extend(d.glob(pat))
            globs.extend(d.rglob(pat))
        for f in globs:
            if not f.is_file():
                continue
            if SKILL_SANYUAN_MARKER in f.name or "自动生成" in f.name:
                continue  # Skill 成品走 find_skill_sanyuan_kyc
            if company in f.name or key in f.name:
                hits.append(f)
    by_key: dict[str, Path] = {}
    for h in hits:
        try:
            s = str(h.resolve())
        except OSError:
            s = str(h)
        prev = by_key.get(s)
        if prev is None or h.stat().st_mtime >= prev.stat().st_mtime:
            by_key[s] = h

    def prefer_rank(p: Path) -> tuple:
        s = str(p).lower()
        if "onedrive" in s or "/desktop/" in s:
            tier = 0
        elif "/downloads/kyc" in s:
            tier = 1
        elif "/credit-platform/data/uploads" in s:
            tier = 3
        else:
            tier = 2
        return (tier, -p.stat().st_mtime)

    return sorted(by_key.values(), key=prefer_rank)


def build_conclusion(
    companies: list[str],
    profiles: list[dict],
    existing_arts: list[dict],
    skill_arts: list[dict] | None = None,
) -> dict:
    """按2026指引公开KYC维度形成简要结论（不替代征信/中登/业务审批）。"""
    points = []
    risk = "低风险"
    for p in profiles:
        name = p.get("name") or "—"
        nature = p.get("enterprise_nature") or ""
        rating = p.get("latest_subject_rating")
        ctrl = p.get("controller") or "—"
        if "国有" in str(nature):
            points.append(f"{name}：企业性质为{nature}，实控人/控制关系为「{ctrl}」，公开身份维度未见直接高风险认定情形。")
        else:
            points.append(f"{name}：企业性质为{nature or '待核'}，须结合非公开材料与征信进一步判断。")
            if risk == "低风险":
                risk = "一般关注"
        if rating:
            org = p.get("latest_subject_rating_organization") or ""
            dt = p.get("latest_subject_rating_date") or ""
            points.append(f"{name}：公开主体评级 {rating}（{org} {dt}）。".strip())
        else:
            points.append(f"{name}：公开主体评级暂缺，建议以跟踪评级/债券披露或征信续核。")
    n_skill = len(skill_arts or [])
    n_exist = len(existing_arts or [])
    if n_skill:
        points.append(f"已挂载/生成 Skill KYC 报告 {n_skill} 份（默认附件4结论置前结构）。")
    if n_exist:
        points.append(f"已关联客户经理既有 KYC 报告 {n_exist} 份，可与 Skill 附件4版交叉核对。")
    if not n_skill and not n_exist:
        points.append("未生成/关联到 KYC Word 报告，建议重新执行自动 KYC。")
        if risk == "低风险":
            risk = "一般关注"

    overall_map = {
        "低风险": "公开KYC维度建议：低风险客户＋简易程序管理；附条件准入（须完成征信/中登/决议等）。",
        "一般关注": "公开KYC维度建议：一般关注；准入前补齐评级/报告与非公开核验，再进入尽调与审批。",
        "高风险": "公开KYC维度提示高风险关注事项；暂停推进并按指引强化审查。",
    }
    return {
        "risk_level": risk,
        "overall": overall_map.get(risk, overall_map["一般关注"]),
        "points": points,
        "next_steps": [
            "进入实体尽调：导入非公开材料与录音转写",
            "放款前补充征信、中登登记与有效决议/保证合同",
            "如需三源债券/财务细核，按 auto-kyc skill 调用 iFinD / Wind 续核",
        ],
        "disclaimer": "本结论仅基于公开数据自动KYC，不替代正式法律意见、征信、审计或业务审批判断。",
    }


def profile_for_company(profiles: list[dict], company: str) -> dict | None:
    for p in profiles:
        n = p.get("name") or ""
        if n == company or company in n or n in company:
            return p
    return None


def _clip(s: Any, n: int = 180) -> str:
    t = " ".join(str(s or "").split())
    return t if len(t) <= n else t[: n - 1] + "…"


def _extract_blob(obj: Any) -> str:
    if obj is None:
        return ""
    if isinstance(obj, str):
        return obj
    if isinstance(obj, dict):
        for k in ("answer", "text", "msg", "message", "data", "content"):
            if k in obj:
                v = obj[k]
                if isinstance(v, str) and v.strip():
                    return v
                got = _extract_blob(v)
                if got:
                    return got
        parts = [_extract_blob(v) for v in obj.values()]
        return "\n".join(p for p in parts if p)
    if isinstance(obj, list):
        return "\n".join(_extract_blob(x) for x in obj if _extract_blob(x))
    return str(obj)


def fetch_kyc_extras(company: str) -> dict:
    """尽力拉取司法/处罚摘要，写入三源核验版第六章。"""
    extras: dict[str, Any] = {
        "judicial_rows": [],
        "penalty_note": "财汇主体监管处罚检索未返回可结构化明细；归档日须保留信用中国等官方留痕。",
        "shareholder_rows": [],
        "control_path": "",
        "finance_note": "财务三源细核（iFinD/Wind）可在 Agent 侧按 skill 续核后补表；结构对齐上海城投三源核验版第五章。",
        "industry_note": "公开工商信息已采集；实质行业与主营构成建议以 iFinD/Wind 及申报书业务描述交叉认定。",
        "news_note": "舆情细核可结合财汇新闻/iFinD 公告检索在 Agent 侧续补。",
        "negative_rows": [],
    }
    # 司法
    for tool_name in (
        "get_company_lawsuit_info",
        "search_company_judicial_cases",
        "get_company_litigation_info",
    ):
        try:
            rpc = caihui_rpc(
                "tools/call",
                {
                    "name": "execute_tool",
                    "arguments": {
                        "tool_name": tool_name,
                        "arguments": {"target_company": [company], "limit": 20},
                    },
                },
                req_id=20,
            )
            parsed = parse_execute_tool_result(rpc)
            blob = _clip(_extract_blob(parsed), 400)
            if not blob or "不存在" in blob or "Unknown" in blob:
                continue
            # 尝试 records 表
            data = ((parsed or {}).get("data") or {}).get("records") or {}
            rows = data.get("data") or []
            heads = data.get("headInfo") or []
            fields = [h.get("field") or h.get("name") for h in heads]
            if rows and fields:
                for row in rows[:10]:
                    item = {fields[i]: row[i] for i in range(min(len(fields), len(row)))}
                    extras["judicial_rows"].append(
                        [
                            str(item.get("case_date") or item.get("date") or item.get("latest_process_date") or "—")[:16],
                            _clip(item.get("cause") or item.get("case_reason") or item.get("title") or blob, 40),
                            str(item.get("party_role") or item.get("role") or "—")[:8],
                            _clip(item.get("case_no") or item.get("case_number") or "", 36),
                        ]
                    )
            elif blob:
                extras["judicial_rows"].append(["—", _clip(blob, 60), "—", tool_name])
            if extras["judicial_rows"]:
                break
        except Exception:  # noqa: BLE001
            continue
    # 处罚
    for tool_name in ("get_company_penalty_info", "search_company_administrative_penalties"):
        try:
            rpc = caihui_rpc(
                "tools/call",
                {
                    "name": "execute_tool",
                    "arguments": {
                        "tool_name": tool_name,
                        "arguments": {"target_company": [company], "limit": 20},
                    },
                },
                req_id=21,
            )
            parsed = parse_execute_tool_result(rpc)
            blob = _clip(_extract_blob(parsed), 280)
            if blob and "不存在" not in blob and "Unknown" not in blob:
                extras["penalty_note"] = f"财汇（{tool_name}）：{_clip(blob, 220)}"
                break
        except Exception:  # noqa: BLE001
            continue
    return extras


def run(companies: list[str], outdir: Path | None, style: str = "att4") -> dict:
    """style: att4=附件4结论置前（默认）；sanyuan=多章三源核验版（仅显式要求时）。"""
    companies = [c.strip() for c in companies if c and c.strip()]
    style = (style or "att4").lower().strip()
    if style not in ("att4", "sanyuan", "both"):
        style = "att4"
    summary = {
        "skill": "auto-kyc",
        "guideline": "KYC工作指引（2026年修订版）",
        "report_style": style,
        "sources": [
            "启信慧眼MCP（Agent侧）",
            "财汇MCP",
            "附件4结构KYC（茅台租赁默认）",
            "客户经理既有KYC报告",
            "iFinD/Wind可在Agent侧续核",
        ],
        "checked_at": now(),
        "companies": companies,
        "profiles": [],
        "artifacts": [],
        "conclusion": {},
        "notes": [],
        "ok": False,
    }
    try:
        payload = query_basic(companies)
        profiles = records_to_profiles(payload)
        if not profiles and payload.get("status", {}).get("message") == "SUCCESS":
            summary["notes"].append("财汇返回成功但未解析出明细，请查看 raw")
            summary["raw_payload_status"] = payload.get("status")
        summary["profiles"] = profiles
        summary["caihui_status"] = (payload.get("status") or {}).get("message") or "OK"
        summary["ok"] = True
    except Exception as exc:  # noqa: BLE001
        summary["notes"].append(f"财汇MCP查询失败：{exc}")
        summary["ok"] = False

    existing_arts: list[dict] = []
    for c in companies:
        files = find_existing_kyc(c)
        seen_names: set[str] = set()
        for f in files:
            if f.name in seen_names:
                continue
            seen_names.add(f.name)
            existing_arts.append(
                {
                    "company": c,
                    "path": str(f),
                    "type": "existing_kyc_report",
                    "note": "客户经理既有KYC报告",
                }
            )
            if sum(1 for a in existing_arts if a["company"] == c) >= 2:
                break

    skill_arts: list[dict] = []
    script_dir = str(Path(__file__).resolve().parent)
    if script_dir not in __import__("sys").path:
        __import__("sys").path.insert(0, script_dir)

    # 生成本次 KYC：默认附件4；仅 style 含 sanyuan 时才出多章版
    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)
        for c in companies:
            prof = profile_for_company(summary["profiles"], c)
            safe = (prof.get("name") if prof else c) or c
            safe = "".join(ch for ch in safe if ch not in '\\/:*?"<>|')
            extras = fetch_kyc_extras(c)
            if prof and prof.get("controller"):
                extras["control_path"] = f"公开控制关系：{prof.get('controller')}"
            if prof and prof.get("uscc"):
                extras.setdefault("uscc", prof["uscc"])

            if style in ("att4", "both"):
                try:
                    from att4_kyc_docx import build_att4_report

                    dest = outdir / f"{safe}_KYC报告.docx"
                    build_att4_report(c, prof, extras, dest, summary["checked_at"])
                    skill_arts.append(
                        {
                            "company": c,
                            "path": str(dest),
                            "copied_to": str(dest),
                            "type": "skill_generated_kyc_report",
                            "note": "Skill附件4版KYC（茅台租赁默认）",
                            "source": "generated_att4",
                        }
                    )
                except Exception as exc:  # noqa: BLE001
                    summary["notes"].append(f"生成附件4版失败（{c}）：{exc}")

            if style in ("sanyuan", "both"):
                try:
                    from sanyuan_kyc_docx import build_sanyuan_report

                    dest = outdir / f"{safe}_KYC报告_2026三源核验版.docx"
                    build_sanyuan_report(c, prof, extras, dest, summary["checked_at"])
                    skill_arts.append(
                        {
                            "company": c,
                            "path": str(dest),
                            "copied_to": str(dest),
                            "type": "skill_generated_kyc_report",
                            "note": "Skill三源核验版KYC（仅显式要求时）",
                            "source": "generated_sanyuan",
                        }
                    )
                except Exception as exc:  # noqa: BLE001
                    summary["notes"].append(f"生成三源核验版失败（{c}）：{exc}")

    # 可选关联既有附件4/三源成品（不覆盖本次生成）
    seen_skill: set[str] = set(a.get("path", "") for a in skill_arts)
    for c in companies:
        # 优先挂本主体非「三源核验」命名的 KYC
        for f in find_existing_kyc(c):
            if "三源核验" in f.name:
                continue
            key = str(f.resolve()) if f.exists() else str(f)
            if key in seen_skill or any(a.get("path") == str(f) for a in existing_arts):
                continue
            # 已在 existing_arts
            break
        if style in ("sanyuan", "both"):
            for f in find_skill_sanyuan_kyc(c):
                key = str(f.resolve()) if f.exists() else str(f)
                if key in seen_skill:
                    continue
                seen_skill.add(key)
                skill_arts.append(
                    {
                        "company": c,
                        "path": str(f),
                        "type": "skill_generated_kyc_report",
                        "note": "Skill三源核验版KYC（既有成品）",
                        "source": "prior_skill_output",
                    }
                )

    summary["conclusion"] = build_conclusion(
        companies, summary["profiles"], existing_arts, skill_arts
    )
    summary["conclusion"]["checked_at"] = summary["checked_at"]
    summary["artifacts"] = skill_arts + existing_arts

    if outdir:
        outdir.mkdir(parents=True, exist_ok=True)
        # 拷贝既有 Skill 成品与客户经理报告到 outdir
        for art in skill_arts + existing_arts:
            if art.get("copied_to"):
                continue
            src = Path(art["path"])
            if not src.exists():
                continue
            dest = outdir / src.name
            if not dest.exists():
                shutil.copy2(src, dest)
            art["copied_to"] = str(dest)
        (outdir / "kyc_summary.json").write_text(
            json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        lines = [f"# 自动KYC摘要（{summary['checked_at']}）", f"指引：{summary['guideline']}", ""]
        conc = summary["conclusion"]
        lines += [
            "## 结论",
            f"- 风险等级（公开KYC）：{conc.get('risk_level')}",
            f"- 总体建议：{conc.get('overall')}",
            "",
        ]
        for pt in conc.get("points") or []:
            lines.append(f"- {pt}")
        lines.append("")
        for p in summary["profiles"]:
            lines += [
                f"## {p.get('name')}",
                f"- 法定代表人：{p.get('legal_representative')}",
                f"- 注册资本：{p.get('registered_capital')}",
                f"- 成立日期：{p.get('establishment_date')}",
                f"- 注册地址：{p.get('registered_address')}",
                f"- 企业性质：{p.get('enterprise_nature')}",
                f"- 实际控制人：{p.get('controller')}",
                f"- 主体评级：{p.get('latest_subject_rating') or '见财汇/跟踪评级续核'} "
                f"{p.get('latest_subject_rating_organization') or ''} {p.get('latest_subject_rating_date') or ''}",
                "",
            ]
        if skill_arts:
            lines.append("## Skill生成KYC")
            for a in skill_arts:
                lines.append(f"- [{a.get('note')}] {a['company']}：{a.get('copied_to') or a['path']}")
        if existing_arts:
            lines.append("## 客户经理既有KYC报告")
            for a in existing_arts:
                lines.append(f"- {a['company']}：{a.get('copied_to') or a['path']}")
        lines.append("")
        lines.append(conc.get("disclaimer") or "")
        (outdir / "kyc_summary.md").write_text("\n".join(lines), encoding="utf-8")

    return summary


def main() -> None:
    ap = argparse.ArgumentParser(
        description="自动KYC（默认附件4结论置前版；可选三源核验多章版）"
    )
    ap.add_argument("--companies", required=True, help='JSON数组，如 ["公司A","公司B"]')
    ap.add_argument("--outdir", default="", help="输出目录")
    ap.add_argument(
        "--style",
        default="att4",
        choices=["att4", "sanyuan", "both"],
        help="报告样式：att4=附件4默认；sanyuan=多章三源核验版；both=两份都出",
    )
    args = ap.parse_args()
    companies = json.loads(args.companies)
    outdir = Path(args.outdir) if args.outdir else None
    result = run(companies, outdir, style=args.style)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
