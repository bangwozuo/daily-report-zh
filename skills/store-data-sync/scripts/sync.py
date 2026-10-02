# -*- coding: utf-8 -*-
"""
门店数据接入 —— 多源接入体检 + 字段映射校验 + 脱敏红线扫描 + 产物生成。

职责边界（重要）：
  本脚本只做**确定性校验与产物生成**：字段完整率、接入方式合规判定、每日到数覆盖率、
  个人信息红线扫描、Excel/PNG 落盘。
  **它不连接任何外部系统**（不调用平台 API、不登录后台、不爬取数据）——
  数据必须由用户自行从官方 API 或商家后台导出后喂进来。

四类校验：
  C1 接入方式合规  官方 API / 用户自行导出 = 放行；爬取、模拟登录、协议机器人 = 红线
  C2 字段映射完整  实际字段 ÷ 应到字段；缺哪个字段、走哪条降级路径
  C3 到数覆盖率    每日实际到数源数 ÷ 应到源数；断供天数显式列出
  C4 脱敏红线      手机号 / 身份证 / 银行卡 / 详细地址 / 邮箱 出现在字段里即报红线

用法：
  python sync.py --input input.json --outdir out
  python sync.py --demo --outdir out

产物：
  out/门店数据接入清单.xlsx   数据源清单 / 字段映射 / 数据质量 / 脱敏检查 / 汇总
  out/接入完整率.png          各数据源字段完整率柱状图
  out/到数趋势图.png          每日到数源数 vs 应到源数
  out/sync.json               机器可读结果（供工作流与下游技能读取）
"""
from __future__ import annotations

import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
REPO = os.path.dirname(os.path.dirname(SKILL_DIR))
sys.path.insert(0, os.path.join(REPO, "lib"))

try:
    import assettools as at
except ImportError:  # pragma: no cover
    print("[错误] 未找到 lib/assettools.py。请确认技能位于 <repo>/skills/<slug>/scripts/ 下。",
          file=sys.stderr)
    sys.exit(2)

at.need("openpyxl", "matplotlib")

# ---------------------------------------------------------------- 规则表

# C1：接入方式白名单 / 黑名单（黑名单命中即红线，不进入后续采集）
ACCESS_OK = ["官方API", "官方 API", "用户导出", "商家后台", "用户自行导出", "手工录入"]
ACCESS_RED = ["爬取", "抓取", "模拟登录", "协议机器人", "破解", "接码", "私域抓取"]

# 本员工需要的最小字段集（缺任一字段 → 下游日报/异常检测会失真）
REQUIRED_FIELDS = {
    "收银系统": ["date", "revenue", "orders", "visitors", "refunds"],
    "美团/点评商家后台": ["date", "团购核销", "评价数", "评分"],
    "抖音开放平台": ["date", "团购核销", "视频播放"],
    "企业微信": ["date", "客户标签"],
}

# C4：脱敏红线
PII_PATTERNS = [
    (r"^1[3-9]\d{9}$", "手机号", "《个人信息保护法》：手机号属个人敏感信息，一律不录入"),
    (r"^\d{17}[\dXx]$", "身份证号", "《个人信息保护法》：身份证号属个人敏感信息，一律不录入"),
    (r"^\d{16,19}$", "银行卡/支付卡号", "支付信息不得进入分析数据集"),
    (r"^[\w.\-]+@[\w\-]+\.[A-Za-z]{2,}$", "邮箱", "邮箱属个人信息，非必要不采集"),
    (r"(省|市|区|县|街道|路|号).{0,12}(室|栋|单元|楼)", "详细地址", "详细地址属个人信息，只保留商圈层级"),
    (r"^[\u4e00-\u9fa5]{2,4}$", "客户姓名/昵称", "客户姓名属个人信息，输出只保留代号（如 客户 A）"),
]

# 降级路径：缺字段时怎么办
DEGRADE = {
    "date": "无日期字段无法做趋势；导出时勾选「日期」列，或补一份后台日报截图人工录入",
    "revenue": "无营收字段改用团购核销金额 × 客单价估算，并在报告中标注「估算值」",
    "orders": "无订单数则客单价不可算，改报「客单价缺失」，不做推算",
    "visitors": "无访客数则转化率不可算，只报订单量与营收",
    "refunds": "无退款额则退款率不可算；改从平台评价与售后单里人工汇总",
    "团购核销": "无核销数据则线上引流效果不可评估，建议开启平台核销回传",
    "评价数": "无评价数则差评监控停摆，改为人工每日查看商家后台",
    "评分": "无评分则口碑趋势断档，改为每周人工记录一次",
    "视频播放": "无播放数据则内容效果不可评估，改用平台创作者后台导出",
    "客户标签": "无客户标签则私域分层停摆，改为人工维护客户 Excel",
}


# ---------------------------------------------------------------- 演示数据

def demo_data():
    """望京 SOHO 店 2026-09 的多源接入现状：含 1 条红线接入、1 处字段缺失、3 天断供、2 条脱敏违规。"""
    daily = []
    for day in range(1, 31):
        d = f"2026-09-{day:02d}"
        arrived = 4
        if day in (15, 16):        # 停电：收银数据未导出
            arrived = 3
        if day == 23:              # 抖音授权到期，未续期
            arrived = 2
        daily.append({"date": d, "arrived": arrived, "expected": 4})
    return {
        "store_name": "望京 SOHO 店",
        "period": {"start": "2026-09-01", "end": "2026-09-30"},
        "sources": [
            {
                "source": "收银系统", "access": "官方API", "authorized": True,
                "owner": "店长 王敏", "last_sync": "2026-09-30T21:32:00",
                "expected_fields": ["date", "revenue", "orders", "visitors", "refunds"],
                "actual_fields": ["date", "revenue", "orders", "visitors", "refunds"],
            },
            {
                "source": "美团/点评商家后台", "access": "用户导出", "authorized": True,
                "owner": "店长 王敏", "last_sync": "2026-09-30T21:40:00",
                "expected_fields": ["date", "团购核销", "评价数", "评分"],
                "actual_fields": ["date", "评价数", "评分"],
            },
            {
                "source": "抖音开放平台", "access": "官方API", "authorized": False,
                "owner": "运营 李强", "last_sync": "2026-09-22T20:00:00",
                "expected_fields": ["date", "团购核销", "视频播放"],
                "actual_fields": ["date"],
            },
            {
                "source": "企业微信", "access": "官方API", "authorized": True,
                "owner": "运营 李强", "last_sync": "2026-09-30T21:45:00",
                "expected_fields": ["date", "客户标签"],
                "actual_fields": ["date", "客户标签"],
            },
            {
                "source": "竞品门店客流", "access": "爬取", "authorized": False,
                "owner": "运营 李强", "last_sync": "2026-09-28T19:00:00",
                "expected_fields": ["date", "客流估算"],
                "actual_fields": ["date", "客流估算"],
            },
        ],
        "daily_arrival": daily,
        "raw_samples": [
            {"source": "企业微信", "field": "customer_phone", "value": "13800001111"},
            {"source": "企业微信", "field": "customer_id_card", "value": "110101199003074512"},
            {"source": "美团/点评商家后台", "field": "customer_name", "value": "张先生"},
            {"source": "收银系统", "field": "member_address", "value": "北京市朝阳区望京街道 12 号 3 栋 501 室"},
        ],
    }


# ---------------------------------------------------------------- 校验层

def check_sources(sources):
    rows, reds = [], []
    for s in sources:
        name = s.get("source", "未命名数据源")
        access = s.get("access", "")
        exp = list(s.get("expected_fields") or [])
        act = list(s.get("actual_fields") or [])
        hit = [f for f in exp if f in act]
        rate = (len(hit) / len(exp) * 100) if exp else 0.0
        missing = [f for f in exp if f not in act]

        is_red = any(k in access for k in ACCESS_RED)
        ok = any(k in access for k in ACCESS_OK) and not is_red
        if is_red:
            verdict, action = "🔴 红线：禁止接入", "立即停止该数据源，改用官方 API 或用户导出"
            reds.append(f"{name}：接入方式「{access}」属禁止路径")
        elif not s.get("authorized", True):
            verdict, action = "🟡 未授权", "先完成账号授权，未授权不得采集"
        elif missing:
            verdict, action = "🟡 字段不全", "按降级路径补齐，或接受降级并标注"
        else:
            verdict, action = "✅ 正常", "可进入下游流程"

        rows.append({
            "数据源": name, "接入方式": access,
            "授权状态": "已授权" if s.get("authorized") else "未授权",
            "责任人": s.get("owner", ""), "最近同步": s.get("last_sync", ""),
            "应到字段": len(exp), "实到字段": len(hit),
            "字段完整率%": round(rate, 1),
            "缺失字段": "、".join(missing) if missing else "—",
            "判定": verdict, "处理动作": action,
        })
    return rows, reds


def check_fields(sources):
    rows = []
    for s in sources:
        name = s.get("source", "")
        exp = list(s.get("expected_fields") or [])
        act = list(s.get("actual_fields") or [])
        for f in exp:
            rows.append({
                "数据源": name, "字段": f,
                "是否到数": "✅ 到数" if f in act else "❌ 缺失",
                "下游影响": _impact(f),
                "降级路径": DEGRADE.get(f, "先与平台确认该字段是否可导出；不可导出则人工记录"),
            })
    return rows


def _impact(f):
    return {
        "date": "无日期 → 趋势、环比全部不可算",
        "revenue": "无营收 → 日报核心指标缺失",
        "orders": "无订单 → 客单价不可算",
        "visitors": "无访客 → 转化率不可算",
        "refunds": "无退款额 → 退款率不可算",
        "团购核销": "无核销 → 线上引流效果不可评估",
        "评价数": "无评价数 → 差评监控停摆",
        "评分": "无评分 → 口碑趋势断档",
        "视频播放": "无播放 → 内容效果不可评估",
        "客户标签": "无标签 → 私域分层停摆",
    }.get(f, "下游指标可能失真")


def check_arrival(daily):
    rows, gaps = [], []
    for d in daily:
        exp = d.get("expected", 0)
        arr = d.get("arrived", 0)
        rate = (arr / exp * 100) if exp else 0.0
        status = "✅ 全到" if arr >= exp else ("🟡 部分到数" if arr > 0 else "🔴 全断")
        if arr < exp:
            gaps.append(d.get("date", ""))
        rows.append({
            "日期": d.get("date", ""), "应到源数": exp, "实到源数": arr,
            "覆盖率%": round(rate, 1), "状态": status,
        })
    return rows, gaps


def check_pii(raw_samples):
    rows, reds = [], []
    for smp in raw_samples or []:
        val = str(smp.get("value", ""))
        hits = [t for pat, t, _ in PII_PATTERNS if re.search(pat, val)]
        if hits:
            kind = hits[0]
            basis = next(b for p, t, b in PII_PATTERNS if t == kind)
            masked = _mask(val)
            rows.append({
                "数据源": smp.get("source", ""), "字段": smp.get("field", ""),
                "个人信息类型": kind, "脱敏后示例": masked, "判定": "🔴 红线：禁止录入",
                "依据": basis, "处理动作": "从数据集中删除该字段，只保留代号/分层标签",
            })
            reds.append(f"{smp.get('source')}.{smp.get('field')} 命中{kind}")
        else:
            rows.append({
                "数据源": smp.get("source", ""), "字段": smp.get("field", ""),
                "个人信息类型": "—", "脱敏后示例": _mask(val), "判定": "✅ 通过",
                "依据": "非个人信息字段，可保留", "处理动作": "无需处理",
            })
    return rows, reds


def _mask(v):
    if len(v) <= 4:
        return "*" * len(v)
    return v[:2] + "*" * (len(v) - 4) + v[-2:]


def summarize(src_rows, gaps, reds, pii_rows):
    ok = [r for r in src_rows if r["判定"].startswith("✅")]
    full_rates = [r["字段完整率%"] for r in src_rows if "爬取" not in r["接入方式"]]
    return {
        "门店": "",
        "数据源总数": len(src_rows),
        "正常数据源": len(ok),
        "红线数据源": len([r for r in src_rows if r["判定"].startswith("🔴")]),
        "未授权/字段不全": len([r for r in src_rows if r["判定"].startswith("🟡")]),
        "合规源平均字段完整率%": round(sum(full_rates) / len(full_rates), 1) if full_rates else 0.0,
        "断供天数": len(gaps),
        "断供日期": "、".join(gaps) if gaps else "—",
        "脱敏红线条数": len([r for r in pii_rows if r["判定"].startswith("🔴")]),
        "阻断项合计": len(reds),
    }


# ---------------------------------------------------------------- 输出层

def write_outputs(cur, src_rows, field_rows, arrival_rows, pii_rows, gaps, reds, outdir):
    at.ensure_outdir(outdir)
    files = []

    s = summarize(src_rows, gaps, reds, pii_rows)
    s["门店"] = cur.get("store_name", "")
    summ = [{"项": k, "内容": str(v)} for k, v in s.items()]

    xlsx = at.write_excel(
        os.path.join(outdir, "门店数据接入清单.xlsx"),
        {
            "数据源清单": src_rows or [{"数据源": "（无）"}],
            "字段映射": field_rows or [{"数据源": "（无）"}],
            "数据质量": arrival_rows or [{"日期": "（无）"}],
            "脱敏检查": pii_rows or [{"数据源": "（无）"}],
            "汇总": summ,
        },
        highlights={"数据源清单": {"判定": "contains:🔴", "字段完整率%": "<80"},
                    "脱敏检查": {"判定": "contains:🔴"},
                    "数据质量": {"状态": "contains:🔴"}},
        widths={"数据源清单": {"处理动作": 34, "缺失字段": 24},
                "字段映射": {"降级路径": 44, "下游影响": 26},
                "脱敏检查": {"依据": 44, "处理动作": 34}},
    )
    files.append(xlsx)

    ok_sources = [r for r in src_rows if "爬取" not in r["接入方式"]]
    png1 = at.bar_chart(
        os.path.join(outdir, "接入完整率.png"),
        [r["数据源"] for r in ok_sources],
        [r["字段完整率%"] for r in ok_sources],
        title=f"{cur.get('store_name', '')} 各数据源字段完整率（%）",
        xlabel="数据源", ylabel="字段完整率%", figsize=(9.4, 4.6),
    )
    files.append(png1)

    if arrival_rows:
        png2 = at.line_chart(
            os.path.join(outdir, "到数趋势图.png"),
            [r["日期"][5:] for r in arrival_rows],
            {"实到源数": [r["实到源数"] for r in arrival_rows],
             "应到源数": [r["应到源数"] for r in arrival_rows]},
            title=f"{cur.get('store_name', '')} 每日数据到数趋势",
            xlabel="日期", ylabel="数据源数量", figsize=(11, 4.4),
        )
        files.append(png2)

    js = at.write_json({
        "store_name": cur.get("store_name", ""),
        "period": cur.get("period", {}),
        "summary": s,
        "sources": src_rows,
        "field_mapping": field_rows,
        "daily_arrival": arrival_rows,
        "pii_check": pii_rows,
        "redlines": reds,
        "generated_at": at.stamp(),
        "note": "本脚本不连接任何外部系统；数据须由用户从官方 API 或商家后台导出后喂入",
    }, os.path.join(outdir, "sync.json"))
    files.append(js)
    return files, s


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser(description="门店数据接入体检（字段/合规/到数/脱敏）")
    ap.add_argument("--input", help="输入 JSON 路径")
    ap.add_argument("--outdir", default="out", help="输出目录")
    ap.add_argument("--demo", action="store_true", help="用内置演示数据运行")
    a = ap.parse_args()

    if a.demo:
        cur = demo_data()
    elif a.input:
        cur = at.read_json(a.input)
    else:
        ap.error("需提供 --input 或 --demo")

    if not cur.get("sources"):
        print("[错误] sources 为空：无数据源清单无法体检，请先补齐接入配置。", file=sys.stderr)
        sys.exit(3)

    src_rows, reds1 = check_sources(cur["sources"])
    field_rows = check_fields(cur["sources"])
    arrival_rows, gaps = check_arrival(cur.get("daily_arrival") or [])
    pii_rows, reds2 = check_pii(cur.get("raw_samples") or [])
    reds = reds1 + reds2

    files, s = write_outputs(cur, src_rows, field_rows, arrival_rows, pii_rows, gaps, reds, a.outdir)
    print(f"数据源 {s['数据源总数']} 个（正常 {s['正常数据源']}，红线 {s['红线数据源']}）"
          f"｜断供 {s['断供天数']} 天｜脱敏红线 {s['脱敏红线条数']} 条"
          f"｜合规源平均字段完整率 {s['合规源平均字段完整率%']}%")
    for f in files:
        print(" 产物:", f, f"({os.path.getsize(f) / 1024:.1f} KB)")
    at.emit({"status": "ok", "files": files, "summary": s, "redlines": reds})


if __name__ == "__main__":
    main()
