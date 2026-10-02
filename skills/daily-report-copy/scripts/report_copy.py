# -*- coding: utf-8 -*-
"""
日报文案生成 —— 经营数据 → Word 日报 + 数据摘要 Excel。

职责边界：
  数值计算、健康区间判定、关键数挑选、异常点排查方向映射由本脚本完成（机器的强项，
  数值以脚本输出为准，不要自己算）；文案润色与归因的业务解释由模型按 prompt.txt 补充。

日报硬结构（与 prompt.txt 一致，结论前置）：
  一、一句话结论（方向 + 幅度 + 归因）
  二、核心指标表（关键数 ≤ 3 个，每个带参照物）
  三、异常点与排查方向（异常必须给排查方向，不写"建议关注"空话）
  四、明日建议（责任人 + 动作 + 范围）

判定基准（量化）：
  转化率 健康 2%~5% / 需关注 1.5%~2% / 需当天处理 < 1.5%
  退款率 健康 < 3% / 需关注 3%~5% / 需当天处理 > 5%
  客单价波动 ±15% 内正常；营收波动 ±10% 一句话带过、10%~20% 提方向、> 20% 必须点出并给排查方向
  单日订单 < 20 单：转化率与退款率只报值不下结论
  特殊期（促销/节假日/新店 < 30 天）：写"活动带动"，不写"异常"

用法：
  python report_copy.py --input examples/input.json --outdir out
  python report_copy.py --demo                # 内置真实样例（含 2 个异常点）

产物：
  out/经营日报.docx       四段式 Word 日报（结论前置）
  out/日报数据摘要.xlsx   核心指标 / 异常清单 / 汇总
  out/report_copy.json    机器可读结果（供智能体读取）
"""
from __future__ import annotations

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
REPO = os.path.dirname(os.path.dirname(SKILL_DIR))
sys.path.insert(0, os.path.join(REPO, "lib"))

try:
    import assettools as at
except ImportError:  # pragma: no cover
    print("[错误] 未找到 lib/assettools.py。请确认技能位于 <repo>/skills/<slug>/scripts/ 下，"
          "且 <repo>/lib/assettools.py 存在。", file=sys.stderr)
    sys.exit(2)

# ---------------------------------------------------------------- 基准（与 prompt.txt 同源）

CONV_HEALTHY = (0.02, 0.05)      # 转化率健康区间
CONV_WATCH = 0.015               # < 1.5% 需当天处理
REFUND_HEALTHY = 0.03            # 退款率健康线
REFUND_BAD = 0.05                # > 5% 需当天处理
ATV_BAND = 0.15                  # 客单价波动带
REV_NORMAL = 0.10                # 营收波动 ±10% 内正常
REV_MENTION = 0.20               # > 20% 必须点出并给排查方向
MIN_ORDERS = 20                  # 小样本线：转化率/退款率不下结论
NEW_STORE_DAYS = 30              # 新店豁免期
MAX_KEY_METRICS = 3              # 关键数上限

# 异常指标 → 排查方向（异常点必须给排查方向，不写空话）
INVESTIGATE = {
    "营收": "按小时拆分营收定位掉量时段，先看外卖曝光与进店客流哪个先掉，再决定调流量还是调活动",
    "订单": "核对渠道订单明细，区分平台限流、库存售罄、出餐速度拖慢三类原因，逐项排除",
    "转化率": "检查流量构成（新客占比与广告渠道质量）与商品页（价格、库存、差评新增），先留住有效流量",
    "退款率": "逐单核对当日退款原因，按出餐慢 / 配送超时 / 质量问题分桶，占 60% 以上的桶先处理",
    "客单价": "对比套餐与满减结构，看高单价商品动销是否下滑、满减门槛是否挡住了加购",
    "访客": "核对平台曝光位变化与竞对活动，确认是流量入口减少还是自然流量波动",
}

DEMO = {
    "period": "2026-09-12",
    "store_name": "望京 SOHO 店",
    "category": "餐饮",
    "metrics": {
        "date": "2026-09-12", "weekday": "周六",
        "revenue": 6895, "orders": 124, "visitors": 5957,
        "refunds": 276, "last_year_revenue": 8120,
    },
    "compare": {
        "上周同星期（2026-09-05 周六）": {"revenue": 11276, "orders": 174, "visitors": 5762, "refunds": 278},
        "同星期均值（9 月全部周六）": {"revenue": 10474},
    },
    "special_days": [],
    "store_age_days": 512,
    "alerts": [
        {"日期": "2026-09-12", "星期": "周六", "指标": "营收", "规则": "R2 硬阈值", "级别": "高",
         "实际值": "¥6,895", "阈值/基线": "环比 ≤ -20%",
         "证据": "上周六 ¥11,276，星期对齐偏离 -34.2%", "判定": "确认异常",
         "星期对齐偏离": "-34.2%", "周环比": "—"},
        {"日期": "2026-09-12", "星期": "周六", "指标": "退款率", "规则": "R2 硬阈值", "级别": "中",
         "实际值": "4.00%", "阈值/基线": "健康线 < 3%",
         "证据": "退款 ¥276 ÷ 营收 ¥6,895 = 4.00%", "判定": "确认异常",
         "星期对齐偏离": "—", "周环比": "—"},
    ],
    "notes": "09-12 无促销、非节假日，门店开业已 512 天",
}


# ---------------------------------------------------------------- 计算

def judge_conv(v):
    if v is None:
        return "数据缺失"
    if v >= CONV_HEALTHY[0] and v <= CONV_HEALTHY[1]:
        return "健康"
    if v >= CONV_WATCH:
        return "需关注"
    return "需当天处理"


def judge_refund(v):
    if v is None:
        return "数据缺失"
    if v < REFUND_HEALTHY:
        return "健康"
    if v <= REFUND_BAD:
        return "需关注"
    return "需当天处理"


def fmt_money(v):
    return f"¥{v:,.0f}" if isinstance(v, (int, float)) else "—"


def fmt_pct(v):
    return f"{v:.2%}" if isinstance(v, (int, float)) else "—"


def fmt_money2(v):
    return f"¥{v:,.2f}" if isinstance(v, (int, float)) else "—"


def weekday_aligned(compare, weekday):
    """在 compare 中找与本期同星期几的口径（键含星期几），无则退化为任意口径。"""
    for k, v in compare.items():
        if weekday and weekday in k:
            return k, v
    for k, v in compare.items():
        return k, v
    return None, {}


def pick_key_metrics(rows):
    """关键数挑选（与 prompt.txt 优先级同源，上限 3 个）。"""
    by = {r["指标"]: r for r in rows}
    picked = []
    if "营收" in by:
        picked.append(by["营收"])
    r = by.get("退款率")
    if r and "健康" not in r["判定"] and "缺失" not in r["判定"]:
        picked.append(r)
    c = by.get("转化率")
    if c and c["判定"] in ("需关注", "需当天处理"):
        picked.append(c)
    a = by.get("客单价")
    if a and "波动带内" not in a["判定"] and len(picked) < MAX_KEY_METRICS:
        picked.append(a)
    for name in ("订单数", "访客数"):
        if len(picked) >= MAX_KEY_METRICS:
            break
        if name in by and all(x["指标"] != name for x in picked):
            picked.append(by[name])
    return picked[:MAX_KEY_METRICS]


def build(payload, outdir):
    m = payload.get("metrics", {})
    compare = payload.get("compare", {}) or {}
    special_days = payload.get("special_days", []) or []
    alerts = [a for a in (payload.get("alerts", []) or [])
              if "确认异常" in str(a.get("判定", ""))]
    store = payload.get("store_name", "门店")
    period = payload.get("period") or m.get("date", "")
    weekday = m.get("weekday", "")
    category = payload.get("category", "通用")
    small_sample = isinstance(m.get("orders"), (int, float)) and m["orders"] < MIN_ORDERS
    special = (m.get("date") in special_days) or payload.get("store_age_days", 999) < NEW_STORE_DAYS

    # ---- 派生指标重算（数值以脚本为准）----
    revenue, orders = m.get("revenue"), m.get("orders")
    visitors, refunds = m.get("visitors"), m.get("refunds")
    conv = orders / visitors if orders is not None and visitors else None
    refund_rate = refunds / revenue if refunds is not None and revenue else None
    atv = revenue / orders if revenue is not None and orders else None

    # ---- 核心指标行 ----
    cmp_key, cmp_v = weekday_aligned(compare, weekday)
    rev_prev = cmp_v.get("revenue") if cmp_v else None
    rev_chg = at.pct_change(revenue, rev_prev) if revenue is not None else None
    order_chg = at.pct_change(orders, cmp_v.get("orders")) if cmp_v.get("orders") else None

    atv_chg = None
    atv_prev = None
    if cmp_v.get("revenue") and cmp_v.get("orders"):
        atv_prev = cmp_v["revenue"] / cmp_v["orders"]
        atv_chg = at.pct_change(atv, atv_prev)

    small_note = f"（订单 < {MIN_ORDERS} 单，小样本只报值）" if small_sample else ""
    rows = [
        {"指标": "营收", "本期": fmt_money(revenue),
         "对比": f"{cmp_key} {fmt_money(rev_prev)}，{rev_chg:+.1f}%" if rev_chg is not None else "无对比口径",
         "判定": ("超出 20% 波动带，需当天复盘" if special is False and rev_chg is not None and abs(rev_chg) > REV_MENTION * 100
                  else "正常波动带内" if rev_chg is not None and abs(rev_chg) <= REV_NORMAL * 100
                  else ("活动带动（特殊期）" if special else "波动 10%~20%，提一句方向" if rev_chg is not None else "—"))},
        {"指标": "转化率", "本期": fmt_pct(conv) + small_note,
         "对比": f"健康区间 {CONV_HEALTHY[0]:.0%}~{CONV_HEALTHY[1]:.0%}",
         "判定": "只报值不下结论" + small_note if small_sample else judge_conv(conv)},
        {"指标": "退款率", "本期": fmt_pct(refund_rate) + small_note,
         "对比": f"健康线 < {REFUND_HEALTHY:.0%}",
         "判定": "只报值不下结论" + small_note if small_sample else judge_refund(refund_rate)},
        {"指标": "客单价", "本期": (f"¥{atv:,.2f}" if atv is not None else "数据缺失"),
         "对比": (f"{cmp_key.split('（')[0]} {fmt_money2(atv_prev)}，"
                  f"{atv_chg:+.1f}%") if atv_chg is not None else "无对比口径",
         "判定": ("超出 ±15% 波动带" if atv_chg is not None and abs(atv_chg) > ATV_BAND * 100
                  else "±15% 波动带内" if atv_chg is not None else "—")},
        {"指标": "订单数", "本期": f"{orders:,.0f} 单" if orders is not None else "数据缺失",
         "对比": (f"{cmp_key.split('（')[0]} {cmp_v.get('orders'):,.0f} 单，{order_chg:+.1f}%"
                  if order_chg is not None else "无对比口径"),
         "判定": "与营收同向" if (rev_chg is not None and order_chg is not None
                                  and rev_chg * order_chg > 0) else "—"},
        {"指标": "访客数", "本期": f"{visitors:,.0f}" if visitors is not None else "数据缺失",
         "对比": (f"{cmp_key.split('（')[0]} {cmp_v.get('visitors'):,.0f}"
                  if cmp_v.get("visitors") else "无对比口径"),
         "判定": "—"},
    ]
    key_rows = pick_key_metrics(rows)

    # ---- 一句话结论：方向 + 幅度 + 归因 ----
    if special:
        headline = (f"今天{period}为特殊期（促销/节假日），营收 {fmt_money(revenue)}，"
                    f"偏离属活动带动，不按异常处理。")
    elif rev_chg is None:
        headline = f"今天营收 {fmt_money(revenue)}，无对比口径，方向待补数据后判断。"
    elif rev_chg > REV_NORMAL * 100:
        headline = f"今天涨了，营收 {fmt_money(revenue)}，比{cmp_key.split('（')[0]}多 {rev_chg:.1f}%，动因见核心指标。"
    elif rev_chg < -REV_MENTION * 100:
        drop = "到店客流" if (order_chg is not None and order_chg < 0) else "客单结构"
        headline = (f"今天掉了，营收 {fmt_money(revenue)}，比{cmp_key.split('（')[0]}少 {abs(rev_chg):.1f}%，"
                    f"问题出在{drop}，不是订单质量。")
    elif rev_chg < -REV_NORMAL * 100:
        headline = f"今天略降，营收 {fmt_money(revenue)}，比{cmp_key.split('（')[0]}少 {abs(rev_chg):.1f}%，幅度可控。"
    else:
        headline = f"今天稳，营收 {fmt_money(revenue)}，比{cmp_key.split('（')[0]}波动 {rev_chg:+.1f}%，在正常带内。"

    # ---- 异常点与排查方向（异常必须给排查方向）----
    anomaly_rows = []
    for a in alerts:
        metric = str(a.get("指标", ""))
        direction = next((v for k, v in INVESTIGATE.items() if k in metric), None)
        if direction is None:
            direction = "先核对数据源无误，再按指标口径拆维度（渠道/时段/商品）定位掉量点"
        anomaly_rows.append({
            "指标": metric, "级别": str(a.get("级别", "—")),
            "实际值": str(a.get("实际值", "—")), "阈值/基线": str(a.get("阈值/基线", "—")),
            "证据": str(a.get("证据", "—")), "排查方向": direction,
        })
    if not anomaly_rows:
        anomaly_rows = [{"指标": "（本期无确认异常）", "级别": "—", "实际值": "—",
                         "阈值/基线": "—", "证据": "—", "排查方向": "—"}]

    # ---- 明日建议（责任人 + 动作 + 范围）----
    suggestions = []
    if not small_sample and refund_rate is not None and refund_rate > REFUND_BAD:
        suggestions.append("让店长当日逐单核对全部退款原因，按出餐慢/配送/质量分桶，明天午市前给出占比最高的桶的处理动作")
    elif not small_sample and refund_rate is not None and refund_rate > REFUND_HEALTHY:
        suggestions.append("让店长明天逐单看近 3 天的退款原因，先确认集中在出餐慢还是配送，锁定主要桶再处理")
    if rev_chg is not None and rev_chg < -REV_MENTION * 100 and not special:
        suggestions.append("让店长复盘高峰时段（10:00—14:00）的外卖曝光与进店转化，先确认是流量没进来还是进来了没下单")
    if not small_sample and conv is not None and conv < CONV_WATCH:
        suggestions.append("让运营检查这两天投的泛流量广告，把预算挪到进店转化最高的渠道，先跑 3 天看转化")
    if atv_chg is not None and abs(atv_chg) > ATV_BAND * 100 and not special:
        suggestions.append("让店长核对套餐与满减结构，检查高单价商品动销，本周内恢复客单价到 ±15% 带内")
    if not suggestions:
        suggestions.append("今日各项在健康区间，维持现有排班与出餐节奏，无专项动作")

    summary = {
        "门店": store, "品类": category, "周期": period, "星期": weekday or "—",
        "对比口径": cmp_key or "无", "数据日期": m.get("date", "—"),
        "营收": fmt_money(revenue), "订单数": orders, "访客数": visitors,
        "客单价": (f"¥{atv:,.2f}" if atv is not None else "—"), "转化率": fmt_pct(conv), "退款率": fmt_pct(refund_rate),
        "小样本豁免": "是" if small_sample else "否", "特殊期豁免": "是" if special else "否",
        "确认异常数": len(anomaly_rows) if alerts else 0,
        "基准": {"转化率": "2%~5%", "退款率": "<3%", "客单价波动": "±15%",
                 "营收波动": "±10% 正常 / >20% 必须点出", "小样本线": f"{MIN_ORDERS} 单"},
        "硬规则": ("结论前置；关键数 ≤ 3；异常必给排查方向；数值以脚本重算为准"
                   "（客单价=营收÷订单、转化率=订单÷访客、退款率=退款额÷营收）"),
    }

    at.ensure_outdir(outdir)

    xlsx = at.write_excel(
        os.path.join(outdir, "日报数据摘要.xlsx"),
        {
            "核心指标": key_rows,
            "全部指标": rows,
            "异常清单": anomaly_rows,
            "明日建议": [{"序号": i, "建议": s} for i, s in enumerate(suggestions, 1)],
            "汇总": [{"项": k, "内容": str(v)} for k, v in summary.items()],
        },
        highlights={"核心指标": {"判定": "contains:需"}, "异常清单": {"级别": "contains:高"}},
        widths={"异常清单": {"证据": 34, "排查方向": 46}, "明日建议": {"建议": 52}},
    )

    sections = [
        {"heading": "一、一句话结论", "paras": [headline]},
        {"heading": "二、核心指标（关键数 ≤ 3，每个带参照物）",
         "table": {"cols": ["指标", "本期", "对比", "判定"],
                   "rows": [[r["指标"], r["本期"], r["对比"], r["判定"]] for r in key_rows]}},
        {"heading": "三、异常点与排查方向",
         "table": {"cols": ["指标", "实际值", "阈值/基线", "排查方向"],
                   "rows": [[r["指标"], r["实际值"], r["阈值/基线"], r["排查方向"]] for r in anomaly_rows]}},
        {"heading": "四、明日建议（责任人 + 动作 + 范围）",
         "bullets": suggestions},
        {"heading": "五、数据口径与人工确认",
         "bullets": [
             f"星期对齐：对比口径为「{cmp_key or '无'}」，不做跨星期比较（周末效应 30%~50% 属正常波动）",
             f"小样本纪律：当日订单 {orders if orders is not None else '—'} 单"
             + (f"（< {MIN_ORDERS} 单，转化率/退款率只报值不下结论）" if small_sample else "（≥ 20 单，判定有效）"),
             f"特殊期检查：{'命中促销/节假日/新店期，偏离写活动带动' if special else '非特殊期，偏离按真实波动处理'}",
             "本日报为 AI 生成内容，数据口径以门店台账为准，对外使用前请人工复核",
         ]},
    ]
    docx = at.write_docx(
        os.path.join(outdir, "经营日报.docx"),
        f"{store} 经营日报（{period}）", sections,
        subtitle=f"{category} · AI 生成内容 · 数值以脚本重算为准",
    )
    js = at.write_json(
        {"summary": summary, "key_metrics": key_rows, "all_metrics": rows,
         "anomalies": anomaly_rows, "suggestions": suggestions, "headline": headline,
         "generated_at": at.stamp(),
         "note": "脚本生成的日报数据；文案润色与业务归因由模型按 prompt.txt 补充"},
        os.path.join(outdir, "report_copy.json"))
    return {"files": [docx, xlsx, js], "summary": summary,
            "headline": headline, "anomaly_count": len(anomaly_rows) if alerts else 0}


def main():
    ap = argparse.ArgumentParser(description="日报文案生成 —— 数据 → Word 日报 + Excel 摘要")
    ap.add_argument("--input", help="输入 JSON（metrics/compare/alerts/...）")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--demo", action="store_true", help="用内置样例（含 2 个异常点）跑一遍")
    a = ap.parse_args()

    if a.demo:
        payload = DEMO
    elif a.input:
        payload = at.read_json(a.input)
    else:
        ap.error("需要 --input / --demo 之一")

    r = build(payload, a.outdir)
    print(r["headline"])
    print(f"确认异常 {r['anomaly_count']} 项，排查方向已附")
    for f in r["files"]:
        print(" 产物:", f)
    at.emit(r)


if __name__ == "__main__":
    main()
