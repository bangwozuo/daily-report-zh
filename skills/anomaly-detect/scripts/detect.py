# -*- coding: utf-8 -*-
"""
经营数据异常检测 —— 四规则量化检测 + 假阳性拦截 + 产物生成。

职责边界（重要）：
  本脚本只做**确定性计算与产物生成**：派生指标、四条规则的量化判定、星期对齐、
  假阳性拦截、Excel/PNG 落盘。
  业务语境解释、排查方向撰写、与店主确认口径，由模型按 prompt.txt 完成。

四条规则（与 prompt.txt 一致）：
  R1 绝对阈值   营收环比 ≤ -20% / 转化率 < 1.5% / 退款率 > 5%
  R2 离群       |x-μ| > 2σ（样本 ≥ 14 天）；样本 < 14 天改用 IQR(1.5)
  R3 连续趋势   连续 ≥ 3 天同向且累计幅度 > 10%
  R4 季节性     有去年同期用同比；同比与环比冲突时以同比为准并标注

假阳性拦截：
  - 促销日 / 节假日 / 新店未满 30 天      → 不判异常，标注「特殊期，需人工判断」
  - 单日订单 < 20 单                      → 转化率/退款率失真，只报值不下结论
  - 区间内缺失日期                        → 不补零，显式标注「数据缺失 N 天」

用法：
  python detect.py --input input.json --outdir out
  python detect.py --demo --outdir out

产物：
  out/异常预警清单.xlsx   预警清单 / 指标明细 / 星期对齐 / 数据质量 / 汇总
  out/趋势图.png          营收与转化率双轴趋势（异常点用红圈标出）
  out/detect.json         机器可读结果（供智能体/工作流读取）
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
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

at.need("pandas")

# ---------------------------------------------------------------- 默认规则与行业基准

DEFAULT_RULES = {
    "revenue_dod_drop": 20.0,     # R1 营收环比跌幅阈值（%）
    "conversion_floor": 1.5,      # R1 转化率下限（%）
    "refund_ceil": 5.0,           # R1 退款率上限（%）
    "z_k": 2.0,                   # R2 离群倍数
    "z_min_sample": 14,           # R2 启用 Z-score 的最小样本天数
    "iqr_k": 1.5,                 # R2 小样本回退系数
    "trend_days": 3,              # R3 连续同向天数
    "trend_total_pct": 10.0,      # R3 累计幅度阈值（%）
    "yoy_conflict_pct": 5.0,      # R4 同比与环比背离多少个百分点算冲突
    "min_orders": 20,             # 小样本失真阈值（单）
    "new_store_days": 30,         # 新店保护期（天）
}

# 餐饮/零售行业基准（用于给"值不看结论"的项补一句健康区间判断）
BENCHMARK = {
    "餐饮": {"conversion": (2.0, 5.0), "refund_rate_ceil": 3.0, "aov_wave": 15.0},
    "零售": {"conversion": (1.5, 4.0), "refund_rate_ceil": 3.0, "aov_wave": 15.0},
    "通用": {"conversion": (1.5, 5.0), "refund_rate_ceil": 3.0, "aov_wave": 15.0},
}

WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

LEVEL_ORDER = {"高": 0, "中": 1, "低": 2}


# ---------------------------------------------------------------- 演示数据

def demo_data():
    """生成 30 天餐饮门店数据，内嵌 5 类真实异常与 1 处缺失，便于验证全部规则。"""
    import random
    random.seed(20260930)
    rows = []
    start = dt.date(2026, 9, 1)
    for i in range(30):
        d = start + dt.timedelta(days=i)
        if d.day == 15:            # 缺失日：9/15 未上传数据
            continue
        weekend = d.weekday() >= 5
        base = 8600 + i * 35
        revenue = base * (1.28 if weekend else 1.0) + random.randint(-420, 420)
        orders = int(revenue / 63) + random.randint(-5, 5)
        visitors = int(orders / 0.030) + random.randint(-60, 60)
        refunds = int(revenue * (0.021 + random.random() * 0.006))

        if d.day == 12:            # 异常 A：营收骤降（绝对阈值 + 离群）
            revenue = int(revenue * 0.62)
            orders = int(orders * 0.70)
        if d.day == 18:            # 特殊期：周年庆促销日（营收冲高，须拦截不判异常）
            revenue = int(revenue * 1.66)
            orders = int(orders * 1.58)
            visitors = int(visitors * 1.45)
        if 20 <= d.day <= 24:      # 异常 B：退款率持续 5 天高于 5%（商品质量事件）
            refunds = int(revenue * 0.079)
        if 22 <= d.day <= 24:      # 异常 C：连续三天转化率下滑（访客翻倍，订单没跟上）
            visitors = int(visitors * 2.05)
        if d.day == 25:            # 特殊期：中秋节假日，写字楼商圈客流骤减（须拦截不判异常）
            revenue = int(revenue * 0.72)
            orders = int(orders * 0.75)
        if d.day == 27:            # 小样本：当天下雨只做了 15 单
            orders = 15
            visitors = int(visitors * 0.28)
            revenue = int(revenue * 0.19)
        rows.append({
            "date": d.isoformat(),
            "revenue": int(revenue),
            "orders": int(orders),
            "visitors": int(visitors),
            "refunds": int(refunds),
            "last_year_revenue": int(revenue / 1.16 + random.randint(-260, 260)),
        })
    # 构造 9/8 的"同比优于环比"样本：环比小幅上升，但同比大幅下滑
    for r in rows:
        if r["date"] == "2026-09-08":
            r["last_year_revenue"] = int(r["revenue"] / 0.75)
    return {
        "store_name": "望京 SOHO 店",
        "category": "餐饮",
        "period": {"start": "2026-09-01", "end": "2026-09-30"},
        "opened_date": "2025-04-18",
        "promo_days": ["2026-09-18"],
        "holidays": ["2026-09-25"],
        "metrics": rows,
        "rules": DEFAULT_RULES,
    }


# ---------------------------------------------------------------- 计算层

def compute(rows, rules):
    """派生指标 + 环比 + 同比 + 星期对齐。数值全部在此算出，模型不得自行计算。"""
    df = _pandas().DataFrame(rows)
    df["date"] = _pandas().to_datetime(df["date"])
    df = df.sort_values("date").reset_index(drop=True)

    for col in ("revenue", "orders", "visitors", "refunds"):
        if col not in df.columns:
            df[col] = 0
    if "last_year_revenue" not in df.columns:
        df["last_year_revenue"] = None

    # 派生指标：客单价 / 转化率 / 退款率（均由脚本计算，禁止用户提供或模型估算）
    df["aov"] = (df["revenue"] / df["orders"]).where(df["orders"] > 0).round(2)
    df["conversion"] = (df["orders"] / df["visitors"] * 100).where(df["visitors"] > 0).round(2)
    df["refund_rate"] = (df["refunds"] / df["revenue"] * 100).where(df["revenue"] > 0).round(2)

    # 环比（日）：补零会制造假跌，因此只在相邻两个"真实存在"的日期之间算环比
    df["revenue_dod"] = df["revenue"].pct_change() * 100
    df["revenue_dod"] = df["revenue_dod"].round(2)

    # 同比：仅有去年同期数据时计算
    df["revenue_yoy"] = [
        at.pct_change(r.revenue, r.last_year_revenue) if r.last_year_revenue else None
        for r in df.itertuples()
    ]

    # 星期对齐：同星期几的历史均值（周末效应会造成 30%+ 正常波动，比较必须按星期几对齐）
    df["weekday"] = df["date"].dt.weekday
    wd_mean = df.groupby("weekday")["revenue"].transform("mean")
    df["weekday_mean"] = wd_mean.round(0)
    df["weekday_dev"] = ((df["revenue"] - wd_mean) / wd_mean * 100).round(2)

    # 周环比（同星期几 vs 上周同星期几）——餐饮/零售最可比的口径
    wow = []
    by_wd = {}
    for r in df.itertuples():
        prev = by_wd.get(r.weekday)
        wow.append(at.pct_change(r.revenue, prev) if prev else None)
        by_wd[r.weekday] = r.revenue
    df["revenue_wow"] = [round(v, 2) if v is not None else None for v in wow]

    df["weekday_cn"] = df["weekday"].map(lambda i: WEEKDAY_CN[i])
    return df


def _pandas():
    import pandas as pd
    return pd


# ---------------------------------------------------------------- 检测层

def detect(cur, df, rules, meta):
    """四规则判定 + 假阳性拦截。返回 (alerts, quality, stats)。"""
    r = {**DEFAULT_RULES, **(rules or {})}
    alerts = []
    period = cur.get("period") or {}
    if isinstance(period, str):          # 兼容只给单日字符串（如 "2026-09-17"）的调用方
        period = {"start": period, "end": period}
    promo = set(cur.get("promo_days") or [])
    holidays = set(cur.get("holidays") or [])
    opened = cur.get("opened_date")

    n = len(df)
    stat = {
        "sample_days": n,
        "outlier_method": "Z-score(2σ)" if n >= r["z_min_sample"] else f"IQR({r['iqr_k']})",
    }

    # ---- 数据质量：缺失日期（不补零）
    quality, missing = [], []
    if period.get("start") and period.get("end"):
        s = dt.date.fromisoformat(period["start"])
        e = dt.date.fromisoformat(period["end"])
        have = {d.date() for d in df["date"]}
        d = s
        while d <= e:
            if d not in have:
                missing.append(d.isoformat())
            d += dt.timedelta(days=1)
        if missing:
            quality.append({
                "项": "区间内缺失日期",
                "值": f"{len(missing)} 天",
                "明细": "、".join(missing),
                "处理": "未补零；环比在缺失处断开，不得跨缺口比较，否则会制造假跌",
            })
    # 新店保护期
    store_age_days = None
    if opened:
        opened_d = dt.date.fromisoformat(opened)
        store_age_days = (dt.date.fromisoformat(period.get("end", dt.date.today().isoformat()))
                          - opened_d).days
        if store_age_days < r["new_store_days"]:
            quality.append({
                "项": "新店保护期",
                "值": f"开业 {store_age_days} 天",
                "明细": f"开业日 {opened}",
                "处理": f"未满 {r['new_store_days']} 天，全部指标不判异常（缺少稳定基线）",
            })
    # 极小样本日
    small = df[df["orders"] < r["min_orders"]]
    for row in small.itertuples():
        quality.append({
            "项": "单日样本量不足",
            "值": f"{row.date:%Y-%m-%d}（{int(row.orders)} 单）",
            "明细": f"转化率 {row.conversion}%、退款率 {row.refund_rate}%",
            "处理": f"订单 < {r['min_orders']} 单，转化率与退款率失真 —— 只报值不下结论",
        })
    stat["missing_days"] = len(missing)
    stat["new_store"] = bool(store_age_days is not None and store_age_days < r["new_store_days"])
    stat["small_sample_days"] = int(len(small))

    # ---- 判异常总闸：新店保护期内不产生任何异常
    def gate(date_str, metric):
        if stat["new_store"]:
            return "已拦截（新店保护期）"
        if date_str in promo:
            return "已拦截（促销日）"
        if date_str in holidays:
            return "已拦截（节假日）"
        return None

    # ---- R2 离群：样本 ≥14 用 Z-score，否则 IQR
    outliers = {}
    for metric, col in (("营收", "revenue"), ("转化率", "conversion"), ("退款率", "refund_rate")):
        xs = [v for v in df[col].tolist() if v is not None]
        if n >= r["z_min_sample"]:
            mu = at.mean(xs)
            sd = at.stdev(xs)
            for row in df.itertuples():
                v = getattr(row, col)
                if v is None or not sd:
                    continue
                z = (v - mu) / sd
                if abs(z) > r["z_k"]:
                    outliers.setdefault((row.date.strftime("%Y-%m-%d"), metric), []).append(
                        f"Z={z:+.2f}（μ={mu:.1f}，σ={sd:.1f}）")
        else:
            lo, hi = at.iqr_outliers(xs, r["iqr_k"])
            if lo is None:
                continue
            for row in df.itertuples():
                v = getattr(row, col)
                if v is None:
                    continue
                if v < lo or v > hi:
                    outliers.setdefault((row.date.strftime("%Y-%m-%d"), metric), []).append(
                        f"IQR 越界（下界 {lo:.1f}／上界 {hi:.1f}）")

    # ---- R1 绝对阈值 + R4 季节性 + R2 离群 逐日归并
    for row in df.itertuples():
        ds = row.date.strftime("%Y-%m-%d")
        hits = []

        dod = row.revenue_dod
        if dod is not None and not _isna(dod) and dod <= -r["revenue_dod_drop"]:
            hits.append({
                "指标": "营收环比", "规则": "R1 绝对阈值",
                "级别": "高" if dod <= -30 else "中",
                "实际值": f"{row.revenue_dod:.1f}%",
                "阈值": f"环比 ≤ -{r['revenue_dod_drop']:.0f}%",
                "证据": f"营收 {int(row.revenue)} 元（前一日 {_prev_revenue(df, row.Index)} 元）",
            })
        conv = row.conversion
        if conv is not None and not _isna(conv) and conv < r["conversion_floor"]:
            low_sample = row.orders < r["min_orders"]
            hits.append({
                "指标": "转化率", "规则": "R1 绝对阈值",
                "级别": "低" if low_sample else ("高" if conv < 1.0 else "中"),
                "实际值": f"{conv:.2f}%",
                "阈值": f"< {r['conversion_floor']}%",
                "证据": (f"访客 {int(row.visitors)}，订单 {int(row.orders)} —— "
                         f"单日 {int(row.orders)} 单 < {r['min_orders']} 单，失真，只报值不下结论"
                         if low_sample else f"访客 {int(row.visitors)}，订单 {int(row.orders)}"),
                "只报值": low_sample,
            })
        rr = row.refund_rate
        if rr is not None and not _isna(rr) and rr > r["refund_ceil"]:
            low_sample = row.orders < r["min_orders"]
            hits.append({
                "指标": "退款率", "规则": "R1 绝对阈值",
                "级别": "低" if low_sample else ("高" if rr > 7 else "中"),
                "实际值": f"{rr:.2f}%",
                "阈值": f"> {r['refund_ceil']}%",
                "证据": (f"退款 {int(row.refunds)} 元 / 营收 {int(row.revenue)} 元 —— "
                         f"订单量不足，失真，只报值不下结论"
                         if low_sample else f"退款 {int(row.refunds)} 元 / 营收 {int(row.revenue)} 元"),
                "只报值": low_sample,
            })

        # R4 季节性：同比与环比冲突以同比为准
        yoy = row.revenue_yoy
        if yoy is not None and not _isna(yoy) and abs(yoy) >= r["revenue_dod_drop"]:
            conflict = (dod is not None and not _isna(dod)
                        and abs(yoy - dod) >= r["yoy_conflict_pct"] * 4
                        and (yoy < 0) != (dod < 0))
            hits.append({
                "指标": "营收同比", "规则": "R4 季节性（同比优先）",
                "级别": "中",
                "实际值": f"{yoy:.1f}%",
                "阈值": f"同比偏离 ≥ {r['revenue_dod_drop']:.0f}%",
                "证据": (f"去年同期 {int(row.last_year_revenue)} 元，本期 {int(row.revenue)} 元；"
                         f"环比 {dod:+.1f}% 与同比 {yoy:+.1f}% 方向冲突，以同比为准"
                         if conflict else
                         f"去年同期 {int(row.last_year_revenue)} 元，本期 {int(row.revenue)} 元"),
            })

        for (ods, metric), ev in outliers.items():
            if ods != ds:
                continue
            base = {"营收": ("revenue", "元", 0), "转化率": ("conversion", "%", 2),
                    "退款率": ("refund_rate", "%", 2)}[metric]
            val = getattr(row, base[0])
            low_sample = metric != "营收" and row.orders < r["min_orders"]
            hits.append({
                "指标": f"{metric}离群", "规则": f"R2 离群（{stat['outlier_method']}）",
                "级别": "低" if low_sample else "中",
                "实际值": f"{val:.{base[2]}f}{base[1]}",
                "阈值": f"|x-μ| > {r['z_k']:.0f}σ" if n >= r["z_min_sample"]
                        else f"IQR × {r['iqr_k']}",
                "证据": "；".join(ev) + (
                    f" —— 当日仅 {int(row.orders)} 单，样本不足，只报值不下结论" if low_sample else ""),
                "只报值": low_sample,
            })

        if not hits:
            continue
        for h in hits:
            blocked = gate(ds, h["指标"])
            h_out = {
                "日期": ds, "星期": row.weekday_cn, "指标": h["指标"], "规则": h["规则"],
                "级别": h["级别"] if not blocked else "拦截",
                "实际值": h["实际值"], "阈值/基线": h["阈值"], "证据": h["证据"],
                "判定": ("特殊期，需人工判断 —— " + blocked) if blocked
                        else ("只报值不下结论" if h.get("只报值") else "确认异常"),
                "星期对齐偏离": f"{row.weekday_dev:+.1f}%",
                "周环比": "—" if _isna(row.revenue_wow) else f"{row.revenue_wow:+.1f}%",
            }
            alerts.append(h_out)
    stat["blocked_promo_holiday"] = len([a for a in alerts if "已拦截（" in a["判定"]])

    # ---- R3 连续趋势：连续 ≥3 天同向且累计 >10%（对营收/转化率/退款率分别扫描）
    dates = [f"{d:%Y-%m-%d}" for d in df["date"]]
    for metric, col in (("营收", "revenue"), ("转化率", "conversion"), ("退款率", "refund_rate")):
        vals = df[col].tolist()
        chg = [None] + [at.pct_change(vals[i], vals[i - 1]) for i in range(1, len(vals))]
        for sign, direction in ((-1, "下滑"), (1, "上升")):
            rows_i = [i for i, v in enumerate(chg) if v is not None and v * sign > 0]
            if not rows_i:
                continue
            flags = [i in set(rows_i) for i in range(len(vals))]
            _, runs = at.consecutive_runs(flags)
            for (a, b) in runs:
                if b - a + 1 < r["trend_days"] + 1:
                    continue
                total = at.pct_change(vals[b], vals[a - 1]) if a >= 1 else None
                if total is None or abs(total) <= r["trend_total_pct"]:
                    continue
                ds = dates[b]
                blocked = gate(ds, metric)
                alerts.append({
                    "日期": ds, "星期": WEEKDAY_CN[df["date"][b].weekday()],
                    "指标": f"{metric}连续趋势", "规则": "R3 连续趋势",
                    "级别": "高" if (direction == "下滑" and metric != "退款率") else "中",
                    "实际值": f"连续 {b - a + 1} 天{direction}",
                    "阈值/基线": f"连续 ≥ {r['trend_days']} 天 且 累计 > {r['trend_total_pct']:.0f}%",
                    "证据": f"{dates[a - 1]} → {ds} 累计 {total:+.1f}%",
                    "判定": ("特殊期，需人工判断 —— " + blocked) if blocked else "确认异常",
                    "星期对齐偏离": f"{df['weekday_dev'][b]:+.1f}%",
                    "周环比": "—" if _isna(df["revenue_wow"][b]) else f"{df['revenue_wow'][b]:+.1f}%",
                })

    alerts.sort(key=lambda a: (LEVEL_ORDER.get(a["级别"], 9), a["日期"]))
    return alerts, quality, stat


def _isna(v):
    return v is None or (isinstance(v, float) and v != v)


def _prev_revenue(df, idx):
    if idx == 0:
        return 0
    return int(df["revenue"][idx - 1])


def summary(df, alerts, stat, cur):
    confirmed = [a for a in alerts if a["判定"] == "确认异常"]
    cat = cur.get("category", "通用")
    bench = BENCHMARK.get(cat, BENCHMARK["通用"])
    return {
        "门店": cur.get("store_name", ""),
        "品类": cat,
        "样本天数": stat["sample_days"],
        "离群方法": stat["outlier_method"],
        "缺失天数": stat["missing_days"],
        "新店保护期": "是" if stat["new_store"] else "否",
        "小样本天数": stat["small_sample_days"],
        "预警总条数": len(alerts),
        "确认异常条数": len(confirmed),
        "特殊期拦截条数": stat["blocked_promo_holiday"],
        "涉及天数": len({a["日期"] for a in confirmed}),
        "总营收": int(df["revenue"].sum()),
        "平均转化率%": round(df["conversion"].mean(), 2),
        "整体退款率%": round(df["refunds"].sum() / df["revenue"].sum() * 100, 2),
        "行业转化率健康区间": f"{bench['conversion'][0]}~{bench['conversion'][1]}%",
        "行业退款率上限": f"{bench['refund_rate_ceil']}%",
        "客单价波动正常带": f"±{bench['aov_wave']}%",
    }


# ---------------------------------------------------------------- 输出层

DETAIL_COLS = {
    "date": "日期", "weekday_cn": "星期", "revenue": "营收", "orders": "订单", "visitors": "访客",
    "refunds": "退款额", "aov": "客单价", "conversion": "转化率%", "refund_rate": "退款率%",
    "revenue_dod": "营收环比%", "revenue_yoy": "营收同比%", "weekday_dev": "同星期偏离%",
}


def write_outputs(cur, df, alerts, quality, stat, outdir):
    at.ensure_outdir(outdir)
    files = []

    detail = []
    for row in df.itertuples():
        d = {}
        for k, cn in DETAIL_COLS.items():
            v = getattr(row, k)
            if k == "date":
                v = f"{v:%Y-%m-%d}"
            elif _isna(v):
                v = "—"
            elif isinstance(v, float):
                v = round(v, 2)
            d[cn] = v
        detail.append(d)

    wd = []
    for row in df.itertuples():
        wd.append({
            "星期": row.weekday_cn, "日期": f"{row.date:%Y-%m-%d}",
            "营收": int(row.revenue), "同星期均值": int(row.weekday_mean),
            "同星期偏离%": row.weekday_dev, "周环比(同星期)%":
                "—" if _isna(row.revenue_wow) else row.revenue_wow,
        })

    alerts_sheet = alerts or [{"日期": "（本期无异常）", "星期": "", "指标": "", "规则": "",
                               "级别": "", "实际值": "", "阈值/基线": "", "证据": "",
                               "判定": "", "星期对齐偏离": "", "周环比": ""}]
    summ = [{"项": k, "内容": str(v)} for k, v in summary(df, alerts, stat, cur).items()]

    xlsx = at.write_excel(
        os.path.join(outdir, "异常预警清单.xlsx"),
        {"预警清单": alerts_sheet, "指标明细": detail, "星期对齐": wd,
         "数据质量": quality or [{"项": "数据质量", "值": "无问题", "明细": "", "处理": ""}],
         "汇总": summ},
        highlights={"预警清单": {"级别": "contains:高", "判定": "contains:确认异常"}},
        widths={"预警清单": {"证据": 46, "判定": 30, "阈值/基线": 22, "规则": 20},
                "指标明细": {"日期": 12}, "数据质量": {"明细": 34, "处理": 46}},
    )
    files.append(xlsx)

    # 趋势图：营收 + 转化率双轴，异常点红圈叠加
    anom_dates = {a["日期"] for a in alerts if a["判定"] == "确认异常"}
    xs = [f"{d:%m-%d}" for d in df["date"]]
    rev = [int(v) for v in df["revenue"]]
    conv = [None if _isna(v) else round(float(v), 2) for v in df["conversion"]]
    rev_pts = [v if xs[i] in anom_dates else None for i, v in enumerate(rev)]
    png = at.line_chart(
        os.path.join(outdir, "趋势图.png"), xs,
        {"营收（元）": rev, "异常点": rev_pts},
        title=f"{cur.get('store_name', '')} 营收与转化率趋势（红点为确认异常）",
        xlabel="日期", ylabel="营收（元）",
        y2={"转化率（%）": conv}, y2label="转化率（%）", figsize=(11.5, 4.8),
    )
    files.append(png)

    js = at.write_json({
        "store_name": cur.get("store_name", ""),
        "category": cur.get("category", ""),
        "period": cur.get("period", {}),
        "rules": {**DEFAULT_RULES, **(cur.get("rules") or {})},
        "summary": summary(df, alerts, stat, cur),
        "alerts": alerts,
        "data_quality": quality,
        "generated_at": at.stamp(),
        "note": "数值均由本脚本判定；语境解释与排查方向由模型按 prompt.txt 完成",
    }, os.path.join(outdir, "detect.json"))
    files.append(js)
    return files


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser(description="经营数据异常检测（四规则 + 假阳性拦截）")
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

    if not cur.get("metrics"):
        print("[错误] metrics 为空：缺失数据不补零、不估算，请先补齐经营数据。", file=sys.stderr)
        sys.exit(3)

    df = compute(cur["metrics"], cur.get("rules") or {})
    alerts, quality, stat = detect(cur, df, cur.get("rules") or {}, {})
    files = write_outputs(cur, df, alerts, quality, stat, a.outdir)
    s = summary(df, alerts, stat, cur)
    print(f"样本 {s['样本天数']} 天｜预警 {s['预警总条数']} 条"
          f"（确认异常 {s['确认异常条数']}，特殊期拦截 {s['特殊期拦截条数']}）"
          f"｜离群方法 {s['离群方法']}")
    for f in files:
        print(" 产物:", f, f"({os.path.getsize(f) / 1024:.1f} KB)")
    at.emit({"status": "ok", "files": files, "summary": s})


if __name__ == "__main__":
    main()
