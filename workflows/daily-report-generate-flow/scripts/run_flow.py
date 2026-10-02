# -*- coding: utf-8 -*-
"""
日报生成流程（daily-report-generate-flow）—— 端到端编排脚本。

编排的步骤：
  1. 异常检测     → 调用 ../../skills/anomaly-detect/scripts/detect.py 取当日确认异常
  2. 指标与结论   → 按 ../../skills/daily-report-copy/ 的口径算派生指标、挑 3 个关键数、写结论
  3. 交付物落盘   → 产出《经营日报》Word + Excel（附趋势图）
交付物：
  out/经营日报.docx          核心指标 / 异常预警 / 趋势图 / 行动建议（可直接汇报）
  out/经营日报.xlsx          关键数 / 指标明细 / 异常引用 / 汇总
  out/report_flow.json       流程结果（结论句 + 三关键数 + 产物路径）
  out/steps/01-异常检测/*     第 1 步原始产物

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo --outdir out
"""
from __future__ import annotations

import argparse
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
FLOW_DIR = os.path.dirname(HERE)
REPO = os.path.dirname(os.path.dirname(FLOW_DIR))
sys.path.insert(0, os.path.join(REPO, "lib"))
try:
    import assettools as at
except ImportError:  # pragma: no cover
    print("[错误] 未找到 lib/assettools.py。", file=sys.stderr)
    sys.exit(2)

STEP1 = os.path.join(REPO, "skills", "anomaly-detect", "scripts", "detect.py")

# 餐饮/零售行业健康区间（与 daily-report-copy 的 prompt 一致）
CONV_HEALTHY = (2.0, 5.0)
REFUND_CEIL = 3.0
AOV_BAND = 15.0


def run_step(cmd, label):
    print(f"[step] {label}")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        print(r.stdout or "", file=sys.stderr)
        print(r.stderr or "", file=sys.stderr)
        raise RuntimeError(f"步骤「{label}」失败，退出码 {r.returncode}；"
                           f"先修复依赖或补齐输入，不要手工编造日报内容")
    return r


def main():
    ap = argparse.ArgumentParser(description="日报生成流程")
    ap.add_argument("--input", help="输入 JSON 路径")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--demo", action="store_true")
    a = ap.parse_args()

    if a.demo:
        cur = at.read_json(os.path.join(FLOW_DIR, "examples", "input.json"))
    elif a.input:
        cur = at.read_json(a.input)
    else:
        ap.error("需提供 --input 或 --demo")

    outdir = at.ensure_outdir(a.outdir)
    step1_dir = os.path.join(outdir, "steps", "01-异常检测")

    py = [sys.executable, STEP1, "--outdir", step1_dir]
    py = py + (["--input", a.input] if a.input else ["--demo"])
    run_step(py, "异常检测 / 当日异常判定")
    det = at.read_json(os.path.join(step1_dir, "detect.json"))

    period = cur.get("period")
    rows = cur["metrics"]
    today = next((r for r in rows if r["date"] == period), rows[-1])
    rev, orders = today["revenue"], today["orders"]
    visitors, refunds = today["visitors"], today["refunds"]
    aov = rev / orders if orders else 0
    conv = orders / visitors * 100 if visitors else 0
    rr = refunds / rev * 100 if rev else 0

    cmp_name, cmp_data = next(iter((cur.get("compare") or {"环比": {}}).items()))
    cmp_rev = cmp_data.get("revenue")
    delta = (rev / cmp_rev - 1) * 100 if cmp_rev else None

    # ---- 挑 3 个关键数（营收必选 → 退款率 >3% → 转化率 <2% → 客单价波动 >15% → 兜底订单数）
    keys = [{"指标": "营收", "本期": f"¥{rev:,}",
             "对比": (f"{cmp_name} ¥{cmp_rev:,}，{delta:+.1f}%" if delta is not None else "无对比口径"),
             "判定": ("超出 20% 波动带，需当天复盘" if delta is not None and abs(delta) > 20
                      else "正常波动带内（±20%）")}]
    if rr > REFUND_CEIL:
        keys.append({"指标": "退款率", "本期": f"{rr:.2f}%", "对比": f"健康线 < {REFUND_CEIL}%",
                     "判定": "超线，需关注"})
    if conv < CONV_HEALTHY[0]:
        keys.append({"指标": "转化率", "本期": f"{conv:.2f}%",
                     "对比": f"健康区间 {CONV_HEALTHY[0]}%~{CONV_HEALTHY[1]}%",
                     "判定": "低于健康下限"})
    if len(keys) < 3:
        keys.append({"指标": "订单数", "本期": f"{orders} 单",
                     "对比": (f"{cmp_name} {cmp_data.get('orders')} 单"
                              if cmp_data.get("orders") else "无对比口径"),
                     "判定": "与营收同向参考"})
    if len(keys) < 3:
        keys.append({"指标": "客单价", "本期": f"¥{aov:.2f}", "对比": "营收 ÷ 订单（脚本重算）",
                     "判定": "结构参考"})
    keys = keys[:3]

    # ---- 结论句：方向 + 幅度 + 归因
    if delta is None:
        direction, reason = "情况", "缺少对比口径，仅报当日值"
    elif delta <= -20:
        direction, reason = "掉了", "问题出在流量与客单价，需要当天复盘"
    elif delta <= -10:
        direction, reason = "小幅回落", "幅度在可控范围，观察一天"
    elif delta < 10:
        direction, reason = "稳", "指标落在正常波动带内"
    else:
        direction, reason = "涨了", "增量需确认是否来自活动，活动期数据不可外推到平日"
    headline = f"今天{direction}，营收 ¥{rev:,}，" + (
        f"比{cmp_name} {'多' if delta and delta > 0 else '少'} {abs(delta):.1f}%，" if delta is not None else ""
    ) + f"{reason}。"

    # ---- 当日确认异常（引用第 1 步结果，不重复判定）
    day_alerts = [al for al in det["alerts"]
                  if al["日期"] == today["date"] and al["判定"] == "确认异常"]
    suppressed = [al for al in det["alerts"]
                  if al["日期"] == today["date"] and al["判定"] != "确认异常"]

    reminder = ("当日无异常，指标平稳。" if not day_alerts else
                f"当日命中 {len(day_alerts)} 条异常：" +
                "；".join(f"{al['指标']} {al['实际值']}" for al in day_alerts[:2]) + "。")
    if suppressed:
        reminder += "　另：" + "；".join(f"{al['指标']}（{al['判定']}）" for al in suppressed[:1])

    # ---- 行动建议：责任人 + 动作 + 范围
    if day_alerts:
        top = day_alerts[0]
        if "退款" in top["指标"]:
            suggestion = "让店长逐单看当日退款原因，确认是集中在某个单品还是某个配送时段。"
        elif "转化" in top["指标"]:
            suggestion = "让店长复盘当日投的流量渠道，把预算挪到进店转化最高的那个渠道。"
        else:
            suggestion = "让店长复盘当日 10:00—14:00 的外卖曝光与进店转化，先看流量端。"
    else:
        suggestion = "维持现有排班与活动节奏，明天同一时间再看一次转化率。"

    # ---- 产物
    img = os.path.join(step1_dir, "趋势图.png")
    xlsx = at.write_excel(
        os.path.join(outdir, "经营日报.xlsx"),
        {"关键数": keys,
         "指标明细": [{"日期": today["date"], "星期": today.get("weekday", ""),
                       "营收": rev, "订单": orders, "访客": visitors, "退款额": refunds,
                       "客单价": round(aov, 2), "转化率%": round(conv, 2), "退款率%": round(rr, 2)}],
         "异常引用": day_alerts or [{"日期": today["date"], "指标": "（当日无确认异常）"}],
         "数据质量": det["data_quality"] or [{"项": "（无）"}]},
        highlights={"关键数": {"判定": "contains:需"}},
        widths={"关键数": {"对比": 30}},
    )
    docx = at.write_docx(
        os.path.join(outdir, "经营日报.docx"),
        f"{cur.get('store_name', '')} 经营日报（{period}）",
        [
            {"heading": "一、一句话结论", "paras": [headline]},
            {"heading": "二、三个关键数",
             "table": {"cols": ["指标", "本期", "对比", "判定"],
                       "rows": [[k["指标"], k["本期"], k["对比"], k["判定"]] for k in keys]}},
            {"heading": "三、一句提醒", "paras": [reminder]},
            {"heading": "四、一个行动建议", "bullets": [suggestion]},
            {"heading": "五、趋势图", "image": img},
            {"heading": "六、当日确认异常",
             "table": {"cols": ["日期", "指标", "规则", "级别", "实际值", "证据"],
                       "rows": [[al["日期"], al["指标"], al["规则"], al["级别"], al["实际值"],
                                 al["证据"]] for al in day_alerts] or [["—", "无", "—", "—", "—", "—"]]}},
        ],
        subtitle=f"统计区间：{period}　|　数据口径以门店台账为准　|　本报告由 AI 生成，对外使用前请人工复核",
    )

    result = {
        "flow": "daily-report-generate-flow", "store_name": cur.get("store_name", ""),
        "period": period, "headline": headline, "key_metrics": keys,
        "day_anomalies": len(day_alerts), "reminder": reminder, "suggestion": suggestion,
        "human_confirm_required": bool((cur.get("delivery") or {}).get("publish_requires_human_confirm", True)),
        "artifacts": {"docx": docx, "xlsx": xlsx, "chart": img, "step1_dir": step1_dir},
        "generated_at": at.stamp(),
    }
    js = at.write_json(result, os.path.join(outdir, "report_flow.json"))
    result["artifacts"]["report_flow_json"] = js

    print(f"结论：{headline}")
    print(f"当日确认异常 {len(day_alerts)} 条｜关键数 {len(keys)} 个")
    for f in (docx, xlsx, js):
        print(" 产物:", f, f"({os.path.getsize(f) / 1024:.1f} KB)")
    at.emit(result)


if __name__ == "__main__":
    main()
