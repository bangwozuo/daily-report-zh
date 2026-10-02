# -*- coding: utf-8 -*-
"""
周度经营建议流程（weekly-ops-advice-flow）—— 端到端编排脚本。

编排的步骤：
  1. 异常检测     → 调用 ../../skills/anomaly-detect/scripts/detect.py 取本周确认异常
  2. 周度汇总     → 本周 vs 上周（营收/订单/访客/退款）+ 派生指标 + 按星期几聚合
  3. 聚焦一个动作 → 按优先级规则从本周问题里选出**唯一一个**重点动作（零预算约束）
交付物：
  out/周度经营建议.docx      周度概览 / 周环比 / 同星期结构 / 本周异常 / 下周一个重点动作
  out/周度汇总.xlsx          周度概览 / 周环比 / 星期结构 / 本周异常 / 汇总
  out/周对比图.png           本周 vs 上周 的周环比（%）
  out/weekly_flow.json       流程结果（周环比 + 重点动作 + 产物路径）
  out/steps/01-异常检测/*     第 1 步原始产物

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo --outdir out
"""
from __future__ import annotations

import argparse
import datetime as dt
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
WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def run_step(cmd, label):
    print(f"[step] {label}")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        print(r.stdout or "", file=sys.stderr)
        print(r.stderr or "", file=sys.stderr)
        raise RuntimeError(f"步骤「{label}」失败，退出码 {r.returncode}；"
                           f"先修复依赖或补齐输入，不要手工拼一个周报")
    return r


def week_agg(rows):
    rev = sum(r["revenue"] for r in rows)
    orders = sum(r["orders"] for r in rows)
    visitors = sum(r["visitors"] for r in rows)
    refunds = sum(r["refunds"] for r in rows)
    return {
        "days": len(rows), "revenue": rev, "orders": orders,
        "visitors": visitors, "refunds": refunds,
        "aov": round(rev / orders, 2) if orders else 0,
        "conversion": round(orders / visitors * 100, 2) if visitors else 0,
        "refund_rate": round(refunds / rev * 100, 2) if rev else 0,
    }


def main():
    ap = argparse.ArgumentParser(description="周度经营建议流程")
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
    run_step(py, "异常检测 / 本周异常判定")
    det = at.read_json(os.path.join(step1_dir, "detect.json"))

    period = cur.get("period") or {}
    s, e = period.get("start"), period.get("end")
    week_rows = [r for r in cur["metrics"] if s <= r["date"] <= e]
    if not week_rows:
        print(f"[错误] 区间 {s}~{e} 内没有任何指标数据，无法出周报。", file=sys.stderr)
        sys.exit(3)

    this_w = week_agg(week_rows)
    lw = cur.get("last_week") or {}
    last_w = {"revenue": lw.get("revenue"), "orders": lw.get("orders"),
              "visitors": lw.get("visitors"), "refunds": lw.get("refunds")}

    wob = {}
    for k, label in (("revenue", "营收"), ("orders", "订单"), ("visitors", "访客"), ("refunds", "退款额")):
        prev = last_w.get(k)
        wob[label] = round(at.pct_change(this_w[k], prev), 2) if prev else None

    # 按星期几聚合（周末效应会造成 30%~50% 的正常波动，必须按星期几对齐）
    wd = {}
    for r in week_rows:
        w = WEEKDAY_CN[dt.date.fromisoformat(r["date"]).weekday()]
        d = wd.setdefault(w, {"营收": 0, "订单": 0, "天数": 0})
        d["营收"] += r["revenue"]
        d["订单"] += r["orders"]
        d["天数"] += 1
    wd_rows = [{"星期": k, "天数": v["天数"], "营收": v["营收"], "订单": v["订单"],
                "日均营收": round(v["营收"] / v["天数"], 0)} for k, v in wd.items()]

    # 本周确认异常（引用第 1 步结果）
    week_alerts = [al for al in det["alerts"]
                   if s <= al["日期"] <= e and al["判定"] == "确认异常"]
    high_alerts = [al for al in week_alerts if al["级别"] == "高"]

    # ---- 聚焦唯一一个重点动作（优先级：高异常 → 退款率 → 转化率 → 营收下滑 → 维持）
    budget = (cur.get("action_constraints") or {}).get("budget_ceiling", 0)
    owner = (cur.get("action_constraints") or {}).get("owner", "店长")
    if high_alerts:
        top = high_alerts[0]
        focus_action = (f"让{owner}把「{top['日期']} {top['指标']}（{top['实际值']}）」当作下周第一件事："
                        f"拉出该指标对应的原始单据逐条核对，先定性再动手。")
        focus_reason = f"本周存在 {len(high_alerts)} 条高等级确认异常，优先级最高"
    elif this_w["refund_rate"] > 3:
        focus_action = (f"让{owner}把退款原因分布做出来（按单品 / 时段 / 配送），"
                        f"下周只改排名第一的那一项。")
        focus_reason = f"本周退款率 {this_w['refund_rate']}% 高于健康线 3%"
    elif this_w["conversion"] < 2:
        focus_action = (f"让{owner}暂停最低效的一个投放渠道，把预算挪到进店转化最高的渠道，先跑一周。")
        focus_reason = f"本周转化率 {this_w['conversion']}% 低于健康下限 2%"
    elif wob["营收"] is not None and wob["营收"] <= -10:
        focus_action = (f"让{owner}对上周到店的老客做一次零成本召回（社群 + 朋友圈文案），"
                        f"目标把复购客单提升 5%。")
        focus_reason = f"本周营收周环比 {wob['营收']}%，跌幅超过 10%"
    else:
        focus_action = f"让{owner}维持现有排班与活动节奏，下周同一时间再对比一次转化率与客单价。"
        focus_reason = "各指标均在健康区间，本周不引入新动作"

    if budget == 0 and "预算" not in focus_action:
        focus_action += "（零预算：不动用投放费用）"

    # ---- 产物
    png = at.bar_chart(
        os.path.join(outdir, "周对比图.png"),
        list(wob.keys()),
        [0 if v is None else v for v in wob.values()],
        title=f"{cur.get('store_name', '')} 本周 vs 上周 周环比（%）",
        xlabel="指标", ylabel="周环比%", figsize=(8.6, 4.4),
    )
    xlsx = at.write_excel(
        os.path.join(outdir, "周度汇总.xlsx"),
        {"周度概览": [
            {"指标": "统计天数", "本周": this_w["days"], "上周": last_w.get("days", "—")},
            {"指标": "营收", "本周": this_w["revenue"], "上周": last_w.get("revenue")},
            {"指标": "订单", "本周": this_w["orders"], "上周": last_w.get("orders")},
            {"指标": "访客", "本周": this_w["visitors"], "上周": last_w.get("visitors")},
            {"指标": "退款额", "本周": this_w["refunds"], "上周": last_w.get("refunds")},
            {"指标": "客单价", "本周": this_w["aov"], "上周": "—"},
            {"指标": "转化率%", "本周": this_w["conversion"], "上周": "—"},
            {"指标": "退款率%", "本周": this_w["refund_rate"], "上周": "—"},
        ],
         "周环比": [{"指标": k, "周环比%": "—" if v is None else v} for k, v in wob.items()],
         "星期结构": wd_rows,
         "本周异常": week_alerts or [{"日期": "（本周无确认异常）"}],
         "汇总": [{"项": "重点动作", "内容": focus_action},
                  {"项": "选取理由", "内容": focus_reason},
                  {"项": "责任人", "内容": owner},
                  {"项": "预算上限", "内容": budget}]},
        highlights={"周环比": {"周环比%": "<-10"}, "本周异常": {"级别": "contains:高"}},
        widths={"汇总": {"内容": 60}},
    )
    docx = at.write_docx(
        os.path.join(outdir, "周度经营建议.docx"),
        f"{cur.get('store_name', '')} 周度经营建议（{s} ~ {e}）",
        [
            {"heading": "一、周度概览",
             "table": {"cols": ["指标", "本周", "上周", "周环比"],
                       "rows": [["营收", f"¥{this_w['revenue']:,}", f"¥{last_w.get('revenue'):,}"
                                 if last_w.get("revenue") else "—",
                                 "—" if wob["营收"] is None else f"{wob['营收']:+.1f}%"],
                                ["订单", f"{this_w['orders']:,} 单",
                                 f"{last_w.get('orders'):,} 单" if last_w.get("orders") else "—",
                                 "—" if wob["订单"] is None else f"{wob['订单']:+.1f}%"],
                                ["访客", f"{this_w['visitors']:,}",
                                 f"{last_w.get('visitors'):,}" if last_w.get("visitors") else "—",
                                 "—" if wob["访客"] is None else f"{wob['访客']:+.1f}%"],
                                ["退款额", f"¥{this_w['refunds']:,}", f"¥{last_w.get('refunds'):,}"
                                 if last_w.get("refunds") else "—",
                                 "—" if wob["退款额"] is None else f"{wob['退款额']:+.1f}%"],
                                ["客单价", f"¥{this_w['aov']}", "—", "—"],
                                ["转化率", f"{this_w['conversion']}%", "—",
                                 f"健康区间 2%~5%"],
                                ["退款率", f"{this_w['refund_rate']}%", "—", "健康线 < 3%"]]}},
            {"heading": "二、按星期几看结构（周末效应会造成 30%+ 正常波动，必须对齐星期）",
             "table": {"cols": ["星期", "天数", "营收", "订单", "日均营收"],
                       "rows": [[r["星期"], r["天数"], f"¥{r['营收']:,}", r["订单"],
                                 f"¥{r['日均营收']:,.0f}"] for r in wd_rows]}},
            {"heading": "三、本周确认异常",
             "table": {"cols": ["日期", "指标", "规则", "级别", "实际值", "证据"],
                       "rows": [[al["日期"], al["指标"], al["规则"], al["级别"], al["实际值"],
                                 al["证据"]] for al in week_alerts] or [["—", "无", "—", "—", "—", "—"]]}},
            {"heading": "四、下周一个重点动作", "bullets": [focus_action],
             "paras": [f"选取理由：{focus_reason}｜责任人：{owner}｜预算上限：{budget}"]},
            {"heading": "五、周环比图", "image": png},
        ],
        subtitle="本建议由 AI 生成，仅给出一个重点动作；执行前请门店负责人人工确认",
    )

    result = {
        "flow": "weekly-ops-advice-flow", "store_name": cur.get("store_name", ""),
        "period": period, "week_metrics": this_w, "wow_pct": wob,
        "week_alerts": len(week_alerts), "high_alerts": len(high_alerts),
        "focus_action": focus_action, "focus_reason": focus_reason,
        "owner": owner, "budget_ceiling": budget, "actions_count": 1,
        "human_confirm_required": True,
        "artifacts": {"docx": docx, "xlsx": xlsx, "chart": png, "step1_dir": step1_dir},
        "generated_at": at.stamp(),
    }
    js = at.write_json(result, os.path.join(outdir, "weekly_flow.json"))
    result["artifacts"]["weekly_flow_json"] = js

    rev_wow = "—" if wob["营收"] is None else f"{wob['营收']:+.1f}%"
    print(f"本周 {this_w['days']} 天｜营收 ¥{this_w['revenue']:,}（周环比 {rev_wow}）"
          f"｜退款率 {this_w['refund_rate']}%｜转化率 {this_w['conversion']}%")
    print(f"本周确认异常 {len(week_alerts)} 条（高 {len(high_alerts)}）｜重点动作 1 个")
    for f in (docx, xlsx, png, js):
        print(" 产物:", f, f"({os.path.getsize(f) / 1024:.1f} KB)")
    at.emit(result)


if __name__ == "__main__":
    main()
