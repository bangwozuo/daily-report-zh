# -*- coding: utf-8 -*-
"""
异常预警流程（anomaly-alert-flow）—— 端到端编排脚本。

编排的步骤：
  1. 异常检测   → 调用 ../../skills/anomaly-detect/scripts/detect.py 做四规则判定
  2. 推送编排   → 按级别过滤 + 静默时段 + 同源去重 + 值班人/升级规则
  3. 人工确认   → 高风险项必须人工确认后才可对外推送（不自动发送）
交付物：
  out/预警推送单.xlsx        推送清单 / 免打扰与升级 / 数据质量 / 汇总
  out/推送消息预览.md        实际要发出去的文案（群机器人 / 短信）
  out/alert_flow.json        流程结果（推送条数 + 静默状态 + 产物路径）
  out/steps/01-异常检测/*    第 1 步原始产物（异常预警清单 xlsx + 趋势图 png + detect.json）

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo --outdir out
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
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

LEVEL_ORDER = {"高": 0, "中": 1, "低": 2, "拦截": 3}
LEVEL_TAG = {"高": "🔴 高", "中": "🟡 中", "低": "🔵 低", "拦截": "⚪ 拦截"}


def run_step(cmd, label):
    print(f"[step] {label}")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        print(r.stdout or "", file=sys.stderr)
        print(r.stderr or "", file=sys.stderr)
        raise RuntimeError(f"步骤「{label}」失败，退出码 {r.returncode}；"
                           f"请先修复依赖或补齐输入后重跑（不跳过、不用旧数据顶替）")
    return r


def in_silent(hhmm_now, window):
    """判断当前是否处于免打扰时段，窗口形如 '23:00-07:00'（可跨零点）。"""
    try:
        s, e = window.split("-")
        now = int(hhmm_now[:2]) * 60 + int(hhmm_now[3:5])
        sh = int(s[:2]) * 60 + int(s[3:5])
        eh = int(e[:2]) * 60 + int(e[3:5])
    except Exception:
        return False
    return (sh <= now or now < eh) if sh > eh else (sh <= now < eh)


def main():
    ap = argparse.ArgumentParser(description="异常预警流程")
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

    # ---- 步骤 1：异常检测
    py = [sys.executable, STEP1, "--outdir", step1_dir]
    py = py + (["--input", a.input] if a.input else ["--demo"])
    run_step(py, "异常检测 / 四规则判定")
    det = at.read_json(os.path.join(step1_dir, "detect.json"))

    policy = cur.get("alert_policy") or {}
    push_levels = policy.get("push_levels") or ["高", "中"]
    channels = policy.get("channels") or {}
    silent = policy.get("silent_hours", "23:00-07:00")
    on_call = policy.get("on_call", "值班人未指定")
    dedup_h = int(policy.get("dedup_window_hours", 12) or 12)
    escalate = int(policy.get("escalate_if_no_ack_minutes", 30) or 30)

    now = dt.datetime.now()
    silent_now = in_silent(now.strftime("%H:%M"), silent)

    # ---- 步骤 2：推送编排（过滤级别 → 去重 → 免打扰判定）
    push, suppressed = [], []
    seen = {}
    for al in det["alerts"]:
        lv = al["级别"]
        if lv == "拦截" or al["判定"].startswith("特殊期"):
            suppressed.append({**al, "抑制原因": "特殊期（促销/节假日），仅记录不推送"})
            continue
        if lv == "低" or al["判定"].startswith("只报值"):
            suppressed.append({**al, "抑制原因": "样本量不足（订单 < 20 单），只报值不推送"})
            continue
        if lv not in push_levels:
            suppressed.append({**al, "抑制原因": f"级别「{lv}」不在推送级别 {push_levels} 内"})
            continue
        key = (al["指标"], al["判定"])
        last = seen.get(key)
        if last and (dt.date.fromisoformat(al["日期"]) - dt.date.fromisoformat(last)).days < dedup_h / 24:
            suppressed.append({**al, "抑制原因": f"同指标 {dedup_h} 小时去重窗口内已推送过"})
            continue
        seen[key] = al["日期"]
        push.append({
            "_lv": lv,
            "推送级别": LEVEL_TAG.get(lv, lv), "日期": al["日期"], "星期": al["星期"],
            "指标": al["指标"], "规则": al["规则"], "实际值": al["实际值"],
            "证据": al["证据"], "推送渠道": "、".join(channels.get(lv, ["企业微信群机器人"])),
            "需人工确认": "是" if lv == "高" else "否",
            "确认后才可对外发送": "是",
        })

    push.sort(key=lambda x: (LEVEL_ORDER.get(x["_lv"], 9), x["日期"]))
    for p in push:
        del p["_lv"]
    if silent_now:
        mode = f"免打扰时段（{silent}）内，改为次日 07:00 汇总推送"
    else:
        mode = "立即推送"

    rule_rows = [
        {"项": "推送级别", "内容": "、".join(push_levels)},
        {"项": "免打扰时段", "内容": silent},
        {"项": "当前是否免打扰", "内容": "是" if silent_now else "否"},
        {"项": "本次推送模式", "内容": mode},
        {"项": "高等级渠道", "内容": "、".join(channels.get("高", [])) or "—"},
        {"项": "中等级渠道", "内容": "、".join(channels.get("中", [])) or "—"},
        {"项": "值班人", "内容": on_call},
        {"项": "去重窗口", "内容": f"{dedup_h} 小时"},
        {"项": "未确认升级", "内容": f"{escalate} 分钟未确认，升级至门店负责人"},
        {"项": "人工确认", "内容": "所有对外推送必须先经人工确认，脚本不自动发送"},
    ]

    xlsx = at.write_excel(
        os.path.join(outdir, "预警推送单.xlsx"),
        {"推送清单": push or [{"推送级别": "（本期无可推送项）"}],
         "抑制清单": suppressed or [{"日期": "（无抑制项）"}],
         "免打扰与升级规则": rule_rows,
         "数据质量": det["data_quality"] or [{"项": "（无）"}],
         "汇总": [{"项": k, "内容": str(v)} for k, v in det["summary"].items()]},
        highlights={"推送清单": {"推送级别": "contains:高"}},
        widths={"推送清单": {"证据": 44}, "抑制清单": {"抑制原因": 34, "证据": 40}},
    )

    # ---- 推送文案预览
    lines = [f"# 推送消息预览（{mode}）", ""]
    if not push:
        lines.append("本期无可推送项：所有命中项均被规则抑制或级别不足。")
    for p in push:
        lines += [
            f"## [{p['推送级别']}] {p['日期']}（{p['星期']}）{p['指标']}",
            "",
            f"- 规则：{p['规则']}　实际值：{p['实际值']}",
            f"- 证据：{p['证据']}",
            f"- 渠道：{p['推送渠道']}　值班：{on_call}",
            f"- 文案：**【{p['日期']} {p['指标']}】{p['实际值']}，{p['证据']}。"
            f"请 {escalate} 分钟内回复处理人，否则升级至门店负责人。**",
            "",
        ]
    lines += ["---", "", "> AI 生成内容｜所有对外推送须人工确认后发送，脚本不自动发送。"]
    md = os.path.join(outdir, "推送消息预览.md")
    with open(md, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    result = {
        "flow": "anomaly-alert-flow",
        "store_name": det["store_name"],
        "period": det["period"],
        "detected": len(det["alerts"]),
        "pushed": len(push),
        "suppressed": len(suppressed),
        "silent_now": silent_now,
        "push_mode": mode,
        "on_call": on_call,
        "human_confirm_required": True,
        "artifacts": {"push_xlsx": xlsx, "push_md": md, "step1_dir": step1_dir,
                      "step1_files": sorted(os.listdir(step1_dir))},
        "generated_at": at.stamp(),
    }
    js = at.write_json(result, os.path.join(outdir, "alert_flow.json"))
    result["artifacts"]["alert_flow_json"] = js

    print(f"检测 {len(det['alerts'])} 条 → 推送 {len(push)} 条，抑制 {len(suppressed)} 条｜{mode}")
    for f in (xlsx, md, js):
        print(" 产物:", f, f"({os.path.getsize(f) / 1024:.1f} KB)")
    at.emit(result)


if __name__ == "__main__":
    main()
