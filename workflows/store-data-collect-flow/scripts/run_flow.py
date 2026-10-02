# -*- coding: utf-8 -*-
"""
门店数据采集流程（store-data-collect-flow）—— 端到端编排脚本。

编排的步骤：
  1. 门店数据对接   → 调用 ../../skills/store-data-sync/scripts/sync.py 做接入体检
  2. 接入门禁       → 有红线源 / 未授权源 → 阻断下游；否则放行
  3. 归集与交付     → 产出《采集归集单》（Excel）+ collect.json（供下游读取）

用法：
  python run_flow.py --input input.json --outdir out
  python run_flow.py --demo --outdir out

产物：
  out/采集归集单.xlsx     归集结果 / 门禁判定 / 源清单 / 断供明细 / 汇总
  out/collect.json        流程结果（门禁状态 + 各源状态 + 产物路径）
  out/steps/01-接入体检/*  第 1 步产出的原始产物（接入清单 xlsx + 2 张 PNG + sync.json）
"""
from __future__ import annotations

import argparse
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

STEP1 = os.path.join(REPO, "skills", "store-data-sync", "scripts", "sync.py")


def run_step(cmd, label):
    """执行一步；失败时打印修复提示并抛出，不静默跳过。"""
    print(f"[step] {label}: {' '.join(cmd)}")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        print(r.stdout or "", file=sys.stderr)
        print(r.stderr or "", file=sys.stderr)
        raise RuntimeError(f"步骤「{label}」失败，退出码 {r.returncode}；"
                           f"请先按上方提示修复依赖或输入后重跑")
    return r


def main():
    ap = argparse.ArgumentParser(description="门店数据采集流程")
    ap.add_argument("--input", help="输入 JSON 路径")
    ap.add_argument("--outdir", default="out")
    ap.add_argument("--demo", action="store_true", help="用内置样例跑一遍")
    a = ap.parse_args()

    if a.demo:
        cur = json.loads(open(os.path.join(FLOW_DIR, "examples", "input.json"),
                              encoding="utf-8").read())
    elif a.input:
        cur = at.read_json(a.input)
    else:
        ap.error("需提供 --input 或 --demo")

    outdir = at.ensure_outdir(a.outdir)
    step1_dir = os.path.join(outdir, "steps", "01-接入体检")

    # ---- 步骤 2：门店数据对接（调用原子技能脚本）
    run_step([sys.executable, STEP1, "--outdir", step1_dir,
              "--input", a.input] if a.input else
             [sys.executable, STEP1, "--demo", "--outdir", step1_dir],
             "门店数据对接 / 接入体检")
    sync = at.read_json(os.path.join(step1_dir, "sync.json"))

    # ---- 步骤 2：接入门禁
    red = [s for s in sync["sources"] if s["判定"].startswith("🔴")]
    unauth = [s for s in sync["sources"] if s["判定"].startswith("🟡")]
    if red:
        gate, gate_reason = "BLOCK", f"存在 {len(red)} 个红线数据源，禁止进入下游流程"
    elif unauth:
        gate, gate_reason = "WARN", f"存在 {len(unauth)} 个未授权/字段不全源，下游须标注口径缺失"
    else:
        gate, gate_reason = "PASS", "全部数据源合规且字段完整"

    policy = cur.get("collect_policy") or {}
    gate_rows = [
        {"门禁项": "红线数据源", "数量": len(red),
         "明细": "、".join(s["数据源"] for s in red) or "—",
         "处置": policy.get("on_redline", "停止该源并告警")},
        {"门禁项": "未授权 / 字段不全", "数量": len(unauth),
         "明细": "、".join(s["数据源"] for s in unauth) or "—",
         "处置": "按降级路径补齐，或在日报中标注缺数"},
        {"门禁项": "断供天数", "数量": sync["summary"]["断供天数"],
         "明细": sync["summary"]["断供日期"], "处置": "不补零；在日报中显式标注"},
        {"门禁项": "脱敏红线", "数量": sync["summary"]["脱敏红线条数"],
         "明细": "；".join(sync["redlines"]), "处置": "从数据集删除字段后重新归集"},
        {"门禁项": "下游是否放行", "数量": 1 if gate != "BLOCK" else 0,
         "明细": gate, "处置": gate_reason},
    ]

    collect_rows = []
    for s in sync["sources"]:
        collect_rows.append({
            "数据源": s["数据源"], "接入方式": s["接入方式"], "字段完整率%": s["字段完整率%"],
            "缺失字段": s["缺失字段"], "归集状态": ("已归集" if s["判定"].startswith("✅")
                                                 else ("待补齐" if s["判定"].startswith("🟡") else "已剔除")),
            "进入下游": "是" if s["判定"].startswith(("✅", "🟡")) else "否（红线）",
        })

    gaps = [d for d in sync["daily_arrival"] if d["状态"] != "✅ 全到"]

    xlsx = at.write_excel(
        os.path.join(outdir, "采集归集单.xlsx"),
        {"归集结果": collect_rows, "门禁判定": gate_rows,
         "源清单": sync["sources"],
         "断供明细": gaps or [{"日期": "（无断供）", "应到源数": "", "实到源数": "",
                              "覆盖率%": "", "状态": ""}],
         "汇总": [{"项": k, "内容": str(v)} for k, v in sync["summary"].items()]},
        highlights={"门禁判定": {"数量": ">0"}, "归集结果": {"归集状态": "contains:待补齐"}},
        widths={"门禁判定": {"明细": 44, "处置": 40}},
    )

    result = {
        "flow": "store-data-collect-flow",
        "store_name": sync["store_name"],
        "gate": gate,
        "gate_reason": gate_reason,
        "downstream_allowed": gate != "BLOCK",
        "redline_sources": [s["数据源"] for s in red],
        "degraded_sources": [s["数据源"] for s in unauth],
        "gap_days": sync["summary"]["断供日期"],
        "pii_redlines": sync["summary"]["脱敏红线条数"],
        "artifacts": {"collect_xlsx": xlsx,
                      "step1_dir": step1_dir,
                      "step1_files": sorted(os.listdir(step1_dir))},
        "generated_at": at.stamp(),
    }
    js = at.write_json(result, os.path.join(outdir, "collect.json"))
    result["artifacts"]["collect_json"] = js

    print(f"门禁：{gate} —— {gate_reason}")
    print(f"归集：{len(collect_rows)} 个数据源（红线剔除 {len(red)}）｜断供 {sync['summary']['断供天数']} 天"
          f"｜脱敏红线 {sync['summary']['脱敏红线条数']} 条")
    print(" 产物:", xlsx, f"({os.path.getsize(xlsx) / 1024:.1f} KB)")
    print(" 产物:", js, f"({os.path.getsize(js) / 1024:.1f} KB)")
    at.emit(result)


if __name__ == "__main__":
    main()
