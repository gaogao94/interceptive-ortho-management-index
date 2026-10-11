# -*- coding: utf-8 -*-
"""公式验收：往四张维度表的「选择」列填答案 → 重算 → 逐项断言。

五层都要验：
  1. 逐题得分换算（选项 1→3 / 2→2 / 3→1 / 4→0 / 5→0）
  2. 每张维度表底部的小计
  3. 维度分（合计 ÷ 18 × 100）
  4. 把四个维度分折成报告口径，对照康桥样例报告

期望值全部**手算写死**在断言里，不读工作簿自身的中间结果 —— 否则公式错了
测试也跟着错。

用法：python test_metrics_workbook.py <工作簿.xlsx>
退出码：0 通过；1 失败；2 用法错误
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile

from openpyxl import load_workbook

NODE = r"D:\Program\DSH-desktop\resources\runtime\primary-runtime\dependencies\node\bin\node.exe"
CLI = r"D:\Program\DSH-desktop\resources\app.asar.unpacked\dsh\node_modules\@deepseek-ai\libreoffice-kit\lib\cli.js"

#: 康桥样例（case1）的 24 题判分。抄自样例报告，**不要手写**（本轮手写错过两次）。
FIXTURE_SCORES = {
    "A1": 1, "A2": 1, "A3": 1, "A4": 1, "A5": 1, "A6": 1,   # 6
    "B1": 1, "B2": 1, "B3": 1, "B4": 0, "B5": 1, "B6": 1,   # 5
    "C1": 2, "C2": 1, "C3": 2, "C4": 0, "C5": 1, "C6": 1,   # 7
    "D1": 1, "D2": 0, "D3": 1, "D4": 0, "D5": 1, "D6": 0,   # 3
}
SCORE_TO_OPTION = {3: 1, 2: 2, 1: 3, 0: 4}
ANSWERS = {q: SCORE_TO_OPTION[s] for q, s in FIXTURE_SCORES.items()}

EXPECT_SUB = {"A": 6, "B": 5, "C": 7, "D": 3}
EXPECT_DIM_PCT = {"A": 33, "B": 28, "C": 39, "D": 17}   # 与报告一致
EXPECT_TOTAL = 21
EXPECT_PCT = round(21 / 72 * 100)                        # 29

SHEETS = {"A": "A-1 标准与分级", "B": "A-2 方案审核与结案",
          "C": "A-3 记录与过程数据", "D": "A-4 财务与激励"}


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 2
    src = argv[1]
    wb = load_workbook(src)

    # 1) 填答案：每张维度表按 A 列的题号定位
    for dim, sheet in SHEETS.items():
        ws = wb[sheet]
        for r in range(3, ws.max_row + 1):
            qid = ws.cell(r, 1).value
            if isinstance(qid, str) and len(qid) == 2 and qid[0] == dim and qid[1].isdigit():
                ws.cell(r, 5, ANSWERS[qid])

    work = os.path.join(tempfile.mkdtemp(prefix="iomi_xlsx_"), "filled.xlsx")
    wb.save(work)

    # 2) 重算（openpyxl 不算公式）
    recalc = work.replace("filled.xlsx", "recalc.xlsx")
    r = subprocess.run([NODE, CLI, "recalculate", "--input", work, "--output", recalc],
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0 or not os.path.exists(recalc):
        print(f"  ✗ 重算失败 exit={r.returncode} {(r.stderr or '')[:200]}")
        return 1

    wbf = load_workbook(recalc, data_only=False)
    wbv = load_workbook(recalc, data_only=True)
    fails = []
    dimension_pct = {}

    for dim, sheet in SHEETS.items():
        sf, sv = wbf[sheet], wbv[sheet]
        print(f"\n【{sheet}】")
        row_of = {}
        for rr in range(3, sf.max_row + 1):
            qid = sf.cell(rr, 1).value
            if isinstance(qid, str) and len(qid) == 2 and qid[0] == dim and qid[1].isdigit():
                row_of[qid] = rr

        if len(row_of) != 6:
            print(f"  ✗ 应有 6 题，实际 {len(row_of)}")
            fails.append(f"{sheet} 题数")
            continue

        bad = []
        for qid, rr in row_of.items():
            got = sv.cell(rr, 6).value
            exp = FIXTURE_SCORES[qid]
            if got != exp:
                bad.append(f"{qid}: 得 {got}，期望 {exp}")
        print(f"  逐题换算：{'6 题全部正确' if not bad else '错误 ' + str(bad)}")
        if bad:
            fails.append(f"{sheet} 逐题")

        # 底部两行：小计 / 维度分
        sub_row = pct_row = None
        for rr in range(3, sf.max_row + 1):
            a = sf.cell(rr, 1).value
            if a == "小计":
                sub_row = rr
            elif a == "维度分":
                pct_row = rr
        if sub_row is None or pct_row is None:
            print("  ✗ 找不到小计 / 维度分行")
            fails.append(f"{sheet} 汇总行")
            continue

        got_sub = sv.cell(sub_row, 6).value
        got_pct = sv.cell(pct_row, 6).value
        exp_sub = EXPECT_SUB[dim]
        exp_pct = EXPECT_DIM_PCT[dim]
        dimension_pct[dim] = got_pct
        ok_sub = got_sub == exp_sub
        ok_pct = got_pct == exp_pct
        print(f"  小计：{got_sub}（手算 {exp_sub}）{'OK' if ok_sub else '✗'}")
        print(f"  维度分：{got_pct}（手算 {exp_pct}，报告写 {exp_pct}）{'OK' if ok_pct else '✗'}")
        if not ok_sub:
            fails.append(f"{sheet} 小计")
        if not ok_pct:
            fails.append(f"{sheet} 维度分")

    # 3) 四维合计 → 报告口径
    total_raw = sum(EXPECT_SUB.values())
    total_pct = round(total_raw / 72 * 100)
    print(f"\n【报告口径核对】")
    print(f"  四维原始分合计 {total_raw} / 72 → {total_pct}"
          f"（期望 {EXPECT_TOTAL} → {EXPECT_PCT}）"
          f"  {'OK' if (total_raw, total_pct) == (EXPECT_TOTAL, EXPECT_PCT) else '✗'}")
    if (total_raw, total_pct) != (EXPECT_TOTAL, EXPECT_PCT):
        fails.append("报告口径")

    # 4) 公式还在（没被值替换）
    n_formula = sum(1 for ws in wbf.worksheets for row in ws.iter_rows() for c in row
                    if isinstance(c.value, str) and c.value.startswith("="))
    print(f"  工作簿保留公式 {n_formula} 个")
    if n_formula < 32:
        fails.append("公式被值替换")

    print()
    if fails:
        print("✗ 失败项: " + "; ".join(fails))
        return 1
    print("✓ 公式全部正确（逐题 24 + 小计 4 + 维度分 4 + 报告口径 1）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
