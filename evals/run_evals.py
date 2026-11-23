#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""回归测试：用 fixture 验证审计器的判定与重排器的效果。

    python evals/make_fixtures.py && python evals/run_evals.py

检查四件事：
1. bad.docx 必须报出严重项（退出码 1）
2. bad.docx 里 T1/T4/T8/T9/W1/W4/W11 都必须命中（不能漏判）
3. good.docx 必须严重项为 0（退出码 0，不能误判合规文档）
4. restyle 之后 bad.docx 的排印类严重项（T*）必须清零，文字类（W*）保留
"""
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
AUDIT = os.path.join(ROOT, "scripts", "audit_docx.py")
RESTYLE = os.path.join(ROOT, "scripts", "restyle_docx.py")
FIX = os.path.join(HERE, "fixtures")


def run(script, *args):
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    proc = subprocess.run([sys.executable, script, *args],
                          capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env)
    return proc.returncode, proc.stdout, proc.stderr


def audit(path):
    code, out, err = run(AUDIT, path, "--json")
    if not out.strip():
        raise RuntimeError(f"audit 无输出：{err}")
    return code, json.loads(out)


def main():
    bad = os.path.join(FIX, "bad.docx")
    good = os.path.join(FIX, "good.docx")
    for f in (bad, good):
        if not os.path.exists(f):
            print(f"缺少 fixture：{f}\n先跑 python evals/make_fixtures.py")
            return 2

    failures = []

    # 1 + 2：坏文档要报出来，且关键编号不能漏
    code, data = audit(bad)
    codes = {f["code"] for f in data["findings"]}
    if code != 1:
        failures.append(f"bad.docx 退出码应为 1，实为 {code}")
    for need in ("T1", "T4", "T8", "T9", "W1", "W4", "W11"):
        if need not in codes:
            failures.append(f"bad.docx 漏判 {need}（实际：{sorted(codes)}）")

    # 3：好文档不能误判
    code2, data2 = audit(good)
    if code2 != 0 or data2["critical"] != 0:
        msgs = [f'{f["code"]} {f["issue"]}' for f in data2["findings"]
                if f["severity"] == "严重"]
        failures.append(f"good.docx 误判严重项：{msgs}")

    # 4：重排只清排印项，文字项保留
    with tempfile.TemporaryDirectory() as td:
        out = os.path.join(td, "restyled.docx")
        rc, so, se = run(RESTYLE, bad, "--style", "board", "--out", out, "--quiet")
        if rc != 0 or not os.path.exists(out):
            failures.append(f"restyle 失败：{se}")
        else:
            _, data3 = audit(out)
            left = {f["code"] for f in data3["findings"]
                    if f["severity"] == "严重"}
            still_t = sorted(c for c in left if c.startswith("T"))
            if still_t:
                failures.append(f"重排后仍有排印严重项：{still_t}")
            if not any(c.startswith("W") for c in left):
                failures.append("重排后文字类问题不该消失（它只改样式层）")

    if failures:
        print("✗ 失败")
        for f in failures:
            print("  -", f)
        return 1

    print("✓ 全部通过")
    print(f"  bad.docx：严重 {data['critical']} / 提示 {data['hint']}"
          f"　命中 {len(codes)} 类")
    print(f"  good.docx：严重 {data2['critical']} / 提示 {data2['hint']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
