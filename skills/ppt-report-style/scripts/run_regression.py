# -*- coding: utf-8 -*-
"""回归测试：正例必须干净，反例必须被检出指定规则。

用法:
  python run_regression.py            # 运行全部检查
  python run_regression.py --keep     # 保留生成的临时文件
退出码非 0 表示回归失败。
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.normpath(os.path.join(HERE, ".."))

EXPECTED_NEGATIVE_RULES = (
    "TYP.MINSIZE", "PAL.COLOR", "PAL.CONTRAST", "FONT.EA",
    "TXT.OVERFLOW", "TXT.GLYPH", "IMG.RES", "LAY.REPEAT", "LAY.PURETEXT",
)


def run(args):
    completed = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", cwd=SKILL_ROOT)
    return completed.returncode, (completed.stdout or "") + (completed.stderr or "")


def check_deck(path, mode, report_path, extra=None):
    args = [os.path.join(HERE, "check_ppt.py"), path, "--mode", mode, "--json", report_path]
    args.extend(extra or [])
    code, output = run(args)
    with open(report_path, encoding="utf-8") as handle:
        return code, output, json.load(handle)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--keep", action="store_true", help="保留临时文件")
    args = parser.parse_args()

    temp_dir = tempfile.mkdtemp(prefix="ppt-regression-")
    failures = []
    print("回归测试目录:", temp_dir)

    positive = os.path.join(SKILL_ROOT, "assets", "example-slide.pptx")
    code, output = run([os.path.join(HERE, "build_example.py")])
    if code != 0:
        failures.append("正例生成失败：" + output.strip()[:200])
    print("[1/4] 重建正例", "OK" if code == 0 else "FAIL")

    for mode in ("balanced", "strict"):
        report = os.path.join(temp_dir, f"positive-{mode}.json")
        code, output, data = check_deck(positive, mode, report)
        counts = data["counts"]
        ok = counts["ERROR"] == 0 and counts["WARNING"] == 0
        print(f"[2/4] 正例 {mode}: {counts['ERROR']} 错误 / {counts['WARNING']} 警告 / "
              f"{counts['SUGGESTION']} 建议", "OK" if ok else "FAIL")
        if not ok:
            failures.append(f"正例 {mode} 不干净：\n" + output.strip()[:800])
        if len(data["layout"]) != 9:
            failures.append("正例页数与预期不一致")
        relations = {row["relation"] for row in data["layout"]}
        if len(relations) < 4:
            failures.append(f"正例关系模型不足：{relations}")

    negative = os.path.join(temp_dir, "negative.pptx")
    code, output = run([os.path.join(HERE, "build_example.py"), "--variant", "negative", "--out", negative])
    if code != 0:
        failures.append("反例生成失败：" + output.strip()[:200])
    report = os.path.join(temp_dir, "negative.json")
    code, output, data = check_deck(negative, "balanced", report)
    found = {item["rule"] for item in data["findings"]} | {item["rule"] for item in data["exceptions"]}
    missing = sorted(set(EXPECTED_NEGATIVE_RULES) - found)
    print("[3/4] 反例检出", "OK" if not missing else "FAIL", "缺失规则:" , missing or "无")
    print("       检出规则:", "、".join(sorted(found)))
    if missing:
        failures.append("反例未被检出：" + "、".join(missing))
    if code == 0:
        failures.append("反例应当返回非 0 退出码")

    preview = os.path.join(temp_dir, "preview.html")
    code, output = run([os.path.join(HERE, "preview_ppt.py"), positive, preview])
    ok = code == 0 and os.path.exists(preview) and os.path.getsize(preview) > 5000
    print("[4/4] 预览生成", "OK" if ok else "FAIL")
    if not ok:
        failures.append("预览生成失败：" + output.strip()[:200])

    print()
    if failures:
        print("回归失败：")
        for item in failures:
            print(" -", item)
        return 1
    print("回归全部通过：正例零错误零警告，反例命中全部预期规则。")
    if not args.keep:
        for name in os.listdir(temp_dir):
            os.remove(os.path.join(temp_dir, name))
        os.rmdir(temp_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())