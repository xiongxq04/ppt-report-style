# -*- coding: utf-8 -*-
"""回归测试：正例必须零错误零警告，反例必须命中全部预期规则。

用法:
  python run_regression.py            # 运行全部检查
  python run_regression.py --keep     # 保留生成的临时文件
退出码非 0 表示回归失败。
"""
import argparse
import os
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

from profile_lib import preset_names

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.normpath(os.path.join(HERE, ".."))

# 反例（--variant negative）在 balanced 档下必须命中的规则；新增机审规则时同步扩充。
EXPECTED_NEGATIVE_RULES = (
    "TYP.MINSIZE", "TYP.LADDER", "TYP.POSTER",
    "PAL.COLOR", "PAL.CONTRAST", "PAL.ACCENT_PAGE", "PAL.CHART_SERIES",
    "FONT.EA", "FONT.LATIN_CJK", "FONT.TABLE_KAI", "FONT.EA_BOLD",
    "TXT.PREFIX", "TXT.CAVEAT", "TXT.OVERFLOW", "TXT.GLYPH",
    "IMG.RES", "LAY.REPEAT", "LAY.PURETEXT", "MED.SMARTART",
)


def run(args):
    completed = subprocess.run([sys.executable] + args, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", cwd=SKILL_ROOT)
    return completed.returncode, (completed.stdout or "") + (completed.stderr or "")


def check_deck(path, tier, report_path, preset="A"):
    args = [os.path.join(HERE, "check_ppt.py"), path, "--tier", tier, "--preset", preset,
            "--json", report_path]
    code, output = run(args)
    import json
    with open(report_path, encoding="utf-8") as handle:
        return code, output, json.load(handle)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--keep", action="store_true", help="保留生成的临时文件")
    args = parser.parse_args()

    temp_dir = tempfile.mkdtemp(prefix="ppt-regression-")
    failures = []
    print("回归测试目录:", temp_dir)

    positive = os.path.join(temp_dir, "positive.pptx")
    code, output = run([os.path.join(HERE, "build_example.py"), positive])
    if code != 0:
        failures.append("正例生成失败：" + output.strip()[:200])
    print("[1/6] 生成正例（academic）", "OK" if code == 0 else "FAIL")

    for tier in ("balanced", "strict"):
        report = os.path.join(temp_dir, f"positive-{tier}.json")
        code, output, data = check_deck(positive, tier, report)
        counts = data["counts"]
        ok = counts["ERROR"] == 0 and counts["WARNING"] == 0
        print(f"[2/6] 正例 {tier}: {counts['ERROR']} 错误 / {counts['WARNING']} 警告 / "
              f"{counts['SUGGESTION']} 建议", "OK" if ok else "FAIL")
        if not ok:
            failures.append(f"正例 {tier} 不干净：\n" + output.strip()[:800])
        if len(data["layout"]) != 6:
            failures.append(f"正例页数与预期不一致：{len(data['layout'])} != 6")
        relations = {row["relation"] for row in data["layout"]}
        if len(relations) < 3:
            failures.append(f"正例关系模型不足：{relations}")

    # 全部配色预设端到端冒烟：A 已在上方主流程覆盖，其余预设每套生成一份并通过 balanced 档
    for preset in tuple(p for p in preset_names() if p != "A"):
        deck = os.path.join(temp_dir, f"positive-{preset}.pptx")
        code, output = run([os.path.join(HERE, "build_example.py"), deck, "--preset", preset])
        report = os.path.join(temp_dir, f"positive-{preset}.json")
        code, output, data = check_deck(deck, "balanced", report, preset=preset) if code == 0 else (1, output, None)
        if data is None:
            failures.append(f"预设 {preset} 生成失败：" + output.strip()[:200])
            print(f"[2/4] 正例 preset {preset}: FAIL（生成失败）")
            continue
        counts = data["counts"]
        ok = counts["ERROR"] == 0 and counts["WARNING"] == 0
        print(f"[2/6] 正例 preset {preset}: {counts['ERROR']} 错误 / {counts['WARNING']} 警告",
              "OK" if ok else "FAIL")
        if not ok:
            failures.append(f"预设 {preset} 正例不干净：\n" + output.strip()[:800])

    negative = os.path.join(temp_dir, "negative.pptx")
    code, output = run([os.path.join(HERE, "build_example.py"), negative, "--variant", "negative"])
    if code != 0:
        failures.append("反例生成失败：" + output.strip()[:200])
    report = os.path.join(temp_dir, "negative.json")
    code, output, data = check_deck(negative, "balanced", report)
    found = {item["rule"] for item in data["findings"]} | {item["rule"] for item in data["exceptions"]}
    missing = sorted(set(EXPECTED_NEGATIVE_RULES) - found)
    print("[3/6] 反例检出", "OK" if not missing else "FAIL", "缺失规则:", missing or "无")
    print("       检出规则:", "、".join(sorted(found)))
    if missing:
        failures.append("反例未被检出：" + "、".join(missing))
    if code == 0:
        failures.append("反例应当返回非 0 退出码")

    guard = os.path.join(temp_dir, "guard.pptx")
    guard_build = run([os.path.join(HERE, "build_example.py"), guard, "--variant", "guard"])
    if guard_build[0] != 0:
        failures.append("正例变体生成失败：" + guard_build[1].strip()[:200])
    for tier in ("balanced", "strict"):
        report = os.path.join(temp_dir, f"guard-{tier}.json")
        if guard_build[0] != 0:
            print(f"[4/6] 正例变体 {tier}: FAIL（生成失败）")
            continue
        code, output, data = check_deck(guard, tier, report)
        counts = data["counts"]
        ok = counts["ERROR"] == 0 and counts["WARNING"] == 0
        print(f"[4/6] 正例变体 {tier}: {counts['ERROR']} 错误 / {counts['WARNING']} 警告",
              "OK" if ok else "FAIL")
        if not ok:
            failures.append(f"正例变体 {tier} 出现误报：\n" + output.strip()[:800])

    safety = os.path.join(temp_dir, "safety.pptx")
    code, output = run([os.path.join(HERE, "build_example.py"), safety, "--variant", "safety"])
    if code != 0:
        failures.append("安全例外反例生成失败：" + output.strip()[:200])
    report = os.path.join(temp_dir, "safety.json")
    code, output, data = check_deck(safety, "strict", report) if code == 0 else (1, output, None)
    if data is None:
        failures.append("安全例外反例无法检查：" + output.strip()[:200])
        print("[5/6] 安全例外拒绝: FAIL（生成失败）")
    else:
        severities = {item["rule"]: item["severity"] for item in data["findings"]}
        ok = (code != 0 and severities.get("EFF.SHADOW") == "ERROR"
              and severities.get("PAL.CONTRAST") == "ERROR"
              and data["counts"].get("例外", 0) == 0)
        print("[5/6] 安全例外拒绝", "OK" if ok else "FAIL")
        if not ok:
            failures.append("安全规则被例外降级或未按 ERROR 报出：\n" + output.strip()[:800])

    preview = os.path.join(temp_dir, "preview.html")
    code, output = run([os.path.join(HERE, "preview_ppt.py"), positive, preview])
    ok = code == 0 and os.path.exists(preview) and os.path.getsize(preview) > 1500
    print("[6/6] 预览生成", "OK" if ok else "FAIL")
    if not ok:
        failures.append("预览生成失败：" + output.strip()[:200])

    print()
    if failures:
        print("回归失败：")
        for item in failures:
            print(" -", item)
        return 1
    print("回归全部通过：正例与变体零误报，反例命中全部预期规则，安全例外被拒绝。")
    if not args.keep:
        for name in os.listdir(temp_dir):
            path = os.path.join(temp_dir, name)
            if os.path.isfile(path):
                os.remove(path)
        os.rmdir(temp_dir)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
