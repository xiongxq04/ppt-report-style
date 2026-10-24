# -*- coding: utf-8 -*-
"""统一执行 PPT 结构与风格检查，并按错误 / 警告 / 建议 / 例外汇总。

职责划分：audit_ppt.py 管几何、布局骨架、配图、禁用对象、字形与备注；
verify_ppt_style.py 管字阶、配色与对比度、中西文双字体槽。本脚本只做调度、去重、例外过滤和汇总。

用法:
  python check_ppt.py <file.pptx> [--mode balanced|strict] [--exceptions exceptions.json]
                      [--preview-out preview.html] [--json report.json] [--layout-report]
主要选项与两个子脚本一致，见各自 --help。
"""
import argparse
import json
import os
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

from audit_ppt import audit, format_findings, format_layout
from palette_utils import ACCENT_DEFAULT
from verify_ppt_style import check as style_check

SEVERITY_ORDER = {"ERROR": 0, "WARNING": 1, "SUGGESTION": 2, "例外": 3}


def load_exceptions(path):
    if not path:
        return []
    with open(path, "r", encoding="utf-8") as handle:
        data = json.load(handle)
    if isinstance(data, dict):
        data = data.get("exceptions", [])
    rules = []
    for item in data:
        if not isinstance(item, dict) or not item.get("rule"):
            continue
        rules.append({"rule": item["rule"],
                      "slides": {int(value) for value in (item.get("slides") or [])},
                      "reason": item.get("reason", "未填写理由")})
    return rules


def apply_exceptions(findings, rules):
    """命中例外的条目降级为「例外」，而不是直接丢弃，保证可追溯。"""
    kept, excepted = [], []
    for item in findings:
        severity, slide_no, rule, kind, detail, extra = item
        hit = None
        for entry in rules:
            if entry["rule"] == rule and (not entry["slides"] or slide_no in entry["slides"]):
                hit = entry
                break
        if hit is None:
            kept.append(item)
        else:
            note = f"{extra}；例外理由：{hit['reason']}" if extra else f"例外理由：{hit['reason']}"
            excepted.append(("例外", slide_no, rule, kind, detail, note))
    return kept, excepted


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("file")
    parser.add_argument("--mode", choices=("balanced", "strict"), default="balanced")
    parser.add_argument("--min-size", type=float, default=16)
    parser.add_argument("--accent", default=ACCENT_DEFAULT)
    parser.add_argument("--contrast-color", default=None)
    parser.add_argument("--support", default=None)
    parser.add_argument("--fonts", default="宋体,楷体,黑体", help="中文字体（a:ea）白名单")
    parser.add_argument("--latin-fonts", default="Arial,Times New Roman", help="西文字体（a:latin）白名单")
    parser.add_argument("--sizes", default="16,18,20,30,36,44", help="允许的字阶档位")
    parser.add_argument("--max-steps", type=int, default=5)
    parser.add_argument("--allow-colors", default="")
    parser.add_argument("--allow-cjk-latin", action="store_true")
    parser.add_argument("--allow-smartart", action="store_true")
    parser.add_argument("--notes-min", type=int, default=None)
    parser.add_argument("--max-per-rule", type=int, default=3)
    parser.add_argument("--exceptions", default=None, help="例外记录 JSON")
    parser.add_argument("--preview-out", default=None)
    parser.add_argument("--json", dest="json_out", default=None)
    parser.add_argument("--layout-report", action="store_true")
    args = parser.parse_args()

    if not os.path.exists(args.file):
        print("文件不存在:", args.file)
        return 2

    notes_min = args.notes_min if args.notes_min is not None else (150 if args.mode == "strict" else 0)
    allow_colors = [item.strip() for item in args.allow_colors.split(",") if item.strip()]

    findings, layout_rows = audit(args.file, min_size=args.min_size, notes_min=notes_min,
                                  mode=args.mode, allow_smartart=args.allow_smartart)
    findings = set(findings)
    findings.update(style_check(
        args.file,
        accent=args.accent,
        min_size=args.min_size,
        fonts=args.fonts,
        allow=allow_colors,
        mode=args.mode,
        contrast_color=args.contrast_color,
        support=args.support,
        max_per_rule=args.max_per_rule,
        latin_fonts=args.latin_fonts,
        sizes=args.sizes,
        max_steps=args.max_steps,
        allow_cjk_latin=args.allow_cjk_latin,
    ))
    ordered = sorted(findings, key=lambda item: (item[1], SEVERITY_ORDER.get(item[0], 9), item[2]))

    excepted = []
    if args.exceptions:
        try:
            ordered, excepted = apply_exceptions(ordered, load_exceptions(args.exceptions))
        except Exception as error:
            print("例外文件读取失败:", error)
            return 2

    counts = {"ERROR": 0, "WARNING": 0, "SUGGESTION": 0}
    for severity, *_ in ordered:
        counts[severity] = counts.get(severity, 0) + 1

    print(f"文件: {os.path.abspath(args.file)}")
    print(f"模式: {args.mode}")
    print(f"结果: {counts['ERROR']} 个错误，{counts['WARNING']} 个警告，"
          f"{counts['SUGGESTION']} 个建议，{len(excepted)} 条例外")
    if ordered or excepted:
        print(format_findings(ordered + excepted))
    if args.layout_report:
        print(format_layout(layout_rows))

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as handle:
            json.dump({
                "file": os.path.abspath(args.file),
                "mode": args.mode,
                "counts": counts,
                "findings": [
                    {"rule": rule, "severity": severity, "slide": slide_no or None,
                     "kind": kind, "detail": detail, "suggestion": extra}
                    for severity, slide_no, rule, kind, detail, extra in ordered],
                "exceptions": [
                    {"rule": rule, "severity": severity, "slide": slide_no or None,
                     "kind": kind, "detail": detail, "note": extra}
                    for severity, slide_no, rule, kind, detail, extra in excepted],
                "layout": [
                    {"slide": slide_no, "relation": relation, "skeleton": skeleton,
                     "means": means, "pure_text": pure_text}
                    for slide_no, relation, skeleton, means, pure_text in layout_rows],
            }, handle, ensure_ascii=False, indent=2)

    exit_code = 1 if counts["ERROR"] else 0
    if args.preview_out:
        preview_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "preview_ppt.py")
        completed = subprocess.run([sys.executable, preview_script, args.file, args.preview_out], check=False)
        if completed.returncode != 0:
            print("预览生成失败")
            exit_code = exit_code or completed.returncode
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())