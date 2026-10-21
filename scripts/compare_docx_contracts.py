#!/usr/bin/env python3
"""Compare stable layout/style invariants between two DOCX contract JSON files."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def add_diff(diffs, severity, path, expected, actual):
    if expected != actual:
        diffs.append({"severity": severity, "path": path, "expected": expected, "actual": actual})


def compare(reference: dict, output: dict) -> dict:
    diffs = []
    add_diff(diffs, "error", "counts.sections", reference["counts"]["sections"], output["counts"]["sections"])
    for i, ref_sec in enumerate(reference.get("sections", [])):
        if i >= len(output.get("sections", [])):
            break
        out_sec = output["sections"][i]
        for key in (
            "orientation", "page_width_pt", "page_height_pt", "top_margin_pt",
            "bottom_margin_pt", "left_margin_pt", "right_margin_pt",
            "header_distance_pt", "footer_distance_pt", "different_first_page",
            "header_text", "footer_text",
        ):
            add_diff(diffs, "error", f"sections[{i}].{key}", ref_sec.get(key), out_sec.get(key))

    for style_name, ref_style in reference.get("styles", {}).items():
        if ref_style.get("used_count", 0) == 0:
            continue
        out_style = output.get("styles", {}).get(style_name)
        if out_style is None:
            add_diff(diffs, "error", f"styles.{style_name}", "present", "missing")
            continue
        for group in ("font", "paragraph_format"):
            if ref_style.get(group) is None:
                continue
            for key, expected in ref_style[group].items():
                add_diff(diffs, "warning", f"styles.{style_name}.{group}.{key}", expected, out_style.get(group, {}).get(key))

    for key in ("tables", "inline_images", "floating_images", "content_controls"):
        add_diff(diffs, "info", f"counts.{key}", reference["counts"].get(key), output["counts"].get(key))
    add_diff(diffs, "info", "cell_shading_counts", reference.get("cell_shading_counts"), output.get("cell_shading_counts"))

    return {
        "reference": reference["source"],
        "output": output["source"],
        "summary": {
            "errors": sum(d["severity"] == "error" for d in diffs),
            "warnings": sum(d["severity"] == "warning" for d in diffs),
            "info": sum(d["severity"] == "info" for d in diffs),
        },
        "diffs": diffs,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("reference", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--fail-on-error", action="store_true")
    args = parser.parse_args()
    ref = json.loads(args.reference.read_text(encoding="utf-8"))
    out = json.loads(args.output.read_text(encoding="utf-8"))
    result = compare(ref, out)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(text, encoding="utf-8")
    print(json.dumps(result["summary"], ensure_ascii=False))
    if args.fail_on_error and result["summary"]["errors"]:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
