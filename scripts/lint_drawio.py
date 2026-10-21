#!/usr/bin/env python3
"""Static geometry checks for uncompressed Drawio XML files."""

from __future__ import annotations

import argparse
import json
import math
import xml.etree.ElementTree as ET
from pathlib import Path


def style_map(text: str) -> dict[str, str]:
    result = {}
    for token in (text or "").split(";"):
        if "=" in token:
            key, value = token.split("=", 1)
            result[key] = value
        elif token:
            result[token] = "1"
    return result


def geometry(cell):
    return cell.find("mxGeometry")


def point(geom, kind):
    p = geom.find(f"mxPoint[@as='{kind}']")
    if p is None:
        return None
    return float(p.get("x", 0)), float(p.get("y", 0))


def edge_points(cell):
    geom = geometry(cell)
    if geom is None:
        return []
    result = []
    src = point(geom, "sourcePoint")
    dst = point(geom, "targetPoint")
    if src:
        result.append(src)
    array = geom.find("Array[@as='points']")
    if array is not None:
        for p in array.findall("mxPoint"):
            result.append((float(p.get("x", 0)), float(p.get("y", 0))))
    if dst:
        result.append(dst)
    # Hand-authored Drawio XML often repeats a bend point. It has no visual
    # meaning and must not become a spurious "very-short-segment" warning.
    deduped = []
    for p in result:
        if not deduped or math.dist(deduped[-1], p) > 1e-6:
            deduped.append(p)
    return deduped


def rect(cell):
    geom = geometry(cell)
    if geom is None:
        return None
    return tuple(float(geom.get(k, 0)) for k in ("x", "y", "width", "height"))


def is_text_only(cell):
    st = style_map(cell.get("style", ""))
    return "text" in st or (st.get("strokeColor") == "none" and st.get("fillColor") == "none")


def segment_crosses_rect(a, b, r, inset=2.0):
    x, y, w, h = r
    left, right = x + inset, x + w - inset
    top, bottom = y + inset, y + h - inset
    if right <= left or bottom <= top:
        return False
    x1, y1 = a
    x2, y2 = b
    if abs(x1 - x2) < 1e-6:
        return left < x1 < right and max(min(y1, y2), top) < min(max(y1, y2), bottom)
    if abs(y1 - y2) < 1e-6:
        return top < y1 < bottom and max(min(x1, x2), left) < min(max(x1, x2), right)
    # Liang-Barsky clipping for diagonal segments.
    dx, dy = x2 - x1, y2 - y1
    p = (-dx, dx, -dy, dy)
    q = (x1 - left, right - x1, y1 - top, bottom - y1)
    u1, u2 = 0.0, 1.0
    for pi, qi in zip(p, q):
        if abs(pi) < 1e-9:
            if qi < 0:
                return False
        else:
            t = qi / pi
            if pi < 0:
                u1 = max(u1, t)
            else:
                u2 = min(u2, t)
    return u1 < u2 and u2 > 0 and u1 < 1


def point_on_boundary(p, r, tolerance=8.0):
    x, y = p
    rx, ry, rw, rh = r
    within_x = rx - tolerance <= x <= rx + rw + tolerance
    within_y = ry - tolerance <= y <= ry + rh + tolerance
    on_vertical = abs(x - rx) <= tolerance or abs(x - (rx + rw)) <= tolerance
    on_horizontal = abs(y - ry) <= tolerance or abs(y - (ry + rh)) <= tolerance
    return (within_y and on_vertical) or (within_x and on_horizontal)


def lint(path: Path) -> dict:
    root = ET.parse(path).getroot()
    model = root.find(".//mxGraphModel")
    canvas = {
        "width": float(model.get("pageWidth", 0)) if model is not None else 0,
        "height": float(model.get("pageHeight", 0)) if model is not None else 0,
    }
    cells = root.findall(".//mxCell")
    vertices = [c for c in cells if c.get("vertex") == "1" and rect(c)]
    shapes = [(c.get("id"), rect(c)) for c in vertices if not is_text_only(c)]
    labels = [(c.get("id"), c.get("value", ""), rect(c), style_map(c.get("style", ""))) for c in vertices if is_text_only(c) and c.get("value", "")]
    edges = [c for c in cells if c.get("edge") == "1"]
    issues = []
    edge_keys = {}
    diagram_kind = "sequence" if "sequence" in path.stem.lower() else "generic"

    for cell in edges:
        cid = cell.get("id")
        st = style_map(cell.get("style", ""))
        pts = edge_points(cell)
        if len(pts) < 2:
            issues.append({"severity": "error", "cell": cid, "kind": "missing-edge-points"})
            continue
        key = tuple(round(v, 2) for p in pts for v in p)
        if key in edge_keys:
            issues.append({"severity": "error", "cell": cid, "kind": "duplicate-edge", "other": edge_keys[key]})
        edge_keys[key] = cid

        if cell.get("value", "").strip():
            issues.append({"severity": "warning", "cell": cid, "kind": "edge-has-inline-label"})
        if st.get("labelBackgroundColor") not in (None, "none") and cell.get("value", "").strip():
            issues.append({"severity": "error", "cell": cid, "kind": "edge-label-can-mask-line"})

        for a, b in zip(pts, pts[1:]):
            if math.dist(a, b) < 10:
                issues.append({"severity": "warning", "cell": cid, "kind": "very-short-segment", "segment": [a, b]})
            for label_id, label_text, label_rect, _ in labels:
                if segment_crosses_rect(a, b, label_rect):
                    issues.append({
                        "severity": "error",
                        "cell": cid,
                        "kind": "connector-crosses-text",
                        "label_cell": label_id,
                        "label": label_text.replace("&#10;", " / ")[:80],
                    })

        # Sequence-message endpoints intentionally terminate on lifelines,
        # which are not ordinary rectangular vertices. Their horizontal
        # alignment is a visual/export QA concern, not a detached-arrow error.
        if st.get("endArrow", "none") != "none" and diagram_kind != "sequence":
            if not any(point_on_boundary(pts[-1], r) for _, r in shapes):
                issues.append({"severity": "error", "cell": cid, "kind": "arrow-target-detached", "point": pts[-1]})
            if not any(point_on_boundary(pts[0], r) for _, r in shapes):
                issues.append({"severity": "warning", "cell": cid, "kind": "arrow-source-detached", "point": pts[0]})

    for label_id, label_text, label_rect, st in labels:
        font_size = float(st.get("fontSize", 0) or 0)
        if font_size and font_size < 14:
            issues.append({"severity": "warning", "cell": label_id, "kind": "small-text", "font_size": font_size, "label": label_text[:80]})
        x, y, w, h = label_rect
        if canvas["width"] and (x < 0 or x + w > canvas["width"] or y < 0 or y + h > canvas["height"]):
            issues.append({"severity": "error", "cell": label_id, "kind": "text-outside-canvas", "label": label_text[:80]})

    return {
        "file": str(path.resolve()),
        "canvas": canvas,
        "counts": {"vertices": len(vertices), "labels": len(labels), "edges": len(edges)},
        "summary": {
            "errors": sum(i["severity"] == "error" for i in issues),
            "warnings": sum(i["severity"] == "warning" for i in issues),
        },
        "issues": issues,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("inputs", nargs="+", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--fail-on-error", action="store_true")
    args = parser.parse_args()
    results = [lint(p) for p in args.inputs]
    payload = results[0] if len(results) == 1 else results
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    for result in results:
        print(f"{Path(result['file']).name}: {result['summary']}")
    if args.fail_on_error and any(r["summary"]["errors"] for r in results):
        raise SystemExit(2)


if __name__ == "__main__":
    main()
