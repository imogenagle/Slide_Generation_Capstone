#!/usr/bin/env python3
"""Build bundle-eval comparison tables for an arbitrary set of methods."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any


ALL_METRICS = {
    "core_coverage_topic_iou": ("core_coverage", "topic_iou"),
    "geometry_aware_density_gad_geom": ("geometry_aware_density", "gad_geom"),
    "visual_appeal_deck_score": ("visual_appeal", "deck_score"),
    "layout_defects_deck_score": ("layout_defects", "deck_score"),
    "logical_flow_deck_score": ("logical_flow", "deck_score"),
    "paper_faithfulness_deck_score": ("paper_faithfulness", "deck_score"),
}


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float | None:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def normalize_paper_key(paper_id: str | None, summary_path: Path) -> str:
    if paper_id:
        return paper_id.replace(":", "_")
    return summary_path.parent.name


def parse_method_spec(value: str) -> tuple[str, Path]:
    if "=" not in value:
        raise argparse.ArgumentTypeError(
            f"Invalid --method value '{value}'. Expected LABEL=/path/to/method_eval_root."
        )
    label, root = value.split("=", 1)
    label = label.strip()
    root = root.strip()
    if not label or not root:
        raise argparse.ArgumentTypeError(
            f"Invalid --method value '{value}'. Expected LABEL=/path/to/method_eval_root."
        )
    return label, Path(root)


def collect_method_rows(method_root: Path, metric_keys: list[str]) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for summary_path in sorted(method_root.glob("*/summary.json")):
        payload = load_json(summary_path)
        paper_id = payload.get("paper_id")
        key = normalize_paper_key(str(paper_id) if paper_id else None, summary_path)
        metrics = payload.get("metrics") or {}
        row = {
            "paper_key": key,
            "paper_id": paper_id,
            "title": payload.get("title"),
            "summary_path": str(summary_path),
        }
        for out_key in metric_keys:
            metric_group, metric_name = ALL_METRICS[out_key]
            row[out_key] = safe_float(((metrics.get(metric_group) or {}).get(metric_name)))
        rows[key] = row
    return rows


def build_wide_rows(
    method_order: list[str],
    method_rows: dict[str, dict[str, dict[str, Any]]],
    metric_keys: list[str],
) -> list[dict[str, Any]]:
    all_keys: set[str] = set()
    for rows in method_rows.values():
        all_keys.update(rows.keys())

    wide_rows: list[dict[str, Any]] = []
    for key in sorted(all_keys):
        row: dict[str, Any] = {"paper_key": key, "paper_id": None, "title": None}
        for method_label in method_order:
            source = method_rows[method_label].get(key, {})
            if row["paper_id"] is None:
                row["paper_id"] = source.get("paper_id")
            if row["title"] is None:
                row["title"] = source.get("title")
            for metric_key in metric_keys:
                row[f"{method_label}__{metric_key}"] = source.get(metric_key)
        wide_rows.append(row)
    return wide_rows


def build_average_row(
    *,
    method_order: list[str],
    wide_rows: list[dict[str, Any]],
    metric_keys: list[str],
) -> dict[str, Any]:
    avg: dict[str, Any] = {"paper_key": "AVERAGE", "paper_id": "", "title": ""}
    for method_label in method_order:
        for metric_key in metric_keys:
            values = [
                safe_float(row.get(f"{method_label}__{metric_key}"))
                for row in wide_rows
            ]
            values = [value for value in values if value is not None]
            avg[f"{method_label}__{metric_key}"] = round(sum(values) / len(values), 4) if values else None
    return avg


def build_method_rows(
    *,
    method_order: list[str],
    wide_rows: list[dict[str, Any]],
    metric_keys: list[str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for method_label in method_order:
        row: dict[str, Any] = {"method": method_label}
        for metric_key in metric_keys:
            values = [
                safe_float(item.get(f"{method_label}__{metric_key}"))
                for item in wide_rows
            ]
            values = [value for value in values if value is not None]
            row[metric_key] = round(sum(values) / len(values), 4) if values else None
        rows.append(row)
    return rows


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def build_wide_fieldnames(method_order: list[str], metric_keys: list[str]) -> list[str]:
    fieldnames = ["paper_key", "paper_id", "title"]
    for method_label in method_order:
        for metric_key in metric_keys:
            fieldnames.append(f"{method_label}__{metric_key}")
    return fieldnames


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare bundle-eval summaries across multiple methods.")
    parser.add_argument(
        "--method",
        action="append",
        required=True,
        type=parse_method_spec,
        help="Method label and eval root in the form LABEL=/path/to/eval_root.",
    )
    parser.add_argument(
        "--metrics",
        nargs="+",
        choices=sorted(ALL_METRICS.keys()),
        default=[
            "core_coverage_topic_iou",
            "geometry_aware_density_gad_geom",
            "visual_appeal_deck_score",
            "layout_defects_deck_score",
            "logical_flow_deck_score",
            "paper_faithfulness_deck_score",
        ],
        help="Subset of metric columns to include in the comparison tables.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="Where to save the comparison tables.",
    )
    args = parser.parse_args()

    metric_keys = list(args.metrics)
    method_specs = list(args.method)
    method_order = [label for label, _ in method_specs]
    if len(set(method_order)) != len(method_order):
        raise SystemExit("Each --method label must be unique.")

    collected: dict[str, dict[str, dict[str, Any]]] = {}
    for method_label, method_root in method_specs:
        if not method_root.exists():
            raise FileNotFoundError(f"Method eval root not found: {method_root}")
        collected[method_label] = collect_method_rows(method_root, metric_keys)

    wide_rows = build_wide_rows(method_order, collected, metric_keys)
    average_row = build_average_row(method_order=method_order, wide_rows=wide_rows, metric_keys=metric_keys)
    method_rows = build_method_rows(method_order=method_order, wide_rows=wide_rows, metric_keys=metric_keys)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(
        args.output_dir / "bundle_eval_comparison_by_paper.csv",
        wide_rows,
        build_wide_fieldnames(method_order, metric_keys),
    )
    write_csv(
        args.output_dir / "bundle_eval_comparison_averages.csv",
        [average_row],
        build_wide_fieldnames(method_order, metric_keys),
    )
    write_csv(
        args.output_dir / "bundle_eval_comparison_method_rows.csv",
        method_rows,
        ["method", *metric_keys],
    )
    (args.output_dir / "bundle_eval_comparison_by_paper.json").write_text(
        json.dumps(wide_rows, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "methods": method_order,
                "papers": len(wide_rows),
                "output_dir": str(args.output_dir),
                "method_rows_csv": str(args.output_dir / "bundle_eval_comparison_method_rows.csv"),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
