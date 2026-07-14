#!/usr/bin/env python3
"""Build comparison tables for SlideGen bundle-eval outputs.

This script compares:
1. SlideGen Original
2. SlideGen baseline
3. SlideGen personalized

It reads each run's `summary.json` files and writes:
- a per-paper wide CSV table
- an averages CSV table
- a method-by-metric averages CSV table
- a JSON dump of the normalized rows
"""

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
    "layout_correctness_deck_score": ("layout_correctness", "deck_score"),
    "logical_flow_deck_score": ("logical_flow", "deck_score"),
    "paper_faithfulness_deck_score": ("paper_faithfulness", "deck_score"),
}

SYSTEM_LABELS = {
    "original": "SlideGen_Original",
    "baseline": "SlideGen_Baseline",
    "personalized": "SlideGen_Personalized",
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


def collect_original(original_root: Path) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for summary_path in sorted(original_root.glob("*/summary.json")):
        payload = load_json(summary_path)
        paper_id = payload.get("paper_id")
        key = normalize_paper_key(str(paper_id) if paper_id else None, summary_path)
        metrics = payload.get("metrics") or {}
        row = {
            "paper_key": key,
            "paper_id": paper_id,
            "title": payload.get("title"),
            "variant": "original",
            "summary_path": str(summary_path),
        }
        for out_key, (metric_group, metric_name) in ALL_METRICS.items():
            row[out_key] = safe_float(((metrics.get(metric_group) or {}).get(metric_name)))
        rows[key] = row
    return rows


def collect_variant(bundle_root: Path, variant: str) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for summary_path in sorted((bundle_root / variant).glob("*/summary.json")):
        payload = load_json(summary_path)
        paper_id = payload.get("paper_id")
        key = normalize_paper_key(str(paper_id) if paper_id else None, summary_path)
        metrics = payload.get("metrics") or {}
        row = {
            "paper_key": key,
            "paper_id": paper_id,
            "title": payload.get("title"),
            "variant": variant,
            "summary_path": str(summary_path),
        }
        for out_key, (metric_group, metric_name) in ALL_METRICS.items():
            row[out_key] = safe_float(((metrics.get(metric_group) or {}).get(metric_name)))
        rows[key] = row
    return rows


def build_wide_rows(
    original_rows: dict[str, dict[str, Any]],
    baseline_rows: dict[str, dict[str, Any]],
    personalized_rows: dict[str, dict[str, Any]],
    metric_keys: list[str],
) -> list[dict[str, Any]]:
    all_keys = sorted(set(original_rows) | set(baseline_rows) | set(personalized_rows))
    wide_rows: list[dict[str, Any]] = []

    for key in all_keys:
        original = original_rows.get(key, {})
        baseline = baseline_rows.get(key, {})
        personalized = personalized_rows.get(key, {})

        row: dict[str, Any] = {
            "paper_key": key,
            "paper_id": original.get("paper_id") or baseline.get("paper_id") or personalized.get("paper_id"),
            "title": original.get("title") or baseline.get("title") or personalized.get("title"),
        }
        for system_key, source in (
            ("original", original),
            ("baseline", baseline),
            ("personalized", personalized),
        ):
            for metric_key in metric_keys:
                row[f"{system_key}__{metric_key}"] = source.get(metric_key)
        wide_rows.append(row)
    return wide_rows


def build_average_row(wide_rows: list[dict[str, Any]], metric_keys: list[str]) -> dict[str, Any]:
    avg: dict[str, Any] = {
        "paper_key": "AVERAGE",
        "paper_id": "",
        "title": "",
    }
    for system_key in SYSTEM_LABELS:
        for metric_key in metric_keys:
            values = [
                safe_float(row.get(f"{system_key}__{metric_key}"))
                for row in wide_rows
            ]
            values = [v for v in values if v is not None]
            avg[f"{system_key}__{metric_key}"] = round(sum(values) / len(values), 4) if values else None
    return avg


def build_method_rows(wide_rows: list[dict[str, Any]], metric_keys: list[str]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for system_key, system_label in SYSTEM_LABELS.items():
        row: dict[str, Any] = {"method": system_label}
        for metric_key in metric_keys:
            values = [
                safe_float(item.get(f"{system_key}__{metric_key}"))
                for item in wide_rows
            ]
            values = [v for v in values if v is not None]
            row[metric_key] = round(sum(values) / len(values), 4) if values else None
        rows.append(row)
    return rows


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def build_fieldnames(metric_keys: list[str]) -> list[str]:
    fields = ["paper_key", "paper_id", "title"]
    for system_key in ("original", "baseline", "personalized"):
        for metric_key in metric_keys:
            fields.append(f"{system_key}__{metric_key}")
    return fields


def build_method_fieldnames(metric_keys: list[str]) -> list[str]:
    return ["method", *metric_keys]


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare bundle-eval outputs across original/baseline/personalized runs.")
    parser.add_argument(
        "--original-root",
        type=Path,
        default=Path("outputs/original_bundle_eval"),
        help="Directory containing one summary.json folder per original deck.",
    )
    parser.add_argument(
        "--bundle-root",
        type=Path,
        default=Path("outputs/retrieval_0702/Capstone/batch_runs/experiments/data_raw_5papers_bundle_eval_only_0707/bundle_eval"),
        help="Bundle eval root containing baseline/ and personalized/ subfolders.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("outputs/bundle_eval_comparison"),
        help="Where to write comparison tables.",
    )
    parser.add_argument(
        "--metrics",
        nargs="+",
        choices=sorted(ALL_METRICS.keys()),
        default=list(ALL_METRICS.keys()),
        help="Subset of metric columns to include in the comparison tables.",
    )
    args = parser.parse_args()
    metric_keys = list(args.metrics)

    original_rows = collect_original(args.original_root)
    baseline_rows = collect_variant(args.bundle_root, "baseline")
    personalized_rows = collect_variant(args.bundle_root, "personalized")

    wide_rows = build_wide_rows(original_rows, baseline_rows, personalized_rows, metric_keys)
    avg_row = build_average_row(wide_rows, metric_keys)
    method_rows = build_method_rows(wide_rows, metric_keys)
    fieldnames = build_fieldnames(metric_keys)
    method_fieldnames = build_method_fieldnames(metric_keys)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(args.output_dir / "bundle_eval_comparison_by_paper.csv", wide_rows, fieldnames)
    write_csv(args.output_dir / "bundle_eval_comparison_averages.csv", [avg_row], fieldnames)
    write_csv(args.output_dir / "bundle_eval_comparison_method_rows.csv", method_rows, method_fieldnames)
    (args.output_dir / "bundle_eval_comparison_by_paper.json").write_text(
        json.dumps(wide_rows, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "papers": len(wide_rows),
                "output_dir": str(args.output_dir),
                "by_paper_csv": str(args.output_dir / "bundle_eval_comparison_by_paper.csv"),
                "averages_csv": str(args.output_dir / "bundle_eval_comparison_averages.csv"),
                "method_rows_csv": str(args.output_dir / "bundle_eval_comparison_method_rows.csv"),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
