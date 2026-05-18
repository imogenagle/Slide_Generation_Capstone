#!/usr/bin/env python3
"""Summarize bundle-eval summary.json files into aggregate JSON and CSV outputs."""

from __future__ import annotations

import argparse
import csv
import json
import statistics
from pathlib import Path
from typing import Any


METRIC_KEYS = [
    "core_coverage_topic_iou",
    "geometry_aware_density_gad_geom",
    "slidetailor_aesthetic_quality_deck_score",
    "slidetailor_content_informativeness_deck_score",
]

PREFERENCE_DEPENDENT_METRIC_KEYS = [
    "slidetailor_structure_similarity_coverage_iou",
    "slidetailor_structure_similarity_flow_ngld",
    "slidetailor_structure_similarity_content_structure_similarity",
]


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def to_pct(value: float) -> float:
    return round(value * 100.0, 2)


def summarize_bundle_eval_root(
    root_dir: Path,
    *,
    include_preference_dependent_slidetailor: bool,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    paths = sorted(root_dir.glob("*/bundle_eval/*/summary.json"))
    if not paths:
        raise SystemExit(f"No bundle-eval summary.json files found under {root_dir}")

    metric_keys = list(METRIC_KEYS)
    if include_preference_dependent_slidetailor:
        metric_keys.extend(PREFERENCE_DEPENDENT_METRIC_KEYS)

    rows: list[dict[str, Any]] = []
    skipped_counts: dict[str, int] = {}

    for path in paths:
        payload = load_json(path)
        metrics = payload.get("metrics") or {}
        skipped = payload.get("skipped") or {}
        paper_id = str(payload.get("paper_id") or path.parents[2].name)
        variant = path.parent.name

        row: dict[str, Any] = {
            "paper_id": paper_id,
            "variant": variant,
            "title": str(payload.get("title") or ""),
            "summary_path": str(path),
        }

        row["core_coverage_topic_iou"] = safe_float(((metrics.get("core_coverage") or {}).get("topic_iou")))
        row["geometry_aware_density_gad_geom"] = safe_float(((metrics.get("geometry_aware_density") or {}).get("gad_geom")))
        row["slidetailor_aesthetic_quality_deck_score"] = safe_float(((metrics.get("slidetailor_aesthetic_quality") or {}).get("deck_score")))
        row["slidetailor_content_informativeness_deck_score"] = safe_float(((metrics.get("slidetailor_content_informativeness") or {}).get("deck_score")))
        row["slidetailor_structure_similarity_coverage_iou"] = safe_float(((metrics.get("slidetailor_structure_similarity") or {}).get("coverage_iou")))
        row["slidetailor_structure_similarity_flow_ngld"] = safe_float(((metrics.get("slidetailor_structure_similarity") or {}).get("flow_ngld")))
        row["slidetailor_structure_similarity_content_structure_similarity"] = safe_float(((metrics.get("slidetailor_structure_similarity") or {}).get("content_structure_similarity")))

        for key, reason in skipped.items():
            skip_key = f"{variant}:{key}:{reason}"
            skipped_counts[skip_key] = skipped_counts.get(skip_key, 0) + 1

        rows.append(row)

    aggregate: dict[str, Any] = {
        "num_reports": len(rows),
        "metric_means_0_to_1": {},
        "metric_means_0_to_100": {},
        "metric_medians_0_to_1": {},
        "metric_medians_0_to_100": {},
        "best_by_metric": {},
        "skipped_counts": skipped_counts,
        "included_metric_keys": metric_keys,
    }

    for key in metric_keys:
        values = [safe_float(row.get(key)) for row in rows]
        mean_value = round(statistics.fmean(values), 4)
        median_value = round(statistics.median(values), 4)
        aggregate["metric_means_0_to_1"][key] = mean_value
        aggregate["metric_means_0_to_100"][key] = to_pct(mean_value)
        aggregate["metric_medians_0_to_1"][key] = median_value
        aggregate["metric_medians_0_to_100"][key] = to_pct(median_value)
        ranked = sorted(rows, key=lambda item: safe_float(item.get(key)), reverse=True)
        aggregate["best_by_metric"][key] = {
            "top": {
                "paper_id": ranked[0]["paper_id"],
                "variant": ranked[0]["variant"],
                "score_0_to_1": round(safe_float(ranked[0].get(key)), 4),
                "score_0_to_100": to_pct(safe_float(ranked[0].get(key))),
            },
            "bottom": {
                "paper_id": ranked[-1]["paper_id"],
                "variant": ranked[-1]["variant"],
                "score_0_to_1": round(safe_float(ranked[-1].get(key)), 4),
                "score_0_to_100": to_pct(safe_float(ranked[-1].get(key))),
            },
        }

    return aggregate, rows


def write_csv(path: Path, rows: list[dict[str, Any]], metric_keys: list[str]) -> None:
    fieldnames = [
        "paper_id",
        "variant",
        "title",
        "summary_path",
        *metric_keys,
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize bundle-eval summary.json files.")
    parser.add_argument("--root-dir", type=Path, required=True, help="Root experiment dir containing per-paper bundle_eval outputs.")
    parser.add_argument("--output-json", type=Path, default=None, help="Optional output JSON summary path.")
    parser.add_argument("--output-csv", type=Path, default=None, help="Optional output CSV detail path.")
    parser.add_argument(
        "--include-preference-dependent-slidetailor",
        action="store_true",
        help="Include preference-based SlideTailor metrics in the aggregate summary.",
    )
    args = parser.parse_args()

    root_dir = args.root_dir.resolve()
    output_json = args.output_json or (root_dir / "BUNDLE_EVAL_SUMMARY.json")
    output_csv = args.output_csv or (root_dir / "BUNDLE_EVAL_DETAILS.csv")

    aggregate, rows = summarize_bundle_eval_root(
        root_dir,
        include_preference_dependent_slidetailor=args.include_preference_dependent_slidetailor,
    )

    output_json.write_text(json.dumps(aggregate, indent=2, ensure_ascii=False), encoding="utf-8")
    write_csv(output_csv, rows, aggregate["included_metric_keys"])

    print(
        json.dumps(
            {
                "output_json": str(output_json),
                "output_csv": str(output_csv),
                "num_reports": aggregate["num_reports"],
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
