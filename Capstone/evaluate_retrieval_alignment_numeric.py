#!/usr/bin/env python3
"""Deterministic evaluator for retrieval-profile personalization runs.

This script is intentionally retrieval-specific: it compares baseline and
personalized slide plans only on the numeric targets used by the retrieval
personalization path.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from SlidesAgent.slide_plan_summary import summarize_slide_plan


RETRIEVAL_METRIC_SPECS = {
    "target_avg_words_per_slide": {
        "label": "avg_words_per_slide",
        "summary_key": "avg_words_per_slide",
        "normalizer_floor": 8.0,
    },
    "target_image_slide_count": {
        "label": "image_slide_count",
        "summary_key": "image_slide_count",
        "normalizer_floor": 1.0,
    },
    "target_table_slide_count": {
        "label": "table_slide_count",
        "summary_key": "table_slide_count",
        "normalizer_floor": 1.0,
    },
    "target_formula_slide_count": {
        "label": "formula_slide_count",
        "summary_key": "formula_slide_count",
        "normalizer_floor": 1.0,
    },
}


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def build_retrieval_target_summary(profile: dict[str, Any]) -> dict[str, float]:
    numeric = dict(profile.get("numeric_preferences") or {})
    summary: dict[str, float] = {}
    for key in RETRIEVAL_METRIC_SPECS:
        raw_value = numeric.get(key)
        if raw_value is None:
            continue
        try:
            summary[key] = float(raw_value)
        except Exception:
            continue
    return summary


def build_metric_comparison(
    *,
    target_summary: dict[str, float],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
) -> dict[str, Any]:
    metrics: dict[str, Any] = {}
    normalized_baseline_distances: list[float] = []
    normalized_personalized_distances: list[float] = []
    raw_baseline_distances: list[float] = []
    raw_personalized_distances: list[float] = []

    for target_key, target_value in target_summary.items():
        spec = RETRIEVAL_METRIC_SPECS[target_key]
        summary_key = str(spec["summary_key"])
        baseline_value = float(baseline_summary.get(summary_key, 0.0))
        personalized_value = float(personalized_summary.get(summary_key, 0.0))

        baseline_distance = abs(baseline_value - target_value)
        personalized_distance = abs(personalized_value - target_value)
        normalizer = max(
            abs(target_value),
            baseline_distance,
            personalized_distance,
            float(spec["normalizer_floor"]),
        )
        baseline_normalized = baseline_distance / normalizer
        personalized_normalized = personalized_distance / normalizer

        if abs(baseline_distance - personalized_distance) <= 1e-9:
            closer = "tie"
        elif personalized_distance < baseline_distance:
            closer = "personalized"
        else:
            closer = "baseline"

        raw_gain = baseline_distance - personalized_distance
        normalized_gain = baseline_normalized - personalized_normalized

        metrics[target_key] = {
            "label": spec["label"],
            "target": round(target_value, 4),
            "baseline": round(baseline_value, 4),
            "personalized": round(personalized_value, 4),
            "baseline_distance": round(baseline_distance, 4),
            "personalized_distance": round(personalized_distance, 4),
            "baseline_normalized_distance": round(baseline_normalized, 4),
            "personalized_normalized_distance": round(personalized_normalized, 4),
            "distance_improvement": round(raw_gain, 4),
            "normalized_distance_improvement": round(normalized_gain, 4),
            "closer_to_target": closer,
        }

        normalized_baseline_distances.append(baseline_normalized)
        normalized_personalized_distances.append(personalized_normalized)
        raw_baseline_distances.append(baseline_distance)
        raw_personalized_distances.append(personalized_distance)

    metric_count = len(metrics)
    baseline_mean_normalized = (
        sum(normalized_baseline_distances) / metric_count if metric_count else 0.0
    )
    personalized_mean_normalized = (
        sum(normalized_personalized_distances) / metric_count if metric_count else 0.0
    )
    baseline_raw_total = sum(raw_baseline_distances)
    personalized_raw_total = sum(raw_personalized_distances)

    if abs(baseline_mean_normalized - personalized_mean_normalized) <= 1e-9:
        winner = "tie"
    elif personalized_mean_normalized < baseline_mean_normalized:
        winner = "personalized"
    else:
        winner = "baseline"

    return {
        "metrics": metrics,
        "aggregate": {
            "metric_count": metric_count,
            "baseline_total_raw_distance": round(baseline_raw_total, 4),
            "personalized_total_raw_distance": round(personalized_raw_total, 4),
            "baseline_mean_normalized_distance": round(baseline_mean_normalized, 4),
            "personalized_mean_normalized_distance": round(personalized_mean_normalized, 4),
            "overall_distance_improvement": round(baseline_raw_total - personalized_raw_total, 4),
            "overall_normalized_lift": round(
                baseline_mean_normalized - personalized_mean_normalized, 4
            ),
            "winner": winner,
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Evaluate baseline vs personalized plans against retrieval-profile numeric targets."
    )
    parser.add_argument("--profile", type=Path, required=True)
    parser.add_argument("--baseline-plan", type=Path, required=True)
    parser.add_argument("--personalized-plan", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args()

    for path_arg in (args.profile, args.baseline_plan, args.personalized_plan):
        if not path_arg.exists():
            raise FileNotFoundError(f"Required input not found: {path_arg}")

    profile = load_json(args.profile)
    baseline_plan = load_json(args.baseline_plan)
    personalized_plan = load_json(args.personalized_plan)

    target_summary = build_retrieval_target_summary(profile)
    baseline_summary = summarize_slide_plan(baseline_plan)
    personalized_summary = summarize_slide_plan(personalized_plan)
    baseline_image_slide_count = sum(1 for slide in list(baseline_plan.get("slides") or []) if isinstance(slide, dict) and slide.get("images"))
    baseline_table_slide_count = sum(1 for slide in list(baseline_plan.get("slides") or []) if isinstance(slide, dict) and slide.get("tables"))
    baseline_formula_slide_count = sum(1 for slide in list(baseline_plan.get("slides") or []) if isinstance(slide, dict) and slide.get("formulas"))
    personalized_image_slide_count = sum(1 for slide in list(personalized_plan.get("slides") or []) if isinstance(slide, dict) and slide.get("images"))
    personalized_table_slide_count = sum(1 for slide in list(personalized_plan.get("slides") or []) if isinstance(slide, dict) and slide.get("tables"))
    personalized_formula_slide_count = sum(1 for slide in list(personalized_plan.get("slides") or []) if isinstance(slide, dict) and slide.get("formulas"))
    baseline_summary["image_slide_count"] = baseline_image_slide_count
    baseline_summary["table_slide_count"] = baseline_table_slide_count
    baseline_summary["formula_slide_count"] = baseline_formula_slide_count
    personalized_summary["image_slide_count"] = personalized_image_slide_count
    personalized_summary["table_slide_count"] = personalized_table_slide_count
    personalized_summary["formula_slide_count"] = personalized_formula_slide_count

    comparison = build_metric_comparison(
        target_summary=target_summary,
        baseline_summary=baseline_summary,
        personalized_summary=personalized_summary,
    )

    report = {
        "inputs": {
            "profile": str(args.profile),
            "baseline_plan": str(args.baseline_plan),
            "personalized_plan": str(args.personalized_plan),
            "scoring_mode": "retrieval_numeric_target_distance",
        },
        "target_summary": target_summary,
        "observed_metrics": {
            "baseline": {
                "avg_words_per_slide": round(float(baseline_summary.get("avg_words_per_slide", 0.0)), 4),
                "image_slide_count": baseline_image_slide_count,
                "table_slide_count": baseline_table_slide_count,
                "formula_slide_count": baseline_formula_slide_count,
            },
            "personalized": {
                "avg_words_per_slide": round(float(personalized_summary.get("avg_words_per_slide", 0.0)), 4),
                "image_slide_count": personalized_image_slide_count,
                "table_slide_count": personalized_table_slide_count,
                "formula_slide_count": personalized_formula_slide_count,
            },
        },
        "comparison": comparison["metrics"],
        "summary": comparison["aggregate"],
    }

    output_text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output_text, encoding="utf-8")
    print(output_text)


if __name__ == "__main__":
    main()
