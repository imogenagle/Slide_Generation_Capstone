#!/usr/bin/env python3
"""Summarize personalization alignment reports into aggregate JSON and CSV outputs."""

from __future__ import annotations

import argparse
import csv
import json
import statistics
from pathlib import Path
from typing import Any


SCORE_KEYS = [
    "section_structure_alignment",
    "bullet_density_alignment",
    "text_density_alignment",
    "figure_usage_alignment",
    "table_usage_alignment",
    "formula_usage_alignment",
    "layout_bias_alignment",
    "overall_style_alignment",
]


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def summarize_eval_dir(eval_dir: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    paths = sorted(eval_dir.glob("*.alignment.json"))
    if not paths:
        raise SystemExit(f"No alignment reports found in {eval_dir}")

    rows: list[dict[str, Any]] = []
    winner_counts: dict[str, int] = {}
    confidence_counts: dict[str, int] = {}

    for path in paths:
        payload = load_json(path)
        winner = str(((payload.get("summary") or {}).get("winner")) or "unknown")
        confidence = str(((payload.get("summary") or {}).get("confidence")) or "unknown")
        winner_counts[winner] = winner_counts.get(winner, 0) + 1
        confidence_counts[confidence] = confidence_counts.get(confidence, 0) + 1

        row: dict[str, Any] = {
            "paper": path.stem.replace(".alignment", ""),
            "winner": winner,
            "confidence": confidence,
            "headline": str(((payload.get("summary") or {}).get("headline")) or ""),
        }
        for key in SCORE_KEYS:
            row[f"baseline_{key}"] = safe_float(((payload.get("baseline") or {}).get("scores") or {}).get(key))
            row[f"personalized_{key}"] = safe_float(((payload.get("personalized") or {}).get("scores") or {}).get(key))
            row[f"lift_{key}"] = safe_float(((payload.get("lift") or {}).get(key)))
        rows.append(row)

    aggregate: dict[str, Any] = {
        "num_reports": len(rows),
        "winner_counts": winner_counts,
        "confidence_counts": confidence_counts,
        "mean_scores": {
            "baseline": {},
            "personalized": {},
        },
        "mean_lift": {},
        "median_lift": {},
        "improved_counts": {},
        "notable_lifts": {},
    }

    for key in SCORE_KEYS:
        baseline_values = [row[f"baseline_{key}"] for row in rows]
        personalized_values = [row[f"personalized_{key}"] for row in rows]
        lift_values = [row[f"lift_{key}"] for row in rows]

        aggregate["mean_scores"]["baseline"][key] = round(statistics.fmean(baseline_values), 4)
        aggregate["mean_scores"]["personalized"][key] = round(statistics.fmean(personalized_values), 4)
        aggregate["mean_lift"][key] = round(statistics.fmean(lift_values), 4)
        aggregate["median_lift"][key] = round(statistics.median(lift_values), 4)
        aggregate["improved_counts"][key] = {
            "improved": sum(value > 0 for value in lift_values),
            "declined": sum(value < 0 for value in lift_values),
            "unchanged": sum(value == 0 for value in lift_values),
            "total": len(lift_values),
        }
        aggregate["notable_lifts"][key] = {
            "top_gains": [
                {"paper": row["paper"], "lift": row[f"lift_{key}"]}
                for row in sorted(rows, key=lambda item: item[f"lift_{key}"], reverse=True)[:5]
            ],
            "top_drops": [
                {"paper": row["paper"], "lift": row[f"lift_{key}"]}
                for row in sorted(rows, key=lambda item: item[f"lift_{key}"])[:5]
            ],
        }

    aggregate["overall_style_winners"] = [
        {
            "paper": row["paper"],
            "winner": row["winner"],
            "confidence": row["confidence"],
            "overall_style_lift": row["lift_overall_style_alignment"],
            "headline": row["headline"],
        }
        for row in sorted(rows, key=lambda item: item["lift_overall_style_alignment"], reverse=True)
    ]

    return aggregate, rows


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    fieldnames = ["paper", "winner", "confidence", "headline"]
    for key in SCORE_KEYS:
        fieldnames.extend(
            [
                f"baseline_{key}",
                f"personalized_{key}",
                f"lift_{key}",
            ]
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize personalization alignment JSON reports.")
    parser.add_argument("--eval-dir", type=Path, required=True, help="Directory containing *.alignment.json files.")
    parser.add_argument("--output-json", type=Path, default=None, help="Optional output JSON summary path.")
    parser.add_argument("--output-csv", type=Path, default=None, help="Optional output CSV detail path.")
    args = parser.parse_args()

    eval_dir = args.eval_dir.resolve()
    output_json = args.output_json or (eval_dir / "SUMMARY.json")
    output_csv = args.output_csv or (eval_dir / "DETAILS.csv")

    aggregate, rows = summarize_eval_dir(eval_dir)

    output_json.write_text(json.dumps(aggregate, indent=2, ensure_ascii=False), encoding="utf-8")
    write_csv(output_csv, rows)

    print(json.dumps({"output_json": str(output_json), "output_csv": str(output_csv), "num_reports": aggregate["num_reports"]}, indent=2))


if __name__ == "__main__":
    main()
