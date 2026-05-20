#!/usr/bin/env python3
"""Summarize pairwise personalization reports into aggregate JSON and CSV outputs."""

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
    paths = sorted({
        *eval_dir.glob("*.pairwise.json"),
        *eval_dir.glob("*_pairwise.json"),
    })
    if not paths:
        raise SystemExit(f"No pairwise personalization reports found in {eval_dir}")

    rows: list[dict[str, Any]] = []
    winner_counts: dict[str, int] = {}
    confidence_counts: dict[str, int] = {}
    dimension_winner_counts: dict[str, dict[str, int]] = {
        key: {"personalized": 0, "baseline": 0, "tie": 0}
        for key in SCORE_KEYS
    }

    for path in paths:
        payload = load_json(path)
        summary = dict(payload.get("summary") or {})
        dimensions = dict(payload.get("dimensions") or {})
        lift = dict(payload.get("lift") or {})
        basis = dict(payload.get("dimension_score_basis") or {})

        winner = str(summary.get("winner") or "unknown")
        confidence = str(summary.get("confidence") or "unknown")
        winner_counts[winner] = winner_counts.get(winner, 0) + 1
        confidence_counts[confidence] = confidence_counts.get(confidence, 0) + 1

        row: dict[str, Any] = {
            "paper": path.stem.removesuffix(".pairwise").removesuffix("_pairwise"),
            "winner": winner,
            "confidence": confidence,
            "headline": str(summary.get("headline") or ""),
            "personalized_win_rate": safe_float(summary.get("personalized_win_rate")),
            "baseline_win_rate": safe_float(summary.get("baseline_win_rate")),
        }

        for key in SCORE_KEYS:
            dim = dict(dimensions.get(key) or {})
            dim_winner = str(dim.get("winner") or "tie")
            if dim_winner not in dimension_winner_counts[key]:
                dim_winner = "tie"
            dimension_winner_counts[key][dim_winner] += 1
            row[f"winner_{key}"] = dim_winner
            row[f"lift_{key}"] = safe_float(lift.get(key))
            row[f"basis_{key}"] = str(basis.get(key) or "")

        rows.append(row)

    aggregate: dict[str, Any] = {
        "num_reports": len(rows),
        "winner_counts": winner_counts,
        "confidence_counts": confidence_counts,
        "dimension_winner_counts": dimension_winner_counts,
        "mean_lift": {},
        "median_lift": {},
        "improved_counts": {},
        "notable_lifts": {},
    }

    for key in SCORE_KEYS:
        lift_values = [row[f"lift_{key}"] for row in rows]
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
    fieldnames = [
        "paper",
        "winner",
        "confidence",
        "headline",
        "personalized_win_rate",
        "baseline_win_rate",
    ]
    for key in SCORE_KEYS:
        fieldnames.extend(
            [
                f"winner_{key}",
                f"lift_{key}",
                f"basis_{key}",
            ]
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize pairwise personalization JSON reports.")
    parser.add_argument("--eval-dir", type=Path, required=True, help="Directory containing *.pairwise.json files.")
    parser.add_argument("--output-json", type=Path, default=None, help="Optional output JSON summary path.")
    parser.add_argument("--output-csv", type=Path, default=None, help="Optional output CSV detail path.")
    args = parser.parse_args()

    eval_dir = args.eval_dir.resolve()
    output_json = args.output_json or (eval_dir / "SUMMARY.json")
    output_csv = args.output_csv or (eval_dir / "DETAILS.csv")

    aggregate, rows = summarize_eval_dir(eval_dir)

    output_json.write_text(json.dumps(aggregate, indent=2, ensure_ascii=False), encoding="utf-8")
    write_csv(output_csv, rows)

    print(
        json.dumps(
            {"output_json": str(output_json), "output_csv": str(output_csv), "num_reports": aggregate["num_reports"]},
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
