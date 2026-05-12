#!/usr/bin/env python3
"""Summarize pair-guideline win-rate reports into aggregate JSON and CSV outputs."""

from __future__ import annotations

import argparse
import csv
import json
import statistics
from pathlib import Path
from typing import Any


SCORE_KEYS = [
    "narrative_flow_alignment",
    "section_emphasis_alignment",
    "content_style_alignment",
    "compression_style_alignment",
    "layout_style_alignment",
    "overall_pair_guideline_match",
]


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def summarize_eval_dir(eval_dir: Path) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    paths = sorted(eval_dir.glob("*.pairwin.json"))
    if not paths:
        raise SystemExit(f"No pair-guideline win-rate reports found in {eval_dir}")

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
            "paper": path.stem.replace(".pairwin", ""),
            "winner": winner,
            "confidence": confidence,
            "headline": str(((payload.get("summary") or {}).get("headline")) or ""),
            "winning_margin": safe_float(((payload.get("summary") or {}).get("winning_margin"))),
        }
        for key in SCORE_KEYS:
            row[f"baseline_{key}"] = safe_float(((payload.get("baseline") or {}).get("scores") or {}).get(key))
            row[f"pair_guided_{key}"] = safe_float(((payload.get("pair_guided") or {}).get("scores") or {}).get(key))
            row[f"lift_{key}"] = safe_float(((payload.get("lift") or {}).get(key)))
        rows.append(row)

    aggregate: dict[str, Any] = {
        "num_reports": len(rows),
        "winner_counts": winner_counts,
        "confidence_counts": confidence_counts,
        "mean_scores": {
            "baseline": {},
            "pair_guided": {},
        },
        "mean_lift": {},
        "median_lift": {},
        "improved_counts": {},
        "notable_lifts": {},
        "mean_winning_margin": round(statistics.fmean(row["winning_margin"] for row in rows), 4),
    }

    for key in SCORE_KEYS:
        baseline_values = [row[f"baseline_{key}"] for row in rows]
        pair_values = [row[f"pair_guided_{key}"] for row in rows]
        lift_values = [row[f"lift_{key}"] for row in rows]

        aggregate["mean_scores"]["baseline"][key] = round(statistics.fmean(baseline_values), 4)
        aggregate["mean_scores"]["pair_guided"][key] = round(statistics.fmean(pair_values), 4)
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

    aggregate["winner_table"] = [
        {
            "paper": row["paper"],
            "winner": row["winner"],
            "confidence": row["confidence"],
            "winning_margin": row["winning_margin"],
            "overall_pair_guideline_lift": row["lift_overall_pair_guideline_match"],
            "headline": row["headline"],
        }
        for row in sorted(rows, key=lambda item: item["lift_overall_pair_guideline_match"], reverse=True)
    ]

    return aggregate, rows


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    fieldnames = ["paper", "winner", "confidence", "headline", "winning_margin"]
    for key in SCORE_KEYS:
        fieldnames.extend(
            [
                f"baseline_{key}",
                f"pair_guided_{key}",
                f"lift_{key}",
            ]
        )

    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize pair-guideline win-rate JSON reports.")
    parser.add_argument("--eval-dir", type=Path, required=True, help="Directory containing *.pairwin.json files.")
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
