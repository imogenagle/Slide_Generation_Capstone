#!/usr/bin/env python3
"""Summarize combined retrieval-eval JSON reports into cohort-level tables."""

from __future__ import annotations

import argparse
import csv
import json
from collections import Counter
from pathlib import Path
from typing import Any


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float | None:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def normalize_winner(value: Any) -> str:
    winner = str(value or "tie").strip().lower()
    if winner not in {"personalized", "baseline", "tie"}:
        return "tie"
    return winner


def collect_reports(eval_dir: Path) -> tuple[list[dict[str, Any]], list[str]]:
    paths = sorted(eval_dir.glob("*.json"))
    if not paths:
        raise SystemExit(f"No retrieval eval JSON files found in {eval_dir}")

    rows: list[dict[str, Any]] = []
    metric_names: list[str] = []

    for path in paths:
        payload = load_json(path)
        numeric_eval = dict(payload.get("numeric_eval") or {})
        section_eval = dict(payload.get("section_eval") or {})
        summary = dict(payload.get("summary") or {})
        comparison = dict(numeric_eval.get("comparison") or {})
        if not metric_names:
            metric_names = list(comparison.keys())

        row: dict[str, Any] = {
            "paper_key": path.stem,
            "numeric_winner": normalize_winner(summary.get("numeric_winner")),
            "section_winner": normalize_winner(summary.get("section_winner")),
            "color_palette_winner": normalize_winner(summary.get("color_palette_winner")),
            "font_winner": normalize_winner(summary.get("font_winner")),
            "personalized_metric_wins": numeric_eval.get("summary", {}).get("personalized_metric_wins"),
            "baseline_metric_wins": numeric_eval.get("summary", {}).get("baseline_metric_wins"),
            "tied_metrics": numeric_eval.get("summary", {}).get("tied_metrics"),
            "section_lift": safe_float(section_eval.get("section_alignment", {}).get("lift")),
        }

        for metric_name, metric_payload in comparison.items():
            metric_payload = dict(metric_payload or {})
            row[f"{metric_name}__winner"] = normalize_winner(metric_payload.get("closer_to_target"))
            row[f"{metric_name}__distance_improvement"] = safe_float(metric_payload.get("distance_improvement"))

        rows.append(row)

    return rows, metric_names


def build_overview(rows: list[dict[str, Any]]) -> dict[str, Any]:
    numeric_counts = Counter(row["numeric_winner"] for row in rows)
    section_counts = Counter(row["section_winner"] for row in rows)
    color_counts = Counter(row["color_palette_winner"] for row in rows)
    font_counts = Counter(row["font_winner"] for row in rows)

    def count(name: str, key: str) -> int:
        return int(sum(1 for row in rows if row.get(key) == name))

    section_lifts = [safe_float(row.get("section_lift")) for row in rows]
    section_lifts = [value for value in section_lifts if value is not None]

    return {
        "num_reports": len(rows),
        "numeric_personalized_wins": numeric_counts["personalized"],
        "numeric_baseline_wins": numeric_counts["baseline"],
        "numeric_ties": numeric_counts["tie"],
        "section_personalized_wins": section_counts["personalized"],
        "section_baseline_wins": section_counts["baseline"],
        "section_ties": section_counts["tie"],
        "color_personalized_wins": color_counts["personalized"],
        "color_baseline_wins": color_counts["baseline"],
        "color_ties": color_counts["tie"],
        "font_personalized_wins": font_counts["personalized"],
        "font_baseline_wins": font_counts["baseline"],
        "font_ties": font_counts["tie"],
        "papers_with_personalized_numeric_win": count("personalized", "numeric_winner"),
        "papers_with_personalized_section_win": count("personalized", "section_winner"),
        "mean_section_lift": round(sum(section_lifts) / len(section_lifts), 4) if section_lifts else None,
    }


def build_metric_summary(rows: list[dict[str, Any]], metric_names: list[str]) -> list[dict[str, Any]]:
    summary_rows: list[dict[str, Any]] = []
    for metric_name in metric_names:
        win_counts = Counter(normalize_winner(row.get(f"{metric_name}__winner")) for row in rows)
        improvements = [safe_float(row.get(f"{metric_name}__distance_improvement")) for row in rows]
        improvements = [value for value in improvements if value is not None]
        summary_rows.append(
            {
                "metric": metric_name,
                "personalized_wins": win_counts["personalized"],
                "baseline_wins": win_counts["baseline"],
                "ties": win_counts["tie"],
                "mean_distance_improvement": round(sum(improvements) / len(improvements), 4) if improvements else None,
            }
        )
    return summary_rows


def build_metric_percentage_table(rows: list[dict[str, Any]], metric_names: list[str]) -> tuple[list[dict[str, Any]], list[str]]:
    attribute_keys = list(metric_names) + ["section", "color_palette", "font"]
    key_to_row_field = {
        **{metric_name: f"{metric_name}__winner" for metric_name in metric_names},
        "section": "section_winner",
        "color_palette": "color_palette_winner",
        "font": "font_winner",
    }

    denominator = max(len(rows), 1)
    table_rows: list[dict[str, Any]] = []
    for winner_name in ("personalized", "baseline", "tie"):
        row: dict[str, Any] = {"winner": winner_name}
        for attribute_key in attribute_keys:
            field_name = key_to_row_field[attribute_key]
            count = sum(1 for item in rows if normalize_winner(item.get(field_name)) == winner_name)
            row[attribute_key] = round(count / denominator, 4)
        table_rows.append(row)
    return table_rows, attribute_keys


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize cohort retrieval-eval reports into CSV/JSON outputs.")
    parser.add_argument("--eval-dir", type=Path, required=True, help="Directory containing per-paper retrieval eval JSON files.")
    parser.add_argument("--output-dir", type=Path, required=True, help="Where to write summary outputs.")
    args = parser.parse_args()

    rows, metric_names = collect_reports(args.eval_dir)
    overview = build_overview(rows)
    metric_summary = build_metric_summary(rows, metric_names)
    metric_percentage_rows, attribute_keys = build_metric_percentage_table(rows, metric_names)

    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(
        args.output_dir / "retrieval_eval_by_paper.csv",
        rows,
        [
            "paper_key",
            "numeric_winner",
            "section_winner",
            "color_palette_winner",
            "font_winner",
            "personalized_metric_wins",
            "baseline_metric_wins",
            "tied_metrics",
            "section_lift",
            *[f"{metric_name}__winner" for metric_name in metric_names],
            *[f"{metric_name}__distance_improvement" for metric_name in metric_names],
        ],
    )
    write_csv(
        args.output_dir / "retrieval_eval_overview.csv",
        [overview],
        list(overview.keys()),
    )
    write_csv(
        args.output_dir / "retrieval_eval_metric_summary.csv",
        metric_summary,
        ["metric", "personalized_wins", "baseline_wins", "ties", "mean_distance_improvement"],
    )
    write_csv(
        args.output_dir / "retrieval_eval_metric_percentages.csv",
        metric_percentage_rows,
        ["winner", *attribute_keys],
    )
    (args.output_dir / "retrieval_eval_summary.json").write_text(
        json.dumps(
            {
                "overview": overview,
                "metric_summary": metric_summary,
                "metric_percentages": metric_percentage_rows,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "num_reports": len(rows),
                "output_dir": str(args.output_dir),
                "overview_csv": str(args.output_dir / "retrieval_eval_overview.csv"),
                "metric_summary_csv": str(args.output_dir / "retrieval_eval_metric_summary.csv"),
                "metric_percentages_csv": str(args.output_dir / "retrieval_eval_metric_percentages.csv"),
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
