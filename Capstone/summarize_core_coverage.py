#!/usr/bin/env python3
"""Aggregate deck-similarity evaluation JSON files into a compact summary."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_EVAL_DIR = REPO_ROOT / "Capstone" / "evaluations"
DEFAULT_OUTPUT_PATH = DEFAULT_EVAL_DIR / "core_coverage_summary.json"


def load_eval(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def safe_topics(value: Any) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(item) for item in value if str(item).strip()]


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize topic-IoU evaluation JSON files.")
    parser.add_argument(
        "--eval-dir",
        type=Path,
        default=DEFAULT_EVAL_DIR,
        help=f"Directory containing *.core_coverage.json files (default: {DEFAULT_EVAL_DIR})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_PATH,
        help=f"Where to save the aggregate summary JSON (default: {DEFAULT_OUTPUT_PATH})",
    )
    args = parser.parse_args()

    eval_paths = sorted(args.eval_dir.glob("*.core_coverage.json"))
    if not eval_paths:
        raise SystemExit(f"No evaluation files found in {args.eval_dir}")

    evals = [load_eval(path) for path in eval_paths]

    per_paper = []
    missing_counter: Counter[str] = Counter()
    extra_counter: Counter[str] = Counter()
    matched_counter: Counter[str] = Counter()
    total_intersection = 0.0
    total_union = 0.0

    for item in evals:
        title = item.get("title", "")
        paper_id = item.get("paper_id", "")
        ratio = item.get("topic_iou", item.get("coverage_ratio"))
        matched_topics = safe_topics(item.get("matched_topics"))
        missing = safe_topics(item.get("reference_only_topics", item.get("missing_from_generated")))
        extra = safe_topics(item.get("generated_only_topics"))

        total_intersection += safe_float(item.get("topic_intersection_count"))
        total_union += safe_float(item.get("topic_union_count"))
        missing_counter.update(missing)
        extra_counter.update(extra)
        matched_counter.update(matched_topics)

        per_paper.append(
            {
                "paper_id": paper_id,
                "title": title,
                "topic_iou": ratio,
                "topic_intersection_count": item.get("topic_intersection_count"),
                "topic_union_count": item.get("topic_union_count"),
                "matched_topics": matched_topics,
                "reference_only_topics": missing,
                "generated_only_topics": extra,
                "overall_summary": (item.get("difference_summary") or {}).get("overall_summary", ""),
            }
        )

    per_paper.sort(
        key=lambda row: (
            safe_float(row.get("topic_iou", 0)),
            row.get("paper_id", ""),
        ),
        reverse=True,
    )

    summary = {
        "num_papers": len(evals),
        "average_topic_iou": round(
            sum(safe_float(item.get("topic_iou", item.get("coverage_ratio"))) for item in evals) / len(evals), 3
        ),
        "overall_topic_iou": round(total_intersection / total_union, 3) if total_union > 0 else None,
        "most_frequently_missing_topics": missing_counter.most_common(),
        "most_frequently_extra_topics": extra_counter.most_common(),
        "most_frequently_matched_topics": matched_counter.most_common(),
        "per_paper": per_paper,
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Saved summary to {args.output}")


if __name__ == "__main__":
    main()
