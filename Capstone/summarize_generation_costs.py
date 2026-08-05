#!/usr/bin/env python3
"""Summarize generation token usage, runtime, and estimated cost for study outputs.

If retrieval-profile token usage is available, this script also reports a
"whole personalized deck" cost that adds profile-building cost to personalized
deck-generation cost. If profile usage was not logged for a study, the script
leaves those fields blank rather than guessing.
"""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from statistics import mean
from typing import Any


DEFAULT_INPUT_RATE_PER_M = 0.20
DEFAULT_OUTPUT_RATE_PER_M = 1.25


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def collect_rows(
    *,
    study_root: Path,
    method: str,
    pattern: str,
    input_rate_per_m: float,
    output_rate_per_m: float,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for path in sorted(study_root.glob(pattern)):
        payload = load_json(path)
        input_tokens_t = float(payload.get("input_tokens_t", 0) or 0)
        output_tokens_t = float(payload.get("output_tokens_t", 0) or 0)
        input_tokens_v = float(payload.get("input_tokens_v", 0) or 0)
        output_tokens_v = float(payload.get("output_tokens_v", 0) or 0)
        total_input_tokens = input_tokens_t + input_tokens_v
        total_output_tokens = output_tokens_t + output_tokens_v
        total_tokens = total_input_tokens + total_output_tokens
        time_taken = float(payload.get("time_taken", 0) or 0)
        estimated_cost = (
            (total_input_tokens / 1_000_000.0) * input_rate_per_m
            + (total_output_tokens / 1_000_000.0) * output_rate_per_m
        )
        rows.append(
            {
                "method": method,
                "deck_dir": path.parent.name,
                "log_path": str(path),
                "input_tokens_t": int(input_tokens_t),
                "output_tokens_t": int(output_tokens_t),
                "input_tokens_v": int(input_tokens_v),
                "output_tokens_v": int(output_tokens_v),
                "total_input_tokens": int(total_input_tokens),
                "total_output_tokens": int(total_output_tokens),
                "total_tokens": int(total_tokens),
                "time_seconds": round(time_taken, 2),
                "estimated_cost_usd": round(estimated_cost, 6),
                "profile_input_tokens": payload.get("profile_input_tokens"),
                "profile_output_tokens": payload.get("profile_output_tokens"),
                "profile_total_tokens": payload.get("profile_total_tokens"),
                "profile_estimated_cost_usd": payload.get("profile_estimated_cost_usd"),
                "full_personalized_input_tokens": payload.get("full_personalized_input_tokens"),
                "full_personalized_output_tokens": payload.get("full_personalized_output_tokens"),
                "full_personalized_total_tokens": payload.get("full_personalized_total_tokens"),
                "full_personalized_estimated_cost_usd": payload.get("full_personalized_estimated_cost_usd"),
                "shared_artifact_input_tokens": payload.get("shared_artifact_input_tokens"),
                "shared_artifact_output_tokens": payload.get("shared_artifact_output_tokens"),
                "shared_artifact_total_tokens": payload.get("shared_artifact_total_tokens"),
                "shared_artifact_time_seconds": payload.get("shared_artifact_time_seconds"),
                "shared_artifact_estimated_cost_usd": payload.get("shared_artifact_estimated_cost_usd"),
                "full_run_input_tokens": payload.get("full_run_input_tokens"),
                "full_run_output_tokens": payload.get("full_run_output_tokens"),
                "full_run_total_tokens": payload.get("full_run_total_tokens"),
                "full_run_estimated_cost_usd": payload.get("full_run_estimated_cost_usd"),
                "full_personalized_with_shared_input_tokens": payload.get("full_personalized_with_shared_input_tokens"),
                "full_personalized_with_shared_output_tokens": payload.get("full_personalized_with_shared_output_tokens"),
                "full_personalized_with_shared_total_tokens": payload.get("full_personalized_with_shared_total_tokens"),
                "full_personalized_with_shared_estimated_cost_usd": payload.get("full_personalized_with_shared_estimated_cost_usd"),
            }
        )
    return rows


def _safe_float(value: Any) -> float | None:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def collect_profile_rows(
    *,
    profiles_root: Path,
    input_rate_per_m: float,
    output_rate_per_m: float,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for path in sorted(profiles_root.glob("*.retrieval.json")):
        payload = load_json(path)
        usage = payload.get("usage") or payload.get("token_usage") or {}
        prompt_tokens = (
            _safe_float(usage.get("prompt_tokens"))
            or _safe_float(usage.get("input_tokens"))
            or _safe_float(payload.get("prompt_tokens"))
            or _safe_float(payload.get("input_tokens"))
        )
        completion_tokens = (
            _safe_float(usage.get("completion_tokens"))
            or _safe_float(usage.get("output_tokens"))
            or _safe_float(payload.get("completion_tokens"))
            or _safe_float(payload.get("output_tokens"))
        )
        total_tokens = _safe_float(usage.get("total_tokens")) or _safe_float(payload.get("total_tokens"))
        if prompt_tokens is None and completion_tokens is None and total_tokens is None:
            rows.append(
                {
                    "profile_file": path.name,
                    "profile_path": str(path),
                    "input_tokens": None,
                    "output_tokens": None,
                    "total_tokens": None,
                    "estimated_cost_usd": None,
                }
            )
            continue

        prompt_tokens = float(prompt_tokens or 0)
        completion_tokens = float(completion_tokens or 0)
        if total_tokens is None:
            total_tokens = prompt_tokens + completion_tokens
        estimated_cost = (
            (prompt_tokens / 1_000_000.0) * input_rate_per_m
            + (completion_tokens / 1_000_000.0) * output_rate_per_m
        )
        rows.append(
            {
                "profile_file": path.name,
                "profile_path": str(path),
                "input_tokens": int(prompt_tokens),
                "output_tokens": int(completion_tokens),
                "total_tokens": int(total_tokens),
                "estimated_cost_usd": round(estimated_cost, 6),
            }
        )
    return rows


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def collect_fieldnames(rows: list[dict[str, Any]]) -> list[str]:
    fieldnames: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row.keys():
            if key not in seen:
                seen.add(key)
                fieldnames.append(key)
    return fieldnames


def build_summary(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    summary_rows: list[dict[str, Any]] = []
    methods = sorted({row["method"] for row in rows})
    for method in methods:
        method_rows = [row for row in rows if row["method"] == method]
        summary_rows.append(
            {
                "method": method,
                "num_decks": len(method_rows),
                "avg_input_tokens": round(mean(row["total_input_tokens"] for row in method_rows), 2),
                "avg_output_tokens": round(mean(row["total_output_tokens"] for row in method_rows), 2),
                "avg_total_tokens": round(mean(row["total_tokens"] for row in method_rows), 2),
                "avg_time_seconds": round(mean(row["time_seconds"] for row in method_rows), 2),
                "avg_estimated_cost_usd": round(mean(row["estimated_cost_usd"] for row in method_rows), 6),
                "total_estimated_cost_usd": round(sum(row["estimated_cost_usd"] for row in method_rows), 6),
            }
        )
    return summary_rows


def maybe_build_personalized_full_summary(
    *,
    generation_rows: list[dict[str, Any]],
    profile_rows: list[dict[str, Any]],
) -> dict[str, Any] | None:
    personalized_rows = [row for row in generation_rows if row["method"] == "SlideGen_Personalized"]
    if not personalized_rows:
        return None

    embedded_full = [
        row
        for row in personalized_rows
        if row.get("full_personalized_input_tokens") is not None
        and row.get("full_personalized_output_tokens") is not None
        and row.get("full_personalized_total_tokens") is not None
        and row.get("profile_estimated_cost_usd") is not None
    ]
    if len(embedded_full) == len(personalized_rows):
        embedded_with_shared = [
            row for row in embedded_full
            if row.get("full_personalized_with_shared_input_tokens") is not None
            and row.get("full_personalized_with_shared_output_tokens") is not None
            and row.get("full_personalized_with_shared_total_tokens") is not None
            and row.get("full_personalized_with_shared_estimated_cost_usd") is not None
        ]
        notes = "Includes profile-building cost plus personalized deck-generation cost."
        if len(embedded_with_shared) == len(embedded_full):
            notes = "Includes shared artifact cost, profile-building cost, and personalized deck-generation cost."
        return {
            "method": "SlideGen_Personalized_Full",
            "num_decks": len(personalized_rows),
            "avg_input_tokens": round(
                mean(
                    float(
                        row["full_personalized_with_shared_input_tokens"]
                        if len(embedded_with_shared) == len(embedded_full)
                        else row["full_personalized_input_tokens"]
                    )
                    for row in embedded_full
                ),
                2,
            ),
            "avg_output_tokens": round(
                mean(
                    float(
                        row["full_personalized_with_shared_output_tokens"]
                        if len(embedded_with_shared) == len(embedded_full)
                        else row["full_personalized_output_tokens"]
                    )
                    for row in embedded_full
                ),
                2,
            ),
            "avg_total_tokens": round(
                mean(
                    float(
                        row["full_personalized_with_shared_total_tokens"]
                        if len(embedded_with_shared) == len(embedded_full)
                        else row["full_personalized_total_tokens"]
                    )
                    for row in embedded_full
                ),
                2,
            ),
            "avg_time_seconds": round(mean(row["time_seconds"] for row in personalized_rows), 2),
            "avg_estimated_cost_usd": round(
                mean(
                    float(
                        row["full_personalized_with_shared_estimated_cost_usd"]
                        if len(embedded_with_shared) == len(embedded_full)
                        else row["full_personalized_estimated_cost_usd"]
                    )
                    for row in embedded_full
                ),
                6,
            ),
            "total_estimated_cost_usd": round(
                sum(
                    float(
                        row["full_personalized_with_shared_estimated_cost_usd"]
                        if len(embedded_with_shared) == len(embedded_full)
                        else row["full_personalized_estimated_cost_usd"]
                    )
                    for row in embedded_full
                ),
                6,
            ),
            "notes": notes,
        }

    profiled = [row for row in profile_rows if row.get("estimated_cost_usd") is not None]
    if not profiled:
        return {
            "method": "SlideGen_Personalized_Full",
            "num_decks": len(personalized_rows),
            "avg_input_tokens": None,
            "avg_output_tokens": None,
            "avg_total_tokens": None,
            "avg_time_seconds": round(mean(row["time_seconds"] for row in personalized_rows), 2),
            "avg_estimated_cost_usd": None,
            "total_estimated_cost_usd": None,
            "notes": "Profile token usage was not logged for this study, so full personalized cost could not be recovered.",
        }

    avg_profile_input = mean(float(row["input_tokens"]) for row in profiled if row.get("input_tokens") is not None)
    avg_profile_output = mean(float(row["output_tokens"]) for row in profiled if row.get("output_tokens") is not None)
    avg_profile_total = mean(float(row["total_tokens"]) for row in profiled if row.get("total_tokens") is not None)
    avg_profile_cost = mean(float(row["estimated_cost_usd"]) for row in profiled if row.get("estimated_cost_usd") is not None)
    total_profile_cost = sum(float(row["estimated_cost_usd"]) for row in profiled if row.get("estimated_cost_usd") is not None)

    return {
        "method": "SlideGen_Personalized_Full",
        "num_decks": len(personalized_rows),
        "avg_input_tokens": round(mean(row["total_input_tokens"] for row in personalized_rows) + avg_profile_input, 2),
        "avg_output_tokens": round(mean(row["total_output_tokens"] for row in personalized_rows) + avg_profile_output, 2),
        "avg_total_tokens": round(mean(row["total_tokens"] for row in personalized_rows) + avg_profile_total, 2),
        "avg_time_seconds": round(mean(row["time_seconds"] for row in personalized_rows), 2),
        "avg_estimated_cost_usd": round(mean(row["estimated_cost_usd"] for row in personalized_rows) + avg_profile_cost, 6),
        "total_estimated_cost_usd": round(sum(row["estimated_cost_usd"] for row in personalized_rows) + total_profile_cost, 6),
        "notes": "Includes profile-building cost plus personalized deck-generation cost.",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize generation token usage and estimated cost.")
    parser.add_argument(
        "--study-root",
        type=Path,
        required=True,
        help="Study directory containing slidegen_outputs/ and original_outputs/.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="Where to save the summary CSVs. Defaults to <study-root>/cost_summary.",
    )
    parser.add_argument(
        "--input-rate-per-m",
        type=float,
        default=DEFAULT_INPUT_RATE_PER_M,
        help="Input token price in USD per 1M tokens.",
    )
    parser.add_argument(
        "--output-rate-per-m",
        type=float,
        default=DEFAULT_OUTPUT_RATE_PER_M,
        help="Output token price in USD per 1M tokens.",
    )
    args = parser.parse_args()

    output_dir = args.output_dir or (args.study_root / "cost_summary")
    rows: list[dict[str, Any]] = []
    rows.extend(
        collect_rows(
            study_root=args.study_root,
            method="SlideGen_Baseline",
            pattern="slidegen_outputs/contents/*_high_level/<gpt-5.4-nano_gpt-5.4-nano>_log_baseline.json",
            input_rate_per_m=args.input_rate_per_m,
            output_rate_per_m=args.output_rate_per_m,
        )
    )
    rows.extend(
        collect_rows(
            study_root=args.study_root,
            method="SlideGen_Personalized",
            pattern="slidegen_outputs/contents/*_high_level_personalized_retrieval/<gpt-5.4-nano_gpt-5.4-nano>_log_personalized_retrieval.json",
            input_rate_per_m=args.input_rate_per_m,
            output_rate_per_m=args.output_rate_per_m,
        )
    )
    rows.extend(
        collect_rows(
            study_root=args.study_root,
            method="SlideGen_Original",
            pattern="original_outputs/contents/*_original/<gpt-5.4-nano_gpt-5.4-nano>_log.json",
            input_rate_per_m=args.input_rate_per_m,
            output_rate_per_m=args.output_rate_per_m,
        )
    )
    profile_rows = collect_profile_rows(
        profiles_root=args.study_root / "retrieval_profiles",
        input_rate_per_m=args.input_rate_per_m,
        output_rate_per_m=args.output_rate_per_m,
    )

    if not rows:
        raise SystemExit(f"No generation log files found under {args.study_root}")

    detail_fieldnames = list(rows[0].keys())
    summary_rows = build_summary(rows)
    personalized_full = maybe_build_personalized_full_summary(
        generation_rows=rows,
        profile_rows=profile_rows,
    )
    if personalized_full is not None:
        summary_rows.append(personalized_full)
    summary_fieldnames = collect_fieldnames(summary_rows)

    write_csv(output_dir / "deck_token_cost_detail.csv", rows, detail_fieldnames)
    write_csv(output_dir / "deck_token_cost_summary.csv", summary_rows, summary_fieldnames)
    if profile_rows:
        write_csv(
            output_dir / "profile_token_cost_detail.csv",
            profile_rows,
            collect_fieldnames(profile_rows),
        )
    (output_dir / "deck_token_cost_summary.json").write_text(
        json.dumps(
            {
                "input_rate_per_m": args.input_rate_per_m,
                "output_rate_per_m": args.output_rate_per_m,
                "detail_rows": len(rows),
                "profile_rows": len(profile_rows),
                "summary": summary_rows,
            },
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print(
        json.dumps(
            {
                "study_root": str(args.study_root),
                "output_dir": str(output_dir),
                "detail_csv": str(output_dir / "deck_token_cost_detail.csv"),
                "summary_csv": str(output_dir / "deck_token_cost_summary.csv"),
                "input_rate_per_m": args.input_rate_per_m,
                "output_rate_per_m": args.output_rate_per_m,
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
