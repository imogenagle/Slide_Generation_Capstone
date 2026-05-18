#!/usr/bin/env python3
"""Pairwise evaluator for baseline vs personalized slide plans against an author profile."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import yaml
from jinja2 import Environment, StrictUndefined

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.evaluate_personalization_alignment import (
    SCORE_KEYS,
    apply_dimension_applicability,
    build_numeric_comparison,
    build_numeric_target_summary,
    call_alignment_judge,
    load_dotenv,
    log,
    summarize_extracted_assets,
    summarize_slide_plan,
)


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "personalization_alignment_evaluator_pairwise.yaml"

DETERMINISTIC_DIMENSION_SPECS = {
    "bullet_density_alignment": {
        "summary_key": "avg_bullets_per_slide",
        "target_key": "target_avg_bullets_per_slide",
        "tie_threshold": 0.05,
        "normalizer_floor": 0.5,
        "label": "avg bullets/slide",
    },
    "text_density_alignment": {
        "summary_key": "avg_words_per_slide",
        "target_key": "target_avg_words_per_slide",
        "tie_threshold": 0.25,
        "normalizer_floor": 2.0,
        "label": "avg words/slide",
    },
    "figure_usage_alignment": {
        "summary_key": "figure_slide_fraction",
        "target_key": "target_fraction_figure_slides",
        "tie_threshold": 0.015,
        "normalizer_floor": 0.05,
        "label": "figure-slide fraction",
    },
    "table_usage_alignment": {
        "summary_key": "table_slide_fraction",
        "target_key": "target_fraction_table_slides",
        "tie_threshold": 0.015,
        "normalizer_floor": 0.05,
        "label": "table-slide fraction",
    },
    "formula_usage_alignment": {
        "summary_key": "formula_slide_fraction",
        "target_key": "target_fraction_formula_slides",
        "tie_threshold": 0.015,
        "normalizer_floor": 0.05,
        "label": "formula-slide fraction",
    },
}


def render_prompt(
    prompt_path: Path,
    author_profile: dict[str, Any],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
    extracted_asset_summary: dict[str, Any],
) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    user_prompt = template.render(
        author_profile_json=author_profile,
        baseline_plan_summary_json=baseline_summary,
        personalized_plan_summary_json=personalized_summary,
        extracted_asset_summary_json=extracted_asset_summary,
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": user_prompt,
    }


def coerce_pairwise_report(report: dict[str, Any]) -> dict[str, Any]:
    dimensions = report.setdefault("dimensions", {})
    lift: dict[str, float] = {}
    rationales: dict[str, str] = {}

    for key in SCORE_KEYS:
        dim = dimensions.setdefault(key, {})
        winner = str(dim.get("winner", "tie")).strip().lower()
        if winner not in {"baseline", "personalized", "tie"}:
            winner = "tie"
        dim["winner"] = winner
        try:
            value = float(dim.get("lift", 0.0))
        except Exception:
            value = 0.0
        value = max(-1.0, min(1.0, value))
        if winner == "baseline" and value > 0:
            value = -value
        if winner == "personalized" and value < 0:
            value = -value
        if winner == "tie":
            value = 0.0
        dim["lift"] = round(value, 4)
        dim["rationale"] = str(dim.get("rationale", "")).strip()
        lift[key] = dim["lift"]
        rationales[key] = dim["rationale"]

    report["lift"] = lift
    report["dimension_rationales"] = rationales
    return report


def compute_deterministic_dimension_results(
    author_profile: dict[str, Any],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
    applicable_dimensions: list[str],
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]], dict[str, str]]:
    target_summary = build_numeric_target_summary(author_profile)
    deterministic_results: dict[str, dict[str, Any]] = {}
    diagnostics: dict[str, dict[str, Any]] = {}
    basis: dict[str, str] = {}

    for metric_key, spec in DETERMINISTIC_DIMENSION_SPECS.items():
        if metric_key not in applicable_dimensions:
            continue
        target_key = spec["target_key"]
        summary_key = spec["summary_key"]
        if target_key not in target_summary:
            continue
        try:
            target_value = float(target_summary[target_key])
            baseline_value = float(baseline_summary.get(summary_key, 0.0))
            personalized_value = float(personalized_summary.get(summary_key, 0.0))
        except Exception:
            continue

        baseline_distance = abs(baseline_value - target_value)
        personalized_distance = abs(personalized_value - target_value)
        distance_gap = baseline_distance - personalized_distance
        if abs(distance_gap) <= float(spec["tie_threshold"]):
            winner = "tie"
            lift_value = 0.0
        else:
            winner = "personalized" if distance_gap > 0 else "baseline"
            normalizer = max(
                baseline_distance,
                personalized_distance,
                float(spec["normalizer_floor"]),
            )
            lift_value = max(-1.0, min(1.0, distance_gap / normalizer))

        label = str(spec["label"])
        rationale = (
            f"Target {label} is {target_value:.4f}; personalized is {personalized_value:.4f} "
            f"(distance {personalized_distance:.4f}) and baseline is {baseline_value:.4f} "
            f"(distance {baseline_distance:.4f})."
        )
        if winner == "tie":
            rationale += " The difference in target distance is negligible, so this dimension is treated as a tie."
        else:
            rationale += f" {winner.capitalize()} is closer to the target on this measurable dimension."

        deterministic_results[metric_key] = {
            "winner": winner,
            "lift": round(lift_value, 4),
            "rationale": rationale,
        }
        diagnostics[metric_key] = {
            "target": round(target_value, 4),
            "baseline": round(baseline_value, 4),
            "personalized": round(personalized_value, 4),
            "baseline_distance": round(baseline_distance, 4),
            "personalized_distance": round(personalized_distance, 4),
        }
        basis[metric_key] = "deterministic_target_distance"

    return deterministic_results, diagnostics, basis


def build_pairwise_headline(
    winner: str,
    applicable_dimensions: list[str],
    dimensions: dict[str, dict[str, Any]],
) -> str:
    applicable = [key for key in applicable_dimensions if key != "overall_style_alignment"]
    if not applicable:
        return "No applicable dimensions were available for pairwise personalization evaluation."

    if winner == "tie":
        return "Baseline and personalized split the applicable dimensions closely in the pairwise profile comparison."

    winning_dims = [
        (key, abs(float((dimensions.get(key) or {}).get("lift", 0.0))))
        for key in applicable
        if str((dimensions.get(key) or {}).get("winner", "tie")) == winner
    ]
    winning_dims.sort(key=lambda item: item[1], reverse=True)
    top_dims = ", ".join(key for key, _ in winning_dims[:2]) or "the applicable dimensions"
    subject = "Personalized" if winner == "personalized" else "Baseline"
    return f"{subject} wins the pairwise profile comparison overall, with its clearest edge on {top_dims}."


def apply_pairwise_applicability(
    report: dict[str, Any],
    author_profile: dict[str, Any],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
    extracted_asset_summary: dict[str, Any],
) -> dict[str, Any]:
    placeholder = {
        "baseline": {"scores": {key: 0.5 for key in SCORE_KEYS}},
        "personalized": {"scores": {key: 0.5 for key in SCORE_KEYS}},
        "lift": {key: 0.0 for key in SCORE_KEYS},
        "summary": {},
    }
    placeholder = apply_dimension_applicability(placeholder, extracted_asset_summary)
    skipped = list(placeholder.get("skipped_dimensions", []))
    applicable = list(placeholder.get("applicable_dimensions", []))

    dimensions = report.setdefault("dimensions", {})
    for key in skipped:
        dim = dimensions.setdefault(key, {})
        dim["winner"] = "tie"
        dim["lift"] = 0.0
        dim["rationale"] = "Skipped from comparison because the target paper did not expose supporting extracted assets for this dimension."

    report = coerce_pairwise_report(report)
    deterministic_results, deterministic_diagnostics, deterministic_basis = compute_deterministic_dimension_results(
        author_profile,
        baseline_summary,
        personalized_summary,
        applicable,
    )
    dimensions = report.setdefault("dimensions", {})
    for metric_key, result in deterministic_results.items():
        dimensions[metric_key] = result
    report = coerce_pairwise_report(report)

    applicable_no_overall = [key for key in applicable if key != "overall_style_alignment"]
    personalized_wins = sum(
        1 for key in applicable_no_overall
        if str((dimensions.get(key) or {}).get("winner", "tie")) == "personalized"
    )
    baseline_wins = sum(
        1 for key in applicable_no_overall
        if str((dimensions.get(key) or {}).get("winner", "tie")) == "baseline"
    )
    tie_count = len(applicable_no_overall) - personalized_wins - baseline_wins
    total_count = max(1, len(applicable_no_overall))

    overall_delta = round((personalized_wins - baseline_wins) / total_count, 4)
    report["lift"]["overall_style_alignment"] = overall_delta
    if overall_delta > 0:
        winner = "personalized"
    elif overall_delta < 0:
        winner = "baseline"
    else:
        winner = "tie"

    summary = report.setdefault("summary", {})
    summary["winner"] = winner
    summary["headline"] = build_pairwise_headline(winner, applicable, dimensions)
    magnitude = abs(overall_delta)
    summary["confidence"] = "high" if magnitude >= 0.5 else "medium" if magnitude >= 0.2 else "low"
    summary["win_counts"] = {
        "personalized": personalized_wins,
        "baseline": baseline_wins,
        "tie": tie_count,
        "applicable": len(applicable_no_overall),
    }
    summary["personalized_win_rate"] = round(personalized_wins / total_count, 4)
    summary["baseline_win_rate"] = round(baseline_wins / total_count, 4)

    report["applicable_dimensions"] = applicable
    report["skipped_dimensions"] = skipped
    report["dimension_score_basis"] = {
        key: deterministic_basis.get(key, "llm_pairwise_judgment")
        for key in SCORE_KEYS
    }
    if deterministic_diagnostics:
        report["deterministic_target_distance_diagnostics"] = deterministic_diagnostics
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate baseline vs personalized with direct pairwise profile comparison.")
    parser.add_argument("--profile", type=Path, required=True)
    parser.add_argument("--baseline-plan", type=Path, required=True)
    parser.add_argument("--personalized-plan", type=Path, required=True)
    parser.add_argument("--model", default="gpt-5")
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--print-prompt", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")

    for path_arg in (args.profile, args.baseline_plan, args.personalized_plan, args.prompt_path):
        if not path_arg.exists():
            raise FileNotFoundError(f"Required input not found: {path_arg}")

    log("Loading profile and slide plans", verbose=args.verbose)
    author_profile = json.loads(args.profile.read_text(encoding="utf-8"))
    baseline_plan = json.loads(args.baseline_plan.read_text(encoding="utf-8"))
    personalized_plan = json.loads(args.personalized_plan.read_text(encoding="utf-8"))

    baseline_summary = summarize_slide_plan(baseline_plan)
    personalized_summary = summarize_slide_plan(personalized_plan)
    extracted_asset_summary = {
        "baseline": summarize_extracted_assets(args.baseline_plan),
        "personalized": summarize_extracted_assets(args.personalized_plan),
    }
    prompt = render_prompt(
        args.prompt_path,
        author_profile,
        baseline_summary,
        personalized_summary,
        extracted_asset_summary,
    )

    if args.print_prompt:
        print(prompt["user_prompt"])
        return

    report = call_alignment_judge(
        args.model,
        prompt["system_prompt"],
        prompt["user_prompt"],
        request_timeout=args.request_timeout,
        verbose=args.verbose,
    )
    report = apply_pairwise_applicability(
        report,
        author_profile,
        baseline_summary,
        personalized_summary,
        extracted_asset_summary,
    )

    numeric_target_summary = build_numeric_target_summary(author_profile)
    numeric_comparison = build_numeric_comparison(
        numeric_target_summary,
        baseline_summary,
        personalized_summary,
    )
    report["inputs"] = {
        "profile": str(args.profile),
        "baseline_plan": str(args.baseline_plan),
        "personalized_plan": str(args.personalized_plan),
        "judge_model": args.model,
        "scoring_mode": "pairwise_profile_comparison",
    }
    report["derived_plan_summaries"] = {
        "baseline": baseline_summary,
        "personalized": personalized_summary,
    }
    report["extracted_asset_summary"] = extracted_asset_summary
    if numeric_target_summary:
        report["numeric_target_summary"] = numeric_target_summary
        report["numeric_comparison"] = numeric_comparison

    output_text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output_text, encoding="utf-8")
        log(f"Wrote report to {args.output}", verbose=args.verbose)
    print(output_text)


if __name__ == "__main__":
    main()
