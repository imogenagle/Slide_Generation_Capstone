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


def apply_pairwise_applicability(
    report: dict[str, Any],
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
    overall_components = [float(report["lift"][key]) for key in applicable if key != "overall_style_alignment"]
    report["lift"]["overall_style_alignment"] = round(sum(overall_components) / len(overall_components), 4) if overall_components else 0.0

    overall_delta = float(report["lift"]["overall_style_alignment"])
    winner = "tie"
    if overall_delta > 0.03:
        winner = "personalized"
    elif overall_delta < -0.03:
        winner = "baseline"

    summary = report.setdefault("summary", {})
    summary["winner"] = winner
    if not summary.get("headline"):
        if winner == "personalized":
            summary["headline"] = "Personalized matches the profile better overall in the pairwise comparison."
        elif winner == "baseline":
            summary["headline"] = "Baseline matches the profile better overall in the pairwise comparison."
        else:
            summary["headline"] = "Baseline and personalized are roughly tied in the pairwise comparison."
    if summary.get("confidence") not in {"low", "medium", "high"}:
        magnitude = abs(overall_delta)
        summary["confidence"] = "high" if magnitude >= 0.2 else "medium" if magnitude >= 0.08 else "low"

    report["applicable_dimensions"] = applicable
    report["skipped_dimensions"] = skipped
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
    report = apply_pairwise_applicability(report, extracted_asset_summary)

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
