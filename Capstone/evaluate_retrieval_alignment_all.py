#!/usr/bin/env python3
"""Combined retrieval evaluator for numeric and qualitative section targets."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.evaluate_retrieval_alignment_numeric import (
    build_color_comparison,
    build_font_comparison,
    build_metric_comparison,
    build_retrieval_target_summary,
)
from Capstone.evaluate_retrieval_alignment_sections_llm import (
    build_section_preferences,
    build_section_summary,
    coerce_report as coerce_section_report,
    render_prompt as render_section_prompt,
)
from Capstone.personalization_eval_common import call_alignment_judge, load_dotenv
from SlidesAgent.slide_plan_summary import summarize_slide_plan


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def compute_numeric_report(
    *,
    profile: dict[str, Any],
    baseline_plan: dict[str, Any],
    personalized_plan: dict[str, Any],
    baseline_plan_path: Path,
    personalized_plan_path: Path,
) -> dict[str, Any]:
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
    color_comparison = build_color_comparison(
        profile=profile,
        baseline_plan_path=baseline_plan_path,
        personalized_plan_path=personalized_plan_path,
    )
    font_comparison = build_font_comparison(
        profile=profile,
        baseline_plan_path=baseline_plan_path,
        personalized_plan_path=personalized_plan_path,
    )

    report = {
        "target_summary": target_summary,
        "observed_metrics": {
            "baseline": {
                "slide_count": int(baseline_summary.get("slide_count", 0)),
                "avg_bullets_per_slide": round(float(baseline_summary.get("avg_bullets_per_slide", 0.0)), 4),
                "avg_words_per_slide": round(float(baseline_summary.get("avg_words_per_slide", 0.0)), 4),
                "image_slide_count": baseline_image_slide_count,
                "table_slide_count": baseline_table_slide_count,
                "formula_slide_count": baseline_formula_slide_count,
            },
            "personalized": {
                "slide_count": int(personalized_summary.get("slide_count", 0)),
                "avg_bullets_per_slide": round(float(personalized_summary.get("avg_bullets_per_slide", 0.0)), 4),
                "avg_words_per_slide": round(float(personalized_summary.get("avg_words_per_slide", 0.0)), 4),
                "image_slide_count": personalized_image_slide_count,
                "table_slide_count": personalized_table_slide_count,
                "formula_slide_count": personalized_formula_slide_count,
            },
        },
        "comparison": comparison["metrics"],
        "color_palette_eval": color_comparison,
        "font_eval": font_comparison,
        "summary": comparison["aggregate"],
    }
    if color_comparison is not None:
        report["summary"]["color_palette_winner"] = color_comparison["closer_to_palette_preferences"]
    if font_comparison is not None:
        report["summary"]["font_winner"] = font_comparison["closer_to_font_preferences"]
    return report


def compute_section_report(
    *,
    profile: dict[str, Any],
    baseline_plan: dict[str, Any],
    personalized_plan: dict[str, Any],
    prompt_path: Path,
    model: str,
    timeout: float,
    verbose: bool,
) -> dict[str, Any]:
    section_preferences = build_section_preferences(profile)
    baseline_section_summary = build_section_summary(baseline_plan)
    personalized_section_summary = build_section_summary(personalized_plan)

    if not section_preferences:
        return {
            "target_section_preferences": section_preferences,
            "observed_sections": {
                "baseline": baseline_section_summary,
                "personalized": personalized_section_summary,
            },
            "section_alignment": {
                "winner": "tie",
                "lift": 0.0,
                "rationale": "Skipped because the retrieval profile did not contain section preferences.",
                "matched_target_sections": {"baseline": [], "personalized": []},
                "missing_target_sections": {"baseline": [], "personalized": []},
                "count_alignment": {
                    "target_section_count": None,
                    "baseline_section_count": baseline_section_summary.get("section_count"),
                    "personalized_section_count": personalized_section_summary.get("section_count"),
                },
            },
            "summary": {
                "winner": "tie",
                "lift": 0.0,
                "headline": "Skipped because the retrieval profile did not contain section preferences.",
            },
        }

    prompt = render_section_prompt(
        prompt_path,
        section_preferences=section_preferences,
        baseline_section_summary=baseline_section_summary,
        personalized_section_summary=personalized_section_summary,
    )
    raw_report = call_alignment_judge(
        model,
        prompt["system_prompt"],
        prompt["user_prompt"],
        request_timeout=timeout,
        verbose=verbose,
    )
    report = coerce_section_report(raw_report)
    return {
        "target_section_preferences": section_preferences,
        "observed_sections": {
            "baseline": baseline_section_summary,
            "personalized": personalized_section_summary,
        },
        "section_alignment": report["section_alignment"],
        "summary": {
            "winner": report["section_alignment"]["winner"],
            "lift": report["section_alignment"]["lift"],
            "headline": report["section_alignment"]["rationale"],
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Evaluate retrieval-profile runs with both numeric and LLM section alignment in one report."
    )
    parser.add_argument("--profile", type=Path, required=True)
    parser.add_argument("--baseline-plan", type=Path, required=True)
    parser.add_argument("--personalized-plan", type=Path, required=True)
    parser.add_argument("--model", default="gpt-5.4-nano")
    parser.add_argument(
        "--section-prompt-path",
        type=Path,
        default=REPO_ROOT / "utils" / "prompt_templates" / "retrieval_section_alignment_evaluator.yaml",
    )
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--timeout", type=float, default=180.0)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    for path_arg in (args.profile, args.baseline_plan, args.personalized_plan, args.section_prompt_path):
        if not path_arg.exists():
            raise FileNotFoundError(f"Required input not found: {path_arg}")

    load_dotenv(REPO_ROOT / ".env")

    profile = load_json(args.profile)
    baseline_plan = load_json(args.baseline_plan)
    personalized_plan = load_json(args.personalized_plan)

    numeric_report = compute_numeric_report(
        profile=profile,
        baseline_plan=baseline_plan,
        personalized_plan=personalized_plan,
        baseline_plan_path=args.baseline_plan,
        personalized_plan_path=args.personalized_plan,
    )
    section_report = compute_section_report(
        profile=profile,
        baseline_plan=baseline_plan,
        personalized_plan=personalized_plan,
        prompt_path=args.section_prompt_path,
        model=args.model,
        timeout=args.timeout,
        verbose=args.verbose,
    )

    numeric_category_winners = {
        target_key: metric_report.get("closer_to_target")
        for target_key, metric_report in (numeric_report.get("comparison") or {}).items()
    }
    all_metric_winners = dict(numeric_category_winners)
    all_metric_winners["section"] = section_report["summary"]["winner"]
    all_metric_winners["color_palette"] = numeric_report["summary"].get("color_palette_winner")
    all_metric_winners["font"] = numeric_report["summary"].get("font_winner")

    final = {
        "inputs": {
            "profile": str(args.profile),
            "baseline_plan": str(args.baseline_plan),
            "personalized_plan": str(args.personalized_plan),
            "judge_model": args.model,
            "scoring_mode": "retrieval_combined_numeric_and_section",
        },
        "numeric_eval": numeric_report,
        "section_eval": section_report,
        "summary": {
            "numeric_winner": numeric_report["summary"]["winner"],
            "section_winner": section_report["summary"]["winner"],
            "numeric_metric_wins": numeric_report["summary"]["personalized_metric_wins"],
            "numeric_metric_losses": numeric_report["summary"]["baseline_metric_wins"],
            "numeric_ties": numeric_report["summary"]["tied_metrics"],
            "numeric_category_winners": numeric_category_winners,
            "color_palette_winner": numeric_report["summary"].get("color_palette_winner"),
            "font_winner": numeric_report["summary"].get("font_winner"),
            "all_metric_winners": all_metric_winners,
        },
    }

    output_text = json.dumps(final, indent=2, ensure_ascii=False)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output_text, encoding="utf-8")
    print(output_text)


if __name__ == "__main__":
    main()
