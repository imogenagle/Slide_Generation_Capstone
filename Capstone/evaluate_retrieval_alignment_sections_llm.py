#!/usr/bin/env python3
"""LLM evaluator for retrieval-profile section-structure alignment."""

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

from Capstone.personalization_eval_common import call_alignment_judge, load_dotenv
from SlidesAgent.slide_plan_summary import summarize_slide_plan


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "retrieval_section_alignment_evaluator.yaml"


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def build_section_preferences(profile: dict[str, Any]) -> dict[str, Any]:
    planning = dict(profile.get("planning_preferences") or {})
    result: dict[str, Any] = {}
    if planning.get("target_section_count") is not None:
        try:
            result["target_section_count"] = float(planning["target_section_count"])
        except Exception:
            pass
    labels = []
    for raw_label in planning.get("preferred_section_labels") or []:
        label = str(raw_label or "").strip().lower()
        if label and label not in labels:
            labels.append(label)
    if labels:
        result["preferred_section_labels"] = labels
    order_style = str(planning.get("section_order_style") or "").strip().lower()
    if order_style in {"canonical", "custom", "mixed"}:
        result["section_order_style"] = order_style
    return result


def build_section_summary(plan: dict[str, Any]) -> dict[str, Any]:
    summary = summarize_slide_plan(plan)
    return {
        "section_count": int(summary.get("section_count", 0)),
        "section_titles": list(summary.get("section_titles") or []),
        "section_slide_counts": dict(summary.get("section_slide_counts") or {}),
        "avg_slides_per_section": float(summary.get("avg_slides_per_section", 0.0)),
        "section_splitting_estimate": str(summary.get("section_splitting_estimate") or ""),
        "prefers_takeaway_like_close": bool(summary.get("prefers_takeaway_like_close")),
        "multi_slide_method_like_sections": list(summary.get("multi_slide_method_like_sections") or []),
        "multi_slide_results_like_sections": list(summary.get("multi_slide_results_like_sections") or []),
    }


def render_prompt(
    prompt_path: Path,
    *,
    section_preferences: dict[str, Any],
    baseline_section_summary: dict[str, Any],
    personalized_section_summary: dict[str, Any],
) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    user_prompt = template.render(
        section_preferences_json=section_preferences,
        baseline_section_summary_json=baseline_section_summary,
        personalized_section_summary_json=personalized_section_summary,
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": user_prompt,
    }


def coerce_report(report: dict[str, Any]) -> dict[str, Any]:
    section_alignment = dict(report.get("section_alignment") or {})
    winner = str(section_alignment.get("winner", "tie")).strip().lower()
    if winner not in {"baseline", "personalized", "tie"}:
        winner = "tie"

    try:
        lift = float(section_alignment.get("lift", 0.0))
    except Exception:
        lift = 0.0
    lift = max(-1.0, min(1.0, lift))
    if winner == "baseline" and lift > 0:
        lift = -lift
    if winner == "personalized" and lift < 0:
        lift = -lift
    if winner == "tie":
        lift = 0.0

    matched = dict(section_alignment.get("matched_target_sections") or {})
    missing = dict(section_alignment.get("missing_target_sections") or {})
    count_alignment = dict(section_alignment.get("count_alignment") or {})

    report["section_alignment"] = {
        "winner": winner,
        "lift": round(lift, 4),
        "rationale": str(section_alignment.get("rationale", "")).strip(),
        "matched_target_sections": {
            "baseline": list(matched.get("baseline") or []),
            "personalized": list(matched.get("personalized") or []),
        },
        "missing_target_sections": {
            "baseline": list(missing.get("baseline") or []),
            "personalized": list(missing.get("personalized") or []),
        },
        "count_alignment": {
            "target_section_count": count_alignment.get("target_section_count"),
            "baseline_section_count": count_alignment.get("baseline_section_count"),
            "personalized_section_count": count_alignment.get("personalized_section_count"),
        },
    }
    return report


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Evaluate baseline vs personalized plans against retrieval-profile section preferences."
    )
    parser.add_argument("--profile", type=Path, required=True)
    parser.add_argument("--baseline-plan", type=Path, required=True)
    parser.add_argument("--personalized-plan", type=Path, required=True)
    parser.add_argument("--model", default="gpt-5.4-nano")
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--timeout", type=float, default=180.0)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    for path_arg in (args.profile, args.baseline_plan, args.personalized_plan, args.prompt_path):
        if not path_arg.exists():
            raise FileNotFoundError(f"Required input not found: {path_arg}")

    load_dotenv(REPO_ROOT / ".env")

    profile = load_json(args.profile)
    baseline_plan = load_json(args.baseline_plan)
    personalized_plan = load_json(args.personalized_plan)

    section_preferences = build_section_preferences(profile)
    baseline_section_summary = build_section_summary(baseline_plan)
    personalized_section_summary = build_section_summary(personalized_plan)

    prompt = render_prompt(
        args.prompt_path,
        section_preferences=section_preferences,
        baseline_section_summary=baseline_section_summary,
        personalized_section_summary=personalized_section_summary,
    )
    raw_report = call_alignment_judge(
        args.model,
        prompt["system_prompt"],
        prompt["user_prompt"],
        request_timeout=args.timeout,
        verbose=args.verbose,
    )
    report = coerce_report(raw_report)

    final = {
        "inputs": {
            "profile": str(args.profile),
            "baseline_plan": str(args.baseline_plan),
            "personalized_plan": str(args.personalized_plan),
            "judge_model": args.model,
            "scoring_mode": "retrieval_section_semantic_pairwise",
        },
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

    output_text = json.dumps(final, indent=2, ensure_ascii=False)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output_text, encoding="utf-8")
    print(output_text)


if __name__ == "__main__":
    main()
