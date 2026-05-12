#!/usr/bin/env python3
"""Evaluate baseline vs pair-guided plans for pair-guideline win-rate."""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any

import yaml
from jinja2 import Environment, StrictUndefined

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from slidegen_openai_utils import build_openai_client, resolve_direct_model_name
from Capstone.evaluate_personalization_alignment import (
    extract_json_object,
    log,
    summarize_extracted_assets,
    summarize_slide_plan,
)


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "pair_guideline_winrate_evaluator.yaml"

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


def dedupe_keep_order(values: list[str], *, limit: int = 10) -> list[str]:
    cleaned: list[str] = []
    for value in values:
        text = re.sub(r"\s+", " ", str(value or "").strip())
        if not text or text in cleaned:
            continue
        cleaned.append(text)
        if len(cleaned) >= limit:
            break
    return cleaned


def majority_label(values: list[str]) -> str:
    cleaned = [str(value or "").strip().lower() for value in values if str(value or "").strip().lower() not in {"", "unknown"}]
    if not cleaned:
        return "unknown"
    counts = Counter(cleaned)
    return counts.most_common(1)[0][0]


def summarize_pair_guideline_context(context: dict[str, Any]) -> dict[str, Any]:
    reference_pairs = list(context.get("reference_pairs") or [])

    narrative_flow: list[str] = []
    section_emphasis: list[str] = []
    content_style: list[str] = []
    compression_patterns: list[str] = []
    visual_usage: list[str] = []
    signature_choices: list[str] = []

    section_split_values: list[str] = []
    bullet_density_values: list[str] = []
    text_density_values: list[str] = []
    visual_density_values: list[str] = []
    layout_bias_values: list[str] = []
    takeaway_values: list[str] = []
    method_values: list[str] = []
    results_values: list[str] = []
    reference_highlights: list[dict[str, Any]] = []

    for pair in reference_pairs:
        presentation = dict(pair.get("presentation_guidelines") or {})
        planning = dict(pair.get("planning_hints") or {})
        structure = dict(planning.get("structure_preferences") or {})
        evidence = dict(pair.get("evidence_summary") or {})

        narrative_flow.extend(presentation.get("narrative_flow_preferences") or [])
        section_emphasis.extend(presentation.get("section_emphasis_patterns") or [])
        content_style.extend(presentation.get("content_style_preferences") or [])
        compression_patterns.extend(presentation.get("compression_patterns") or [])
        visual_usage.extend(presentation.get("visual_usage_patterns") or [])
        signature_choices.extend(presentation.get("signature_choices") or [])

        section_split_values.append(str(planning.get("section_splitting_preference") or "unknown"))
        bullet_density_values.append(str(planning.get("bullet_density_preference") or "unknown"))
        text_density_values.append(str(planning.get("text_density_preference") or "unknown"))
        visual_density_values.append(str(planning.get("visual_density_preference") or "unknown"))
        layout_bias_values.extend(planning.get("layout_bias") or [])
        takeaway_values.append(str(structure.get("prefers_takeaway_slide") or "unknown"))
        method_values.append(str(structure.get("prefers_multi_slide_method_section") or "unknown"))
        results_values.append(str(structure.get("prefers_multi_slide_results_section") or "unknown"))

        reference_highlights.append(
            {
                "paper_id": pair.get("paper_id"),
                "notes": str(evidence.get("notes") or "").strip(),
                "signature_choices": dedupe_keep_order(list(presentation.get("signature_choices") or []), limit=3),
                "section_emphasis_patterns": dedupe_keep_order(list(presentation.get("section_emphasis_patterns") or []), limit=3),
            }
        )

    layout_bias_counts = Counter(str(value).strip() for value in layout_bias_values if str(value).strip())

    return {
        "author_id": context.get("author_id"),
        "author_display_name": context.get("author_display_name"),
        "target_paper_id": context.get("target_paper_id"),
        "reference_pair_ids": list(context.get("reference_pair_ids") or []),
        "reference_pair_count": len(reference_pairs),
        "target_paper_context": {
            "title_guess": ((context.get("target_paper_context") or {}).get("title_guess") or ""),
            "section_heading_candidates": list(((context.get("target_paper_context") or {}).get("section_heading_candidates") or []))[:10],
        },
        "aggregated_planning_hints": {
            "section_splitting_preference": majority_label(section_split_values),
            "bullet_density_preference": majority_label(bullet_density_values),
            "text_density_preference": majority_label(text_density_values),
            "visual_density_preference": majority_label(visual_density_values),
            "layout_bias": [key for key, _count in layout_bias_counts.most_common()],
            "layout_bias_counts": dict(layout_bias_counts),
            "structure_preferences": {
                "prefers_takeaway_slide": majority_label(takeaway_values),
                "prefers_multi_slide_method_section": majority_label(method_values),
                "prefers_multi_slide_results_section": majority_label(results_values),
            },
        },
        "merged_presentation_guidelines": {
            "narrative_flow_preferences": dedupe_keep_order(narrative_flow, limit=8),
            "section_emphasis_patterns": dedupe_keep_order(section_emphasis, limit=8),
            "content_style_preferences": dedupe_keep_order(content_style, limit=8),
            "compression_patterns": dedupe_keep_order(compression_patterns, limit=8),
            "visual_usage_patterns": dedupe_keep_order(visual_usage, limit=8),
            "signature_choices": dedupe_keep_order(signature_choices, limit=8),
        },
        "reference_pair_highlights": reference_highlights[:4],
    }


def locate_raw_content(plan_path: Path) -> Path | None:
    match = re.match(r"(<[^>]+>)_slide_plan(?:_[^.]+)?\.json$", plan_path.name)
    if not match:
        return None
    raw_content_path = plan_path.parent / f"{match.group(1)}_raw_content.json"
    return raw_content_path if raw_content_path.exists() else None


def summarize_target_paper_outline(plan_paths: list[Path], pair_context: dict[str, Any]) -> dict[str, Any]:
    raw_content_path = next((path for path in (locate_raw_content(plan_path) for plan_path in plan_paths) if path), None)
    target_summary = {
        "target_paper_id": pair_context.get("target_paper_id"),
        "target_paper_path": pair_context.get("target_paper_path"),
        "title_guess": ((pair_context.get("target_paper_context") or {}).get("title_guess") or ""),
        "section_heading_candidates": list(((pair_context.get("target_paper_context") or {}).get("section_heading_candidates") or []))[:10],
        "raw_content_path": str(raw_content_path) if raw_content_path else None,
    }

    if not raw_content_path:
        return target_summary

    raw_content = load_json(raw_content_path)
    sections = list(raw_content.get("sections") or [])
    section_summaries = []
    total_subsections = 0

    for section in sections:
        title = str(section.get("title") or "").strip()
        subsections = list(section.get("subsections") or [])
        subsection_titles = [str(item.get("title") or "").strip() for item in subsections if isinstance(item, dict)]
        total_subsections += len(subsection_titles)
        section_summaries.append(
            {
                "title": title,
                "subsection_count": len(subsection_titles),
                "subsection_titles": subsection_titles[:6],
            }
        )

    target_summary.update(
        {
            "metadata": dict(raw_content.get("metadata") or {}),
            "section_titles": [item["title"] for item in section_summaries],
            "section_count": len(section_summaries),
            "subsection_count": total_subsections,
            "sections": section_summaries,
        }
    )
    return target_summary


def compare_preference_label(actual: str, preferred: str) -> float | None:
    actual = str(actual or "unknown").strip().lower()
    preferred = str(preferred or "unknown").strip().lower()
    if preferred in {"", "unknown"}:
        return None
    if actual == preferred:
        return 1.0

    order = {"low": 0, "medium": 1, "high": 2}
    if preferred in order and actual in order:
        return 0.5 if abs(order[preferred] - order[actual]) == 1 else 0.0
    if preferred == "balanced":
        return 0.5 if actual in {"coarse", "fine_grained"} else 0.0
    if preferred in {"coarse", "fine_grained"} and actual == "balanced":
        return 0.5
    return 0.0


def compare_layout_bias(observed: list[str], preferred: list[str]) -> float | None:
    observed_set = set(str(item).strip() for item in observed if str(item).strip())
    preferred_set = set(str(item).strip() for item in preferred if str(item).strip())
    if not preferred_set:
        return None
    if not observed_set:
        return 0.0
    return round(len(observed_set & preferred_set) / len(observed_set | preferred_set), 4)


def compare_structure_signal(observed: bool, preference_label: str) -> float | None:
    preference_label = str(preference_label or "unknown").strip().lower()
    if preference_label in {"", "unknown"}:
        return None
    if preference_label == "high":
        return 1.0 if observed else 0.0
    if preference_label == "medium":
        return 1.0 if observed else 0.5
    if preference_label == "low":
        return 0.5 if observed else 1.0
    return None


def build_heuristic_signal_summary(plan_summary: dict[str, Any], pair_guideline_summary: dict[str, Any]) -> dict[str, Any]:
    planning_hints = dict(pair_guideline_summary.get("aggregated_planning_hints") or {})
    structure_prefs = dict(planning_hints.get("structure_preferences") or {})

    signal_scores = {
        "section_splitting_signal": compare_preference_label(
            plan_summary.get("section_splitting_estimate", "unknown"),
            planning_hints.get("section_splitting_preference", "unknown"),
        ),
        "bullet_density_signal": compare_preference_label(
            plan_summary.get("bullet_density_estimate", "unknown"),
            planning_hints.get("bullet_density_preference", "unknown"),
        ),
        "text_density_signal": compare_preference_label(
            plan_summary.get("text_density_estimate", "unknown"),
            planning_hints.get("text_density_preference", "unknown"),
        ),
        "figure_usage_signal": compare_preference_label(
            plan_summary.get("figure_usage_estimate", "unknown"),
            planning_hints.get("visual_density_preference", "unknown"),
        ),
        "layout_bias_signal": compare_layout_bias(
            list(plan_summary.get("layout_bias_observed") or []),
            list(planning_hints.get("layout_bias") or []),
        ),
        "takeaway_signal": compare_structure_signal(
            bool(plan_summary.get("prefers_takeaway_like_close")),
            structure_prefs.get("prefers_takeaway_slide", "unknown"),
        ),
        "method_expansion_signal": compare_structure_signal(
            bool(plan_summary.get("multi_slide_method_like_sections")),
            structure_prefs.get("prefers_multi_slide_method_section", "unknown"),
        ),
        "results_expansion_signal": compare_structure_signal(
            bool(plan_summary.get("multi_slide_results_like_sections")),
            structure_prefs.get("prefers_multi_slide_results_section", "unknown"),
        ),
    }

    numeric_values = [value for value in signal_scores.values() if value is not None]
    mean_signal = round(sum(numeric_values) / len(numeric_values), 4) if numeric_values else None

    return {
        "signal_scores": signal_scores,
        "mean_signal_score": mean_signal,
        "observations": {
            "section_splitting_estimate": plan_summary.get("section_splitting_estimate"),
            "bullet_density_estimate": plan_summary.get("bullet_density_estimate"),
            "text_density_estimate": plan_summary.get("text_density_estimate"),
            "layout_bias_observed": plan_summary.get("layout_bias_observed"),
            "has_takeaway_like_close": bool(plan_summary.get("prefers_takeaway_like_close")),
            "multi_slide_method_like_sections": plan_summary.get("multi_slide_method_like_sections"),
            "multi_slide_results_like_sections": plan_summary.get("multi_slide_results_like_sections"),
        },
    }


def render_prompt(
    prompt_path: Path,
    *,
    pair_guideline_summary: dict[str, Any],
    target_paper_outline_summary: dict[str, Any],
    baseline_plan_summary: dict[str, Any],
    pair_guided_plan_summary: dict[str, Any],
    extracted_asset_summary: dict[str, Any],
    heuristic_signal_comparison: dict[str, Any],
) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    user_prompt = template.render(
        pair_guideline_summary_json=pair_guideline_summary,
        target_paper_outline_summary_json=target_paper_outline_summary,
        baseline_plan_summary_json=baseline_plan_summary,
        pair_guided_plan_summary_json=pair_guided_plan_summary,
        extracted_asset_summary_json=extracted_asset_summary,
        heuristic_signal_comparison_json=heuristic_signal_comparison,
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": user_prompt,
    }


def call_judge(
    model: str,
    system_prompt: str,
    user_prompt: str,
    *,
    request_timeout: float,
    verbose: bool,
) -> dict[str, Any]:
    client = build_openai_client()
    resolved_model = resolve_direct_model_name(model)

    request_kwargs: dict[str, Any] = {
        "model": resolved_model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.1,
        "timeout": request_timeout,
    }
    if "gpt-5" in resolved_model.lower():
        request_kwargs["max_completion_tokens"] = 2200
    else:
        request_kwargs["max_tokens"] = 2200

    log(
        f"Sending pair win-rate judge request to model={resolved_model!r} with timeout={request_timeout:.1f}s",
        verbose=verbose,
    )
    started_at = time.time()
    response = client.chat.completions.create(**request_kwargs)
    elapsed = time.time() - started_at
    log(f"Judge response received in {elapsed:.2f}s", verbose=verbose)
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def coerce_scores(report: dict[str, Any]) -> dict[str, Any]:
    for bucket in ("baseline", "pair_guided"):
        section = report.setdefault(bucket, {})
        scores = section.setdefault("scores", {})
        for key in SCORE_KEYS:
            value = scores.get(key, 0.0)
            try:
                value = float(value)
            except Exception:
                value = 0.0
            scores[key] = max(0.0, min(1.0, value))

    lift = report.setdefault("lift", {})
    for key in SCORE_KEYS:
        base_val = report["baseline"]["scores"][key]
        pair_val = report["pair_guided"]["scores"][key]
        lift[key] = round(pair_val - base_val, 4)

    summary = report.setdefault("summary", {})
    try:
        margin = float(summary.get("winning_margin", 0.0))
    except Exception:
        margin = 0.0
    summary["winning_margin"] = round(max(-1.0, min(1.0, margin)), 4)
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate baseline vs pair-guided win-rate from pair-guideline context.")
    parser.add_argument("--pair-context", type=Path, required=True, help="Path to target-specific pair-guideline context JSON.")
    parser.add_argument("--baseline-plan", type=Path, required=True, help="Path to baseline slide_plan JSON.")
    parser.add_argument("--pair-guided-plan", type=Path, required=True, help="Path to pair-guided slide_plan JSON.")
    parser.add_argument("--model", default="gpt-5", help="Judge model identifier.")
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--output", type=Path, default=None, help="Optional output JSON path.")
    parser.add_argument(
        "--request-timeout",
        type=float,
        default=180.0,
        help="Timeout in seconds for the judge model request.",
    )
    parser.add_argument("--print-prompt", action="store_true")
    parser.add_argument("--verbose", action="store_true", help="Print progress messages to stderr.")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")

    for path_arg in (args.pair_context, args.baseline_plan, args.pair_guided_plan, args.prompt_path):
        if not path_arg.exists():
            raise FileNotFoundError(f"Required input not found: {path_arg}")

    log("Loading pair-guideline context and slide plans", verbose=args.verbose)
    pair_context = load_json(args.pair_context)
    baseline_plan = load_json(args.baseline_plan)
    pair_guided_plan = load_json(args.pair_guided_plan)

    log("Summarizing plans and pair-guideline context", verbose=args.verbose)
    pair_guideline_summary = summarize_pair_guideline_context(pair_context)
    baseline_summary = summarize_slide_plan(baseline_plan)
    pair_guided_summary = summarize_slide_plan(pair_guided_plan)
    target_paper_outline_summary = summarize_target_paper_outline(
        [args.baseline_plan, args.pair_guided_plan],
        pair_context,
    )
    extracted_asset_summary = {
        "baseline": summarize_extracted_assets(args.baseline_plan),
        "pair_guided": summarize_extracted_assets(args.pair_guided_plan),
    }
    heuristic_signal_comparison = {
        "baseline": build_heuristic_signal_summary(baseline_summary, pair_guideline_summary),
        "pair_guided": build_heuristic_signal_summary(pair_guided_summary, pair_guideline_summary),
    }

    log("Rendering judge prompt", verbose=args.verbose)
    prompt = render_prompt(
        args.prompt_path,
        pair_guideline_summary=pair_guideline_summary,
        target_paper_outline_summary=target_paper_outline_summary,
        baseline_plan_summary=baseline_summary,
        pair_guided_plan_summary=pair_guided_summary,
        extracted_asset_summary=extracted_asset_summary,
        heuristic_signal_comparison=heuristic_signal_comparison,
    )

    if args.print_prompt:
        print(prompt["user_prompt"])
        return

    report = call_judge(
        args.model,
        prompt["system_prompt"],
        prompt["user_prompt"],
        request_timeout=args.request_timeout,
        verbose=args.verbose,
    )
    report = coerce_scores(report)
    report.setdefault("summary", {})
    report["summary"]["guardrail_triggered"] = bool(report["summary"].get("guardrail_triggered", False))
    report["inputs"] = {
        "pair_context": str(args.pair_context),
        "baseline_plan": str(args.baseline_plan),
        "pair_guided_plan": str(args.pair_guided_plan),
        "judge_model": args.model,
    }
    report["derived_pair_guideline_summary"] = pair_guideline_summary
    report["derived_target_paper_outline_summary"] = target_paper_outline_summary
    report["derived_plan_summaries"] = {
        "baseline": baseline_summary,
        "pair_guided": pair_guided_summary,
    }
    report["extracted_asset_summary"] = extracted_asset_summary
    report["heuristic_signal_comparison"] = heuristic_signal_comparison

    output_text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output_text, encoding="utf-8")
        log(f"Wrote report to {args.output}", verbose=args.verbose)
    print(output_text)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
