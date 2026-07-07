#!/usr/bin/env python3
"""Shared helpers for personalization evaluation scripts."""

from __future__ import annotations

import json
import re
import sys
import time
from pathlib import Path
from typing import Any

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from SlidesAgent.slide_plan_summary import mean
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


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


def log(message: str, *, verbose: bool) -> None:
    if verbose:
        print(f"[personalization-eval] {message}", file=sys.stderr, flush=True)


def extract_json_object(raw_text: str) -> dict[str, Any]:
    text = (raw_text or "").strip()
    if not text:
        raise ValueError("Model returned an empty response.")

    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, flags=re.DOTALL)
    if fenced:
        text = fenced.group(1)

    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError(f"Could not find JSON object in response: {raw_text[:400]}")

    return json.loads(text[start : end + 1])


def _count_visual_refs(value: Any, prefix: str) -> int:
    count = 0
    if isinstance(value, dict):
        for key, item in value.items():
            if isinstance(key, str) and key.startswith(prefix):
                count += 1
            count += _count_visual_refs(item, prefix)
    elif isinstance(value, list):
        for item in value:
            count += _count_visual_refs(item, prefix)
    return count


def summarize_extracted_assets(plan_path: Path) -> dict[str, Any]:
    parent = plan_path.parent
    name = plan_path.name
    model_match = re.match(r"(<[^>]+>)_slide_plan(?:_[^.]+)?\.json$", name)
    if not model_match:
        return {}

    prefix = model_match.group(1)
    figures_path = parent / f"{prefix}_figures.json"
    formula_path = parent / f"{prefix}_formula_match.json"

    summary = {
        "figures_path": str(figures_path),
        "formula_path": str(formula_path),
        "image_refs_extracted": 0,
        "table_refs_extracted": 0,
        "formula_sections_extracted": 0,
        "table_penalty_applicable": True,
        "formula_penalty_applicable": True,
    }

    if figures_path.exists():
        figures_json = json.loads(figures_path.read_text(encoding="utf-8"))
        summary["image_refs_extracted"] = _count_visual_refs(figures_json, "image")
        summary["table_refs_extracted"] = _count_visual_refs(figures_json, "table")

    if formula_path.exists():
        formula_json = json.loads(formula_path.read_text(encoding="utf-8"))
        sections = formula_json.get("sections", []) if isinstance(formula_json, dict) else []
        if isinstance(sections, list):
            summary["formula_sections_extracted"] = len(sections)

    summary["table_penalty_applicable"] = summary["table_refs_extracted"] > 0
    summary["formula_penalty_applicable"] = summary["formula_sections_extracted"] > 0
    return summary


def build_numeric_comparison(
    target_summary: dict[str, Any],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
) -> dict[str, Any]:
    if not target_summary:
        return {}

    metric_specs = {
        "slide_count": ("target_slide_count", "slide_count"),
        "avg_slides_per_section": ("target_avg_slides_per_section", "avg_slides_per_section"),
        "avg_bullets_per_slide": ("target_avg_bullets_per_slide", "avg_bullets_per_slide"),
        "avg_words_per_slide": ("target_avg_words_per_slide", "avg_words_per_slide"),
        "figure_slide_fraction": ("target_fraction_figure_slides", "figure_slide_fraction"),
        "table_slide_fraction": ("target_fraction_table_slides", "table_slide_fraction"),
        "formula_slide_fraction": ("target_fraction_formula_slides", "formula_slide_fraction"),
        "text_only_fraction": ("target_fraction_text_only_slides", "text_only_fraction"),
        "multi_visual_fraction": ("target_fraction_multi_visual_slides", "multi_visual_fraction"),
        "formula_capable_fraction": ("target_fraction_formula_capable_slides", "formula_capable_fraction"),
        "image_right_fraction": ("target_fraction_image_right_slides", "image_right_fraction"),
        "image_left_fraction": ("target_fraction_image_left_slides", "image_left_fraction"),
        "image_top_fraction": ("target_fraction_image_top_slides", "image_top_fraction"),
    }

    comparison: dict[str, Any] = {}
    for output_key, (target_key, summary_key) in metric_specs.items():
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
        if abs(baseline_distance - personalized_distance) <= 1e-6:
            closer = "tie"
        elif personalized_distance < baseline_distance:
            closer = "personalized"
        else:
            closer = "baseline"

        comparison[output_key] = {
            "target": round(target_value, 4),
            "baseline": round(baseline_value, 4),
            "personalized": round(personalized_value, 4),
            "baseline_distance": round(baseline_distance, 4),
            "personalized_distance": round(personalized_distance, 4),
            "closer_to_target": closer,
        }

    return comparison


def call_alignment_judge(
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
        f"Sending judge request to model={resolved_model!r} with timeout={request_timeout:.1f}s",
        verbose=verbose,
    )
    started_at = time.time()
    response = client.chat.completions.create(**request_kwargs)
    elapsed = time.time() - started_at
    log(f"Judge response received in {elapsed:.2f}s", verbose=verbose)
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def determine_applicable_dimensions(extracted_asset_summary: dict[str, Any]) -> tuple[list[str], list[str]]:
    applicable = {
        "section_structure_alignment",
        "bullet_density_alignment",
        "text_density_alignment",
        "figure_usage_alignment",
        "layout_bias_alignment",
    }
    skipped: list[str] = []

    baseline_assets = dict(extracted_asset_summary.get("baseline") or {})
    personalized_assets = dict(extracted_asset_summary.get("personalized") or {})

    table_applicable = bool(
        baseline_assets.get("table_penalty_applicable") or personalized_assets.get("table_penalty_applicable")
    )
    formula_applicable = bool(
        baseline_assets.get("formula_penalty_applicable") or personalized_assets.get("formula_penalty_applicable")
    )

    if table_applicable:
        applicable.add("table_usage_alignment")
    else:
        skipped.append("table_usage_alignment")

    if formula_applicable:
        applicable.add("formula_usage_alignment")
    else:
        skipped.append("formula_usage_alignment")

    return [key for key in SCORE_KEYS if key != "overall_style_alignment" and key in applicable], skipped


def apply_dimension_applicability(
    report: dict[str, Any],
    extracted_asset_summary: dict[str, Any],
) -> dict[str, Any]:
    applicable_dims, skipped_dims = determine_applicable_dimensions(extracted_asset_summary)

    for metric_key in skipped_dims:
        for bucket in ("baseline", "personalized"):
            report[bucket]["scores"][metric_key] = 0.5
            rationale = report[bucket].setdefault("rationale", {})
            rationale[metric_key] = (
                "Skipped from comparison because the target paper did not expose supporting extracted assets "
                "for this dimension."
            )

    for bucket in ("baseline", "personalized"):
        component_values = [float(report[bucket]["scores"][key]) for key in applicable_dims]
        report[bucket]["scores"]["overall_style_alignment"] = round(mean(component_values), 4) if component_values else 0.5

    lift = report.setdefault("lift", {})
    for key in SCORE_KEYS:
        base_val = report["baseline"]["scores"][key]
        pers_val = report["personalized"]["scores"][key]
        lift[key] = round(pers_val - base_val, 4)

    overall_delta = float(lift.get("overall_style_alignment", 0.0))
    winner = "tie"
    if overall_delta > 0.03:
        winner = "personalized"
    elif overall_delta < -0.03:
        winner = "baseline"

    summary = report.setdefault("summary", {})
    summary["winner"] = winner
    if skipped_dims:
        summary["headline"] = (
            summary.get("headline")
            or "Comparison excludes unsupported asset-dependent dimensions for this target paper."
        )
    report["applicable_dimensions"] = applicable_dims
    report["skipped_dimensions"] = skipped_dims
    return report
