from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any, Dict, List
import yaml
from jinja2 import Environment, StrictUndefined
from utils.src.utils import   get_json_from_response
from utils.wei_utils import *
from utils.pptx_utils import extract_text_from_responses
from SlidesAgent.personalization_targets import (
    build_numeric_target_summary,
    profile_target_tolerance_multiplier,
)
from SlidesAgent.output_paths import (
    figures_json_path,
    formula_match_path,
    formula_mode3_index_path,
    images_filtered_path,
    personalization_trace_path,
    raw_content_path,
    slide_plan_draft_path,
    slide_plan_path,
    slide_plan_repair_report_path,
    tables_filtered_path,
)
from SlidesAgent.slide_plan_summary import summarize_slide_plan
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name
from camel.models import ModelFactory          
from camel.agents import ChatAgent     
from pptx.util import Cm, Pt
import time

FORMULA_TEMPLATE_IDS = {
    "T14_ImageRight_1Formula",
    "T15_ImageLeft_1Formula",
    "T16_1Img_2formula_TopTextBottom",
    "T17_2Img_1formula_TopTextBottom",
    "T18_2formula_TopTextBottom",
}

STRONG_PROMPT_PATH = Path("utils/prompt_templates/layout_agent_xin_strong.yaml")
DEFAULT_STRONG_REPAIR_ROUNDS = 3
NUMERIC_TARGET_SUMMARY_KEYS = {
    "target_avg_slides_per_section": "avg_slides_per_section",
    "target_avg_bullets_per_slide": "avg_bullets_per_slide",
    "target_avg_words_per_slide": "avg_words_per_slide",
    "target_fraction_figure_slides": "figure_slide_fraction",
    "target_fraction_table_slides": "table_slide_fraction",
    "target_fraction_formula_slides": "formula_slide_fraction",
    "target_fraction_text_only_slides": "text_only_fraction",
    "target_fraction_multi_visual_slides": "multi_visual_fraction",
    "target_fraction_formula_capable_slides": "formula_capable_fraction",
    "target_fraction_image_right_slides": "image_right_fraction",
    "target_fraction_image_left_slides": "image_left_fraction",
    "target_fraction_image_top_slides": "image_top_fraction",
}


def plan_variant_suffix(args) -> str:
    return getattr(
        args,
        "output_variant_suffix",
        "_personalized" if getattr(args, "use_author_preferences", False) else "_baseline",
    )


def _fallback_template_without_formulas(slide: Dict[str, Any]) -> str:
    image_count = len(slide.get("images") or [])
    table_count = len(slide.get("tables") or [])
    visual_count = image_count + table_count

    if visual_count <= 0:
        return "T1_TextOnly"
    if visual_count >= 2:
        return "T5_TwoImages2"

    template_id = str(slide.get("template_id") or "")
    if "ImageLeft" in template_id:
        return "T3_ImageLeft"
    if table_count > 0:
        return "T4_ImageTop"
    return "T2_ImageRight"


def sanitize_slide_plan_templates(slide_plan: Dict[str, Any]) -> Dict[str, Any]:
    slides = list(slide_plan.get("slides") or [])
    for slide in slides:
        if not isinstance(slide, dict):
            continue
        template_id = str(slide.get("template_id") or "")
        formulas = list(slide.get("formulas") or [])
        if template_id in FORMULA_TEMPLATE_IDS and not formulas:
            slide["template_id"] = _fallback_template_without_formulas(slide)
    slide_plan["slides"] = slides
    return slide_plan


def _collect_formula_asset_refs(node: Any, found: set[str]) -> None:
    if isinstance(node, dict):
        for value in node.values():
            _collect_formula_asset_refs(value, found)
        return
    if isinstance(node, list):
        for value in node:
            _collect_formula_asset_refs(value, found)
        return
    if isinstance(node, str):
        value = node.strip()
        if value.endswith(".png") and "formula" in value.lower():
            found.add(value)


def derive_asset_support(
    *,
    formulas_json: Dict[str, Any] | List[Any],
    images: Dict[str, Any],
    tables: Dict[str, Any],
) -> Dict[str, Any]:
    formula_refs: set[str] = set()
    _collect_formula_asset_refs(formulas_json, formula_refs)
    return {
        "supports_figures": bool(images),
        "supports_tables": bool(tables),
        "supports_formulas": bool(formula_refs),
        "supports_visual_layout_changes": bool(images) or bool(tables),
    }


def infer_priority_metric_keys(
    author_preference_profile: Dict[str, Any] | None,
    numeric: Dict[str, Any],
) -> List[str]:
    if not author_preference_profile:
        return []

    ordered_metric_keys = [
        "target_avg_slides_per_section",
        "target_avg_bullets_per_slide",
        "target_avg_words_per_slide",
        "target_fraction_figure_slides",
        "target_fraction_formula_slides",
        "target_fraction_table_slides",
        "target_fraction_multi_visual_slides",
        "target_fraction_formula_capable_slides",
        "target_fraction_image_top_slides",
        "target_fraction_image_right_slides",
        "target_fraction_image_left_slides",
        "target_fraction_text_only_slides",
        "slide_count",
    ]
    priority = [metric_key for metric_key in ordered_metric_keys if metric_key in numeric]
    if "slide_count_range" in numeric and "slide_count" not in priority:
        priority.append("slide_count")
    return priority


def target_delta_threshold(metric_key: str) -> float:
    if metric_key == "slide_count":
        return 1.0
    if metric_key == "target_avg_slides_per_section":
        return 0.2
    if metric_key == "target_avg_bullets_per_slide":
        return 0.3
    if metric_key == "target_avg_words_per_slide":
        return 3.0
    if "fraction" in metric_key:
        return 0.05
    return 0.1


def compute_target_distance_breakdown(
    plan_summary: Dict[str, Any],
    target_summary: Dict[str, Any] | None,
) -> Dict[str, float]:
    target_summary = target_summary or {}
    distances: Dict[str, float] = {}
    for target_key, summary_key in NUMERIC_TARGET_SUMMARY_KEYS.items():
        if target_key not in target_summary or summary_key not in plan_summary:
            continue
        try:
            target_value = float(target_summary[target_key])
            observed_value = float(plan_summary[summary_key])
        except Exception:
            continue
        distances[target_key] = round(abs(observed_value - target_value), 4)
    return distances


def average_target_distance(
    plan_summary: Dict[str, Any],
    target_summary: Dict[str, Any] | None,
) -> float:
    distances = compute_target_distance_breakdown(plan_summary, target_summary)
    if not distances:
        return 0.0
    return round(sum(distances.values()) / len(distances), 4)


def build_planner_target_brief(
    author_preference_profile: Dict[str, Any] | None,
    asset_support: Dict[str, Any],
) -> Dict[str, Any]:
    numeric = build_numeric_target_summary(author_preference_profile)
    priority_metric_keys = infer_priority_metric_keys(author_preference_profile, numeric)
    actionable_priority_metrics = [
        metric_key for metric_key in priority_metric_keys if is_actionable_mismatch(metric_key, asset_support)
    ]
    blocked_priority_metrics = [
        metric_key for metric_key in priority_metric_keys if metric_key not in actionable_priority_metrics
    ]
    return {
        "numeric_target_summary": numeric,
        "priority_metrics": actionable_priority_metrics,
        "blocked_priority_metrics": blocked_priority_metrics,
    }


def mismatch_actionability_details(metric_key: str, asset_support: Dict[str, Any]) -> tuple[bool, str]:
    if metric_key == "slide_count":
        return True, "always_actionable"
    if metric_key in {"target_fraction_formula_slides", "target_fraction_formula_capable_slides"}:
        return asset_support.get("supports_formulas", False), "missing_formula_assets"
    if metric_key == "target_fraction_table_slides":
        return asset_support.get("supports_tables", False), "missing_table_assets"
    if metric_key in {
        "target_fraction_figure_slides",
        "target_fraction_multi_visual_slides",
        "target_fraction_image_top_slides",
        "target_fraction_image_right_slides",
        "target_fraction_image_left_slides",
    }:
        return asset_support.get("supports_visual_layout_changes", False), "missing_visual_assets"
    return True, "always_actionable"


def is_actionable_mismatch(metric_key: str, asset_support: Dict[str, Any]) -> bool:
    actionable, _ = mismatch_actionability_details(metric_key, asset_support)
    return actionable


def select_actionable_targets(
    mismatches: List[Dict[str, Any]],
    asset_support: Dict[str, Any],
    *,
    max_targets: int | None = None,
) -> List[Dict[str, Any]]:
    selected: List[Dict[str, Any]] = []
    for mismatch in mismatches:
        metric_key = str(mismatch.get("metric") or "")
        if not bool(mismatch.get("actionable", is_actionable_mismatch(metric_key, asset_support))):
            continue
        enriched = dict(mismatch)
        enriched["minimum_delta"] = target_delta_threshold(metric_key)
        selected.append(enriched)
        if max_targets is not None and len(selected) >= max_targets:
            break
    return selected


def _section_title_drift_count(before: List[str], after: List[str]) -> int:
    limit = min(len(before), len(after))
    drift = sum(1 for idx in range(limit) if str(before[idx]) != str(after[idx]))
    drift += abs(len(before) - len(after))
    return drift


def compute_anchor_drift_score(
    anchor_summary: Dict[str, Any],
    candidate_summary: Dict[str, Any],
) -> float:
    drift = 0.0
    drift += abs(float(candidate_summary.get("slide_count", 0.0)) - float(anchor_summary.get("slide_count", 0.0))) * 0.15
    drift += abs(float(candidate_summary.get("avg_slides_per_section", 0.0)) - float(anchor_summary.get("avg_slides_per_section", 0.0))) * 0.35
    drift += abs(float(candidate_summary.get("avg_bullets_per_slide", 0.0)) - float(anchor_summary.get("avg_bullets_per_slide", 0.0))) * 0.2
    drift += abs(float(candidate_summary.get("avg_words_per_slide", 0.0)) - float(anchor_summary.get("avg_words_per_slide", 0.0))) / 10.0
    drift += abs(float(candidate_summary.get("figure_slide_fraction", 0.0)) - float(anchor_summary.get("figure_slide_fraction", 0.0))) * 1.5
    drift += abs(float(candidate_summary.get("table_slide_fraction", 0.0)) - float(anchor_summary.get("table_slide_fraction", 0.0))) * 1.5
    drift += abs(float(candidate_summary.get("formula_slide_fraction", 0.0)) - float(anchor_summary.get("formula_slide_fraction", 0.0))) * 1.5
    drift += abs(float(candidate_summary.get("text_only_fraction", 0.0)) - float(anchor_summary.get("text_only_fraction", 0.0))) * 1.2
    drift += abs(float(candidate_summary.get("multi_visual_fraction", 0.0)) - float(anchor_summary.get("multi_visual_fraction", 0.0))) * 1.2
    drift += abs(float(candidate_summary.get("image_top_fraction", 0.0)) - float(anchor_summary.get("image_top_fraction", 0.0))) * 1.0
    drift += abs(float(candidate_summary.get("image_right_fraction", 0.0)) - float(anchor_summary.get("image_right_fraction", 0.0))) * 1.0
    drift += abs(float(candidate_summary.get("image_left_fraction", 0.0)) - float(anchor_summary.get("image_left_fraction", 0.0))) * 1.0
    drift += _section_title_drift_count(
        list(anchor_summary.get("section_titles") or []),
        list(candidate_summary.get("section_titles") or []),
    ) * 0.35
    return round(drift, 4)


def build_repair_directives(
    author_preference_profile: Dict[str, Any] | None,
    plan_summary: Dict[str, Any],
    asset_support: Dict[str, Any] | None = None,
) -> Dict[str, Any]:
    if not author_preference_profile:
        return {
            "needs_repair": False,
            "target_summary": {},
            "goals": [],
            "mismatches": [],
            "total_mismatch_score": 0.0,
            "metric_mismatch_scores": {},
            "priority_metrics": [],
            "edit_brief": [],
        }

    if not isinstance(plan_summary, dict):
        return {
            "needs_repair": False,
            "target_summary": {},
            "goals": [],
            "mismatches": [],
            "total_mismatch_score": 0.0,
            "metric_mismatch_scores": {},
            "priority_metrics": [],
            "edit_brief": [],
            "reason": "missing_plan_summary",
        }

    numeric = build_numeric_target_summary(author_preference_profile)
    priority_metric_keys = infer_priority_metric_keys(author_preference_profile, numeric)
    asset_support = asset_support or {}
    tolerance_multiplier = profile_target_tolerance_multiplier(author_preference_profile)
    mismatches: List[Dict[str, Any]] = []
    goals: List[str] = []
    blocked_goals: List[str] = []
    metric_goals: Dict[str, str] = {}
    metric_mismatch_scores: Dict[str, float] = {}
    blocked_metric_mismatch_scores: Dict[str, float] = {}
    total_mismatch_score = 0.0
    total_blocked_mismatch_score = 0.0

    actionable_priority_metrics = [
        metric_key for metric_key in priority_metric_keys if is_actionable_mismatch(metric_key, asset_support)
    ]
    blocked_priority_metrics = [
        metric_key for metric_key in priority_metric_keys if metric_key not in actionable_priority_metrics
    ]

    def record_range_mismatch(key: str, observed: float, allowed_range: List[Any], *, goal_low: str, goal_high: str) -> None:
        nonlocal total_mismatch_score, total_blocked_mismatch_score
        if not isinstance(allowed_range, list) or len(allowed_range) != 2:
            return
        try:
            low = float(allowed_range[0])
            high = float(allowed_range[1])
        except Exception:
            return
        if observed < low:
            delta = low - observed
            actionable, blocked_reason = mismatch_actionability_details(key, asset_support)
            mismatch = {
                "metric": key,
                "observed": observed,
                "target": [low, high],
                "delta": round(delta, 4),
                "priority": key in actionable_priority_metrics,
                "priority_requested": key in priority_metric_keys,
                "actionable": actionable,
                "blocked_reason": None if actionable else blocked_reason,
                "recommended_edit": goal_low,
            }
            mismatches.append(mismatch)
            metric_goals[key] = goal_low
            if actionable:
                total_mismatch_score += delta
                metric_mismatch_scores[key] = round(delta, 4)
                goals.append(goal_low)
            else:
                total_blocked_mismatch_score += delta
                blocked_metric_mismatch_scores[key] = round(delta, 4)
                blocked_goals.append(goal_low)
        elif observed > high:
            delta = observed - high
            actionable, blocked_reason = mismatch_actionability_details(key, asset_support)
            mismatch = {
                "metric": key,
                "observed": observed,
                "target": [low, high],
                "delta": round(delta, 4),
                "priority": key in actionable_priority_metrics,
                "priority_requested": key in priority_metric_keys,
                "actionable": actionable,
                "blocked_reason": None if actionable else blocked_reason,
                "recommended_edit": goal_high,
            }
            mismatches.append(mismatch)
            metric_goals[key] = goal_high
            if actionable:
                total_mismatch_score += delta
                metric_mismatch_scores[key] = round(delta, 4)
                goals.append(goal_high)
            else:
                total_blocked_mismatch_score += delta
                blocked_metric_mismatch_scores[key] = round(delta, 4)
                blocked_goals.append(goal_high)

    def record_scalar_mismatch(key: str, observed: float, tolerance: float, goal_low: str, goal_high: str) -> None:
        nonlocal total_mismatch_score, total_blocked_mismatch_score
        if key not in numeric:
            return
        try:
            target = float(numeric[key])
        except Exception:
            return
        effective_tolerance = tolerance * tolerance_multiplier
        delta = observed - target
        if abs(delta) <= effective_tolerance:
            return
        score = abs(delta) / max(effective_tolerance, 1e-6)
        chosen_goal = goal_low if delta < 0 else goal_high
        metric_goals[key] = chosen_goal
        actionable, blocked_reason = mismatch_actionability_details(key, asset_support)
        mismatch = {
            "metric": key,
            "observed": observed,
            "target": target,
            "delta": round(delta, 4),
            "priority": key in actionable_priority_metrics,
            "priority_requested": key in priority_metric_keys,
            "actionable": actionable,
            "blocked_reason": None if actionable else blocked_reason,
            "tolerance": round(effective_tolerance, 4),
            "recommended_edit": chosen_goal,
        }
        mismatches.append(mismatch)
        if actionable:
            total_mismatch_score += score
            metric_mismatch_scores[key] = round(score, 4)
            goals.append(chosen_goal)
        else:
            total_blocked_mismatch_score += score
            blocked_metric_mismatch_scores[key] = round(score, 4)
            blocked_goals.append(chosen_goal)

    record_range_mismatch(
        "slide_count",
        float(plan_summary.get("slide_count", 0.0)),
        numeric.get("slide_count_range", []),
        goal_low="Increase the total slide count toward the author's typical range by splitting only the most overloaded slides.",
        goal_high="Reduce the total slide count toward the author's typical range by merging thin adjacent slides.",
    )
    record_scalar_mismatch(
        "target_avg_slides_per_section",
        float(plan_summary.get("avg_slides_per_section", 0.0)),
        0.35,
        "Spread content a bit more within sections so the deck is less compressed.",
        "Compress section-level splitting so each section uses fewer slides on average.",
    )
    record_scalar_mismatch(
        "target_avg_bullets_per_slide",
        float(plan_summary.get("avg_bullets_per_slide", 0.0)),
        0.6,
        "Increase bullet density modestly by merging sparse neighboring slides or enriching thin slides.",
        "Reduce bullet density by trimming overly busy slides or splitting only when necessary.",
    )
    record_scalar_mismatch(
        "target_avg_words_per_slide",
        float(plan_summary.get("avg_words_per_slide", 0.0)),
        4.0,
        "Increase text density modestly while keeping slides legible.",
        "Reduce text density modestly while preserving the core message of each slide.",
    )
    record_scalar_mismatch(
        "target_fraction_figure_slides",
        float(plan_summary.get("figure_slide_fraction", 0.0)),
        0.12,
        "Use figures on a larger fraction of slides when supporting assets already exist.",
        "Use figures more selectively and favor text-led slides when the current plan is too image-heavy.",
    )
    record_scalar_mismatch(
        "target_fraction_table_slides",
        float(plan_summary.get("table_slide_fraction", 0.0)),
        0.1,
        "Use tables slightly more often when the paper provides them and they support the story.",
        "Use tables more selectively when they are overrepresented relative to the target style.",
    )
    record_scalar_mismatch(
        "target_fraction_formula_slides",
        float(plan_summary.get("formula_slide_fraction", 0.0)),
        0.1,
        "Use formula-capable slides more often when equations are available and relevant.",
        "Use formulas more selectively when the current draft leans too heavily on them.",
    )
    record_scalar_mismatch(
        "target_fraction_text_only_slides",
        float(plan_summary.get("text_only_fraction", 0.0)),
        0.12,
        "Allow slightly more text-only slides where visuals are not essential.",
        "Convert some text-only slides into visual-supported slides when assets are available.",
    )
    record_scalar_mismatch(
        "target_fraction_multi_visual_slides",
        float(plan_summary.get("multi_visual_fraction", 0.0)),
        0.12,
        "Use multi-visual layouts more often when legibility and available assets support them.",
        "Reduce multi-visual layouts when the deck is visually denser than the author's style target.",
    )
    record_scalar_mismatch(
        "target_fraction_formula_capable_slides",
        float(plan_summary.get("formula_capable_fraction", 0.0)),
        0.12,
        "Use formula-capable layouts more often when equations are available.",
        "Use formula-capable layouts more selectively when they are overused.",
    )
    record_scalar_mismatch(
        "target_fraction_image_right_slides",
        float(plan_summary.get("image_right_fraction", 0.0)),
        0.1,
        "Use right-image layouts slightly more often when they fit the content.",
        "Use right-image layouts less often so the deck better matches the target layout mix.",
    )
    record_scalar_mismatch(
        "target_fraction_image_left_slides",
        float(plan_summary.get("image_left_fraction", 0.0)),
        0.1,
        "Use left-image layouts slightly more often when they fit the content.",
        "Use left-image layouts less often so the deck better matches the target layout mix.",
    )
    record_scalar_mismatch(
        "target_fraction_image_top_slides",
        float(plan_summary.get("image_top_fraction", 0.0)),
        0.1,
        "Use top-image layouts slightly more often for wide visuals or tables when supported.",
        "Use top-image layouts less often when they exceed the target layout mix.",
    )

    mismatches.sort(
        key=lambda item: (
            not item.get("actionable", False),
            not item.get("priority", False),
            -abs(float(item.get("delta", 0.0))),
        )
    )
    actionable_targets = select_actionable_targets(mismatches, asset_support)

    deduped_goals: List[str] = []
    for goal in goals:
        if goal not in deduped_goals:
            deduped_goals.append(goal)

    deduped_blocked_goals: List[str] = []
    for goal in blocked_goals:
        if goal not in deduped_blocked_goals:
            deduped_blocked_goals.append(goal)

    edit_brief = [item["recommended_edit"] for item in actionable_targets]

    return {
        "needs_repair": bool(actionable_targets),
        "target_summary": numeric,
        "goals": deduped_goals[:8],
        "blocked_goals": deduped_blocked_goals[:8],
        "mismatches": mismatches,
        "total_mismatch_score": round(total_mismatch_score, 4),
        "metric_mismatch_scores": metric_mismatch_scores,
        "blocked_total_mismatch_score": round(total_blocked_mismatch_score, 4),
        "blocked_metric_mismatch_scores": blocked_metric_mismatch_scores,
        "priority_metrics": actionable_priority_metrics,
        "blocked_priority_metrics": blocked_priority_metrics,
        "requested_priority_metrics": priority_metric_keys,
        "tolerance_multiplier": round(tolerance_multiplier, 4),
        "edit_brief": edit_brief,
        "actionable_targets": actionable_targets,
    }


def evaluate_repair_acceptance(
    draft_summary: Dict[str, Any],
    repaired_summary: Dict[str, Any],
    draft_directives: Dict[str, Any],
    repaired_directives: Dict[str, Any],
) -> Dict[str, Any]:
    target_summary = dict(draft_directives.get("target_summary") or {})
    draft_score = float(draft_directives.get("total_mismatch_score", 0.0))
    repaired_score = float(repaired_directives.get("total_mismatch_score", 0.0))
    score_improvement = round(draft_score - repaired_score, 4)
    required_improvement = round(max(0.2, draft_score * 0.04), 4) if draft_score > 0 else 0.0
    draft_avg_target_distance = average_target_distance(draft_summary, target_summary)
    repaired_avg_target_distance = average_target_distance(repaired_summary, target_summary)
    target_distance_improvement = round(draft_avg_target_distance - repaired_avg_target_distance, 4)

    draft_metric_scores = {
        key: float(value)
        for key, value in (draft_directives.get("metric_mismatch_scores") or {}).items()
    }
    repaired_metric_scores = {
        key: float(value)
        for key, value in (repaired_directives.get("metric_mismatch_scores") or {}).items()
    }
    priority_metrics = list(draft_directives.get("priority_metrics") or [])
    actionable_targets = list(draft_directives.get("actionable_targets") or [])
    actionable_thresholds = {
        str(target.get("metric") or ""): float(target.get("minimum_delta") or 0.0)
        for target in actionable_targets
    }
    actionable_metric_keys = set(actionable_thresholds)
    measurable_improvement_count = 0

    improved_priority_metrics: List[str] = []
    worsened_priority_metrics: List[str] = []
    stable_priority_metrics: List[str] = []
    for metric_key in priority_metrics:
        before = draft_metric_scores.get(metric_key, 0.0)
        after = repaired_metric_scores.get(metric_key, 0.0)
        delta = round(before - after, 4)
        if delta > 0.15:
            improved_priority_metrics.append(metric_key)
        elif delta < -0.15:
            worsened_priority_metrics.append(metric_key)
        else:
            stable_priority_metrics.append(metric_key)

    measurable_target_improvements: List[str] = []
    missed_target_improvements: List[str] = []
    for target in actionable_targets:
        metric_key = str(target.get("metric") or "")
        minimum_delta = float(target.get("minimum_delta") or target_delta_threshold(metric_key))
        before = draft_metric_scores.get(metric_key, 0.0)
        after = repaired_metric_scores.get(metric_key, 0.0)
        delta = round(before - after, 4)
        if delta >= minimum_delta:
            measurable_target_improvements.append(metric_key)
            measurable_improvement_count += 1
        else:
            missed_target_improvements.append(metric_key)

    weighted_total_net_gain = 0.0
    weighted_priority_net_gain = 0.0
    weighted_actionable_net_gain = 0.0
    weighted_regression = 0.0
    severe_worsened_priority_metrics: List[str] = []
    metric_keys = sorted(set(draft_metric_scores) | set(repaired_metric_scores))
    for metric_key in metric_keys:
        before = draft_metric_scores.get(metric_key, 0.0)
        after = repaired_metric_scores.get(metric_key, 0.0)
        delta = before - after
        if metric_key in actionable_metric_keys:
            weight = 1.35
        elif metric_key in priority_metrics:
            weight = 1.15
        else:
            weight = 1.0
        weighted_total_net_gain += delta * weight
        if metric_key in priority_metrics:
            weighted_priority_net_gain += delta * weight
        if metric_key in actionable_metric_keys:
            weighted_actionable_net_gain += delta * weight
        if delta < 0:
            weighted_regression += (-delta) * weight
            severe_threshold = max(0.6, actionable_thresholds.get(metric_key, target_delta_threshold(metric_key)) * 0.6)
            if metric_key in priority_metrics and (-delta) >= severe_threshold:
                severe_worsened_priority_metrics.append(metric_key)

    drift_score = compute_anchor_drift_score(draft_summary, repaired_summary)
    max_allowed_drift = round(max(3.0, draft_score * 0.4), 4) if draft_score > 0 else 3.0
    slide_count_delta = abs(
        float(repaired_summary.get("slide_count", 0.0)) - float(draft_summary.get("slide_count", 0.0))
    )
    max_allowed_slide_count_delta = 5.0
    required_weighted_net_gain = round(max(0.2, required_improvement * 0.65), 4)
    max_allowed_weighted_regression = round(max(1.8, score_improvement * 1.15), 4) if score_improvement > 0 else 1.8
    min_measurable_improvements = 1 if len(actionable_targets) <= 2 else 2
    max_priority_worsen_count = 0 if len(priority_metrics) <= 2 else 1

    accepted = (
        score_improvement >= required_improvement
        and measurable_improvement_count >= min_measurable_improvements
        and weighted_total_net_gain >= required_weighted_net_gain
        and weighted_priority_net_gain > 0.0
        and target_distance_improvement > 0.0
        and weighted_regression <= max_allowed_weighted_regression
        and not severe_worsened_priority_metrics
        and len(worsened_priority_metrics) <= max_priority_worsen_count
        and len(improved_priority_metrics) >= len(worsened_priority_metrics)
        and drift_score <= max_allowed_drift
        and slide_count_delta <= max_allowed_slide_count_delta
    )

    reason_parts: List[str] = []
    if score_improvement < required_improvement:
        reason_parts.append(
            f"repair improvement {score_improvement} did not reach required threshold {required_improvement}"
        )
    if measurable_improvement_count < min_measurable_improvements:
        reason_parts.append(
            f"only {measurable_improvement_count} actionable target(s) improved; required {min_measurable_improvements}"
        )
    if weighted_total_net_gain < required_weighted_net_gain:
        reason_parts.append(
            f"weighted net gain {round(weighted_total_net_gain, 4)} did not reach required threshold {required_weighted_net_gain}"
        )
    if weighted_priority_net_gain <= 0.0:
        reason_parts.append("priority-metric net gain was not positive")
    if target_distance_improvement <= 0.0:
        reason_parts.append(
            f"average numeric target distance did not improve ({draft_avg_target_distance} -> {repaired_avg_target_distance})"
        )
    if weighted_regression > max_allowed_weighted_regression:
        reason_parts.append(
            f"weighted regression {round(weighted_regression, 4)} exceeds threshold {max_allowed_weighted_regression}"
        )
    if severe_worsened_priority_metrics:
        reason_parts.append("severely worsened priority metrics: " + ", ".join(severe_worsened_priority_metrics))
    elif worsened_priority_metrics:
        reason_parts.append("some priority metrics worsened but stayed within the allowed tradeoff budget")
    if len(worsened_priority_metrics) > max_priority_worsen_count:
        reason_parts.append(
            f"too many priority metrics worsened ({len(worsened_priority_metrics)} > {max_priority_worsen_count})"
        )
    if drift_score > max_allowed_drift:
        reason_parts.append(
            f"repair drift {drift_score} exceeds threshold {max_allowed_drift}"
        )
    if slide_count_delta > max_allowed_slide_count_delta:
        reason_parts.append(
            f"slide-count change {slide_count_delta} exceeds threshold {max_allowed_slide_count_delta}"
        )
    if accepted and not reason_parts:
        reason_parts.append("repair made a meaningful improvement without excessive structural drift")

    return {
        "accepted": accepted,
        "score_improvement": score_improvement,
        "required_improvement": required_improvement,
        "draft_avg_target_distance": draft_avg_target_distance,
        "repaired_avg_target_distance": repaired_avg_target_distance,
        "target_distance_improvement": target_distance_improvement,
        "improved_priority_metrics": improved_priority_metrics,
        "worsened_priority_metrics": worsened_priority_metrics,
        "severe_worsened_priority_metrics": severe_worsened_priority_metrics,
        "stable_priority_metrics": stable_priority_metrics,
        "measurable_target_improvements": measurable_target_improvements,
        "missed_target_improvements": missed_target_improvements,
        "measurable_improvement_count": measurable_improvement_count,
        "min_measurable_improvements": min_measurable_improvements,
        "weighted_total_net_gain": round(weighted_total_net_gain, 4),
        "weighted_priority_net_gain": round(weighted_priority_net_gain, 4),
        "weighted_actionable_net_gain": round(weighted_actionable_net_gain, 4),
        "weighted_regression": round(weighted_regression, 4),
        "required_weighted_net_gain": required_weighted_net_gain,
        "max_allowed_weighted_regression": max_allowed_weighted_regression,
        "max_priority_worsen_count": max_priority_worsen_count,
        "drift_score": drift_score,
        "max_allowed_drift": max_allowed_drift,
        "slide_count_delta": slide_count_delta,
        "max_allowed_slide_count_delta": max_allowed_slide_count_delta,
        "reason": "; ".join(reason_parts),
    }


def call_layout_model(
    prompt_text: str,
    system_prompt: str,
    *,
    args: Any,
    cfg: Dict[str, Any],
    use_gpt5_responses: bool,
    client: Any = None,
    agent: Any = None,
) -> tuple[str, int, int, float]:
    start_time = time.time()

    if use_gpt5_responses:
        raw_text, in_tok, out_tok = openai_chat_text(
            client=client,
            model=resolve_direct_model_name(args.model_name_v),
            user_prompt=prompt_text,
            system_prompt=system_prompt,
            prefer_responses=True,
        )
    elif str(args.model_name_v).startswith('vllm_qwen'):
        response = chat_via_vllm(prompt_text, cfg, agent, system_prompt)
        raw_text = response.choices[0].message.content
        in_tok = response.usage.prompt_tokens
        out_tok = response.usage.completion_tokens
    else:
        agent.reset()
        response = agent.step(prompt_text)
        raw_text = response.msgs[0].content
        in_tok, out_tok = account_token(response)

    return raw_text, in_tok, out_tok, time.time() - start_time


def render_planner_prompt(
    *,
    template: Any,
    raw_json: Dict[str, Any],
    figures_json: Dict[str, Any],
    formulas_json: Dict[str, Any],
    images: Dict[str, Any],
    tables: Dict[str, Any],
    use_author_preferences: bool,
    author_preference_profile: Dict[str, Any] | None,
    planner_target_brief: Dict[str, Any] | None,
) -> str:
    return template.render(
        raw_result_json=raw_json,
        figures_json=figures_json,
        formulas_json=formulas_json,
        image_informations_json=images,
        table_informations_json=tables,
        use_author_preferences=use_author_preferences,
        author_preference_profile_json=author_preference_profile,
        planner_target_brief_json=planner_target_brief or {},
        numeric_target_summary_json=(planner_target_brief or {}).get("numeric_target_summary", {}),
        priority_metrics_json=(planner_target_brief or {}).get("priority_metrics", []),
        blocked_priority_metrics_json=(planner_target_brief or {}).get("blocked_priority_metrics", []),
    )


def build_profile_trace(author_preference_profile: Dict[str, Any] | None) -> Dict[str, Any] | None:
    if not author_preference_profile:
        return None
    return {
        "author_id": author_preference_profile.get("author_id"),
        "profile_version": author_preference_profile.get("profile_version"),
        "distilled_from": author_preference_profile.get("distilled_from"),
        "planning_preferences": author_preference_profile.get("planning_preferences"),
        "numeric_preferences": author_preference_profile.get("numeric_preferences"),
    }


def generate_slide_plan(
    args 
) -> Dict[str, Any]: 
    paper_outline_json = raw_content_path(args)
    figures_path = figures_json_path(args)
    
    if args.formula_mode == 1 or args.formula_mode == 2:
        print("👉 Using Docling bbox crop method...") 
        formulas_path = formula_match_path(args)
    elif args.formula_mode == 3:
        print("👉 Using user-marked boxes method...")
        formulas_path = formula_mode3_index_path(args)

    raw_json = json.loads(Path(paper_outline_json).read_text(encoding="utf-8"))
    figures_json = json.loads(Path(figures_path).read_text(encoding="utf-8"))
    formulas_json = json.loads(Path(formulas_path).read_text(encoding="utf-8"))
    images = json.loads(images_filtered_path(args).read_text(encoding="utf-8"))
    tables = json.loads(tables_filtered_path(args).read_text(encoding="utf-8"))
    asset_support = derive_asset_support(
        formulas_json=formulas_json,
        images=images,
        tables=tables,
    )
    author_preference_profile = None
    if getattr(args, "use_author_preferences", False):
        profile_path = Path(getattr(args, "author_profile_path", ""))
        if not profile_path.exists():
            raise FileNotFoundError(f"Author preference profile not found: {profile_path}")
        author_preference_profile = json.loads(profile_path.read_text(encoding="utf-8"))
    planner_target_brief = None
    if getattr(args, "use_author_preferences", False) and author_preference_profile:
        planner_target_brief = build_planner_target_brief(
            author_preference_profile,
            asset_support,
        )
    with open(STRONG_PROMPT_PATH, "r", encoding="utf-8") as f:
        prompt_cfg =  yaml.safe_load(f) 
    use_gpt5_responses = False
    cfg = get_agent_config(args.model_name_v)
    if "gpt-5" in str(args.model_name_v).lower():  
        client = build_openai_client()
        use_gpt5_responses = True
    else:
        if str(args.model_name_v).startswith('vllm_qwen'):
            model = ModelFactory.create(
                model_platform=cfg['model_platform'],
                model_type=cfg['model_type'],
                model_config_dict=cfg['model_config'],
                url=cfg['url'],
            )
        else:
            model = ModelFactory.create(
                model_platform=cfg["model_platform"],
                model_type=cfg["model_type"],
                model_config_dict=cfg["model_config"],
                url=cfg.get("url"),
            )  
        if not str(args.model_name_v).startswith('vllm_qwen'):
            agent = ChatAgent(
                system_message=prompt_cfg['system_prompt'],
                model=model,
                message_window_size=5,
            )
            repair_agent = ChatAgent(
                system_message=prompt_cfg.get('repair_system_prompt', prompt_cfg['system_prompt']),
                model=model,
                message_window_size=5,
            )
        else:
            agent = model
            repair_agent = model

    jinja_env = Environment(undefined=StrictUndefined)
    template =  jinja_env.from_string(prompt_cfg["template"]) 
    in_tok, out_tok, time_taken = 0, 0, 0.0
    slide_plan: Dict[str, Any]
    draft_plan_source = "direct_personalized_planner" if getattr(args, "use_author_preferences", False) and author_preference_profile else "direct_planner"

    planner_prompt = render_planner_prompt(
        template=template,
        raw_json=raw_json,
        figures_json=figures_json,
        formulas_json=formulas_json,
        images=images,
        tables=tables,
        use_author_preferences=getattr(args, "use_author_preferences", False),
        author_preference_profile=author_preference_profile,
        planner_target_brief=planner_target_brief,
    )
    raw_text, fresh_in_tok, fresh_out_tok, fresh_time_taken = call_layout_model(
        planner_prompt,
        prompt_cfg['system_prompt'],
        args=args,
        cfg=cfg,
        use_gpt5_responses=use_gpt5_responses,
        client=client if use_gpt5_responses else None,
        agent=agent if not use_gpt5_responses else None,
    )
    print(f"[layout-agent] tokens: in={fresh_in_tok} out={fresh_out_tok}")
    print("time_taken:",fresh_time_taken)
    slide_plan = get_json_from_response(raw_text)
    slide_plan = sanitize_slide_plan_templates(slide_plan)
    in_tok += fresh_in_tok
    out_tok += fresh_out_tok
    time_taken += fresh_time_taken
    draft_summary = summarize_slide_plan(slide_plan)
    draft_repair_directives: Dict[str, Any] | None = None
    repaired_summary: Dict[str, Any] | None = None
    repaired_repair_directives: Dict[str, Any] | None = None
    acceptance: Dict[str, Any] | None = None
    repair_attempted = False
    repair_report_path: str | None = None

    plan_debug_path = slide_plan_draft_path(args, plan_variant_suffix(args))
    plan_debug_path.parent.mkdir(parents=True, exist_ok=True)
    with open(plan_debug_path, 'w', encoding="utf-8") as f:
        json.dump(slide_plan, f, indent=4)

    if getattr(args, "use_author_preferences", False) and author_preference_profile:
        draft_repair_directives = build_repair_directives(author_preference_profile, draft_summary, asset_support)
        if draft_repair_directives["needs_repair"] and isinstance(draft_summary, dict):
            repair_attempted = True
            repair_template = jinja_env.from_string(prompt_cfg["repair_template"])
            max_repair_rounds = max(1, int(getattr(args, "personalization_repair_rounds", DEFAULT_STRONG_REPAIR_ROUNDS) or DEFAULT_STRONG_REPAIR_ROUNDS))
            repair_round_reports: List[Dict[str, Any]] = []
            current_plan = slide_plan
            current_summary = draft_summary
            current_directives = draft_repair_directives
            current_source = draft_plan_source
            best_plan = slide_plan
            best_summary = draft_summary
            best_directives = draft_repair_directives
            best_score = float(draft_repair_directives.get("total_mismatch_score", 0.0))
            best_round_index = 0

            for repair_round_index in range(1, max_repair_rounds + 1):
                repair_prompt = repair_template.render(
                    raw_result_json=raw_json,
                    figures_json=figures_json,
                    formulas_json=formulas_json,
                    author_preference_profile_json=author_preference_profile,
                    numeric_target_summary_json=current_directives["target_summary"],
                    current_plan_summary_json=current_summary,
                    repair_directives_json=current_directives,
                    priority_metrics_json=current_directives.get("priority_metrics", []),
                    edit_brief_json=current_directives.get("edit_brief", []),
                    actionable_targets_json=current_directives.get("actionable_targets", []),
                    current_slide_plan_json=current_plan,
                    anchor_plan_source=current_source,
                )
                repair_raw_text, repair_in_tok, repair_out_tok, repair_time_taken = call_layout_model(
                    repair_prompt,
                    prompt_cfg.get('repair_system_prompt', prompt_cfg['system_prompt']),
                    args=args,
                    cfg=cfg,
                    use_gpt5_responses=use_gpt5_responses,
                    client=client if use_gpt5_responses else None,
                    agent=repair_agent if not use_gpt5_responses else None,
                )
                print(f"[layout-repair round {repair_round_index}] tokens: in={repair_in_tok} out={repair_out_tok}")
                print("repair_time_taken:", repair_time_taken)
                in_tok += repair_in_tok
                out_tok += repair_out_tok
                time_taken += repair_time_taken
                repaired_plan = get_json_from_response(repair_raw_text)
                repaired_plan = sanitize_slide_plan_templates(repaired_plan)
                round_repaired_summary = summarize_slide_plan(repaired_plan)
                round_repaired_directives = build_repair_directives(author_preference_profile, round_repaired_summary, asset_support)
                round_acceptance = evaluate_repair_acceptance(
                    current_summary,
                    round_repaired_summary,
                    current_directives,
                    round_repaired_directives,
                )
                round_score = float(round_repaired_directives.get("total_mismatch_score", 0.0))
                current_score = float(current_directives.get("total_mismatch_score", 0.0))
                score_delta = current_score - round_score
                soft_accept = (
                    score_delta > 0.1
                    and float(round_acceptance.get("target_distance_improvement", 0.0)) > 0.0
                    and bool(round_acceptance.get("measurable_target_improvements"))
                    and float(round_acceptance.get("weighted_total_net_gain", 0.0)) > 0.0
                    and float(round_acceptance.get("drift_score", 0.0)) <= float(round_acceptance.get("max_allowed_drift", 0.0)) * 1.25
                )
                adopted = bool(round_acceptance.get("accepted")) or soft_accept
                repair_round_reports.append(
                    {
                        "round_index": repair_round_index,
                        "anchor_summary": current_summary,
                        "anchor_repair_directives": current_directives,
                        "repaired_summary": round_repaired_summary,
                        "repaired_repair_directives": round_repaired_directives,
                        "acceptance": round_acceptance,
                        "soft_accept": soft_accept,
                        "adopted": adopted,
                    }
                )
                if round_score < best_score:
                    best_plan = repaired_plan
                    best_summary = round_repaired_summary
                    best_directives = round_repaired_directives
                    best_score = round_score
                    best_round_index = repair_round_index
                    repaired_summary = round_repaired_summary
                    repaired_repair_directives = round_repaired_directives
                    acceptance = round_acceptance
                if not adopted:
                    break
                current_plan = repaired_plan
                current_summary = round_repaired_summary
                current_directives = round_repaired_directives
                current_source = f"repair_round_{repair_round_index}"
                if not current_directives.get("needs_repair"):
                    break

            repair_report_path = slide_plan_repair_report_path(args, plan_variant_suffix(args))
            repair_report = {
                "anchor_plan_source": draft_plan_source,
                "draft_summary": draft_summary,
                "draft_repair_directives": draft_repair_directives,
                "repair_rounds": repair_round_reports,
                "repaired_summary": best_summary,
                "repaired_repair_directives": best_directives,
                "acceptance": acceptance,
                "accepted_repair": bool(acceptance and acceptance.get("accepted")),
                "best_round_index": best_round_index,
            }
            with open(repair_report_path, 'w', encoding="utf-8") as f:
                json.dump(repair_report, f, indent=2)
            if best_round_index > 0 and best_score < float(draft_repair_directives.get("total_mismatch_score", 0.0)):
                best_avg_target_distance = average_target_distance(best_summary, best_directives.get("target_summary"))
                draft_avg_target_distance = average_target_distance(draft_summary, draft_repair_directives.get("target_summary"))
                if best_avg_target_distance < draft_avg_target_distance - 0.02:
                    slide_plan = best_plan
                    repaired_summary = best_summary
                    repaired_repair_directives = best_directives
                    final_round_acceptance = acceptance or {}
                    if not final_round_acceptance.get("accepted"):
                        final_round_acceptance = dict(final_round_acceptance)
                        final_round_acceptance["accepted"] = True
                        final_round_acceptance["reason"] = (
                            str(final_round_acceptance.get("reason", "")).strip() + "; accepted best iterative repair under strong mode after numeric target-distance improvement"
                        ).strip("; ")
                        acceptance = final_round_acceptance

    final_slide_plan_path = slide_plan_path(args, plan_variant_suffix(args))
    final_slide_plan_path.parent.mkdir(parents=True, exist_ok=True)
    with open(final_slide_plan_path, 'w', encoding="utf-8") as f:
        json.dump(slide_plan, f, indent=4)
    final_summary = summarize_slide_plan(slide_plan)
    trace_path = personalization_trace_path(args, plan_variant_suffix(args))
    final_plan_source = draft_plan_source
    if acceptance and acceptance.get("accepted"):
        final_plan_source = "accepted_iterative_repair"
    trace_payload = {
        "paper_name": args.paper_name,
        "model_name_t": args.model_name_t,
        "model_name_v": args.model_name_v,
        "plan_variant_suffix": plan_variant_suffix(args),
        "use_author_preferences": bool(getattr(args, "use_author_preferences", False)),
        "author_profile_path": getattr(args, "author_profile_path", None),
        "author_profile_summary": build_profile_trace(author_preference_profile),
        "asset_support": asset_support,
        "paths": {
            "paper_outline_json": str(paper_outline_json),
            "figures_json": str(figures_path),
            "formulas_json": str(formulas_path),
            "images_json": str(images_filtered_path(args)),
            "tables_json": str(tables_filtered_path(args)),
            "draft_plan_json": str(plan_debug_path),
            "repair_report_json": str(repair_report_path) if repair_report_path else None,
            "final_plan_json": str(final_slide_plan_path),
        },
        "planner": {
            "source": draft_plan_source,
            "profile_injected": bool(author_preference_profile),
            "target_brief": planner_target_brief,
            "draft_summary": draft_summary,
            "draft_repair_directives": draft_repair_directives,
        },
        "repair": {
            "attempted": repair_attempted,
            "accepted": bool(acceptance and acceptance.get("accepted")),
            "acceptance": acceptance,
            "repaired_summary": repaired_summary,
            "repaired_repair_directives": repaired_repair_directives,
        },
        "final": {
            "selected_plan_source": final_plan_source,
            "final_summary": final_summary,
        },
    }
    with open(trace_path, 'w', encoding="utf-8") as f:
        json.dump(trace_payload, f, indent=2)
    print("slide_plan")
    print(slide_plan)
    return in_tok, out_tok,time_taken 
  
if __name__ == "__main__":  # pragma: no cover — keeps CLI convenience
    import argparse
    p = argparse.ArgumentParser(description="Generate slide-layout plan JSON via LLM.")
    p.add_argument("--raw", required=True, help="Path to raw_result.json")
    p.add_argument("--figures", required=True, help="Path to figures.json")
    p.add_argument("--formulas", required=True, help="Path to formula_index.json")
    p.add_argument("--prompt", default="prompt.yaml", help="Prompt YAML path")
    p.add_argument("--output", default="slide_plan.json", help="Where to save plan JSON")
    p.add_argument("--model_name_v", default="gpt-4o-mini", help="Model identifier")
    args = p.parse_args()

    plan = generate_slide_plan_from_files(
        raw_path=args.raw,
        figures_path=args.figures,
        formulas_path=args.formulas,
        prompt_path=args.prompt,
        model_name_v=args.model_name_v,
    )

    Path(args.output).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f" Saved {len(plan['slides'])}-slide plan → {args.output}")
