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
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name
from camel.models import ModelFactory          
from camel.agents import ChatAgent     
from pptx.util import Cm, Pt
import time


def plan_variant_suffix(args) -> str:
    return getattr(
        args,
        "output_variant_suffix",
        "_personalized" if getattr(args, "use_author_preferences", False) else "_baseline",
    )


def anchor_variant_suffix(args) -> str:
    return getattr(args, "anchor_variant_suffix", "_anchor")


def neutral_paper_name(paper_name: str) -> str:
    base = str(paper_name or "").strip()
    if base.endswith("_personalized"):
        return base[: -len("_personalized")]
    return base


def slide_plan_path_for(paper_name: str, model_name_t: str, model_name_v: str, variant_suffix: str) -> Path:
    return Path(
        f"contents/{paper_name}/"
        f"<{model_name_t}_{model_name_v}>_slide_plan{variant_suffix}.json"
    )


def load_existing_anchor_plan(args: Any) -> tuple[Dict[str, Any] | None, str | None]:
    local_anchor_path = slide_plan_path_for(
        args.paper_name,
        args.model_name_t,
        args.model_name_v,
        anchor_variant_suffix(args),
    )
    candidate_paths: List[Path] = [local_anchor_path]

    base_paper_name = neutral_paper_name(args.paper_name)
    if base_paper_name and base_paper_name != args.paper_name:
        candidate_paths.append(
            slide_plan_path_for(
                base_paper_name,
                args.model_name_t,
                args.model_name_v,
                "_baseline",
            )
        )

    for candidate_path in candidate_paths:
        if not candidate_path.exists():
            continue
        try:
            plan = json.loads(candidate_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(plan, dict) and isinstance(plan.get("slides"), list) and plan["slides"]:
            return plan, str(candidate_path)
    return None, None


def mean(values: List[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def classify_level(value: float, low: float, high: float) -> str:
    if value < low:
        return "low"
    if value < high:
        return "medium"
    return "high"


def classify_structure(value: float) -> str:
    if value < 1.5:
        return "coarse"
    if value > 2.5:
        return "fine_grained"
    return "balanced"


def summarize_slide_plan(plan: Dict[str, Any]) -> Dict[str, Any]:
    slides = list(plan.get("slides") or [])
    slide_count = len(slides)

    section_order: List[str] = []
    section_counts: Dict[str, int] = {}
    bullets_per_slide: List[int] = []
    words_per_slide: List[int] = []
    figure_flags: List[int] = []
    table_flags: List[int] = []
    formula_flags: List[int] = []
    layout_bias_counts = {
        "text_only": 0,
        "image_right": 0,
        "image_left": 0,
        "image_top": 0,
        "multi_visual": 0,
        "formula_capable": 0,
    }

    for slide in slides:
        section = str(slide.get("section") or "").strip() or "UNKNOWN"
        if section not in section_counts:
            section_order.append(section)
            section_counts[section] = 0
        section_counts[section] += 1

        bullets = list(slide.get("bullets") or [])
        bullets_per_slide.append(len(bullets))

        words = 0
        for bullet in bullets:
            if not isinstance(bullet, dict):
                continue
            words += len(str(bullet.get("text") or "").split())
            for sub in bullet.get("sub") or []:
                words += len(str(sub).split())
        words_per_slide.append(words)

        image_count = len(slide.get("images") or [])
        table_count = len(slide.get("tables") or [])
        formula_count = len(slide.get("formulas") or [])
        figure_flags.append(1 if image_count > 0 else 0)
        table_flags.append(1 if table_count > 0 else 0)
        formula_flags.append(1 if formula_count > 0 else 0)

        template_id = str(slide.get("template_id") or "")
        if template_id == "T1_TextOnly":
            layout_bias_counts["text_only"] += 1
        if "ImageRight" in template_id:
            layout_bias_counts["image_right"] += 1
        if "ImageLeft" in template_id:
            layout_bias_counts["image_left"] += 1
        if "ImageTop" in template_id:
            layout_bias_counts["image_top"] += 1
        if image_count + table_count + formula_count >= 2 or "TwoImages" in template_id or "2x2" in template_id or "3Img" in template_id:
            layout_bias_counts["multi_visual"] += 1
        if formula_count > 0:
            layout_bias_counts["formula_capable"] += 1

    section_count = len(section_order)
    avg_slides_per_section = (slide_count / section_count) if section_count else 0.0

    def fraction(count: int) -> float:
        return round((count / slide_count), 4) if slide_count else 0.0

    return {
        "slide_count": slide_count,
        "section_count": section_count,
        "section_titles": section_order,
        "section_slide_counts": section_counts,
        "avg_slides_per_section": round(avg_slides_per_section, 3),
        "section_splitting_estimate": classify_structure(avg_slides_per_section) if section_count else "unknown",
        "avg_bullets_per_slide": round(mean([float(v) for v in bullets_per_slide]), 3),
        "avg_words_per_slide": round(mean([float(v) for v in words_per_slide]), 3),
        "bullet_density_estimate": classify_level(mean([float(v) for v in bullets_per_slide]), 2.0, 4.0) if slides else "unknown",
        "text_density_estimate": classify_level(mean([float(v) for v in words_per_slide]), 18.0, 38.0) if slides else "unknown",
        "figure_slide_fraction": round(sum(figure_flags) / slide_count, 4) if slide_count else 0.0,
        "table_slide_fraction": round(sum(table_flags) / slide_count, 4) if slide_count else 0.0,
        "formula_slide_fraction": round(sum(formula_flags) / slide_count, 4) if slide_count else 0.0,
        "text_only_fraction": fraction(layout_bias_counts["text_only"]),
        "multi_visual_fraction": fraction(layout_bias_counts["multi_visual"]),
        "formula_capable_fraction": fraction(layout_bias_counts["formula_capable"]),
        "image_right_fraction": fraction(layout_bias_counts["image_right"]),
        "image_left_fraction": fraction(layout_bias_counts["image_left"]),
        "image_top_fraction": fraction(layout_bias_counts["image_top"]),
        "layout_bias_counts": layout_bias_counts,
    }


def build_numeric_target_summary(author_preference_profile: Dict[str, Any] | None) -> Dict[str, Any]:
    if not author_preference_profile:
        return {}
    numeric = dict(author_preference_profile.get("numeric_preferences") or {})
    return {key: value for key, value in numeric.items() if value not in (None, [], {}, "")}


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


def is_actionable_mismatch(metric_key: str, asset_support: Dict[str, Any]) -> bool:
    if metric_key == "slide_count":
        return False
    if metric_key in {"target_fraction_formula_slides", "target_fraction_formula_capable_slides"}:
        return asset_support.get("supports_formulas", False)
    if metric_key == "target_fraction_table_slides":
        return asset_support.get("supports_tables", False)
    if metric_key in {
        "target_fraction_figure_slides",
        "target_fraction_multi_visual_slides",
        "target_fraction_image_top_slides",
        "target_fraction_image_right_slides",
        "target_fraction_image_left_slides",
    }:
        return asset_support.get("supports_visual_layout_changes", False)
    return True


def select_actionable_targets(
    mismatches: List[Dict[str, Any]],
    asset_support: Dict[str, Any],
    *,
    max_targets: int | None = None,
) -> List[Dict[str, Any]]:
    selected: List[Dict[str, Any]] = []
    for mismatch in mismatches:
        metric_key = str(mismatch.get("metric") or "")
        if not is_actionable_mismatch(metric_key, asset_support):
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

    numeric = build_numeric_target_summary(author_preference_profile)
    priority_metric_keys = infer_priority_metric_keys(author_preference_profile, numeric)
    asset_support = asset_support or {}
    mismatches: List[Dict[str, Any]] = []
    goals: List[str] = []
    metric_goals: Dict[str, str] = {}
    metric_mismatch_scores: Dict[str, float] = {}
    total_mismatch_score = 0.0

    def record_range_mismatch(key: str, observed: float, allowed_range: List[Any], *, goal_low: str, goal_high: str) -> None:
        nonlocal total_mismatch_score
        if not isinstance(allowed_range, list) or len(allowed_range) != 2:
            return
        try:
            low = float(allowed_range[0])
            high = float(allowed_range[1])
        except Exception:
            return
        if observed < low:
            delta = low - observed
            total_mismatch_score += delta
            metric_mismatch_scores[key] = round(delta, 4)
            metric_goals[key] = goal_low
            mismatches.append({
                "metric": key,
                "observed": observed,
                "target": [low, high],
                "delta": round(delta, 4),
                "priority": key in priority_metric_keys,
                "recommended_edit": goal_low,
            })
            goals.append(goal_low)
        elif observed > high:
            delta = observed - high
            total_mismatch_score += delta
            metric_mismatch_scores[key] = round(delta, 4)
            metric_goals[key] = goal_high
            mismatches.append({
                "metric": key,
                "observed": observed,
                "target": [low, high],
                "delta": round(delta, 4),
                "priority": key in priority_metric_keys,
                "recommended_edit": goal_high,
            })
            goals.append(goal_high)

    def record_scalar_mismatch(key: str, observed: float, tolerance: float, goal_low: str, goal_high: str) -> None:
        nonlocal total_mismatch_score
        if key not in numeric:
            return
        try:
            target = float(numeric[key])
        except Exception:
            return
        delta = observed - target
        if abs(delta) <= tolerance:
            return
        score = abs(delta) / max(tolerance, 1e-6)
        total_mismatch_score += score
        metric_mismatch_scores[key] = round(score, 4)
        chosen_goal = goal_low if delta < 0 else goal_high
        metric_goals[key] = chosen_goal
        mismatches.append({
            "metric": key,
            "observed": observed,
            "target": target,
            "delta": round(delta, 4),
            "priority": key in priority_metric_keys,
            "recommended_edit": chosen_goal,
        })
        goals.append(chosen_goal)

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

    mismatches.sort(key=lambda item: (not item.get("priority", False), -abs(float(item.get("delta", 0.0)))))
    actionable_targets = select_actionable_targets(mismatches, asset_support)

    deduped_goals: List[str] = []
    for goal in goals:
        if goal not in deduped_goals:
            deduped_goals.append(goal)

    edit_brief = [item["recommended_edit"] for item in actionable_targets]

    return {
        "needs_repair": bool(actionable_targets),
        "target_summary": numeric,
        "goals": deduped_goals[:8],
        "mismatches": mismatches,
        "total_mismatch_score": round(total_mismatch_score, 4),
        "metric_mismatch_scores": metric_mismatch_scores,
        "priority_metrics": priority_metric_keys,
        "edit_brief": edit_brief,
        "actionable_targets": actionable_targets,
    }


def evaluate_repair_acceptance(
    draft_summary: Dict[str, Any],
    repaired_summary: Dict[str, Any],
    draft_directives: Dict[str, Any],
    repaired_directives: Dict[str, Any],
) -> Dict[str, Any]:
    draft_score = float(draft_directives.get("total_mismatch_score", 0.0))
    repaired_score = float(repaired_directives.get("total_mismatch_score", 0.0))
    score_improvement = round(draft_score - repaired_score, 4)
    required_improvement = round(max(0.35, draft_score * 0.08), 4) if draft_score > 0 else 0.0

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
        else:
            missed_target_improvements.append(metric_key)

    drift_score = compute_anchor_drift_score(draft_summary, repaired_summary)
    max_allowed_drift = round(max(1.75, draft_score * 0.25), 4) if draft_score > 0 else 1.75

    # Accept any repair that reduces total profile mismatch. Drift and
    # per-metric regressions are still reported for inspection, but they no
    # longer block adoption because personalization alignment is the primary
    # objective.
    accepted = score_improvement > 0.0

    reason_parts: List[str] = []
    if score_improvement <= 0.0:
        reason_parts.append(
            f"repair did not improve total profile mismatch (delta={score_improvement})"
        )
    if worsened_priority_metrics:
        reason_parts.append(
            "diagnostic only: worsened priority metrics: " + ", ".join(worsened_priority_metrics)
        )
    if not measurable_target_improvements:
        reason_parts.append(
            "diagnostic only: no actionable target improved by its minimum delta"
        )
    if drift_score > max_allowed_drift:
        reason_parts.append(
            f"diagnostic only: anchor drift {drift_score} exceeds prior threshold {max_allowed_drift}"
        )
    if accepted and not reason_parts:
        reason_parts.append("repair reduced total profile mismatch")

    return {
        "accepted": accepted,
        "score_improvement": score_improvement,
        "required_improvement": required_improvement,
        "improved_priority_metrics": improved_priority_metrics,
        "worsened_priority_metrics": worsened_priority_metrics,
        "stable_priority_metrics": stable_priority_metrics,
        "measurable_target_improvements": measurable_target_improvements,
        "missed_target_improvements": missed_target_improvements,
        "drift_score": drift_score,
        "max_allowed_drift": max_allowed_drift,
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
    use_pair_guidelines: bool,
    pair_guidelines: Dict[str, Any] | None,
) -> str:
    return template.render(
        raw_result_json=raw_json,
        figures_json=figures_json,
        formulas_json=formulas_json,
        image_informations_json=images,
        table_informations_json=tables,
        use_author_preferences=use_author_preferences,
        author_preference_profile_json=author_preference_profile,
        use_pair_guidelines=use_pair_guidelines,
        pair_guidelines_json=pair_guidelines,
    )


def generate_slide_plan(
    args 
) -> Dict[str, Any]: 
    paper_outline_json = f'contents/{args.paper_name}/<{args.model_name_t}_{args.model_name_v}>_raw_content.json' 
    figures_path=f'contents/{args.paper_name}/<{args.model_name_t}_{args.model_name_v}>_figures.json'
    
    if args.formula_mode == 1 or args.formula_mode == 2:
        print("👉 Using Docling bbox crop method...") 
        formulas_path=f'contents/{args.paper_name}/<{args.model_name_t}_{args.model_name_v}>_formula_match.json'
    elif args.formula_mode == 3:
        print("👉 Using user-marked boxes method...")
        formulas_path=f'contents/{args.paper_name}/formula_index_formula_mode3.json'

    raw_json = json.loads(Path(paper_outline_json).read_text(encoding="utf-8"))
    figures_json = json.loads(Path(figures_path).read_text(encoding="utf-8"))
    formulas_json = json.loads(Path(formulas_path).read_text(encoding="utf-8"))
    images = json.loads(Path(f'<{args.model_name_t}_{args.model_name_v}>_images_and_tables/{args.paper_name}/images_filtered.json').read_text(encoding="utf-8"))
    tables = json.loads(Path(f'<{args.model_name_t}_{args.model_name_v}>_images_and_tables/{args.paper_name}/tables_filtered.json' ).read_text(encoding="utf-8"))
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
    pair_guideline_context = None
    if getattr(args, "use_pair_guidelines", False):
        pair_guidelines_path = Path(getattr(args, "pair_guidelines_path", ""))
        if not pair_guidelines_path.exists():
            raise FileNotFoundError(f"Pair-guideline context not found: {pair_guidelines_path}")
        pair_guideline_context = json.loads(pair_guidelines_path.read_text(encoding="utf-8"))
    with open(f'utils/prompt_templates/layout_agent_xin.yaml', "r", encoding="utf-8") as f:
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
    anchor_plan_source: str | None = None

    if getattr(args, "use_author_preferences", False) and author_preference_profile:
        slide_plan, anchor_plan_source = load_existing_anchor_plan(args)
        if slide_plan is None:
            anchor_prompt = render_planner_prompt(
                template=template,
                raw_json=raw_json,
                figures_json=figures_json,
                formulas_json=formulas_json,
                images=images,
                tables=tables,
                use_author_preferences=False,
                author_preference_profile=None,
                use_pair_guidelines=False,
                pair_guidelines=None,
            )
            anchor_raw_text, anchor_in_tok, anchor_out_tok, anchor_time_taken = call_layout_model(
                anchor_prompt,
                prompt_cfg['system_prompt'],
                args=args,
                cfg=cfg,
                use_gpt5_responses=use_gpt5_responses,
                client=client if use_gpt5_responses else None,
                agent=agent if not use_gpt5_responses else None,
            )
            print(f"[layout-anchor] tokens: in={anchor_in_tok} out={anchor_out_tok}")
            print("anchor_time_taken:", anchor_time_taken)
            slide_plan = get_json_from_response(anchor_raw_text)
            in_tok += anchor_in_tok
            out_tok += anchor_out_tok
            time_taken += anchor_time_taken
            anchor_plan_source = "generated_internal_anchor"
            anchor_plan_path = slide_plan_path_for(
                args.paper_name,
                args.model_name_t,
                args.model_name_v,
                anchor_variant_suffix(args),
            )
            anchor_plan_path.parent.mkdir(parents=True, exist_ok=True)
            with open(anchor_plan_path, 'w', encoding="utf-8") as f:
                json.dump(slide_plan, f, indent=4)
        else:
            print(f"[layout-anchor] reusing anchor plan from {anchor_plan_source}")
    else:
        planner_prompt = render_planner_prompt(
            template=template,
            raw_json=raw_json,
            figures_json=figures_json,
            formulas_json=formulas_json,
            images=images,
            tables=tables,
            use_author_preferences=getattr(args, "use_author_preferences", False),
            author_preference_profile=author_preference_profile,
            use_pair_guidelines=getattr(args, "use_pair_guidelines", False),
            pair_guidelines=pair_guideline_context,
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
        in_tok += fresh_in_tok
        out_tok += fresh_out_tok
        time_taken += fresh_time_taken

    plan_debug_path = (
        f'contents/{args.paper_name}/'
        f'<{args.model_name_t}_{args.model_name_v}>_slide_plan_draft{plan_variant_suffix(args)}.json'
    )
    with open(plan_debug_path, 'w', encoding="utf-8") as f:
        json.dump(slide_plan, f, indent=4)

    if getattr(args, "use_author_preferences", False) and author_preference_profile:
        draft_summary = summarize_slide_plan(slide_plan)
        repair_directives = build_repair_directives(author_preference_profile, draft_summary, asset_support)
        if repair_directives["needs_repair"]:
            repair_template = jinja_env.from_string(prompt_cfg["repair_template"])
            repair_prompt = repair_template.render(
                raw_result_json=raw_json,
                figures_json=figures_json,
                formulas_json=formulas_json,
                author_preference_profile_json=author_preference_profile,
                numeric_target_summary_json=repair_directives["target_summary"],
                current_plan_summary_json=draft_summary,
                repair_directives_json=repair_directives,
                priority_metrics_json=repair_directives.get("priority_metrics", []),
                edit_brief_json=repair_directives.get("edit_brief", []),
                actionable_targets_json=repair_directives.get("actionable_targets", []),
                current_slide_plan_json=slide_plan,
                anchor_plan_source=anchor_plan_source or "generated_internal_anchor",
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
            print(f"[layout-repair] tokens: in={repair_in_tok} out={repair_out_tok}")
            print("repair_time_taken:", repair_time_taken)
            repaired_plan = get_json_from_response(repair_raw_text)
            repaired_summary = summarize_slide_plan(repaired_plan)
            repaired_directives = build_repair_directives(author_preference_profile, repaired_summary, asset_support)
            acceptance = evaluate_repair_acceptance(
                draft_summary,
                repaired_summary,
                repair_directives,
                repaired_directives,
            )
            repair_report_path = (
                f'contents/{args.paper_name}/'
                f'<{args.model_name_t}_{args.model_name_v}>_slide_plan_repair_report{plan_variant_suffix(args)}.json'
            )
            repair_report = {
                "anchor_plan_source": anchor_plan_source or "generated_internal_anchor",
                "draft_summary": draft_summary,
                "draft_repair_directives": repair_directives,
                "repaired_summary": repaired_summary,
                "repaired_repair_directives": repaired_directives,
                "acceptance": acceptance,
                "accepted_repair": acceptance["accepted"],
            }
            with open(repair_report_path, 'w', encoding="utf-8") as f:
                json.dump(repair_report, f, indent=2)
            if acceptance["accepted"]:
                slide_plan = repaired_plan
                in_tok += repair_in_tok
                out_tok += repair_out_tok
                time_taken += repair_time_taken

    slide_plan_path = (
        f'contents/{args.paper_name}/'
        f'<{args.model_name_t}_{args.model_name_v}>_slide_plan{plan_variant_suffix(args)}.json'
    )
    with open(slide_plan_path, 'w', encoding="utf-8") as f:
        json.dump(slide_plan, f, indent=4)
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
