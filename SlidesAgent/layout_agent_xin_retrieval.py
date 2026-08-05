from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any, Dict, List

import yaml
from jinja2 import Environment, StrictUndefined

from SlidesAgent.layout_agent_xin import (
    FORMULA_TEMPLATE_IDS,
    build_profile_trace,
    call_layout_model,
    derive_asset_support,
    plan_variant_suffix,
    sanitize_slide_plan_templates,
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
from slidegen_openai_utils import build_openai_client
from utils.src.utils import get_json_from_response
from utils.wei_utils import account_token, chat_via_vllm, get_agent_config
from camel.models import ModelFactory
from camel.agents import ChatAgent


BASE_PROMPT_PATH = Path("utils/prompt_templates/layout_agent_xin_retrieval.yaml")
MAX_RETRIEVAL_REPAIR_ROUNDS = 2

RETRIEVAL_TARGET_KEYS = (
    "target_slide_count",
    "target_avg_bullets_per_slide",
    "target_avg_words_per_slide",
    "target_image_slide_count",
    "target_table_slide_count",
    "target_formula_slide_count",
)

COUNT_TARGET_TO_OBSERVED_KEY = {
    "target_image_slide_count": "image_slide_count",
    "target_table_slide_count": "table_slide_count",
    "target_formula_slide_count": "formula_slide_count",
}


def sanitize_retrieval_profile_for_generation(profile: Dict[str, Any] | None) -> Dict[str, Any] | None:
    if not isinstance(profile, dict):
        return None

    planning = dict(profile.get("planning_preferences") or {})
    numeric = dict(profile.get("numeric_preferences") or {})
    fonts = dict(profile.get("font_preferences") or {})
    colors = dict(profile.get("color_preferences") or {})

    sanitized_planning: Dict[str, Any] = {}
    if planning.get("target_section_count") is not None:
        try:
            sanitized_planning["target_section_count"] = round(float(planning["target_section_count"]), 4)
        except Exception:
            pass
    labels: List[str] = []
    for raw_label in planning.get("preferred_section_labels") or []:
        label = str(raw_label or "").strip().lower()
        if label and label not in labels:
            labels.append(label)
    if labels:
        sanitized_planning["preferred_section_labels"] = labels
    order_style = str(planning.get("section_order_style") or "").strip().lower()
    if order_style in {"canonical", "custom", "mixed"}:
        sanitized_planning["section_order_style"] = order_style

    sanitized_numeric: Dict[str, Any] = {}
    for key in RETRIEVAL_TARGET_KEYS:
        value = numeric.get(key)
        if value is None:
            continue
        try:
            sanitized_numeric[key] = round(float(value), 4)
        except Exception:
            continue

    sanitized_fonts: Dict[str, Any] = {}
    for key in ("title_font_name", "body_font_name"):
        value = str(fonts.get(key) or "").strip()
        if value:
            sanitized_fonts[key] = value

    sanitized_colors: Dict[str, Any] = {}
    for key in (
        "target_theme_hex",
        "target_base_hex",
        "color_source_paper_id",
        "color_source_paper_title",
        "color_source_raw_dir",
        "color_sample_slide_count",
    ):
        value = colors.get(key)
        if value is not None and value != "":
            sanitized_colors[key] = value

    return {
        "author_id": profile.get("author_id"),
        "profile_version": profile.get("profile_version"),
        "profile_method": profile.get("profile_method"),
        "distilled_from": profile.get("distilled_from") or {},
        "planning_preferences": sanitized_planning,
        "numeric_preferences": sanitized_numeric,
        "font_preferences": sanitized_fonts,
        "color_preferences": sanitized_colors,
        "retrieval_context": profile.get("retrieval_context") or {},
    }


def build_retrieval_section_preferences(profile: Dict[str, Any] | None) -> Dict[str, Any]:
    if not isinstance(profile, dict):
        return {}
    profile = sanitize_retrieval_profile_for_generation(profile) or {}
    planning = dict(profile.get("planning_preferences") or {})
    section_prefs: Dict[str, Any] = {}
    target_section_count = planning.get("target_section_count")
    if target_section_count is not None:
        try:
            section_prefs["target_section_count"] = round(float(target_section_count), 4)
        except Exception:
            pass
    labels = []
    for raw_label in planning.get("preferred_section_labels") or []:
        label = str(raw_label or "").strip().lower()
        if label and label not in labels:
            labels.append(label)
    if labels:
        section_prefs["preferred_section_labels"] = labels
    order_style = str(planning.get("section_order_style") or "").strip().lower()
    if order_style in {"canonical", "custom", "mixed"}:
        section_prefs["section_order_style"] = order_style
    return section_prefs

def summarize_retrieval_targets(
    plan: Dict[str, Any],
) -> Dict[str, Any]:
    base_summary = summarize_slide_plan(plan)
    slides = list(plan.get("slides") or [])
    image_slide_count = 0
    table_slide_count = 0
    formula_slide_count = 0
    for slide in slides:
        if not isinstance(slide, dict):
            continue
        if slide.get("images"):
            image_slide_count += 1
        if slide.get("tables"):
            table_slide_count += 1
        if slide.get("formulas"):
            formula_slide_count += 1
    retrieval_summary = dict(base_summary)
    retrieval_summary.update(
        {
            "fraction_slides_with_images": round(float(base_summary.get("figure_slide_fraction", 0.0)), 4),
            "fraction_slides_with_tables": round(float(base_summary.get("table_slide_fraction", 0.0)), 4),
            "fraction_slides_with_formulas": round(float(base_summary.get("formula_slide_fraction", 0.0)), 4),
            "image_slide_count": image_slide_count,
            "table_slide_count": table_slide_count,
            "formula_slide_count": formula_slide_count,
        }
    )
    return retrieval_summary


def build_retrieval_target_summary(profile: Dict[str, Any] | None) -> Dict[str, Any]:
    if not isinstance(profile, dict):
        return {}
    profile = sanitize_retrieval_profile_for_generation(profile) or {}
    numeric = dict(profile.get("numeric_preferences") or {})
    target_summary: Dict[str, Any] = {}
    for key in RETRIEVAL_TARGET_KEYS:
        value = numeric.get(key)
        if value is None:
            continue
        try:
            target_summary[key] = round(float(value), 4)
        except Exception:
            continue
    return target_summary


def _target_prompt_label(target_key: str) -> str:
    labels = {
        "target_slide_count": "total slide count",
        "target_avg_bullets_per_slide": "average bullets per slide",
        "target_avg_words_per_slide": "average words per slide",
        "target_image_slide_count": "number of slides containing images",
        "target_table_slide_count": "number of slides containing tables",
        "target_formula_slide_count": "number of slides containing formulas",
    }
    return labels.get(target_key, target_key)


def build_retrieval_budget(
    target_summary: Dict[str, Any],
    asset_support: Dict[str, Any] | None = None,
    observed_summary: Dict[str, Any] | None = None,
) -> Dict[str, Any]:
    asset_support = asset_support or {}
    observed_summary = observed_summary or {}
    budget_items: List[Dict[str, Any]] = []
    for target_key, observed_key in COUNT_TARGET_TO_OBSERVED_KEY.items():
        if target_key not in target_summary:
            continue
        rounded_target = max(0, int(round(float(target_summary[target_key]))))
        support_key = {
            "target_image_slide_count": "supports_figures",
            "target_table_slide_count": "supports_tables",
            "target_formula_slide_count": "supports_formulas",
        }[target_key]
        supported = bool(asset_support.get(support_key, True))
        observed_value = observed_summary.get(observed_key)
        item = {
            "metric": target_key,
            "label": _target_prompt_label(target_key),
            "target_count": rounded_target,
            "observed_count": None if observed_value is None else int(round(float(observed_value))),
            "support_available": supported,
            "status": "supported" if supported else "unsupported",
        }
        budget_items.append(item)
    return {
        "count_targets_present": bool(budget_items),
        "budget_items": budget_items,
    }


def build_retrieval_planner_brief(
    *,
    target_summary: Dict[str, Any],
    budget_summary: Dict[str, Any] | None = None,
    section_preferences: Dict[str, Any] | None = None,
    asset_support: Dict[str, Any] | None = None,
    observed_summary: Dict[str, Any] | None = None,
) -> str:
    budget_summary = budget_summary or {"budget_items": []}
    section_preferences = section_preferences or {}
    asset_support = asset_support or {}
    lines = [
        "Retrieval-profile personalization targets:",
        "Use these targets directly during whole-deck planning and refinement.",
        "If multiple faithful slide plans are possible, prefer the one that gets closer to these numeric targets.",
        "Treat count targets as a whole-deck visual budget.",
        "Target definitions:",
        "- target_slide_count: total number of slides in the deck.",
        "- target_avg_bullets_per_slide: average number of top-level bullets per slide over the full deck.",
        "- target_avg_words_per_slide: average words per slide over the full deck.",
        "- target_image_slide_count: number of slides containing one or more image assets.",
        "- target_table_slide_count: number of slides containing tables.",
        "- target_formula_slide_count: number of slides containing formulas.",
        "- A slide may count toward multiple categories.",
    ]
    if section_preferences:
        lines.append("Section-structure preferences:")
        if section_preferences.get("target_section_count") is not None:
            lines.append(f"- typical number of major sections = {section_preferences['target_section_count']}")
        if section_preferences.get("preferred_section_labels"):
            lines.append(
                "- preferred major section labels = "
                + ", ".join(str(label) for label in section_preferences["preferred_section_labels"])
            )
        if section_preferences.get("section_order_style"):
            lines.append(f"- typical section order style = {section_preferences['section_order_style']}")
    if target_summary:
        lines.append("Numeric targets:")
        for key, value in target_summary.items():
            lines.append(f"- {_target_prompt_label(key)} = {value}")
    if budget_summary.get("budget_items"):
        lines.append("Whole-deck modality budget:")
        for item in budget_summary["budget_items"]:
            lines.append(
                f"- {item['label']}: target exactly {item['target_count']} slide(s)"
                + ("" if item.get("support_available", True) else " if supported by available assets; support appears unavailable")
            )
    if observed_summary:
        lines.append("Current observed retrieval-target metrics:")
        lines.append(f"- total slide count = {observed_summary.get('slide_count')}")
        lines.append(f"- average bullets per slide = {observed_summary.get('avg_bullets_per_slide')}")
        lines.append(f"- average words per slide = {observed_summary.get('avg_words_per_slide')}")
        lines.append(f"- image slide count = {observed_summary.get('image_slide_count')}")
        lines.append(f"- table slide count = {observed_summary.get('table_slide_count')}")
        lines.append(f"- formula slide count = {observed_summary.get('formula_slide_count')}")
    lines.extend(
        [
            "Optimization guidance:",
            "- Before drafting the final JSON, decide which specific slides will spend the image/table/formula budget.",
            "- Use images/tables/formulas only when supported by the paper assets.",
            "- Try to hit the exact count targets whenever the paper assets and content support it.",
            "- If exact counts are impossible, get as close as possible and spend scarce assets on the highest-value supported slides first.",
            "- Use target_slide_count as a whole-deck pacing budget: merge, split, add, or remove slides only when faithful and helpful for target fit.",
            "- For bullet density, control the number of bullet points carried by each slide.",
            "- If average bullets per slide is too high, trim, merge, or simplify bullet-heavy slides when faithful.",
            "- If average bullets per slide is too low, add faithful bullet structure or split prose into concise bullets where appropriate.",
            "- For text density, control the amount of bullet text and explanatory detail per slide.",
            "- If average words per slide is too high, condense, trim, or split dense slides when faithful.",
            "- If average words per slide is too low, merge thin slides or add faithful explanatory detail where needed.",
            "- A single slide may satisfy multiple count budgets at once when that slide legitimately contains multiple modalities.",
            "- Prefer minimum faithful edits that move the deck measurably closer to target.",
            "- Treat section preferences as early outline constraints, not cosmetic rename requests.",
            "- When the source paper supports it, try to include the preferred section types in the deck's major narrative structure.",
            "- Match the typical number of major sections when faithful, but do not invent unsupported sections.",
            "- Use preferred section labels as soft priors for organizing slides such as motivation, method, experiments, ablation, analysis, limitations, and conclusion.",
        ]
    )
    if asset_support:
        lines.append("Asset support summary:")
        lines.append(f"- figures supported = {asset_support.get('supports_figures')}")
        lines.append(f"- tables supported = {asset_support.get('supports_tables')}")
        lines.append(f"- formulas supported = {asset_support.get('supports_formulas')}")
    return "\n".join(lines)


def build_retrieval_mismatches(
    target_summary: Dict[str, Any],
    observed_summary: Dict[str, Any],
    asset_support: Dict[str, Any],
) -> List[Dict[str, Any]]:
    metric_map = {
        "target_slide_count": ("slide_count", 1.0),
        "target_avg_bullets_per_slide": ("avg_bullets_per_slide", 0.75),
        "target_avg_words_per_slide": ("avg_words_per_slide", 4.0),
        "target_image_slide_count": ("image_slide_count", 1.0),
        "target_table_slide_count": ("table_slide_count", 1.0),
        "target_formula_slide_count": ("formula_slide_count", 1.0),
    }
    mismatches: List[Dict[str, Any]] = []
    for target_key, (observed_key, tolerance) in metric_map.items():
        if target_key not in target_summary:
            continue
        target_value = float(target_summary[target_key])
        observed_value = float(observed_summary.get(observed_key, 0.0))
        delta = round(observed_value - target_value, 4)
        if abs(delta) <= tolerance:
            continue
        actionable = True
        blocked_reason = None
        if target_key == "target_table_slide_count":
            actionable = asset_support.get("supports_tables", False)
            blocked_reason = None if actionable else "missing_table_assets"
        elif target_key == "target_formula_slide_count":
            actionable = asset_support.get("supports_formulas", False)
            blocked_reason = None if actionable else "missing_formula_assets"
        elif target_key == "target_image_slide_count":
            actionable = asset_support.get("supports_figures", False)
            blocked_reason = None if actionable else "missing_image_assets"
        goal = "increase" if delta < 0 else "decrease"
        mismatches.append(
            {
                "metric": target_key,
                "observed_metric": observed_key,
                "observed": observed_value,
                "target": target_value,
                "delta": delta,
                "tolerance": tolerance,
                "goal": goal,
                "actionable": actionable,
                "blocked_reason": blocked_reason,
            }
        )
    mismatches.sort(key=lambda item: (not item.get("actionable", False), -abs(float(item.get("delta", 0.0)))))
    return mismatches


def build_retrieval_repair_directives(
    profile: Dict[str, Any] | None,
    observed_summary: Dict[str, Any],
    asset_support: Dict[str, Any],
) -> Dict[str, Any]:
    target_summary = build_retrieval_target_summary(profile)
    budget_summary = build_retrieval_budget(
        target_summary,
        asset_support=asset_support,
        observed_summary=observed_summary,
    )
    mismatches = build_retrieval_mismatches(target_summary, observed_summary, asset_support)
    actionable_targets = [item for item in mismatches if item.get("actionable")]

    edit_brief: List[str] = []
    for item in actionable_targets:
        metric = item["metric"]
        goal = item["goal"]
        change_needed = int(round(abs(float(item.get("delta", 0.0)))))
        if metric == "target_slide_count":
            if goal == "increase":
                edit_brief.append(f"Increase deck length by about {max(change_needed, 1)} slide(s) using faithful splits or additional necessary slides.")
            else:
                edit_brief.append(f"Reduce deck length by about {max(change_needed, 1)} slide(s) using faithful merges or by removing low-value fragmentation.")
        elif metric == "target_avg_bullets_per_slide":
            if goal == "increase":
                edit_brief.append("Increase bullet structure on sparse slides so the deck moves toward the target average bullets per slide.")
            else:
                edit_brief.append("Reduce bullet density on overly busy slides so the deck moves toward the target average bullets per slide.")
        elif metric == "target_avg_words_per_slide":
            if goal == "increase":
                edit_brief.append("Increase textual detail on underfilled slides so the deck moves toward the target average words per slide.")
            else:
                edit_brief.append("Trim overly dense slides so the deck moves toward the target average words per slide.")
        elif metric == "target_table_slide_count":
            if goal == "increase":
                edit_brief.append(f"Promote about {max(change_needed, 1)} additional supported slide(s) to include tables where they directly support quantitative claims.")
            else:
                edit_brief.append(f"Reduce table-bearing slides by about {max(change_needed, 1)}; remove lower-value table placements first.")
        elif metric == "target_formula_slide_count":
            if goal == "increase":
                edit_brief.append(f"Promote about {max(change_needed, 1)} additional supported slide(s) to include formulas where equations are available and important.")
            else:
                edit_brief.append(f"Reduce formula-bearing slides by about {max(change_needed, 1)}; keep formulas only on the most important equation slides.")
        elif metric == "target_image_slide_count":
            if goal == "increase":
                edit_brief.append(f"Promote about {max(change_needed, 1)} additional supported slide(s) to include image assets when available.")
            else:
                edit_brief.append(f"Reduce image-bearing slides by about {max(change_needed, 1)}; keep visuals on the highest-communication-value slides.")

    return {
        "needs_repair": bool(actionable_targets),
        "target_summary": target_summary,
        "budget_summary": budget_summary,
        "observed_summary": observed_summary,
        "mismatches": mismatches,
        "actionable_targets": actionable_targets,
        "edit_brief": edit_brief,
        "total_mismatch_score": round(sum(abs(float(item.get("delta", 0.0))) for item in actionable_targets), 4),
    }


def evaluate_retrieval_repair_acceptance(
    draft_summary: Dict[str, Any],
    repaired_summary: Dict[str, Any],
    draft_directives: Dict[str, Any],
    repaired_directives: Dict[str, Any],
) -> Dict[str, Any]:
    draft_score = float(draft_directives.get("total_mismatch_score", 0.0))
    repaired_score = float(repaired_directives.get("total_mismatch_score", 0.0))
    score_improvement = round(draft_score - repaired_score, 4)
    measurable_improvements: List[str] = []
    metric_map = {
        "target_slide_count": "slide_count",
        "target_avg_bullets_per_slide": "avg_bullets_per_slide",
        "target_avg_words_per_slide": "avg_words_per_slide",
        "target_image_slide_count": "image_slide_count",
        "target_table_slide_count": "table_slide_count",
        "target_formula_slide_count": "formula_slide_count",
    }
    target_summary = dict(draft_directives.get("target_summary") or {})
    for mismatch in list(draft_directives.get("actionable_targets") or []):
        metric = str(mismatch.get("metric") or "")
        observed_key = metric_map.get(metric)
        if not observed_key or metric not in target_summary:
            continue
        target = float(target_summary[metric])
        before = abs(float(draft_summary.get(observed_key, 0.0)) - target)
        after = abs(float(repaired_summary.get(observed_key, 0.0)) - target)
        if before - after >= float(mismatch.get("tolerance", 0.05)) * 0.5:
            measurable_improvements.append(metric)

    slide_count_delta = abs(float(repaired_summary.get("slide_count", 0.0)) - float(draft_summary.get("slide_count", 0.0)))
    accepted = score_improvement > 0.0 and bool(measurable_improvements) and slide_count_delta <= 3.0
    reason = (
        "repair improved direct retrieval targets"
        if accepted
        else "repair did not measurably improve the direct retrieval targets enough to justify replacement"
    )
    return {
        "accepted": accepted,
        "score_improvement": score_improvement,
        "measurable_target_improvements": measurable_improvements,
        "slide_count_delta": slide_count_delta,
        "reason": reason,
    }


def render_retrieval_planner_prompt(
    *,
    base_template: Any,
    raw_json: Dict[str, Any],
    figures_json: Dict[str, Any],
    formulas_json: Dict[str, Any],
    images: Dict[str, Any],
    tables: Dict[str, Any],
    profile: Dict[str, Any] | None,
) -> str:
    base_prompt = base_template.render(
        raw_result_json=raw_json,
        figures_json=figures_json,
        formulas_json=formulas_json,
        image_informations_json=images,
        table_informations_json=tables,
        use_author_preferences=False,
        author_preference_profile_json=None,
    )
    retrieval_target_summary = build_retrieval_target_summary(profile)
    retrieval_budget_summary = build_retrieval_budget(retrieval_target_summary)
    retrieval_section_preferences = build_retrieval_section_preferences(profile)
    retrieval_brief = build_retrieval_planner_brief(
        target_summary=retrieval_target_summary,
        budget_summary=retrieval_budget_summary,
        section_preferences=retrieval_section_preferences,
    )
    return base_prompt + "\n\n" + retrieval_brief


def build_retrieval_repair_prompt(
    *,
    raw_json: Dict[str, Any],
    figures_json: Dict[str, Any],
    formulas_json: Dict[str, Any],
    profile: Dict[str, Any],
    retrieval_target_summary: Dict[str, Any],
    current_plan_summary: Dict[str, Any],
    repair_directives: Dict[str, Any],
    current_slide_plan: Dict[str, Any],
) -> str:
    explicit_gap_lines: List[str] = []
    for item in list(repair_directives.get("actionable_targets") or []):
        metric = str(item.get("metric") or "")
        observed = float(item.get("observed", 0.0))
        target = float(item.get("target", 0.0))
        delta = float(item.get("delta", 0.0))
        support_note = ""
        if not item.get("actionable", True):
            support_note = f" (blocked: {item.get('blocked_reason')})"
        if metric == "target_slide_count":
            if delta < 0:
                explicit_gap_lines.append(
                    f"- Slide count is too low: current={int(round(observed))}, target={int(round(target))}. Add about {max(int(round(abs(delta))), 1)} slide(s) while staying faithful{support_note}."
                )
            else:
                explicit_gap_lines.append(
                    f"- Slide count is too high: current={int(round(observed))}, target={int(round(target))}. Remove or merge about {max(int(round(abs(delta))), 1)} slide(s) while staying faithful{support_note}."
                )
        elif metric == "target_avg_bullets_per_slide":
            if delta < 0:
                explicit_gap_lines.append(
                    f"- Average bullets per slide is too low: current={observed:.4f}, target={target:.4f}. Increase average bullets per slide by about {abs(delta):.4f}{support_note}."
                )
            else:
                explicit_gap_lines.append(
                    f"- Average bullets per slide is too high: current={observed:.4f}, target={target:.4f}. Decrease average bullets per slide by about {abs(delta):.4f}{support_note}."
                )
        elif metric == "target_avg_words_per_slide":
            if delta < 0:
                explicit_gap_lines.append(
                    f"- Average words per slide is too low: current={observed:.4f}, target={target:.4f}. Increase average words per slide by about {abs(delta):.4f}{support_note}."
                )
            else:
                explicit_gap_lines.append(
                    f"- Average words per slide is too high: current={observed:.4f}, target={target:.4f}. Decrease average words per slide by about {abs(delta):.4f}{support_note}."
                )
        elif metric == "target_image_slide_count":
            if delta < 0:
                explicit_gap_lines.append(
                    f"- Image slide count is too low: current={int(round(observed))}, target={int(round(target))}. Add images to about {max(int(round(abs(delta))), 1)} more slide(s){support_note}."
                )
            else:
                explicit_gap_lines.append(
                    f"- Image slide count is too high: current={int(round(observed))}, target={int(round(target))}. Remove images from about {max(int(round(abs(delta))), 1)} slide(s){support_note}."
                )
        elif metric == "target_table_slide_count":
            if delta < 0:
                explicit_gap_lines.append(
                    f"- Table slide count is too low: current={int(round(observed))}, target={int(round(target))}. Add tables to about {max(int(round(abs(delta))), 1)} more slide(s){support_note}."
                )
            else:
                explicit_gap_lines.append(
                    f"- Table slide count is too high: current={int(round(observed))}, target={int(round(target))}. Remove tables from about {max(int(round(abs(delta))), 1)} slide(s){support_note}."
                )
        elif metric == "target_formula_slide_count":
            if delta < 0:
                explicit_gap_lines.append(
                    f"- Formula slide count is too low: current={int(round(observed))}, target={int(round(target))}. Add formulas to about {max(int(round(abs(delta))), 1)} more slide(s){support_note}."
                )
            else:
                explicit_gap_lines.append(
                    f"- Formula slide count is too high: current={int(round(observed))}, target={int(round(target))}. Remove formulas from about {max(int(round(abs(delta))), 1)} slide(s){support_note}."
                )

    explicit_gap_block = "\n".join(explicit_gap_lines) if explicit_gap_lines else "- No explicit actionable gaps were computed."

    return f"""
Revise the existing slide plan so it better matches the retrieval-profile targets.

Hard constraints:
- Keep the deck faithful to the paper.
- Revise the current draft directly rather than regenerating from scratch.
- Preserve section order unless a small local change clearly improves target fit.
- Preserve the deck's overall fidelity while nudging section structure toward the preferred section labels and typical section count when the source paper supports those changes.
- Merge/split slides only when it directly improves target fit while staying faithful.
- Do not invent visuals, tables, formulas, or claims.
- Use the retrieval-profile targets directly during repair.
- Treat image/table/formula count targets as a whole-deck budget that should be matched as exactly as possible.
- Reassign modality usage across slides deliberately: decide which slides should carry images, tables, and formulas so the final deck spends the budget well.
- If slide count is off target, use faithful merges or splits to move the deck toward the target total number of slides.
- If average bullets per slide is off target, add, remove, merge, or simplify bullet structure to move toward the target while staying faithful.
- If average words per slide is off target, condense, expand, merge, or split slides when faithful to the source content.
- If preferred section types are missing but clearly supported by the source paper, adjust section grouping or section titles to better reflect them.
- Make concrete, target-seeking edits. Do not just make small stylistic changes; close the measured gaps below as directly as possible.
- When a modality count is below target, explicitly choose additional slides to carry that modality.
- When a modality count is above target, explicitly remove that modality from lower-value slides first.
- When text density is above target, shorten dense slides or split them if faithful. When it is below target, merge thin slides or add faithful explanatory detail.

Measured target gaps to fix:
{explicit_gap_block}

Retrieval profile:
{json.dumps(profile, ensure_ascii=False, indent=2)}

Numeric target summary:
{json.dumps(retrieval_target_summary, ensure_ascii=False, indent=2)}

Current observed summary:
{json.dumps(current_plan_summary, ensure_ascii=False, indent=2)}

Repair directives:
{json.dumps(repair_directives, ensure_ascii=False, indent=2)}

Existing slide plan:
{json.dumps(current_slide_plan, ensure_ascii=False, indent=2)}

Available raw_result:
{json.dumps(raw_json, ensure_ascii=False, indent=2)}

Available figures:
{json.dumps(figures_json, ensure_ascii=False, indent=2)}

Available formulas:
{json.dumps(formulas_json, ensure_ascii=False, indent=2)}

Return the repaired slide-plan JSON only.
""".strip()


def generate_slide_plan(args: Any):
    paper_outline_json = raw_content_path(args)
    figures_path = figures_json_path(args)
    if args.formula_mode in (1, 2):
        formulas_path = formula_match_path(args)
    else:
        formulas_path = formula_mode3_index_path(args)

    raw_json = json.loads(Path(paper_outline_json).read_text(encoding="utf-8"))
    figures_json = json.loads(Path(figures_path).read_text(encoding="utf-8"))
    formulas_json = json.loads(Path(formulas_path).read_text(encoding="utf-8"))
    images = json.loads(images_filtered_path(args).read_text(encoding="utf-8"))
    tables = json.loads(tables_filtered_path(args).read_text(encoding="utf-8"))
    asset_support = derive_asset_support(formulas_json=formulas_json, images=images, tables=tables)

    author_profile_arg = getattr(args, "author_profile_path", None)
    author_preference_profile: Dict[str, Any] | None = None
    if author_profile_arg:
        profile_path = Path(author_profile_arg)
        if not profile_path.exists():
            raise FileNotFoundError(f"Retrieval preference profile not found: {profile_path}")
        author_preference_profile = sanitize_retrieval_profile_for_generation(
            json.loads(profile_path.read_text(encoding="utf-8"))
        )
    else:
        profile_path = None

    with open(BASE_PROMPT_PATH, "r", encoding="utf-8") as handle:
        prompt_cfg = yaml.safe_load(handle)

    use_gpt5_responses = False
    cfg = get_agent_config(args.model_name_v)
    if "gpt-5" in str(args.model_name_v).lower():
        client = build_openai_client()
        use_gpt5_responses = True
        agent = None
        repair_agent = None
    else:
        if str(args.model_name_v).startswith("vllm_qwen"):
            model = ModelFactory.create(
                model_platform=cfg["model_platform"],
                model_type=cfg["model_type"],
                model_config_dict=cfg["model_config"],
                url=cfg["url"],
            )
            agent = model
            repair_agent = model
        else:
            model = ModelFactory.create(
                model_platform=cfg["model_platform"],
                model_type=cfg["model_type"],
                model_config_dict=cfg["model_config"],
                url=cfg.get("url"),
            )
            agent = ChatAgent(system_message=prompt_cfg["system_prompt"], model=model, message_window_size=5)
            repair_agent = ChatAgent(system_message=prompt_cfg.get("repair_system_prompt", prompt_cfg["system_prompt"]), model=model, message_window_size=5)

    jinja_env = Environment(undefined=StrictUndefined)
    base_template = jinja_env.from_string(prompt_cfg["template"])

    planner_prompt = render_retrieval_planner_prompt(
        base_template=base_template,
        raw_json=raw_json,
        figures_json=figures_json,
        formulas_json=formulas_json,
        images=images,
        tables=tables,
        profile=author_preference_profile,
    )

    raw_text, in_tok, out_tok, time_taken = call_layout_model(
        planner_prompt,
        prompt_cfg["system_prompt"],
        args=args,
        cfg=cfg,
        use_gpt5_responses=use_gpt5_responses,
        client=client if use_gpt5_responses else None,
        agent=agent if not use_gpt5_responses else None,
    )
    slide_plan = sanitize_slide_plan_templates(get_json_from_response(raw_text))
    draft_summary = summarize_retrieval_targets(slide_plan)
    draft_directives = build_retrieval_repair_directives(author_preference_profile, draft_summary, asset_support)

    draft_path = slide_plan_draft_path(args, plan_variant_suffix(args))
    draft_path.parent.mkdir(parents=True, exist_ok=True)
    draft_path.write_text(json.dumps(slide_plan, indent=2, ensure_ascii=False), encoding="utf-8")

    repaired_summary: Dict[str, Any] | None = None
    repaired_directives: Dict[str, Any] | None = None
    acceptance: Dict[str, Any] | None = None
    repair_attempted = False
    repair_report_path: Path | None = None
    repair_round_reports: List[Dict[str, Any]] = []
    accepted_round_index: int | None = None
    if author_preference_profile and draft_directives.get("needs_repair"):
        repair_attempted = True
        current_plan = slide_plan
        current_summary = draft_summary
        current_directives = draft_directives
        for repair_round in range(1, MAX_RETRIEVAL_REPAIR_ROUNDS + 1):
            if not current_directives.get("needs_repair"):
                break
            repair_prompt = build_retrieval_repair_prompt(
                raw_json=raw_json,
                figures_json=figures_json,
                formulas_json=formulas_json,
                profile=author_preference_profile,
                retrieval_target_summary=current_directives["target_summary"],
                current_plan_summary=current_summary,
                repair_directives=current_directives,
                current_slide_plan=current_plan,
            )
            repair_raw_text, repair_in_tok, repair_out_tok, repair_time_taken = call_layout_model(
                repair_prompt,
                prompt_cfg.get("repair_system_prompt", prompt_cfg["system_prompt"]),
                args=args,
                cfg=cfg,
                use_gpt5_responses=use_gpt5_responses,
                client=client if use_gpt5_responses else None,
                agent=repair_agent if not use_gpt5_responses else None,
            )
            candidate_plan = sanitize_slide_plan_templates(get_json_from_response(repair_raw_text))
            candidate_summary = summarize_retrieval_targets(candidate_plan)
            candidate_directives = build_retrieval_repair_directives(author_preference_profile, candidate_summary, asset_support)
            round_acceptance = evaluate_retrieval_repair_acceptance(
                current_summary,
                candidate_summary,
                current_directives,
                candidate_directives,
            )
            repair_round_reports.append(
                {
                    "round": repair_round,
                    "input_summary": current_summary,
                    "input_repair_directives": current_directives,
                    "candidate_summary": candidate_summary,
                    "candidate_repair_directives": candidate_directives,
                    "acceptance": round_acceptance,
                }
            )
            if round_acceptance.get("accepted"):
                current_plan = candidate_plan
                current_summary = candidate_summary
                current_directives = candidate_directives
                accepted_round_index = repair_round
                in_tok += repair_in_tok
                out_tok += repair_out_tok
                time_taken += repair_time_taken
                if not current_directives.get("needs_repair"):
                    break
            else:
                break
        slide_plan = current_plan
        draft_summary = current_summary
        draft_directives = current_directives
        if repair_round_reports:
            last_round = repair_round_reports[-1]
            repaired_summary = last_round.get("candidate_summary")
            repaired_directives = last_round.get("candidate_repair_directives")
            acceptance = last_round.get("acceptance")
        repair_report_path = slide_plan_repair_report_path(args, plan_variant_suffix(args))
        repair_report_path.write_text(
            json.dumps(
                {
                    "accepted_round_index": accepted_round_index,
                    "repair_round_reports": repair_round_reports,
                    "final_summary_after_repairs": draft_summary,
                    "final_repair_directives_after_repairs": draft_directives,
                },
                indent=2,
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

    final_plan_path = slide_plan_path(args, plan_variant_suffix(args))
    final_plan_path.parent.mkdir(parents=True, exist_ok=True)
    final_plan_path.write_text(json.dumps(slide_plan, indent=2, ensure_ascii=False), encoding="utf-8")

    trace_path = personalization_trace_path(args, plan_variant_suffix(args))
    trace_payload = {
        "paper_name": args.paper_name,
        "model_name_t": args.model_name_t,
        "model_name_v": args.model_name_v,
        "plan_variant_suffix": plan_variant_suffix(args),
        "use_author_preferences": bool(author_preference_profile),
        "author_profile_path": str(profile_path) if profile_path else None,
        "author_profile_summary": build_profile_trace(author_preference_profile),
        "retrieval_target_summary": build_retrieval_target_summary(author_preference_profile),
        "retrieval_budget_summary": build_retrieval_budget(
            build_retrieval_target_summary(author_preference_profile),
            asset_support=asset_support,
            observed_summary=draft_summary,
        ),
        "asset_support": asset_support,
        "planner": {
            "source": "retrieval_direct_planner",
            "draft_summary": draft_summary,
            "draft_repair_directives": draft_directives,
        },
        "repair": {
            "attempted": repair_attempted,
            "accepted": bool(acceptance and acceptance.get("accepted")),
            "accepted_round_index": accepted_round_index,
            "max_rounds": MAX_RETRIEVAL_REPAIR_ROUNDS,
            "round_count": len(repair_round_reports),
            "acceptance": acceptance,
            "repaired_summary": repaired_summary,
            "repaired_repair_directives": repaired_directives,
            "round_reports": repair_round_reports,
        },
        "paths": {
            "paper_outline_json": str(paper_outline_json),
            "figures_json": str(figures_path),
            "formulas_json": str(formulas_path),
            "images_json": str(images_filtered_path(args)),
            "tables_json": str(tables_filtered_path(args)),
            "draft_plan_json": str(draft_path),
            "repair_report_json": str(repair_report_path) if repair_report_path else None,
            "final_plan_json": str(final_plan_path),
        },
        "final": {
            "selected_plan_source": "accepted_repair" if acceptance and acceptance.get("accepted") else "retrieval_direct_planner",
            "final_summary": draft_summary,
        },
    }
    trace_path.write_text(json.dumps(trace_payload, indent=2, ensure_ascii=False), encoding="utf-8")
    return in_tok, out_tok, time_taken
