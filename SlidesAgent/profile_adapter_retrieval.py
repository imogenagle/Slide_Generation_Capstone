from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def _clamp01(value: float) -> float:
    return max(0.0, min(1.0, value))


def _coerce_float(value: Any) -> float | None:
    try:
        number = float(value)
    except Exception:
        return None
    if number != number or number in (float("inf"), float("-inf")):
        return None
    return number


def _usage_label(value: float | None) -> str:
    if value is None:
        return "unknown"
    if value < 0.2:
        return "low"
    if value < 0.5:
        return "medium"
    return "high"


def _text_density_label(words_per_slide: float | None) -> str:
    if words_per_slide is None:
        return "unknown"
    if words_per_slide < 18.0:
        return "low"
    if words_per_slide < 38.0:
        return "medium"
    return "high"


def is_retrieval_profile(profile: dict[str, Any] | None) -> bool:
    if not isinstance(profile, dict):
        return False
    method = str(profile.get("profile_method") or "").lower()
    if "retrieval" in method:
        return True
    numeric = dict(profile.get("numeric_preferences") or {})
    return any(
        key in numeric
        for key in (
            "target_fraction_slides_with_images",
            "target_fraction_slides_with_tables",
            "target_fraction_slides_with_formulas",
            "target_fraction_slides_with_figures",
        )
    )


def adapt_retrieval_profile(profile: dict[str, Any]) -> dict[str, Any]:
    numeric = dict(profile.get("numeric_preferences") or {})
    retrieval_context = dict(profile.get("retrieval_context") or {})

    avg_words = _coerce_float(numeric.get("target_avg_words_per_slide"))
    image_fraction = _coerce_float(numeric.get("target_fraction_slides_with_images"))
    table_fraction = _coerce_float(numeric.get("target_fraction_slides_with_tables"))
    formula_fraction = _coerce_float(numeric.get("target_fraction_slides_with_formulas"))
    figure_fraction = _coerce_float(numeric.get("target_fraction_slides_with_figures"))

    collapsed_visual_fraction = None
    visual_components = [v for v in (image_fraction, figure_fraction) if v is not None]
    if visual_components:
        collapsed_visual_fraction = _clamp01(sum(visual_components))

    adapted_numeric: dict[str, Any] = {}
    if avg_words is not None:
        adapted_numeric["target_avg_words_per_slide"] = round(avg_words, 4)
    if table_fraction is not None:
        adapted_numeric["target_fraction_table_slides"] = round(_clamp01(table_fraction), 4)
    if formula_fraction is not None:
        adapted_numeric["target_fraction_formula_slides"] = round(_clamp01(formula_fraction), 4)
    if collapsed_visual_fraction is not None:
        adapted_numeric["target_fraction_figure_slides"] = round(collapsed_visual_fraction, 4)

    adapted_planning = {
        "section_splitting_preference": "unknown",
        "bullet_density_preference": "unknown",
        "text_density_preference": _text_density_label(avg_words),
        "visual_density_preference": _usage_label(collapsed_visual_fraction),
        "figure_usage_preference": _usage_label(collapsed_visual_fraction),
        "table_usage_preference": _usage_label(table_fraction),
        "formula_usage_preference": _usage_label(formula_fraction),
        "layout_bias": [],
        "typical_section_categories": [],
        "structure_preferences": {
            "prefers_agenda_slide": "unknown",
            "prefers_takeaway_slide": "unknown",
            "prefers_multi_slide_method_section": "unknown",
            "prefers_multi_slide_results_section": "unknown",
        },
    }

    adapted = {
        "author_id": profile.get("author_id"),
        "profile_version": 6,
        "profile_method": "retrieval_runtime_adapter",
        "distilled_from": profile.get("distilled_from") or {},
        "planning_preferences": adapted_planning,
        "numeric_preferences": adapted_numeric,
        "evidence_summary": profile.get("evidence_summary") or {},
        "retrieval_context": retrieval_context,
        "retrieval_adapter_metadata": {
            "source_profile_method": profile.get("profile_method"),
            "source_numeric_preferences": numeric,
            "collapsed_visual_fraction_from": {
                "target_fraction_slides_with_images": image_fraction,
                "target_fraction_slides_with_figures": figure_fraction,
            },
            "mapped_numeric_preferences": adapted_numeric,
        },
    }
    return adapted


def write_adapted_profile(source_path: Path, adapted_profile: dict[str, Any]) -> Path:
    adapted_path = source_path.parent / f"{source_path.stem}.legacy_adapter.json"
    adapted_path.write_text(json.dumps(adapted_profile, indent=2, ensure_ascii=False), encoding="utf-8")
    return adapted_path
