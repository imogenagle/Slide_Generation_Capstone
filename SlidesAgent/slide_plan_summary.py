from __future__ import annotations

from typing import Any


def mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def stdev(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0
    avg = mean(values)
    variance = sum((v - avg) ** 2 for v in values) / len(values)
    return variance ** 0.5


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


def summarize_slide_plan(plan: dict[str, Any]) -> dict[str, Any]:
    slides = list(plan.get("slides") or [])
    slide_count = len(slides)

    section_order: list[str] = []
    section_counts: dict[str, int] = {}
    bullets_per_slide: list[int] = []
    sub_bullets_per_slide: list[int] = []
    words_per_slide: list[int] = []
    figure_flags: list[int] = []
    table_flags: list[int] = []
    formula_flags: list[int] = []
    layout_bias_counts = {
        "text_only": 0,
        "image_right": 0,
        "image_left": 0,
        "image_top": 0,
        "multi_visual": 0,
        "formula_capable": 0,
    }

    for slide in slides:
        if not isinstance(slide, dict):
            continue

        section = str(slide.get("section") or "").strip() or "UNKNOWN"
        if section not in section_counts:
            section_order.append(section)
            section_counts[section] = 0
        section_counts[section] += 1

        bullets = list(slide.get("bullets") or [])
        bullet_count = len(bullets)
        sub_bullet_count = sum(len(b.get("sub") or []) for b in bullets if isinstance(b, dict))
        bullets_per_slide.append(bullet_count)
        sub_bullets_per_slide.append(sub_bullet_count)

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
        if image_count + table_count + formula_count >= 2 or "2Text" in template_id:
            layout_bias_counts["multi_visual"] += 1
        if formula_count > 0:
            layout_bias_counts["formula_capable"] += 1

    section_slide_counts = [section_counts[name] for name in section_order]
    section_count = len(section_order)
    slides_per_section = (slide_count / section_count) if section_count else 0.0

    summary = {
        "slide_count": slide_count,
        "section_count": section_count,
        "section_titles": section_order,
        "section_slide_counts": section_counts,
        "avg_slides_per_section": round(slides_per_section, 3),
        "section_splitting_estimate": classify_structure(slides_per_section) if section_count else "unknown",
        "avg_bullets_per_slide": round(mean(bullets_per_slide), 3),
        "avg_sub_bullets_per_slide": round(mean(sub_bullets_per_slide), 3),
        "avg_words_per_slide": round(mean(words_per_slide), 3),
        "bullet_density_estimate": classify_level(mean(bullets_per_slide), 2.0, 4.0) if slides else "unknown",
        "text_density_estimate": classify_level(mean(words_per_slide), 18.0, 38.0) if slides else "unknown",
        "figure_usage_estimate": classify_level(sum(figure_flags) / slide_count, 0.2, 0.55) if slide_count else "unknown",
        "table_usage_estimate": classify_level(sum(table_flags) / slide_count, 0.08, 0.22) if slide_count else "unknown",
        "formula_usage_estimate": classify_level(sum(formula_flags) / slide_count, 0.08, 0.22) if slide_count else "unknown",
        "figure_slide_fraction": round(sum(figure_flags) / slide_count, 4) if slide_count else 0.0,
        "table_slide_fraction": round(sum(table_flags) / slide_count, 4) if slide_count else 0.0,
        "formula_slide_fraction": round(sum(formula_flags) / slide_count, 4) if slide_count else 0.0,
        "layout_bias_observed": [key for key, count in layout_bias_counts.items() if count > 0],
        "layout_bias_counts": layout_bias_counts,
        "text_only_fraction": round(layout_bias_counts["text_only"] / slide_count, 4) if slide_count else 0.0,
        "multi_visual_fraction": round(layout_bias_counts["multi_visual"] / slide_count, 4) if slide_count else 0.0,
        "formula_capable_fraction": round(layout_bias_counts["formula_capable"] / slide_count, 4) if slide_count else 0.0,
        "image_right_fraction": round(layout_bias_counts["image_right"] / slide_count, 4) if slide_count else 0.0,
        "image_left_fraction": round(layout_bias_counts["image_left"] / slide_count, 4) if slide_count else 0.0,
        "image_top_fraction": round(layout_bias_counts["image_top"] / slide_count, 4) if slide_count else 0.0,
        "prefers_takeaway_like_close": any(
            "conclusion" in title.lower() or "takeaway" in title.lower() or "future" in title.lower()
            for title in section_order
        ),
        "multi_slide_method_like_sections": [
            title
            for title, count in section_counts.items()
            if count >= 2 and any(token in title.lower() for token in ("method", "approach", "system", "model", "core idea"))
        ],
        "multi_slide_results_like_sections": [
            title
            for title, count in section_counts.items()
            if count >= 2 and any(token in title.lower() for token in ("result", "evaluation", "experiment", "analysis", "comparison"))
        ],
        "section_slide_count_std": round(stdev(section_slide_counts), 3),
    }
    return summary
