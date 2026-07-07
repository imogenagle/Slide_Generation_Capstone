from __future__ import annotations

import math
from typing import Any


def profile_source_paper_count(author_profile: dict[str, Any] | None) -> int:
    if not author_profile:
        return 0
    distilled_from = dict(author_profile.get("distilled_from") or {})
    try:
        return max(0, int(distilled_from.get("paper_count", 0) or 0))
    except Exception:
        return 0


def profile_target_tolerance_multiplier(author_profile: dict[str, Any] | None) -> float:
    paper_count = profile_source_paper_count(author_profile)
    if paper_count <= 0:
        return 1.15
    if paper_count == 1:
        return 1.35
    if paper_count == 2:
        return 1.15
    return 1.0


def _slide_count_softening_radius(target_slide_count: float) -> int:
    return max(1, int(round(max(1.0, target_slide_count * 0.12))))


def _text_proxy_softening_radius(proxy_value: float) -> float:
    return round(max(0.02, proxy_value * 0.18), 4)


def build_numeric_target_summary(author_profile: dict[str, Any] | None) -> dict[str, Any]:
    if not author_profile:
        return {}

    numeric = dict(author_profile.get("numeric_preferences") or {})
    if not numeric:
        return {}

    summary = {key: value for key, value in numeric.items() if value not in (None, [], {}, "")}
    paper_count = profile_source_paper_count(author_profile)
    summary["source_paper_count"] = paper_count

    avg_text_density_proxy = numeric.get("target_avg_text_density_proxy", numeric.get("avg_text_density_proxy"))
    text_density_proxy_std = numeric.get("text_density_proxy_std")
    if avg_text_density_proxy is not None:
        avg_text_density_proxy = float(avg_text_density_proxy or 0.0)
        summary["target_avg_text_density_proxy"] = avg_text_density_proxy
    if text_density_proxy_std is not None:
        text_density_proxy_std = float(text_density_proxy_std or 0.0)
        summary["text_density_proxy_std"] = text_density_proxy_std
    if avg_text_density_proxy is not None and text_density_proxy_std is not None:
        proxy_range = [
            round(max(0.0, avg_text_density_proxy - text_density_proxy_std), 4),
            round(avg_text_density_proxy + text_density_proxy_std, 4),
        ]
        summary["text_density_proxy_target_range"] = proxy_range
        summary["target_text_density_proxy_range"] = list(proxy_range)

    if paper_count == 1:
        summary["softened_for_sparse_profile"] = True

        if "target_slide_count" in summary:
            target_slide_count = float(summary["target_slide_count"])
            radius = _slide_count_softening_radius(target_slide_count)
            low = max(1, int(math.floor(target_slide_count - radius)))
            high = max(low + 1, int(math.ceil(target_slide_count + radius)))
            summary["slide_count_range"] = [low, high]
            current_std = float(summary.get("slide_count_std", 0.0) or 0.0)
            summary["slide_count_std"] = round(max(current_std, radius / 1.5), 2)

        if avg_text_density_proxy is not None:
            proxy_radius = _text_proxy_softening_radius(avg_text_density_proxy)
            current_proxy_std = float(summary.get("text_density_proxy_std", 0.0) or 0.0)
            effective_proxy_std = round(max(current_proxy_std, proxy_radius), 4)
            summary["text_density_proxy_std"] = effective_proxy_std
            proxy_range = [
                round(max(0.0, avg_text_density_proxy - effective_proxy_std), 4),
                round(avg_text_density_proxy + effective_proxy_std, 4),
            ]
            summary["text_density_proxy_target_range"] = proxy_range
            summary["target_text_density_proxy_range"] = list(proxy_range)
    else:
        summary["softened_for_sparse_profile"] = False

    return summary
