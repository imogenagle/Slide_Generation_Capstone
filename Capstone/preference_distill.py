#!/usr/bin/env python3
"""Standalone author preference distiller for SlideGen.

This script does not modify or integrate with the existing SlideGen generation
pipeline. It builds a planning-preference profile JSON for a single author by
analyzing prior paper/deck history from the normalized author tables plus
representative slide images from `data_raw`.
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import math
import os
import re
import sys
from pathlib import Path
from typing import Any

import cv2
import numpy as np
import yaml
from dotenv import load_dotenv
from jinja2 import Environment, StrictUndefined

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


DEFAULT_AUTHORS_CSV = REPO_ROOT / "Capstone" / "author_tables" / "authors.csv"
DEFAULT_PAPER_AUTHORS_CSV = REPO_ROOT / "Capstone" / "author_tables" / "paper_authors.csv"
DEFAULT_PAPERS_CSV = REPO_ROOT / "Capstone" / "author_tables" / "papers.csv"
DEFAULT_OUTPUT_DIR = REPO_ROOT / "Capstone" / "profiles"
DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "preference_distiller.yaml"

LAYOUT_BIAS_VALUES = [
    "text_only",
    "image_right",
    "image_left",
    "image_top",
    "multi_visual",
    "formula_capable",
]
SECTION_CATEGORY_VALUES = [
    "opening_context",
    "agenda_roadmap",
    "background_context",
    "motivation",
    "problem_statement",
    "prior_work",
    "approach_overview",
    "methodology_process",
    "system_architecture",
    "implementation_details",
    "data_inputs",
    "evaluation_validation",
    "results_findings",
    "analysis_discussion",
    "comparison_benchmark",
    "case_study_example",
    "limitations_risks",
    "recommendations_implications",
    "conclusion_takeaways",
    "future_directions",
    "appendix_qa",
]
PREFERENCE_LABELS = ["low", "medium", "high", "coarse", "balanced", "fine_grained", "unknown"]
SLIDE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}


def load_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def parse_bool(value: str) -> bool:
    return str(value).strip().lower() == "true"


def parse_int(value: str, default: int = 0) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return default


def resolve_repo_path(raw_path_value: str) -> Path:
    path = Path(raw_path_value)
    if path.exists():
        return path
    if path.is_absolute():
        return path
    if path.parts and path.parts[0] == "SlideGen":
        candidate = REPO_ROOT / Path(*path.parts[1:])
        if candidate.exists():
            return candidate
    return REPO_ROOT / path


def normalize_path_for_compare(path_like: str | Path | None) -> str:
    if not path_like:
        return ""
    try:
        return str(resolve_repo_path(str(path_like)).resolve())
    except Exception:
        return str(Path(path_like)).strip()


def numeric_slide_sort_key(path: Path) -> tuple[int, str]:
    try:
        return (int(path.stem), path.name)
    except ValueError:
        return (10**9, path.name)


def clamp01(value: float) -> float:
    return max(0.0, min(1.0, value))


def clamp(value: float, low: float, high: float) -> float:
    return max(low, min(high, value))


def bucket_score(value: float, low: float = 0.33, high: float = 0.66) -> str:
    if value < low:
        return "low"
    if value < high:
        return "medium"
    return "high"


def mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def stdev(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0
    avg = mean(values)
    variance = sum((value - avg) ** 2 for value in values) / len(values)
    return variance ** 0.5


def compute_colorfulness(image_bgr: Any) -> float:
    (b_channel, g_channel, r_channel) = cv2.split(image_bgr.astype("float"))
    rg = cv2.absdiff(r_channel, g_channel)
    yb = cv2.absdiff(0.5 * (r_channel + g_channel), b_channel)
    std_root = math.sqrt(float(rg.std() ** 2 + yb.std() ** 2))
    mean_root = math.sqrt(float(rg.mean() ** 2 + yb.mean() ** 2))
    return float(std_root + 0.3 * mean_root)


def compute_slide_metrics(slide_path: Path) -> dict[str, Any]:
    image = cv2.imread(str(slide_path))
    if image is None:
        return {
            "slide_path": str(slide_path),
            "readable": False,
        }

    original_height, original_width = image.shape[:2]
    max_side = max(original_width, original_height)
    if max_side > 512:
        scale = 512.0 / max_side
        image = cv2.resize(
            image,
            (max(1, int(original_width * scale)), max(1, int(original_height * scale))),
            interpolation=cv2.INTER_AREA,
        )

    height, width = image.shape[:2]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 100, 200)
    edge_density = float((edges > 0).mean())
    brightness = float(gray.mean() / 255.0)

    colorfulness_raw = compute_colorfulness(image)
    colorfulness = clamp01(colorfulness_raw / 100.0)

    mser = cv2.MSER_create()
    regions, _ = mser.detectRegions(gray)
    mser_mask = np.zeros(gray.shape, dtype=np.uint8)
    for region in regions:
        hull = cv2.convexHull(region.reshape(-1, 1, 2))
        cv2.drawContours(mser_mask, [hull], -1, 255, -1)
    text_region_ratio = float((mser_mask > 0).mean())

    return {
        "slide_path": str(slide_path),
        "readable": True,
        "width": original_width,
        "height": original_height,
        "analysis_width": width,
        "analysis_height": height,
        "aspect_ratio": round(float(original_width / max(original_height, 1)), 3),
        "brightness": round(brightness, 4),
        "edge_density": round(edge_density, 4),
        "colorfulness": round(colorfulness, 4),
        "text_region_ratio": round(text_region_ratio, 4),
    }


def aggregate_deck_metrics(metrics: list[dict[str, Any]]) -> dict[str, Any]:
    readable = [metric for metric in metrics if metric.get("readable")]
    if not readable:
        return {
            "text_density_estimate": "unknown",
            "avg_slide_brightness": "unknown",
            "avg_slide_colorfulness": "unknown",
            "numeric_stats": {
                "avg_text_region_ratio": 0.0,
                "avg_brightness": 0.0,
                "avg_colorfulness": 0.0,
            },
        }

    avg_text_ratio = sum(metric["text_region_ratio"] for metric in readable) / len(readable)
    avg_brightness = sum(metric["brightness"] for metric in readable) / len(readable)
    avg_colorfulness = sum(metric["colorfulness"] for metric in readable) / len(readable)

    return {
        "text_density_estimate": bucket_score(clamp01(avg_text_ratio * 4.0)),
        "avg_slide_brightness": bucket_score(avg_brightness),
        "avg_slide_colorfulness": bucket_score(avg_colorfulness),
        "numeric_stats": {
            "avg_text_region_ratio": round(avg_text_ratio, 4),
            "avg_brightness": round(avg_brightness, 4),
            "avg_colorfulness": round(avg_colorfulness, 4),
        },
    }


def encode_image_data_uri(image_path: Path) -> str:
    mime_type = "image/jpeg"
    if image_path.suffix.lower() == ".png":
        mime_type = "image/png"
    elif image_path.suffix.lower() == ".webp":
        mime_type = "image/webp"

    data = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{data}"


def extract_json_object(raw_text: str) -> dict[str, Any]:
    text = (raw_text or "").strip()
    if not text:
        raise ValueError("Model returned an empty response.")

    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, flags=re.DOTALL)
    if fenced:
        text = fenced.group(1).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        first = text.find("{")
        last = text.rfind("}")
        if first != -1 and last != -1 and last > first:
            return json.loads(text[first : last + 1])
        raise


def sample_representative_slide_indices(metrics: list[dict[str, Any]], max_samples: int = 6) -> list[int]:
    if not metrics:
        return []

    count = len(metrics)
    readable_indices = [idx for idx, metric in enumerate(metrics) if metric.get("readable")]
    if count <= max_samples:
        return readable_indices

    indices: list[int] = []

    def add_index(idx: int) -> None:
        if 0 <= idx < count and idx in readable_indices and idx not in indices and len(indices) < max_samples:
            indices.append(idx)

    # Cover opening, early, middle, late, and closing deck regions.
    for candidate in [0, 1, count // 3, count // 2, (2 * count) // 3, count - 1]:
        add_index(candidate)

    remaining = [idx for idx in readable_indices if idx not in indices]

    # If deck length causes anchor collisions, backfill with slides that are
    # unusually text-heavy or visually distinctive.
    def add_extreme(metric_key: str) -> None:
        nonlocal remaining
        if len(indices) >= max_samples or not remaining:
            return
        chosen = max(remaining, key=lambda idx: metrics[idx].get(metric_key, 0.0))
        add_index(chosen)
        remaining = [idx for idx in remaining if idx != chosen]

    add_extreme("text_region_ratio")
    add_extreme("colorfulness")
    add_extreme("edge_density")

    while len(indices) < max_samples and remaining:
        # Evenly fill any leftover slots if duplicates reduced anchor coverage.
        step = max(1, len(remaining) // max(1, (max_samples - len(indices))))
        chosen = remaining[0]
        add_index(chosen)
        remaining = remaining[step:]

    return sorted(indices)


def build_numeric_preferences(deck_evidence: list[dict[str, Any]]) -> dict[str, Any]:
    slide_counts: list[int] = []
    text_ratios: list[float] = []
    brightness: list[float] = []
    colorfulness: list[float] = []

    for deck in deck_evidence:
        if "slide_count" in deck:
            slide_counts.append(int(deck["slide_count"]))
        numeric_stats = ((deck.get("deck_stats") or {}).get("numeric_stats") or {})
        if "avg_text_region_ratio" in numeric_stats:
            text_ratios.append(float(numeric_stats["avg_text_region_ratio"]))
        if "avg_brightness" in numeric_stats:
            brightness.append(float(numeric_stats["avg_brightness"]))
        if "avg_colorfulness" in numeric_stats:
            colorfulness.append(float(numeric_stats["avg_colorfulness"]))

    slide_count_mean = mean([float(v) for v in slide_counts])
    slide_count_std = stdev([float(v) for v in slide_counts])
    slide_count_low = int(round(max(min(slide_counts, default=0), slide_count_mean - max(1.0, slide_count_std)))) if slide_counts else 0
    slide_count_high = int(round(min(max(slide_counts, default=0), slide_count_mean + max(1.0, slide_count_std)))) if slide_counts else 0
    if slide_counts and slide_count_low >= slide_count_high:
        center = int(round(slide_count_mean))
        slide_count_low = max(min(slide_counts), center - 1)
        slide_count_high = min(max(slide_counts), center + 1)

    avg_text_density_proxy = round(mean(text_ratios), 4)
    text_density_proxy_std = round(stdev(text_ratios), 4)

    return {
        "target_slide_count": round(slide_count_mean, 2),
        "slide_count_std": round(slide_count_std, 2),
        "slide_count_range": [slide_count_low, slide_count_high] if slide_counts else [],
        "target_avg_text_density_proxy": avg_text_density_proxy,
        "text_density_proxy_std": text_density_proxy_std,
        "target_text_density_proxy_range": [
            round(max(0.0, avg_text_density_proxy - text_density_proxy_std), 4),
            round(avg_text_density_proxy + text_density_proxy_std, 4),
        ] if text_ratios else [],
        "avg_brightness_proxy": round(mean(brightness), 4),
        "avg_colorfulness_proxy": round(mean(colorfulness), 4),
    }


def _coerce_float(raw_value: Any, *, low: float, high: float) -> float | None:
    if raw_value is None:
        return None
    try:
        value = float(raw_value)
    except Exception:
        return None
    if math.isnan(value) or math.isinf(value):
        return None
    return round(clamp(value, low, high), 4)


def build_fallback_numeric_targets(
    planning_preferences: dict[str, Any],
    deterministic_numeric_preferences: dict[str, Any],
) -> dict[str, Any]:
    bullet_map = {"low": 2.0, "medium": 3.5, "high": 4.8}
    word_map = {"low": 16.0, "medium": 26.0, "high": 36.0}
    usage_map = {"low": 0.18, "medium": 0.35, "high": 0.55}
    splitting_map = {"coarse": 1.3, "balanced": 2.0, "fine_grained": 2.8}

    layout_bias = list(planning_preferences.get("layout_bias") or [])
    bullet_pref = str(planning_preferences.get("bullet_density_preference", "unknown"))
    text_pref = str(planning_preferences.get("text_density_preference", "unknown"))
    figure_pref = str(planning_preferences.get("figure_usage_preference", "unknown"))
    table_pref = str(planning_preferences.get("table_usage_preference", "unknown"))
    formula_pref = str(planning_preferences.get("formula_usage_preference", "unknown"))
    split_pref = str(planning_preferences.get("section_splitting_preference", "unknown"))

    fallback: dict[str, Any] = {
        "target_avg_bullets_per_slide": bullet_map.get(bullet_pref),
        "target_avg_words_per_slide": word_map.get(text_pref),
        "target_fraction_figure_slides": usage_map.get(figure_pref),
        "target_fraction_table_slides": usage_map.get(table_pref),
        "target_fraction_formula_slides": usage_map.get(formula_pref),
        "target_avg_slides_per_section": splitting_map.get(split_pref),
    }

    if "text_only" in layout_bias:
        fallback["target_fraction_text_only_slides"] = 0.35
    if "multi_visual" in layout_bias:
        fallback["target_fraction_multi_visual_slides"] = 0.32
    if "formula_capable" in layout_bias:
        fallback["target_fraction_formula_capable_slides"] = max(
            usage_map.get(formula_pref, 0.22),
            0.28,
        )
    if "image_right" in layout_bias:
        fallback["target_fraction_image_right_slides"] = 0.2
    if "image_left" in layout_bias:
        fallback["target_fraction_image_left_slides"] = 0.14
    if "image_top" in layout_bias:
        fallback["target_fraction_image_top_slides"] = 0.12

    text_proxy_range = deterministic_numeric_preferences.get("target_text_density_proxy_range") or []
    if text_proxy_range:
        midpoint = sum(text_proxy_range) / 2.0
        fallback["target_avg_words_per_slide"] = round(
            clamp(8.0 + midpoint * 120.0, 10.0, 48.0),
            2,
        )

    return fallback


def sanitize_numeric_preferences(
    raw_numeric_preferences: dict[str, Any] | None,
    planning_preferences: dict[str, Any],
    deterministic_numeric_preferences: dict[str, Any],
) -> dict[str, Any]:
    raw_numeric_preferences = raw_numeric_preferences or {}
    sanitized = dict(deterministic_numeric_preferences)
    fallback_numeric = build_fallback_numeric_targets(
        planning_preferences,
        deterministic_numeric_preferences,
    )

    scalar_specs = {
        "target_avg_bullets_per_slide": (1.0, 6.0),
        "target_avg_words_per_slide": (8.0, 55.0),
        "target_fraction_figure_slides": (0.0, 1.0),
        "target_fraction_table_slides": (0.0, 1.0),
        "target_fraction_formula_slides": (0.0, 1.0),
        "target_fraction_text_only_slides": (0.0, 1.0),
        "target_fraction_multi_visual_slides": (0.0, 1.0),
        "target_fraction_formula_capable_slides": (0.0, 1.0),
        "target_fraction_image_right_slides": (0.0, 1.0),
        "target_fraction_image_left_slides": (0.0, 1.0),
        "target_fraction_image_top_slides": (0.0, 1.0),
        "target_avg_slides_per_section": (1.0, 4.0),
    }

    for key, (low, high) in scalar_specs.items():
        raw_value = _coerce_float(raw_numeric_preferences.get(key), low=low, high=high)
        fallback_value = fallback_numeric.get(key)
        if raw_value is not None:
            sanitized[key] = raw_value
        elif fallback_value is not None:
            sanitized[key] = round(float(fallback_value), 4)

    raw_range = raw_numeric_preferences.get("slide_count_range")
    if isinstance(raw_range, list) and len(raw_range) == 2:
        low = _coerce_float(raw_range[0], low=1.0, high=200.0)
        high = _coerce_float(raw_range[1], low=1.0, high=200.0)
        if low is not None and high is not None:
            low_i = int(round(min(low, high)))
            high_i = int(round(max(low, high)))
            if low_i < high_i:
                sanitized["slide_count_range"] = [low_i, high_i]

    if "target_slide_count" in sanitized and sanitized.get("slide_count_range"):
        low_i, high_i = sanitized["slide_count_range"]
        sanitized["target_slide_count"] = round(
            clamp(float(sanitized["target_slide_count"]), float(low_i), float(high_i)),
            2,
        )

    return sanitized


def sanitize_profile(
    raw_profile: dict[str, Any],
    author_id: str,
    paper_ids: list[str],
    max_papers: int,
    deck_evidence: list[dict[str, Any]],
) -> dict[str, Any]:
    planning = raw_profile.get("planning_preferences", {}) or {}
    structure = planning.get("structure_preferences", {}) or {}
    layout_bias = planning.get("layout_bias", []) or []
    typical_section_categories = planning.get("typical_section_categories", []) or []

    allowed_preference_values = {"low", "medium", "high", "unknown"}
    allowed_split_values = {"coarse", "balanced", "fine_grained", "unknown"}

    def pref(key: str) -> str:
        value = str(planning.get(key, "unknown")).strip().lower()
        return value if value in allowed_preference_values else "unknown"

    def split_pref(key: str) -> str:
        value = str(planning.get(key, "unknown")).strip().lower()
        return value if value in allowed_split_values else "unknown"

    def structure_pref(key: str) -> str:
        value = str(structure.get(key, "unknown")).strip().lower()
        return value if value in allowed_preference_values else "unknown"

    cleaned_layout_bias = []
    for item in layout_bias:
        value = str(item).strip()
        if value in LAYOUT_BIAS_VALUES and value not in cleaned_layout_bias:
            cleaned_layout_bias.append(value)

    cleaned_section_categories = []
    for item in typical_section_categories:
        value = str(item).strip()
        if value in SECTION_CATEGORY_VALUES and value not in cleaned_section_categories:
            cleaned_section_categories.append(value)

    notes = raw_profile.get("evidence_summary", {}).get("notes", "")
    deterministic_numeric_preferences = build_numeric_preferences(deck_evidence)

    cleaned_planning_preferences = {
        "section_splitting_preference": split_pref("section_splitting_preference"),
        "bullet_density_preference": pref("bullet_density_preference"),
        "text_density_preference": pref("text_density_preference"),
        "visual_density_preference": pref("visual_density_preference"),
        "figure_usage_preference": pref("figure_usage_preference"),
        "table_usage_preference": pref("table_usage_preference"),
        "formula_usage_preference": pref("formula_usage_preference"),
        "layout_bias": cleaned_layout_bias,
        "typical_section_categories": cleaned_section_categories,
        "structure_preferences": {
            "prefers_agenda_slide": structure_pref("prefers_agenda_slide"),
            "prefers_takeaway_slide": structure_pref("prefers_takeaway_slide"),
            "prefers_multi_slide_method_section": structure_pref("prefers_multi_slide_method_section"),
            "prefers_multi_slide_results_section": structure_pref("prefers_multi_slide_results_section"),
        },
    }
    numeric_preferences = sanitize_numeric_preferences(
        raw_profile.get("numeric_preferences"),
        cleaned_planning_preferences,
        deterministic_numeric_preferences,
    )

    return {
        "author_id": author_id,
        "profile_version": 4,
        "distilled_from": {
            "paper_count": len(paper_ids),
            "paper_ids": paper_ids,
            "deck_sample_policy": {
                "max_papers": max_papers,
                "slides_per_deck": "representative",
            },
        },
        "planning_preferences": cleaned_planning_preferences,
        "numeric_preferences": numeric_preferences,
        "evidence_summary": {
            "notes": str(notes).strip() or "No summary provided.",
        },
    }


def select_papers_for_author(
    author_id: str,
    paper_author_rows: list[dict[str, str]],
    paper_rows: list[dict[str, str]],
    max_papers: int,
    *,
    exclude_paper_ids: set[str] | None = None,
    exclude_pdf_paths: set[str] | None = None,
) -> list[dict[str, Any]]:
    exclude_paper_ids = exclude_paper_ids or set()
    exclude_pdf_paths = exclude_pdf_paths or set()
    paper_ids_for_author = {
        row["paper_id"]
        for row in paper_author_rows
        if row["author_id"].strip() == author_id
    }
    candidates = []
    for row in paper_rows:
        if row["paper_id"] not in paper_ids_for_author:
            continue
        if row["paper_id"] in exclude_paper_ids:
            continue
        raw_dir = resolve_repo_path(row["raw_dir"])
        if not raw_dir.exists():
            continue
        normalized_pdf_path = normalize_path_for_compare(row["paper_pdf_path"])
        if normalized_pdf_path in exclude_pdf_paths:
            continue
        candidates.append(
            {
                "paper_id": row["paper_id"],
                "paper_title": row["paper_title"],
                "record_id": row["record_id"],
                "split": row["split"],
                "raw_dir": raw_dir,
                "paper_pdf_path": str(resolve_repo_path(row["paper_pdf_path"])),
                "slide_image_count": parse_int(row.get("slide_image_count", "0")),
            }
        )

    candidates.sort(key=lambda row: (-row["slide_image_count"], row["paper_id"]))
    return candidates[:max_papers]


def build_author_metadata(author_row: dict[str, str], selected_papers: list[dict[str, Any]], max_papers: int) -> dict[str, Any]:
    return {
        "author_id": author_row["author_id"],
        "display_name": author_row["display_name"],
        "paper_count": len(selected_papers),
        "paper_ids": [paper["paper_id"] for paper in selected_papers],
        "deck_sample_policy": {
            "max_papers": max_papers,
            "slides_per_deck": "representative",
        },
    }


def build_deck_evidence(selected_papers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    deck_evidence = []
    for paper in selected_papers:
        slide_paths = sorted(
            [path for path in paper["raw_dir"].iterdir() if path.suffix.lower() in SLIDE_EXTENSIONS],
            key=numeric_slide_sort_key,
        )
        slide_metrics = [compute_slide_metrics(path) for path in slide_paths]
        sampled_indices = sample_representative_slide_indices(slide_metrics)
        aggregate = aggregate_deck_metrics(slide_metrics)

        sampled_slides = []
        for index in sampled_indices:
            metric = slide_metrics[index]
            sampled_slides.append(
                {
                    "slide_index": index,
                    "slide_path": str(slide_paths[index]),
                    "metrics": metric,
                }
            )

        deck_evidence.append(
            {
                "paper_id": paper["paper_id"],
                "paper_title": paper["paper_title"],
                "slide_count": len(slide_paths),
                "raw_dir": str(paper["raw_dir"]),
                "sampled_slide_indices": sampled_indices,
                "sampled_slides": sampled_slides,
                "deck_stats": aggregate,
            }
        )
    return deck_evidence


def render_prompt(prompt_path: Path, author_metadata: dict[str, Any], deck_evidence: list[dict[str, Any]]) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    rendered = template.render(
        author_metadata=author_metadata,
        deck_evidence=deck_evidence,
        layout_bias_values=LAYOUT_BIAS_VALUES,
        section_category_values=SECTION_CATEGORY_VALUES,
        preference_labels=PREFERENCE_LABELS,
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": rendered,
    }


def call_distiller_model(model_name: str, system_prompt: str, user_prompt: str, deck_evidence: list[dict[str, Any]]) -> dict[str, Any]:
    client = build_openai_client()
    resolved_model_name = resolve_direct_model_name(model_name)

    content: list[dict[str, Any]] = [
        {"type": "text", "text": user_prompt},
    ]
    for deck in deck_evidence:
        for sampled in deck["sampled_slides"]:
            image_path = Path(sampled["slide_path"])
            content.append(
                {
                    "type": "text",
                    "text": (
                        f"Paper {deck['paper_id']} / slide {sampled['slide_index']}\n"
                        f"Metrics: {json.dumps(sampled['metrics'], ensure_ascii=False)}"
                    ),
                }
            )
            content.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": encode_image_data_uri(image_path),
                    },
                }
            )

    request_kwargs = {
        "model": resolved_model_name,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
        "temperature": 0.2,
    }
    if "gpt-5" in resolved_model_name.lower():
        request_kwargs["max_completion_tokens"] = 1400
    else:
        request_kwargs["max_tokens"] = 1400

    print(
        f"[preferences] Sending distiller request to model={resolved_model_name} "
        f"with {len(deck_evidence)} reference deck(s)",
        flush=True,
    )
    response = client.chat.completions.create(**request_kwargs)
    print("[preferences] Distiller response received", flush=True)
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def distill_author_profile(
    author_id: str,
    *,
    authors_csv: Path = DEFAULT_AUTHORS_CSV,
    paper_authors_csv: Path = DEFAULT_PAPER_AUTHORS_CSV,
    papers_csv: Path = DEFAULT_PAPERS_CSV,
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    prompt_path: Path = DEFAULT_PROMPT_PATH,
    max_papers: int = 5,
    model: str = "4o-mini",
    force_refresh: bool = False,
    exclude_paper_ids: set[str] | None = None,
    exclude_pdf_paths: set[str] | None = None,
) -> dict[str, Any]:
    """Build or load a distilled author profile for planner personalization."""
    output_dir.mkdir(parents=True, exist_ok=True)
    profile_path = output_dir / f"{author_id}.json"

    if profile_path.exists() and not force_refresh:
        cached_profile = json.loads(profile_path.read_text(encoding="utf-8"))
        cached_numeric = dict(cached_profile.get("numeric_preferences") or {})
        if (
            int(cached_profile.get("profile_version", 0) or 0) >= 4
            and "target_slide_count" in cached_numeric
            and "target_avg_words_per_slide" in cached_numeric
        ):
            return cached_profile

    load_dotenv(REPO_ROOT / ".env")

    authors_rows = load_csv_rows(authors_csv)
    paper_author_rows = load_csv_rows(paper_authors_csv)
    paper_rows = load_csv_rows(papers_csv)

    author_row = next((row for row in authors_rows if row["author_id"].strip() == author_id), None)
    if author_row is None:
        raise ValueError(f"Unknown author_id: {author_id}")

    print(f"[preferences] Selecting source decks for author_id={author_id}", flush=True)
    selected_papers = select_papers_for_author(
        author_id,
        paper_author_rows,
        paper_rows,
        max_papers,
        exclude_paper_ids=exclude_paper_ids,
        exclude_pdf_paths=exclude_pdf_paths,
    )
    if not selected_papers:
        raise ValueError(f"No eligible papers with raw decks found for author_id: {author_id}")

    author_metadata = build_author_metadata(author_row, selected_papers, max_papers)
    print(
        f"[preferences] Using papers: {', '.join(author_metadata['paper_ids'])}",
        flush=True,
    )
    print("[preferences] Building deck evidence", flush=True)
    deck_evidence = build_deck_evidence(selected_papers)
    print("[preferences] Deck evidence ready", flush=True)

    rendered_prompt = render_prompt(prompt_path, author_metadata, deck_evidence)
    print("[preferences] Distiller prompt rendered", flush=True)

    input_bundle_path = output_dir / f"{author_id}.input.json"
    input_bundle = {
        "author_metadata": author_metadata,
        "deck_evidence": deck_evidence,
        "prompt_preview": rendered_prompt["user_prompt"],
    }
    input_bundle_path.write_text(json.dumps(input_bundle, indent=2, ensure_ascii=False), encoding="utf-8")

    raw_profile = call_distiller_model(
        model_name=model,
        system_prompt=rendered_prompt["system_prompt"],
        user_prompt=rendered_prompt["user_prompt"],
        deck_evidence=deck_evidence,
    )
    profile = sanitize_profile(
        raw_profile=raw_profile,
        author_id=author_id,
        paper_ids=author_metadata["paper_ids"],
        max_papers=max_papers,
        deck_evidence=deck_evidence,
    )
    profile_path.write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")
    return profile


def main() -> None:
    parser = argparse.ArgumentParser(description="Distill standalone author planning profiles from prior decks.")
    parser.add_argument("--author-id", required=True, help="Canonical author_id from Capstone/author_tables/authors.csv")
    parser.add_argument("--authors-csv", type=Path, default=DEFAULT_AUTHORS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--max-papers", type=int, default=5)
    parser.add_argument("--model", default="4o-mini", help="Model name or deployment alias. Default: 4o-mini")
    parser.add_argument(
        "--dry-run-metadata-only",
        action="store_true",
        help="Assemble evidence and save the input bundle without calling the model.",
    )
    args = parser.parse_args()

    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.dry_run_metadata_only:
        load_dotenv(REPO_ROOT / ".env")
        authors_rows = load_csv_rows(args.authors_csv)
        paper_author_rows = load_csv_rows(args.paper_authors_csv)
        paper_rows = load_csv_rows(args.papers_csv)
        author_row = next((row for row in authors_rows if row["author_id"].strip() == args.author_id), None)
        if author_row is None:
            raise SystemExit(f"Unknown author_id: {args.author_id}")
        selected_papers = select_papers_for_author(args.author_id, paper_author_rows, paper_rows, args.max_papers)
        if not selected_papers:
            raise SystemExit(f"No eligible papers with raw decks found for author_id: {args.author_id}")
        author_metadata = build_author_metadata(author_row, selected_papers, args.max_papers)
        deck_evidence = build_deck_evidence(selected_papers)
        rendered_prompt = render_prompt(args.prompt_path, author_metadata, deck_evidence)
        input_bundle_path = output_dir / f"{args.author_id}.input.json"
        input_bundle = {
            "author_metadata": author_metadata,
            "deck_evidence": deck_evidence,
            "prompt_preview": rendered_prompt["user_prompt"],
        }
        input_bundle_path.write_text(json.dumps(input_bundle, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Saved assembled evidence bundle to {input_bundle_path}")
        print(f"Selected papers: {', '.join(author_metadata['paper_ids'])}")
        return

    profile = distill_author_profile(
        args.author_id,
        authors_csv=args.authors_csv,
        paper_authors_csv=args.paper_authors_csv,
        papers_csv=args.papers_csv,
        output_dir=args.output_dir,
        prompt_path=args.prompt_path,
        max_papers=args.max_papers,
        model=args.model,
    )
    profile_path = output_dir / f"{args.author_id}.json"
    input_bundle_path = output_dir / f"{args.author_id}.input.json"
    print(f"Saved author profile to {profile_path}")
    print(f"Saved assembled evidence bundle to {input_bundle_path}")


if __name__ == "__main__":
    main()
