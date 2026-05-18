#!/usr/bin/env python3
"""Experimental paper-slide pair guideline distiller for SlideGen.

This module keeps pair-derived personalization separate from the existing
author-profile JSONs. It derives transformation-style guidance from prior
paper/deck pairs for a given author, then bundles those guidelines into a
target-specific context artifact that downstream pipelines can consume.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import fitz
import yaml
from dotenv import load_dotenv
from jinja2 import Environment, StrictUndefined

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from slidegen_openai_utils import build_openai_client, resolve_direct_model_name
from Capstone.preference_distill import (
    DEFAULT_AUTHORS_CSV,
    DEFAULT_PAPER_AUTHORS_CSV,
    DEFAULT_PAPERS_CSV,
    LAYOUT_BIAS_VALUES,
    PREFERENCE_LABELS,
    SECTION_CATEGORY_VALUES,
    SLIDE_EXTENSIONS,
    aggregate_deck_metrics,
    compute_slide_metrics,
    encode_image_data_uri,
    extract_json_object,
    load_csv_rows,
    normalize_path_for_compare,
    numeric_slide_sort_key,
    resolve_repo_path,
    sample_representative_slide_indices,
    select_papers_for_author,
)


DEFAULT_PAIR_GUIDELINE_DIR = REPO_ROOT / "Capstone" / "pair_guidelines"
DEFAULT_CONTEXT_DIR = REPO_ROOT / "Capstone" / "pair_guideline_contexts"
DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "pair_guideline_distiller.yaml"

ALLOWED_PREF_VALUES = {"low", "medium", "high", "unknown"}
ALLOWED_SPLIT_VALUES = {"coarse", "balanced", "fine_grained", "unknown"}


def output_key_from_paper_id(paper_id: str | None) -> str | None:
    if not paper_id:
        return None
    key = str(paper_id).strip().replace(":", "_")
    key = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in key)
    key = key.strip("_")
    return key or None


def infer_output_key_from_paper_path(paper_path: str) -> str:
    path = Path(paper_path)
    if path.parent.name and path.parent.name.isdigit():
        split = path.parent.parent.name or "paper"
        return f"{split}_{path.parent.name}"
    stem = path.stem.replace(" ", "_")
    return stem or "paper"


def normalize_text_block(text: str) -> str:
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def looks_like_heading(line: str) -> bool:
    candidate = re.sub(r"\s+", " ", line.strip())
    if not candidate or len(candidate) < 3 or len(candidate) > 80:
        return False
    if candidate.lower() in {"abstract", "references", "appendix"}:
        return True
    if len(candidate.split()) > 12:
        return False
    if re.fullmatch(r"\d+", candidate):
        return False
    if candidate.endswith("."):
        return False
    if re.match(r"^\d+(\.\d+)*\s+[A-Z]", candidate):
        return True
    alpha_ratio = sum(ch.isalpha() for ch in candidate) / max(1, len(candidate))
    if alpha_ratio < 0.6:
        return False
    if candidate.isupper():
        return True
    return candidate == candidate.title()


def guess_title(lines: list[str]) -> str:
    for line in lines[:20]:
        candidate = re.sub(r"\s+", " ", line.strip())
        if not candidate:
            continue
        if len(candidate.split()) > 24:
            continue
        if candidate.lower() in {"abstract", "introduction"}:
            continue
        if candidate.startswith("arXiv") or candidate.startswith("Proceedings"):
            continue
        return candidate
    return lines[0].strip() if lines else ""


def extract_lightweight_paper_context(pdf_path: Path, *, max_pages: int = 4, max_chars: int = 12000) -> dict[str, Any]:
    pdf_path = resolve_repo_path(str(pdf_path))
    doc = fitz.open(str(pdf_path))
    try:
        page_texts: list[str] = []
        heading_candidates: list[str] = []
        for page_index in range(min(max_pages, len(doc))):
            raw_text = doc.load_page(page_index).get_text("text")
            normalized = normalize_text_block(raw_text)
            if not normalized:
                continue
            page_texts.append(normalized)
            for line in normalized.splitlines():
                if looks_like_heading(line):
                    heading = re.sub(r"\s+", " ", line.strip())
                    if heading not in heading_candidates:
                        heading_candidates.append(heading)

        combined_text = "\n\n".join(page_texts)
        lines = [line.strip() for line in combined_text.splitlines() if line.strip()]
        title = guess_title(lines)
        return {
            "paper_pdf_path": str(pdf_path),
            "page_count": len(doc),
            "title_guess": title,
            "section_heading_candidates": heading_candidates[:12],
            "excerpt": combined_text[:max_chars],
        }
    finally:
        doc.close()


def build_pair_reference_evidence(reference_paper: dict[str, Any]) -> dict[str, Any]:
    slide_paths = sorted(
        [path for path in reference_paper["raw_dir"].iterdir() if path.suffix.lower() in SLIDE_EXTENSIONS],
        key=numeric_slide_sort_key,
    )
    slide_metrics = [compute_slide_metrics(path) for path in slide_paths]
    sampled_indices = sample_representative_slide_indices(slide_metrics, max_samples=4)
    sampled_slides = []
    for index in sampled_indices:
        sampled_slides.append(
            {
                "slide_index": index,
                "slide_path": str(slide_paths[index]),
                "metrics": slide_metrics[index],
            }
        )

    return {
        "paper_id": reference_paper["paper_id"],
        "paper_title": reference_paper["paper_title"],
        "paper_pdf_path": reference_paper["paper_pdf_path"],
        "raw_dir": str(reference_paper["raw_dir"]),
        "slide_count": len(slide_paths),
        "sampled_slide_indices": sampled_indices,
        "sampled_slides": sampled_slides,
        "deck_stats": aggregate_deck_metrics(slide_metrics),
        "paper_context": extract_lightweight_paper_context(Path(reference_paper["paper_pdf_path"])),
    }


def render_pair_prompt(prompt_path: Path, author_id: str, reference_pair: dict[str, Any]) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    rendered = template.render(
        author_id=author_id,
        reference_pair=reference_pair,
        layout_bias_values=LAYOUT_BIAS_VALUES,
        section_category_values=SECTION_CATEGORY_VALUES,
        preference_labels=PREFERENCE_LABELS,
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": rendered,
    }


def call_pair_guideline_model(
    *,
    model_name: str,
    system_prompt: str,
    user_prompt: str,
    reference_pair: dict[str, Any],
) -> dict[str, Any]:
    client = build_openai_client()
    resolved_model_name = resolve_direct_model_name(model_name)

    content: list[dict[str, Any]] = [{"type": "text", "text": user_prompt}]
    for sampled in reference_pair["sampled_slides"]:
        slide_path = Path(sampled["slide_path"])
        content.append(
            {
                "type": "text",
                "text": (
                    f"Reference pair {reference_pair['paper_id']} / slide {sampled['slide_index']}\n"
                    f"Metrics: {json.dumps(sampled['metrics'], ensure_ascii=False)}"
                ),
            }
        )
        content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": encode_image_data_uri(slide_path),
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
        request_kwargs["max_completion_tokens"] = 1600
    else:
        request_kwargs["max_tokens"] = 1600

    print(
        f"[pair-guidelines] Sending pair distiller request for {reference_pair['paper_id']} "
        f"to model={resolved_model_name}",
        flush=True,
    )
    response = client.chat.completions.create(**request_kwargs)
    print(f"[pair-guidelines] Pair distiller response received for {reference_pair['paper_id']}", flush=True)
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def sanitize_pref_value(raw_value: Any) -> str:
    value = str(raw_value or "unknown").strip().lower()
    return value if value in ALLOWED_PREF_VALUES else "unknown"


def sanitize_split_value(raw_value: Any) -> str:
    value = str(raw_value or "unknown").strip().lower()
    return value if value in ALLOWED_SPLIT_VALUES else "unknown"


def sanitize_string_list(raw_values: Any, *, limit: int = 8) -> list[str]:
    if not isinstance(raw_values, list):
        return []
    cleaned: list[str] = []
    for raw_value in raw_values:
        text = re.sub(r"\s+", " ", str(raw_value or "").strip())
        if not text or text in cleaned:
            continue
        cleaned.append(text[:220])
        if len(cleaned) >= limit:
            break
    return cleaned


def sanitize_layout_bias(raw_values: Any) -> list[str]:
    cleaned: list[str] = []
    if not isinstance(raw_values, list):
        return cleaned
    for raw_value in raw_values:
        text = str(raw_value or "").strip()
        if text in LAYOUT_BIAS_VALUES and text not in cleaned:
            cleaned.append(text)
    return cleaned


def sanitize_pair_guideline(raw_guideline: dict[str, Any], reference_pair: dict[str, Any], author_id: str) -> dict[str, Any]:
    presentation = raw_guideline.get("presentation_guidelines") or {}
    planning_hints = raw_guideline.get("planning_hints") or {}
    structure_preferences = planning_hints.get("structure_preferences") or {}
    evidence_summary = raw_guideline.get("evidence_summary") or {}

    return {
        "author_id": author_id,
        "paper_id": reference_pair["paper_id"],
        "guideline_version": 1,
        "derived_from": {
            "paper_pdf_path": reference_pair["paper_pdf_path"],
            "raw_dir": reference_pair["raw_dir"],
            "slide_count": reference_pair["slide_count"],
            "sampled_slide_indices": reference_pair["sampled_slide_indices"],
        },
        "presentation_guidelines": {
            "narrative_flow_preferences": sanitize_string_list(presentation.get("narrative_flow_preferences")),
            "section_emphasis_patterns": sanitize_string_list(presentation.get("section_emphasis_patterns")),
            "content_style_preferences": sanitize_string_list(presentation.get("content_style_preferences")),
            "compression_patterns": sanitize_string_list(presentation.get("compression_patterns")),
            "visual_usage_patterns": sanitize_string_list(presentation.get("visual_usage_patterns")),
            "signature_choices": sanitize_string_list(presentation.get("signature_choices")),
        },
        "planning_hints": {
            "section_splitting_preference": sanitize_split_value(planning_hints.get("section_splitting_preference")),
            "bullet_density_preference": sanitize_pref_value(planning_hints.get("bullet_density_preference")),
            "text_density_preference": sanitize_pref_value(planning_hints.get("text_density_preference")),
            "visual_density_preference": sanitize_pref_value(planning_hints.get("visual_density_preference")),
            "layout_bias": sanitize_layout_bias(planning_hints.get("layout_bias")),
            "structure_preferences": {
                "prefers_takeaway_slide": sanitize_pref_value(structure_preferences.get("prefers_takeaway_slide")),
                "prefers_multi_slide_method_section": sanitize_pref_value(
                    structure_preferences.get("prefers_multi_slide_method_section")
                ),
                "prefers_multi_slide_results_section": sanitize_pref_value(
                    structure_preferences.get("prefers_multi_slide_results_section")
                ),
            },
        },
        "evidence_summary": {
            "notes": str(evidence_summary.get("notes") or "").strip() or "No summary provided.",
            "notable_transformations": sanitize_string_list(evidence_summary.get("notable_transformations")),
        },
    }


def distill_reference_pair_guideline(
    *,
    author_id: str,
    reference_paper: dict[str, Any],
    model: str,
    prompt_path: Path,
    cache_dir: Path,
    force_refresh: bool,
) -> dict[str, Any]:
    author_cache_dir = cache_dir / author_id
    author_cache_dir.mkdir(parents=True, exist_ok=True)
    pair_key = output_key_from_paper_id(reference_paper["paper_id"]) or infer_output_key_from_paper_path(
        reference_paper["paper_pdf_path"]
    )
    guideline_path = author_cache_dir / f"{pair_key}.json"

    if guideline_path.exists() and not force_refresh:
        return json.loads(guideline_path.read_text(encoding="utf-8"))

    print(f"[pair-guidelines] Building evidence for reference pair {reference_paper['paper_id']}", flush=True)
    reference_pair = build_pair_reference_evidence(reference_paper)
    rendered_prompt = render_pair_prompt(prompt_path, author_id, reference_pair)

    debug_input_path = author_cache_dir / f"{pair_key}.input.json"
    debug_input = {
        "author_id": author_id,
        "reference_pair": reference_pair,
        "prompt_preview": rendered_prompt["user_prompt"],
    }
    debug_input_path.write_text(json.dumps(debug_input, indent=2, ensure_ascii=False), encoding="utf-8")

    raw_guideline = call_pair_guideline_model(
        model_name=model,
        system_prompt=rendered_prompt["system_prompt"],
        user_prompt=rendered_prompt["user_prompt"],
        reference_pair=reference_pair,
    )
    guideline = sanitize_pair_guideline(raw_guideline, reference_pair, author_id)
    guideline_path.write_text(json.dumps(guideline, indent=2, ensure_ascii=False), encoding="utf-8")
    return guideline


def build_pair_guideline_context(
    author_id: str,
    *,
    target_paper_path: str,
    target_paper_id: str | None = None,
    authors_csv: Path = DEFAULT_AUTHORS_CSV,
    paper_authors_csv: Path = DEFAULT_PAPER_AUTHORS_CSV,
    papers_csv: Path = DEFAULT_PAPERS_CSV,
    pair_guideline_dir: Path = DEFAULT_PAIR_GUIDELINE_DIR,
    context_dir: Path = DEFAULT_CONTEXT_DIR,
    prompt_path: Path = DEFAULT_PROMPT_PATH,
    max_pairs: int = 2,
    candidate_pool: int = 5,
    model: str = "4o-mini",
    force_refresh: bool = False,
) -> dict[str, Any]:
    """Build or load a target-specific pair-guideline context."""
    load_dotenv(REPO_ROOT / ".env")
    context_dir.mkdir(parents=True, exist_ok=True)
    author_context_dir = context_dir / author_id
    author_context_dir.mkdir(parents=True, exist_ok=True)

    target_key = output_key_from_paper_id(target_paper_id) or infer_output_key_from_paper_path(target_paper_path)
    context_path = author_context_dir / f"{target_key}.json"

    if context_path.exists() and not force_refresh:
        return json.loads(context_path.read_text(encoding="utf-8"))

    authors_rows = load_csv_rows(authors_csv)
    paper_author_rows = load_csv_rows(paper_authors_csv)
    paper_rows = load_csv_rows(papers_csv)

    author_row = next((row for row in authors_rows if row["author_id"].strip() == author_id), None)
    if author_row is None:
        raise ValueError(f"Unknown author_id: {author_id}")

    exclude_pdf_paths = {normalize_path_for_compare(target_paper_path)}
    exclude_paper_ids = {target_paper_id} if target_paper_id else set()

    print(f"[pair-guidelines] Selecting reference paper/deck pairs for author_id={author_id}", flush=True)
    candidates = select_papers_for_author(
        author_id,
        paper_author_rows,
        paper_rows,
        max_papers=max(candidate_pool, max_pairs),
        exclude_paper_ids=exclude_paper_ids,
        exclude_pdf_paths=exclude_pdf_paths,
    )
    if not candidates:
        raise ValueError(f"No eligible paper/deck pairs found for author_id: {author_id}")

    reference_papers = candidates[:max_pairs]
    print(
        f"[pair-guidelines] Using reference pairs: {', '.join(paper['paper_id'] for paper in reference_papers)}",
        flush=True,
    )

    reference_guidelines = []
    for reference_paper in reference_papers:
        guideline = distill_reference_pair_guideline(
            author_id=author_id,
            reference_paper=reference_paper,
            model=model,
            prompt_path=prompt_path,
            cache_dir=pair_guideline_dir,
            force_refresh=force_refresh,
        )
        reference_guidelines.append(guideline)

    target_paper_context = extract_lightweight_paper_context(Path(target_paper_path))
    context = {
        "author_id": author_id,
        "context_version": 1,
        "preference_source": "paper_slide_pairs",
        "author_display_name": author_row["display_name"],
        "target_paper_id": target_paper_id,
        "target_paper_path": str(resolve_repo_path(target_paper_path)),
        "target_paper_context": target_paper_context,
        "selection_policy": {
            "strategy": "top_slide_count",
            "candidate_pool": max(candidate_pool, max_pairs),
            "max_pairs": max_pairs,
        },
        "reference_pair_ids": [paper["paper_id"] for paper in reference_papers],
        "reference_pairs": reference_guidelines,
    }
    context_path.write_text(json.dumps(context, indent=2, ensure_ascii=False), encoding="utf-8")
    return context


def main() -> None:
    parser = argparse.ArgumentParser(description="Build experimental pair-guideline contexts from paper/deck pairs.")
    parser.add_argument("--author-id", required=True, help="Canonical author_id from Capstone/author_tables/authors.csv")
    parser.add_argument("--target-paper-path", required=True, help="PDF path for the paper you want to personalize.")
    parser.add_argument("--target-paper-id", default=None, help="Optional paper_id if known.")
    parser.add_argument("--authors-csv", type=Path, default=DEFAULT_AUTHORS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--model", default="4o-mini", help="Model name or deployment alias.")
    parser.add_argument("--max-pairs", type=int, default=2, help="How many prior paper/deck pairs to use.")
    parser.add_argument(
        "--candidate-pool",
        type=int,
        default=5,
        help="How many candidate prior pairs to consider before selecting the top max-pairs.",
    )
    parser.add_argument("--pair-guideline-dir", type=Path, default=DEFAULT_PAIR_GUIDELINE_DIR)
    parser.add_argument("--context-dir", type=Path, default=DEFAULT_CONTEXT_DIR)
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--force-refresh", action="store_true")
    args = parser.parse_args()

    context = build_pair_guideline_context(
        args.author_id,
        target_paper_path=args.target_paper_path,
        target_paper_id=args.target_paper_id,
        authors_csv=args.authors_csv,
        paper_authors_csv=args.paper_authors_csv,
        papers_csv=args.papers_csv,
        pair_guideline_dir=args.pair_guideline_dir,
        context_dir=args.context_dir,
        prompt_path=args.prompt_path,
        max_pairs=args.max_pairs,
        candidate_pool=args.candidate_pool,
        model=args.model,
        force_refresh=args.force_refresh,
    )
    target_key = output_key_from_paper_id(args.target_paper_id) or infer_output_key_from_paper_path(args.target_paper_path)
    context_path = args.context_dir / args.author_id / f"{target_key}.json"
    print(f"Saved pair-guideline context to {context_path}")
    print(f"Reference pairs: {', '.join(context['reference_pair_ids'])}")


if __name__ == "__main__":
    main()
