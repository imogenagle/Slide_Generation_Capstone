#!/usr/bin/env python3
"""SlideTailor-style structure similarity evaluation for SlideGen decks."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

if __package__ in {None, ""}:
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
    from SlideGen.Capstone.slidetailor_eval.common import (
        DEFAULT_CATEGORY_LIST,
        DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES,
        REPO_ROOT,
        _normalize_score,
        add_shared_args,
        call_json_judge,
        collect_slide_images,
        condensed_standard_flow,
        coverage_iou,
        extract_outline_from_presentation_text,
        extract_outline_from_reference_images,
        flow_ngld,
        load_runtime_env,
        render_prompt,
        render_pptx_to_images,
        resolve_output_path,
        standardize_narrative_items,
    )
else:
    from .common import (
        DEFAULT_CATEGORY_LIST,
        DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES,
        REPO_ROOT,
        _normalize_score,
        add_shared_args,
        call_json_judge,
        collect_slide_images,
        condensed_standard_flow,
        coverage_iou,
        extract_outline_from_presentation_text,
        extract_outline_from_reference_images,
        flow_ngld,
        load_runtime_env,
        render_prompt,
        render_pptx_to_images,
        resolve_output_path,
        standardize_narrative_items,
    )


OUTLINE_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "slidetailor_eval" / "outline_extractor.yaml"
REFERENCE_OUTLINE_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "slidetailor_eval" / "reference_outline_from_images.yaml"
STANDARDIZE_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "slidetailor_eval" / "standardize_sections.yaml"
CONTENT_STRUCTURE_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "slidetailor_eval" / "content_structure_similarity.yaml"


def _outline_items_to_strings(outline: dict[str, Any]) -> list[str]:
    values: list[str] = []
    for item in list(outline.get("slide_descriptions") or []):
        title = str(item.get("title") or "").strip()
        description = str(item.get("description") or "").strip()
        category = str(item.get("category_guess") or "").strip()
        text = " | ".join(part for part in (title, description, category) if part)
        if text:
            values.append(text)
    return values


def evaluate_structure_similarity(
    *,
    generated_pptx: Path,
    reference_slide_images: list[Path],
    model: str,
    request_timeout: float,
    verbose: bool,
    categories: list[str],
    max_reference_slides: int,
) -> dict[str, Any]:
    generated_outline = extract_outline_from_presentation_text(
        pptx_path=generated_pptx,
        categories=categories,
        prompt_path=OUTLINE_PROMPT_PATH,
        model=model,
        request_timeout=request_timeout,
        verbose=verbose,
    )
    reference_outline = extract_outline_from_reference_images(
        image_paths=reference_slide_images,
        categories=categories,
        prompt_path=REFERENCE_OUTLINE_PROMPT_PATH,
        model=model,
        request_timeout=request_timeout,
        verbose=verbose,
        max_images=max_reference_slides,
    )
    generated_standardized = standardize_narrative_items(
        narrative_items=_outline_items_to_strings(generated_outline),
        categories=categories,
        prompt_path=STANDARDIZE_PROMPT_PATH,
        model=model,
        request_timeout=request_timeout,
        verbose=verbose,
    )
    reference_standardized = standardize_narrative_items(
        narrative_items=_outline_items_to_strings(reference_outline),
        categories=categories,
        prompt_path=STANDARDIZE_PROMPT_PATH,
        model=model,
        request_timeout=request_timeout,
        verbose=verbose,
    )
    generated_flow = condensed_standard_flow(generated_standardized)
    reference_flow = condensed_standard_flow(reference_standardized)

    prompt = render_prompt(
        CONTENT_STRUCTURE_PROMPT_PATH,
        target_outline_json=json.dumps(generated_outline, ensure_ascii=False, indent=2),
        reference_outline_json=json.dumps(reference_outline, ensure_ascii=False, indent=2),
    )
    judged = call_json_judge(
        model=model,
        system_prompt=prompt["system_prompt"],
        user_prompt=prompt["user_prompt"],
        request_timeout=request_timeout,
        verbose=verbose,
    )

    return {
        "source": "SlideTailor-derived",
        "metric": "structure_similarity",
        "generated_pptx": str(generated_pptx),
        "coverage_iou": coverage_iou(generated_flow, reference_flow),
        "flow_ngld": flow_ngld(generated_flow, reference_flow),
        "content_structure_similarity": _normalize_score(judged.get("score", 0.0)),
        "content_structure_reason": str(judged.get("reason", "")).strip(),
        "categories": categories,
        "generated_outline": generated_outline,
        "reference_outline": reference_outline,
        "generated_standardized_flow": generated_standardized,
        "reference_standardized_flow": reference_standardized,
        "generated_flow_condensed": generated_flow,
        "reference_flow_condensed": reference_flow,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate SlideTailor-style structure similarity.")
    parser.add_argument("--generated-pptx", type=Path, required=True)
    parser.add_argument("--reference-slide-dir", type=Path, required=True)
    parser.add_argument("--max-reference-slides", type=int, default=DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES)
    parser.add_argument("--categories-json", type=Path, default=None, help="Optional JSON file containing category list.")
    add_shared_args(parser, metric_name="structure_similarity")
    args = parser.parse_args()

    load_runtime_env()
    categories = DEFAULT_CATEGORY_LIST
    if args.categories_json:
        categories = list(json.loads(args.categories_json.read_text(encoding="utf-8")))
    reference_slide_images = collect_slide_images(args.reference_slide_dir)
    result = evaluate_structure_similarity(
        generated_pptx=args.generated_pptx,
        reference_slide_images=reference_slide_images,
        model=args.model,
        request_timeout=args.request_timeout,
        verbose=args.verbose,
        categories=categories,
        max_reference_slides=args.max_reference_slides,
    )
    output_path = resolve_output_path("structure_similarity", args.output, args.generated_pptx.stem)
    output_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
