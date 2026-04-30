#!/usr/bin/env python3
"""Batch runner for SlideTailor-derived evaluation metrics on SlideGen decks."""

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
        DEFAULT_MAX_VISION_IMAGES,
        DEFAULT_PAPERS_CSV,
        collect_slide_images,
        load_paper_metadata,
        load_runtime_env,
        normalize_paper_name_from_paper_id,
        render_pptx_to_images,
    )
    from SlideGen.Capstone.slidetailor_eval.evaluate_aesthetic_quality import evaluate_aesthetic_quality
    from SlideGen.Capstone.slidetailor_eval.evaluate_content_informativeness import evaluate_content_informativeness
    from SlideGen.Capstone.slidetailor_eval.evaluate_structure_similarity import evaluate_structure_similarity
    from SlideGen.Capstone.slidetailor_eval.evaluate_template_similarity import evaluate_template_similarity
    from SlideGen.Capstone.slidetailor_eval.paths import metric_output_dir
else:
    from .common import (
        DEFAULT_CATEGORY_LIST,
        DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES,
        DEFAULT_MAX_VISION_IMAGES,
        DEFAULT_PAPERS_CSV,
        collect_slide_images,
        load_paper_metadata,
        load_runtime_env,
        normalize_paper_name_from_paper_id,
        render_pptx_to_images,
    )
    from .evaluate_aesthetic_quality import evaluate_aesthetic_quality
    from .evaluate_content_informativeness import evaluate_content_informativeness
    from .evaluate_structure_similarity import evaluate_structure_similarity
    from .evaluate_template_similarity import evaluate_template_similarity
    from .paths import metric_output_dir


DEFAULT_CONTENTS_DIR = Path(__file__).resolve().parents[2] / "contents"


def find_generated_pptx_dirs(contents_dir: Path, model_name_t: str, model_name_v: str) -> list[tuple[str, Path]]:
    prefix = f"{model_name_t}_{model_name_v}_output_slides"
    results: list[tuple[str, Path]] = []
    for paper_dir in sorted(path for path in contents_dir.iterdir() if path.is_dir()):
        candidates = sorted(
            path
            for path in paper_dir.glob(f"{prefix}*.pptx")
            if not path.name.endswith("_themed.pptx") and not path.name.startswith("~$")
        )
        if not candidates:
            continue
        preferred = None
        for candidate in candidates:
            if candidate.name == f"{prefix}_baseline.pptx":
                preferred = candidate
                break
        if preferred is None:
            for candidate in candidates:
                if candidate.name == f"{prefix}.pptx":
                    preferred = candidate
                    break
        if preferred is None:
            preferred = candidates[0]
        results.append((paper_dir.name, preferred))
    return results


def main() -> None:
    parser = argparse.ArgumentParser(description="Run SlideTailor-derived evaluation metrics on generated decks.")
    parser.add_argument("--contents-dir", type=Path, default=DEFAULT_CONTENTS_DIR)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--model-name-t", default="4o-mini")
    parser.add_argument("--model-name-v", default="4o-mini")
    parser.add_argument("--judge-model", default="gpt-5")
    parser.add_argument("--compute-aesthetic", action="store_true")
    parser.add_argument("--compute-content", action="store_true")
    parser.add_argument("--compute-structure", action="store_true")
    parser.add_argument("--compute-template", action="store_true")
    parser.add_argument("--template-pptx", type=Path, default=None, help="Optional common template PPTX for template similarity.")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    load_runtime_env()
    metadata_by_name = load_paper_metadata(args.papers_csv)
    generated = find_generated_pptx_dirs(args.contents_dir, args.model_name_t, args.model_name_v)
    if args.limit > 0:
        generated = generated[: args.limit]

    if not any((args.compute_aesthetic, args.compute_content, args.compute_structure, args.compute_template)):
        raise SystemExit("Select at least one metric flag.")

    for paper_name, generated_pptx in generated:
        metadata = metadata_by_name.get(paper_name)
        if metadata is None:
            print(f"[skip] No metadata match for {paper_name}")
            continue
        paper_id = str(metadata["paper_id"])
        stem = paper_id.replace(":", "_")
        render_dir = generated_pptx.parent / f"{generated_pptx.stem}_slidetailor_images"
        slide_images = render_pptx_to_images(generated_pptx, render_dir)
        combined: dict[str, Any] = {
            "source": "SlideTailor-derived",
            "paper_id": paper_id,
            "generated_pptx": str(generated_pptx),
            "metrics": {},
        }

        if args.compute_aesthetic:
            aesthetic_result = evaluate_aesthetic_quality(
                pptx_path=generated_pptx,
                slide_images=slide_images,
                model=args.judge_model,
                request_timeout=args.request_timeout,
                verbose=args.verbose,
            )
            metric_output_dir("aesthetic_quality").mkdir(parents=True, exist_ok=True)
            (metric_output_dir("aesthetic_quality") / f"{stem}.aesthetic_quality.json").write_text(
                json.dumps(aesthetic_result, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
            combined["metrics"]["aesthetic_quality"] = aesthetic_result["deck_score"]

        if args.compute_content:
            source_document = metadata.get("paper_pdf_path")
            if source_document and Path(source_document).exists():
                content_result = evaluate_content_informativeness(
                    pptx_path=generated_pptx,
                    slide_images=slide_images,
                    source_document_path=Path(source_document),
                    model=args.judge_model,
                    request_timeout=args.request_timeout,
                    verbose=args.verbose,
                )
                metric_output_dir("content_informativeness").mkdir(parents=True, exist_ok=True)
                (metric_output_dir("content_informativeness") / f"{stem}.content_informativeness.json").write_text(
                    json.dumps(content_result, indent=2, ensure_ascii=False),
                    encoding="utf-8",
                )
                combined["metrics"]["content_informativeness"] = content_result["deck_score"]
            else:
                print(f"[skip] Missing paper PDF for {paper_id}")

        if args.compute_structure:
            reference_slide_dir = metadata.get("reference_slide_dir")
            if reference_slide_dir and Path(reference_slide_dir).exists():
                structure_result = evaluate_structure_similarity(
                    generated_pptx=generated_pptx,
                    reference_slide_images=collect_slide_images(Path(reference_slide_dir)),
                    model=args.judge_model,
                    request_timeout=args.request_timeout,
                    verbose=args.verbose,
                    categories=DEFAULT_CATEGORY_LIST,
                    max_reference_slides=DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES,
                )
                metric_output_dir("structure_similarity").mkdir(parents=True, exist_ok=True)
                (metric_output_dir("structure_similarity") / f"{stem}.structure_similarity.json").write_text(
                    json.dumps(structure_result, indent=2, ensure_ascii=False),
                    encoding="utf-8",
                )
                combined["metrics"]["structure_similarity"] = {
                    "coverage_iou": structure_result["coverage_iou"],
                    "flow_ngld": structure_result["flow_ngld"],
                    "content_structure_similarity": structure_result["content_structure_similarity"],
                }
            else:
                print(f"[skip] Missing reference slide dir for {paper_id}")

        if args.compute_template:
            if args.template_pptx is None:
                print(f"[skip] No --template-pptx provided for template similarity on {paper_id}")
            else:
                template_slide_dir = args.template_pptx.parent / f"{args.template_pptx.stem}_slidetailor_images"
                template_slide_images = render_pptx_to_images(args.template_pptx, template_slide_dir)
                template_result = evaluate_template_similarity(
                    generated_pptx=generated_pptx,
                    generated_slide_images=slide_images,
                    template_slide_images=template_slide_images,
                    model=args.judge_model,
                    request_timeout=args.request_timeout,
                    verbose=args.verbose,
                    max_generated_slides=DEFAULT_MAX_VISION_IMAGES,
                    max_template_slides=DEFAULT_MAX_VISION_IMAGES,
                )
                metric_output_dir("template_similarity").mkdir(parents=True, exist_ok=True)
                (metric_output_dir("template_similarity") / f"{stem}.template_similarity.json").write_text(
                    json.dumps(template_result, indent=2, ensure_ascii=False),
                    encoding="utf-8",
                )
                combined["metrics"]["template_similarity"] = template_result["score"]

        metric_output_dir("combined").mkdir(parents=True, exist_ok=True)
        (metric_output_dir("combined") / f"{stem}.slidetailor_eval.json").write_text(
            json.dumps(combined, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print(f"[done] {paper_id}")


if __name__ == "__main__":
    main()
