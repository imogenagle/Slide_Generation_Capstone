#!/usr/bin/env python3
"""Run all applicable evaluation metrics for a generated PPTX into one folder."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.evaluate_core_coverage import MAX_IMAGES_PER_REQUEST, evaluate_core_coverage
from Capstone.evaluate_geometry_aware_density import evaluate_geometry_aware_density
from Capstone.slidetailor_eval.common import (
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
from Capstone.slidetailor_eval.evaluate_aesthetic_quality import evaluate_aesthetic_quality
from Capstone.slidetailor_eval.evaluate_content_informativeness import evaluate_content_informativeness
from Capstone.slidetailor_eval.evaluate_structure_similarity import evaluate_structure_similarity
from Capstone.slidetailor_eval.evaluate_template_similarity import evaluate_template_similarity


DEFAULT_BUNDLE_ROOT = REPO_ROOT / "Capstone" / "evaluations" / "deck_bundles"


def infer_metadata(
    generated_pptx: Path,
    *,
    papers_csv: Path,
    paper_id: str | None,
    paper_name: str | None,
) -> tuple[str | None, dict[str, Any] | None]:
    metadata_by_name = load_paper_metadata(papers_csv)
    if paper_id:
        return paper_id, metadata_by_name.get(normalize_paper_name_from_paper_id(paper_id))
    if paper_name:
        return paper_name, metadata_by_name.get(paper_name)
    inferred_name = generated_pptx.parent.name
    return inferred_name, metadata_by_name.get(inferred_name)


def default_output_dir(generated_pptx: Path, resolved_key: str | None) -> Path:
    folder_name = resolved_key or generated_pptx.stem
    safe = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in folder_name).strip("_")
    return DEFAULT_BUNDLE_ROOT / safe


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def maybe_path(value: Any) -> Path | None:
    if not value:
        return None
    path = Path(value)
    return path if path.exists() else None


def main() -> None:
    parser = argparse.ArgumentParser(description="Run all applicable evaluation metrics for one generated PPTX.")
    parser.add_argument("--generated-pptx", type=Path, required=True)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--paper-id", default=None, help="Optional explicit paper id, e.g. eccv20:589")
    parser.add_argument("--paper-name", default=None, help="Optional explicit contents-folder style name.")
    parser.add_argument("--title", default=None, help="Optional explicit paper title override.")
    parser.add_argument("--original-slide-dir", type=Path, default=None, help="Optional explicit reference slide image directory.")
    parser.add_argument("--source-document", type=Path, default=None, help="Optional explicit source document path for content informativeness.")
    parser.add_argument("--template-pptx", type=Path, default=None, help="Optional template PPTX for template similarity.")
    parser.add_argument("--output-dir", type=Path, default=None, help="Optional per-deck output directory.")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--core-coverage-model", default="4o-mini")
    parser.add_argument("--judge-model", default="gpt-5")
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--max-original-slides", type=int, default=MAX_IMAGES_PER_REQUEST)
    parser.add_argument("--skip-core-coverage", action="store_true")
    parser.add_argument("--skip-gad", action="store_true")
    parser.add_argument("--skip-aesthetic", action="store_true")
    parser.add_argument("--skip-content", action="store_true")
    parser.add_argument("--skip-structure", action="store_true")
    parser.add_argument("--skip-template", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")
    load_runtime_env()

    resolved_key, metadata = infer_metadata(
        args.generated_pptx,
        papers_csv=args.papers_csv,
        paper_id=args.paper_id,
        paper_name=args.paper_name,
    )

    paper_id = args.paper_id or (str(metadata["paper_id"]) if metadata else None)
    title = args.title or (str(metadata["title"]) if metadata else args.generated_pptx.stem)
    original_slide_dir = args.original_slide_dir or maybe_path(metadata.get("reference_slide_dir") if metadata else None)
    source_document = args.source_document or maybe_path(metadata.get("paper_pdf_path") if metadata else None)

    output_dir = args.output_dir or default_output_dir(args.generated_pptx, paper_id or resolved_key)
    output_dir.mkdir(parents=True, exist_ok=True)

    rendered_slide_dir = output_dir / "rendered_slides"
    slide_images = render_pptx_to_images(args.generated_pptx, rendered_slide_dir, dpi=args.render_dpi)

    summary: dict[str, Any] = {
        "generated_pptx": str(args.generated_pptx),
        "paper_id": paper_id,
        "resolved_key": resolved_key,
        "title": title,
        "output_dir": str(output_dir),
        "inputs": {
            "original_slide_dir": str(original_slide_dir) if original_slide_dir else None,
            "source_document": str(source_document) if source_document else None,
            "template_pptx": str(args.template_pptx) if args.template_pptx else None,
        },
        "metrics": {},
        "skipped": {},
    }

    if not args.skip_core_coverage:
        if paper_id and original_slide_dir and original_slide_dir.exists():
            result = evaluate_core_coverage(
                paper_id=paper_id,
                title=title,
                original_slide_dir=original_slide_dir,
                generated_pptx=args.generated_pptx,
                model=args.core_coverage_model,
                max_original_slides=args.max_original_slides,
            )
            write_json(output_dir / "core_coverage.json", result)
            summary["metrics"]["core_coverage"] = {
                "topic_iou": result.get("topic_iou"),
                "path": "core_coverage.json",
            }
        else:
            summary["skipped"]["core_coverage"] = "Missing paper_id or original_slide_dir."

    if not args.skip_gad:
        result = evaluate_geometry_aware_density(
            pptx_path=args.generated_pptx,
            tau=0.5,
            m_star=4.0,
            kappa=6.3,
            lambda_occupancy=0.5,
            lambda_fragmentation=0.5,
            area_min_ratio=0.005,
            background_area_ratio=0.9,
        )
        write_json(output_dir / "geometry_aware_density.json", result)
        summary["metrics"]["geometry_aware_density"] = {
            "gad_geom": result.get("gad_geom"),
            "path": "geometry_aware_density.json",
        }

    if not args.skip_aesthetic:
        result = evaluate_aesthetic_quality(
            pptx_path=args.generated_pptx,
            slide_images=slide_images,
            model=args.judge_model,
            request_timeout=args.request_timeout,
            verbose=args.verbose,
        )
        write_json(output_dir / "slidetailor_aesthetic_quality.json", result)
        summary["metrics"]["slidetailor_aesthetic_quality"] = {
            "deck_score": result.get("deck_score"),
            "path": "slidetailor_aesthetic_quality.json",
        }

    if not args.skip_content:
        if source_document and source_document.exists():
            result = evaluate_content_informativeness(
                pptx_path=args.generated_pptx,
                slide_images=slide_images,
                source_document_path=source_document,
                model=args.judge_model,
                request_timeout=args.request_timeout,
                verbose=args.verbose,
            )
            write_json(output_dir / "slidetailor_content_informativeness.json", result)
            summary["metrics"]["slidetailor_content_informativeness"] = {
                "deck_score": result.get("deck_score"),
                "path": "slidetailor_content_informativeness.json",
            }
        else:
            summary["skipped"]["slidetailor_content_informativeness"] = "Missing source_document."

    if not args.skip_structure:
        if original_slide_dir and original_slide_dir.exists():
            result = evaluate_structure_similarity(
                generated_pptx=args.generated_pptx,
                reference_slide_images=collect_slide_images(original_slide_dir),
                model=args.judge_model,
                request_timeout=args.request_timeout,
                verbose=args.verbose,
                categories=DEFAULT_CATEGORY_LIST,
                max_reference_slides=DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES,
            )
            write_json(output_dir / "slidetailor_structure_similarity.json", result)
            summary["metrics"]["slidetailor_structure_similarity"] = {
                "coverage_iou": result.get("coverage_iou"),
                "flow_ngld": result.get("flow_ngld"),
                "content_structure_similarity": result.get("content_structure_similarity"),
                "path": "slidetailor_structure_similarity.json",
            }
        else:
            summary["skipped"]["slidetailor_structure_similarity"] = "Missing original_slide_dir."

    if not args.skip_template:
        if args.template_pptx and args.template_pptx.exists():
            template_slide_dir = output_dir / "template_rendered_slides"
            template_slide_images = render_pptx_to_images(args.template_pptx, template_slide_dir, dpi=args.render_dpi)
            result = evaluate_template_similarity(
                generated_pptx=args.generated_pptx,
                generated_slide_images=slide_images,
                template_slide_images=template_slide_images,
                model=args.judge_model,
                request_timeout=args.request_timeout,
                verbose=args.verbose,
                max_generated_slides=DEFAULT_MAX_VISION_IMAGES,
                max_template_slides=DEFAULT_MAX_VISION_IMAGES,
            )
            write_json(output_dir / "slidetailor_template_similarity.json", result)
            summary["metrics"]["slidetailor_template_similarity"] = {
                "score": result.get("score"),
                "path": "slidetailor_template_similarity.json",
            }
        else:
            summary["skipped"]["slidetailor_template_similarity"] = "Missing template_pptx."

    write_json(output_dir / "summary.json", summary)
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
