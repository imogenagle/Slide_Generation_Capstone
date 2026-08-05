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
    DEFAULT_MAX_VISION_IMAGES,
    DEFAULT_PAPERS_CSV,
    load_paper_metadata,
    load_runtime_env,
    normalize_paper_name_from_paper_id,
    render_pptx_to_images,
)
from Capstone.slidetailor_eval.evaluate_logical_flow import evaluate_logical_flow
from Capstone.slidetailor_eval.evaluate_layout_correctness import evaluate_layout_correctness
from Capstone.slidetailor_eval.evaluate_paper_faithfulness import evaluate_paper_faithfulness
from Capstone.slidetailor_eval.evaluate_visual_appeal import evaluate_visual_appeal


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
    parser.add_argument("--source-document", type=Path, default=None, help="Optional explicit source document path for source-grounded evals.")
    parser.add_argument("--output-dir", type=Path, default=None, help="Optional per-deck output directory.")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--max-original-slides", type=int, default=MAX_IMAGES_PER_REQUEST)
    parser.add_argument("--skip-core-coverage", action="store_true")
    parser.add_argument("--skip-gad", action="store_true")
    parser.add_argument("--skip-visual-appeal", action="store_true")
    parser.add_argument("--skip-layout-correctness", action="store_true")
    parser.add_argument("--skip-layout-defects", action="store_true")
    parser.add_argument("--skip-logical-flow", action="store_true")
    parser.add_argument("--skip-faithfulness", action="store_true")
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
            if result.get("skipped"):
                summary["skipped"]["core_coverage"] = str(result.get("notes") or result.get("skip_reason") or "Skipped.")
            else:
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

    if not args.skip_visual_appeal:
        result = evaluate_visual_appeal(
            pptx_path=args.generated_pptx,
            slide_images=slide_images,
            model=args.judge_model,
            request_timeout=args.request_timeout,
            verbose=args.verbose,
        )
        write_json(output_dir / "visual_appeal.json", result)
        summary["metrics"]["visual_appeal"] = {
            "deck_score": result.get("deck_score"),
            "path": "visual_appeal.json",
        }

    if not args.skip_layout_correctness and not args.skip_layout_defects:
        result = evaluate_layout_correctness(
            pptx_path=args.generated_pptx,
        )
        write_json(output_dir / "layout_defects.json", result)
        summary["metrics"]["layout_defects"] = {
            "deck_score": result.get("deck_score"),
            "path": "layout_defects.json",
        }

    if not args.skip_logical_flow:
        result = evaluate_logical_flow(
            pptx_path=args.generated_pptx,
            slide_images=slide_images,
            model=args.judge_model,
            request_timeout=args.request_timeout,
            verbose=args.verbose,
        )
        write_json(output_dir / "logical_flow.json", result)
        summary["metrics"]["logical_flow"] = {
            "deck_score": result.get("deck_score"),
            "path": "logical_flow.json",
        }

    if not args.skip_faithfulness:
        if source_document and source_document.exists():
            result = evaluate_paper_faithfulness(
                pptx_path=args.generated_pptx,
                slide_images=slide_images,
                source_document_path=source_document,
                model=args.judge_model,
                request_timeout=args.request_timeout,
                verbose=args.verbose,
            )
            write_json(output_dir / "paper_faithfulness.json", result)
            summary["metrics"]["paper_faithfulness"] = {
                "deck_score": result.get("deck_score"),
                "path": "paper_faithfulness.json",
            }
        else:
            summary["skipped"]["paper_faithfulness"] = "Missing source_document."

    write_json(output_dir / "summary.json", summary)
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
