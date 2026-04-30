#!/usr/bin/env python3
"""SlideTailor-style content informativeness evaluation for SlideGen decks."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

if __package__ in {None, ""}:
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
    from SlideGen.Capstone.slidetailor_eval.common import (
        REPO_ROOT,
        add_shared_args,
        collect_slide_images,
        evaluate_single_slide_images,
        extract_source_document_text,
        load_runtime_env,
        render_pptx_to_images,
        resolve_output_path,
        summarize_scores,
    )
else:
    from .common import (
        REPO_ROOT,
        add_shared_args,
        collect_slide_images,
        evaluate_single_slide_images,
        extract_source_document_text,
        load_runtime_env,
        render_pptx_to_images,
        resolve_output_path,
        summarize_scores,
    )


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "slidetailor_eval" / "content_informativeness.yaml"


def evaluate_content_informativeness(
    *,
    pptx_path: Path,
    slide_images: list[Path],
    source_document_path: Path,
    model: str,
    request_timeout: float,
    verbose: bool,
) -> dict:
    source_text = extract_source_document_text(source_document_path)
    per_slide = evaluate_single_slide_images(
        metric_name="content_informativeness",
        slide_images=slide_images,
        prompt_path=DEFAULT_PROMPT_PATH,
        model=model,
        request_timeout=request_timeout,
        verbose=verbose,
        prompt_values={"source_document_text": source_text},
    )
    return {
        "source": "SlideTailor-derived",
        "metric": "content_informativeness",
        "pptx_path": str(pptx_path),
        "source_document_path": str(source_document_path),
        "source_document_excerpt_chars": len(source_text),
        "deck_score": summarize_scores(per_slide),
        "per_slide": per_slide,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate SlideTailor-style content informativeness.")
    parser.add_argument("--pptx-path", type=Path, required=True)
    parser.add_argument("--source-document", type=Path, required=True)
    parser.add_argument("--slide-image-dir", type=Path, default=None)
    add_shared_args(parser, metric_name="content_informativeness")
    args = parser.parse_args()

    load_runtime_env()
    slide_dir = args.slide_image_dir or args.pptx_path.with_suffix("")
    if args.slide_image_dir is None:
        slide_dir = slide_dir.parent / f"{slide_dir.name}_slidetailor_images"
        slide_images = render_pptx_to_images(args.pptx_path, slide_dir)
    else:
        slide_images = collect_slide_images(slide_dir)
    result = evaluate_content_informativeness(
        pptx_path=args.pptx_path,
        slide_images=slide_images,
        source_document_path=args.source_document,
        model=args.model,
        request_timeout=args.request_timeout,
        verbose=args.verbose,
    )
    output_path = resolve_output_path("content_informativeness", args.output, args.pptx_path.stem)
    output_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
