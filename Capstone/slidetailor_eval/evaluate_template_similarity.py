#!/usr/bin/env python3
"""SlideTailor-style template similarity evaluation for SlideGen decks."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

if __package__ in {None, ""}:
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
    from SlideGen.Capstone.slidetailor_eval.common import (
        DEFAULT_MAX_VISION_IMAGES,
        REPO_ROOT,
        _normalize_score,
        add_shared_args,
        call_json_judge,
        collect_slide_images,
        load_runtime_env,
        render_pptx_to_images,
        render_prompt,
        resolve_output_path,
        sample_paths,
    )
else:
    from .common import (
        DEFAULT_MAX_VISION_IMAGES,
        REPO_ROOT,
        _normalize_score,
        add_shared_args,
        call_json_judge,
        collect_slide_images,
        load_runtime_env,
        render_pptx_to_images,
        render_prompt,
        resolve_output_path,
        sample_paths,
    )


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "slidetailor_eval" / "template_similarity.yaml"


def evaluate_template_similarity(
    *,
    generated_pptx: Path,
    generated_slide_images: list[Path],
    template_slide_images: list[Path],
    model: str,
    request_timeout: float,
    verbose: bool,
    max_generated_slides: int,
    max_template_slides: int,
) -> dict:
    sampled_generated = sample_paths(generated_slide_images, max_generated_slides)
    sampled_template = sample_paths(template_slide_images, max_template_slides)
    prompt = render_prompt(
        DEFAULT_PROMPT_PATH,
        generated_slide_count=len(generated_slide_images),
        template_slide_count=len(template_slide_images),
    )
    response = call_json_judge(
        model=model,
        system_prompt=prompt["system_prompt"],
        user_prompt=prompt["user_prompt"],
        image_paths=[*sampled_generated, *sampled_template],
        image_labels=[
            *[f"Generated slide: {path.name}" for path in sampled_generated],
            *[f"Template slide: {path.name}" for path in sampled_template],
        ],
        request_timeout=request_timeout,
        verbose=verbose,
    )
    return {
        "source": "SlideTailor-derived",
        "metric": "template_similarity",
        "generated_pptx": str(generated_pptx),
        "score": _normalize_score(response.get("score", 0.0)),
        "reason": str(response.get("reason", "")).strip(),
        "generated_slide_count": len(generated_slide_images),
        "template_slide_count": len(template_slide_images),
        "sampled_generated_slides": [path.name for path in sampled_generated],
        "sampled_template_slides": [path.name for path in sampled_template],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate SlideTailor-style template similarity.")
    parser.add_argument("--generated-pptx", type=Path, required=True)
    parser.add_argument("--generated-slide-dir", type=Path, default=None)
    parser.add_argument("--template-pptx", type=Path, default=None)
    parser.add_argument("--template-slide-dir", type=Path, default=None)
    parser.add_argument("--max-generated-slides", type=int, default=DEFAULT_MAX_VISION_IMAGES)
    parser.add_argument("--max-template-slides", type=int, default=DEFAULT_MAX_VISION_IMAGES)
    add_shared_args(parser, metric_name="template_similarity")
    args = parser.parse_args()

    if args.template_pptx is None and args.template_slide_dir is None:
        raise SystemExit("Provide either --template-pptx or --template-slide-dir.")

    load_runtime_env()
    generated_slide_dir = args.generated_slide_dir or args.generated_pptx.parent / f"{args.generated_pptx.stem}_slidetailor_images"
    if args.generated_slide_dir is None:
        generated_slide_images = render_pptx_to_images(args.generated_pptx, generated_slide_dir)
    else:
        generated_slide_images = collect_slide_images(generated_slide_dir)

    if args.template_slide_dir is not None:
        template_slide_images = collect_slide_images(args.template_slide_dir)
    else:
        template_slide_dir = args.template_pptx.parent / f"{args.template_pptx.stem}_slidetailor_images"
        template_slide_images = render_pptx_to_images(args.template_pptx, template_slide_dir)

    result = evaluate_template_similarity(
        generated_pptx=args.generated_pptx,
        generated_slide_images=generated_slide_images,
        template_slide_images=template_slide_images,
        model=args.model,
        request_timeout=args.request_timeout,
        verbose=args.verbose,
        max_generated_slides=args.max_generated_slides,
        max_template_slides=args.max_template_slides,
    )
    output_path = resolve_output_path("template_similarity", args.output, args.generated_pptx.stem)
    output_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
