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
        DEFAULT_PAPERS_CSV,
        load_paper_metadata,
        load_runtime_env,
        render_pptx_to_images,
    )
    from SlideGen.Capstone.slidetailor_eval.paths import metric_output_dir
else:
    from .common import (
        DEFAULT_PAPERS_CSV,
        load_paper_metadata,
        load_runtime_env,
        render_pptx_to_images,
    )
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
    parser = argparse.ArgumentParser(
        description="Legacy SlideTailor batch runner. Similarity, aesthetic, and content metrics have been removed."
    )
    parser.add_argument("--contents-dir", type=Path, default=DEFAULT_CONTENTS_DIR)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--model-name-t", default="4o-mini")
    parser.add_argument("--model-name-v", default="4o-mini")
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()

    load_runtime_env()
    metadata_by_name = load_paper_metadata(args.papers_csv)
    generated = find_generated_pptx_dirs(args.contents_dir, args.model_name_t, args.model_name_v)
    if args.limit > 0:
        generated = generated[: args.limit]

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

        metric_output_dir("combined").mkdir(parents=True, exist_ok=True)
        (metric_output_dir("combined") / f"{stem}.slidetailor_eval.json").write_text(
            json.dumps(combined, indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        print(f"[done] {paper_id}")


if __name__ == "__main__":
    main()
