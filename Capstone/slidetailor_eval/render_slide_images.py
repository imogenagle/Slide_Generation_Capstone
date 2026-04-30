#!/usr/bin/env python3
"""Render PPTX slides to per-slide JPG images for SlideTailor-derived evaluation."""

from __future__ import annotations

import argparse
from pathlib import Path

if __package__ in {None, ""}:
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
    from SlideGen.Capstone.slidetailor_eval.common import load_runtime_env, render_pptx_to_images
else:
    from .common import load_runtime_env, render_pptx_to_images


def main() -> None:
    parser = argparse.ArgumentParser(description="Render PPTX slides to JPG images.")
    parser.add_argument("--pptx-path", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--dpi", type=int, default=120)
    args = parser.parse_args()

    load_runtime_env()
    images = render_pptx_to_images(args.pptx_path, args.output_dir, force=args.force, dpi=args.dpi)
    print("\n".join(str(path) for path in images))


if __name__ == "__main__":
    main()
