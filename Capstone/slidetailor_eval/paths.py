"""Shared paths for SlideTailor-derived evaluation scripts."""

from __future__ import annotations

from pathlib import Path


SLIDETAILOR_EVAL_ROOT = Path(__file__).resolve().parent
SLIDETAILOR_OUTPUT_ROOT = SLIDETAILOR_EVAL_ROOT.parent / "evaluations" / "slidetailor"


def metric_output_dir(metric_name: str) -> Path:
    return SLIDETAILOR_OUTPUT_ROOT / metric_name
