"""SlideTailor-derived evaluation utilities."""

from pathlib import Path


SLIDETAILOR_EVAL_ROOT = Path(__file__).resolve().parent
SLIDETAILOR_OUTPUT_ROOT = SLIDETAILOR_EVAL_ROOT.parent / "evaluations" / "slidetailor"
