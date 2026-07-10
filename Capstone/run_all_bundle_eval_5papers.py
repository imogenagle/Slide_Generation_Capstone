#!/usr/bin/env python3
"""Run bundle eval for the standard 5-paper comparison set in one command.

This covers:
- SlideGen Original
- SlideGen baseline
- SlideGen personalized

Then refreshes the comparison tables.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
PYTHON = REPO_ROOT / ".venv" / "bin" / "python3"
EVAL_SCRIPT = REPO_ROOT / "Capstone" / "evaluate_pptx_bundle.py"
COMPARE_SCRIPT = REPO_ROOT / "Capstone" / "compare_bundle_eval_tables.py"

PAPERS = [
    {"paper_id": "acl18:74", "paper_key": "acl18_74", "original_key": "acl18_74_original"},
    {"paper_id": "acl20:317", "paper_key": "acl20_317", "original_key": "acl20_317_original"},
    {"paper_id": "cvpr20:1183", "paper_key": "cvpr20_1183", "original_key": "cvpr20_1183_original"},
    {"paper_id": "eccv20:47", "paper_key": "eccv20_47", "original_key": "eccv20_47_original"},
    {"paper_id": "icml20:398", "paper_key": "icml20_398", "original_key": "icml20_398_original"},
]


def pick_existing(paths: list[Path]) -> Path:
    for path in paths:
        if path.exists():
            return path
    tried = "\n".join(str(path) for path in paths)
    raise FileNotFoundError(f"None of the candidate PPTX files exist:\n{tried}")


def run_command(cmd: list[str]) -> None:
    print("Running:", " ".join(cmd), flush=True)
    subprocess.run(cmd, check=True, cwd=REPO_ROOT)


def build_eval_skip_flags(args: argparse.Namespace) -> list[str]:
    flags: list[str] = []
    if args.skip_core_coverage:
        flags.append("--skip-core-coverage")
    if args.skip_gad:
        flags.append("--skip-gad")
    if args.skip_aesthetic:
        flags.append("--skip-aesthetic")
    if args.skip_visual_appeal:
        flags.append("--skip-visual-appeal")
    if args.skip_logical_flow:
        flags.append("--skip-logical-flow")
    if args.skip_faithfulness:
        flags.append("--skip-faithfulness")
    if args.skip_content:
        flags.append("--skip-content")
    return flags


def build_compare_metric_args(args: argparse.Namespace) -> list[str]:
    metric_keys: list[str] = []
    if not args.skip_core_coverage:
        metric_keys.append("core_coverage_topic_iou")
    if not args.skip_gad:
        metric_keys.append("geometry_aware_density_gad_geom")
    if not args.skip_aesthetic:
        metric_keys.append("slidetailor_aesthetic_quality_deck_score")
    if not args.skip_visual_appeal:
        metric_keys.append("visual_appeal_deck_score")
    if not args.skip_logical_flow:
        metric_keys.append("logical_flow_deck_score")
    if not args.skip_faithfulness:
        metric_keys.append("paper_faithfulness_deck_score")
    if not args.skip_content:
        metric_keys.append("slidetailor_content_informativeness_deck_score")
    if not metric_keys:
        raise ValueError("At least one evaluation metric must be included.")
    return ["--metrics", *metric_keys]


def original_pptx(original_runs_root: Path, original_key: str) -> Path:
    base = original_runs_root / "contents" / original_key
    return pick_existing(
        [
            base / "gpt-5.4-nano_gpt-5.4-nano_output_slides_themed.pptx",
            base / "gpt-5.4-nano_gpt-5.4-nano_output_slides.pptx",
        ]
    )


def baseline_pptx(slidegen_runs_root: Path, paper_key: str) -> Path:
    base = slidegen_runs_root / "contents" / f"{paper_key}_high_level"
    return pick_existing(
        [
            base / "gpt-5.4-nano_gpt-5.4-nano_output_slides_baseline_themed.pptx",
            base / "gpt-5.4-nano_gpt-5.4-nano_output_slides_baseline.pptx",
        ]
    )


def personalized_pptx(slidegen_runs_root: Path, paper_key: str) -> Path:
    base = slidegen_runs_root / "contents" / f"{paper_key}_high_level_personalized_retrieval"
    return pick_existing(
        [
            base / "gpt-5.4-nano_gpt-5.4-nano_output_slides_personalized_retrieval_themed.pptx",
            base / "gpt-5.4-nano_gpt-5.4-nano_output_slides_personalized_retrieval.pptx",
        ]
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Run bundle eval for original/baseline/personalized 5-paper decks.")
    parser.add_argument(
        "--slidegen-runs-root",
        type=Path,
        default=REPO_ROOT / "outputs" / "retrieval_0702",
        help="Root directory containing SlideGen baseline/personalized contents folders.",
    )
    parser.add_argument(
        "--original-runs-root",
        type=Path,
        default=REPO_ROOT / "SlideGen_Original" / "my_original_runs",
        help="Root directory containing SlideGen Original contents folders.",
    )
    parser.add_argument(
        "--original-eval-root",
        type=Path,
        default=REPO_ROOT / "outputs" / "original_bundle_eval",
        help="Where to save per-paper original bundle eval outputs.",
    )
    parser.add_argument(
        "--bundle-eval-root",
        type=Path,
        default=REPO_ROOT / "outputs" / "retrieval_0702" / "Capstone" / "batch_runs" / "experiments" / "data_raw_5papers_bundle_eval_only_0707" / "bundle_eval",
        help="Where to save per-paper baseline/personalized bundle eval outputs.",
    )
    parser.add_argument(
        "--comparison-output-dir",
        type=Path,
        default=REPO_ROOT / "outputs" / "bundle_eval_comparison",
        help="Where to save the comparison tables.",
    )
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--skip-core-coverage", action="store_true")
    parser.add_argument("--skip-gad", action="store_true")
    parser.add_argument("--skip-aesthetic", action="store_true")
    parser.add_argument("--skip-visual-appeal", action="store_true")
    parser.add_argument("--skip-logical-flow", action="store_true")
    parser.add_argument("--skip-faithfulness", action="store_true")
    parser.add_argument("--skip-content", action="store_true")
    args = parser.parse_args()

    if not PYTHON.exists():
        raise FileNotFoundError(f"Expected Python interpreter not found: {PYTHON}")
    eval_skip_flags = build_eval_skip_flags(args)
    compare_metric_args = build_compare_metric_args(args)

    for paper in PAPERS:
        paper_id = paper["paper_id"]
        paper_key = paper["paper_key"]
        original_key = paper["original_key"]

        run_command(
            [
                str(PYTHON),
                str(EVAL_SCRIPT),
                "--generated-pptx",
                str(original_pptx(args.original_runs_root, original_key)),
                "--paper-id",
                paper_id,
                "--judge-model",
                args.judge_model,
                "--core-coverage-model",
                args.core_coverage_model,
                "--output-dir",
                str(args.original_eval_root / original_key),
                *eval_skip_flags,
            ]
        )

        run_command(
            [
                str(PYTHON),
                str(EVAL_SCRIPT),
                "--generated-pptx",
                str(baseline_pptx(args.slidegen_runs_root, paper_key)),
                "--paper-id",
                paper_id,
                "--judge-model",
                args.judge_model,
                "--core-coverage-model",
                args.core_coverage_model,
                "--output-dir",
                str(args.bundle_eval_root / "baseline" / paper_key),
                *eval_skip_flags,
            ]
        )

        run_command(
            [
                str(PYTHON),
                str(EVAL_SCRIPT),
                "--generated-pptx",
                str(personalized_pptx(args.slidegen_runs_root, paper_key)),
                "--paper-id",
                paper_id,
                "--judge-model",
                args.judge_model,
                "--core-coverage-model",
                args.core_coverage_model,
                "--output-dir",
                str(args.bundle_eval_root / "personalized" / paper_key),
                *eval_skip_flags,
            ]
        )

    run_command(
        [
            str(PYTHON),
            str(COMPARE_SCRIPT),
            "--original-root",
            str(args.original_eval_root),
            "--bundle-root",
            str(args.bundle_eval_root),
            "--output-dir",
            str(args.comparison_output_dir),
            *compare_metric_args,
        ]
    )

    print("\nFinished bundle eval for original, baseline, and personalized decks.", flush=True)
    print(f"Comparison table: {args.comparison_output_dir / 'bundle_eval_comparison_method_rows.csv'}", flush=True)


if __name__ == "__main__":
    main()
