#!/usr/bin/env python3
"""Run bundle evaluation for baseline and pair-guided decks from a pair-guideline manifest."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from Capstone.generate_random_decks import append_outline_mode_suffix, output_dir_key, resolve_cli_path


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def contents_name_for_manifest_item(item: dict[str, Any], outline_mode: str) -> str:
    paper_id = str(item["paper_id"])
    paper_path = Path(str(item["paper_path"]))
    return append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)


def baseline_pptx_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = contents_name_for_manifest_item(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"{model_name_t}_{model_name_v}_output_slides_baseline.pptx"


def pair_guided_pptx_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = contents_name_for_manifest_item(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"{model_name_t}_{model_name_v}_output_slides_pair_guidelines.pptx"


def build_bundle_eval_command(
    *,
    generated_pptx: Path,
    paper_id: str,
    output_dir: Path,
    core_coverage_model: str,
    judge_model: str,
    render_dpi: int,
    request_timeout: float,
    include_preference_dependent_slidetailor: bool,
    verbose: bool,
) -> list[str]:
    command = [
        sys.executable,
        "Capstone/evaluate_pptx_bundle.py",
        "--generated-pptx",
        str(generated_pptx),
        "--paper-id",
        paper_id,
        "--output-dir",
        str(output_dir),
        "--core-coverage-model",
        core_coverage_model,
        "--judge-model",
        judge_model,
        "--render-dpi",
        str(render_dpi),
        "--request-timeout",
        str(request_timeout),
    ]
    if include_preference_dependent_slidetailor:
        command.append("--include-preference-dependent-slidetailor")
    if verbose:
        command.append("--verbose")
    return command


def build_summary_command(
    *,
    root_dir: Path,
    include_preference_dependent_slidetailor: bool,
) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "Capstone.summarize_bundle_eval",
        "--root-dir",
        str(root_dir),
    ]
    if include_preference_dependent_slidetailor:
        command.append("--include-preference-dependent-slidetailor")
    return command


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    rendered = " ".join(command)
    print(rendered)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run bundle eval for baseline and pair-guided outputs in a pair-guideline manifest."
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        default=PROJECT_ROOT / "Capstone" / "batch_runs" / "pair_guideline_runs" / "pair_guided_12_run" / "manifest.json",
        help="Manifest JSON produced by Capstone.run_pair_guideline_batch.",
    )
    parser.add_argument(
        "--bundle-root",
        type=Path,
        default=None,
        help="Directory for per-paper bundle eval outputs. Defaults to <manifest_dir>.",
    )
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument(
        "--include-preference-dependent-slidetailor",
        action="store_true",
        help="Include preference-based SlideTailor metrics in the bundle evaluation.",
    )
    parser.add_argument(
        "--generate-missing-baselines",
        action="store_true",
        help="If a baseline PPTX is missing, create it first using the standard SlideGen pipeline.",
    )
    parser.add_argument("--include-existing", action="store_true", help="Re-run evals even when summary.json already exists.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    args.manifest = resolve_cli_path(args.manifest)
    args.bundle_root = resolve_cli_path(args.bundle_root)

    if not args.manifest.exists():
        raise SystemExit(f"Manifest not found: {args.manifest}")

    manifest = load_json(args.manifest)
    selected = list(manifest.get("selected_papers") or [])
    if not selected:
        raise SystemExit(f"No selected_papers found in manifest: {args.manifest}")

    outline_mode = str(manifest.get("outline_mode") or "high_level")
    model_name_t = str(manifest.get("model_name_t") or "gpt-5.4-nano")
    model_name_v = str(manifest.get("model_name_v") or "gpt-5.4-nano")
    formula_mode = int(manifest.get("formula_mode") or 1)

    bundle_root = args.bundle_root or args.manifest.parent
    bundle_root.mkdir(parents=True, exist_ok=True)

    for index, item in enumerate(selected, start=1):
        paper_id = str(item["paper_id"])
        print(f"[{index}/{len(selected)}] {paper_id}")

        baseline_pptx = baseline_pptx_path(
            item=item,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        )
        pair_guided_pptx = pair_guided_pptx_path(
            item=item,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        )

        if not pair_guided_pptx.exists():
            raise SystemExit(f"Missing pair-guided PPTX for {paper_id}: {pair_guided_pptx}")

        if not baseline_pptx.exists():
            if not args.generate_missing_baselines:
                raise SystemExit(
                    f"Missing baseline PPTX for {paper_id}: {baseline_pptx}\n"
                    "Rerun with --generate-missing-baselines to create it first."
                )
            baseline_command = [
                sys.executable,
                "-m",
                "SlidesAgent.new_pipeline",
                "--paper_path",
                str(item["paper_path"]),
                "--model_name_t",
                model_name_t,
                "--model_name_v",
                model_name_v,
                "--formula_mode",
                str(formula_mode),
                "--outline_mode",
                outline_mode,
            ]
            run_command(baseline_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)

        for variant, pptx_path in (("baseline", baseline_pptx), ("pair_guided", pair_guided_pptx)):
            output_dir = bundle_root / paper_id.replace(":", "_") / "bundle_eval" / variant
            summary_path = output_dir / "summary.json"
            if summary_path.exists() and not args.include_existing and not args.dry_run:
                print(f"  {variant}: skipping existing {summary_path}")
                continue
            command = build_bundle_eval_command(
                generated_pptx=pptx_path,
                paper_id=paper_id,
                output_dir=output_dir,
                core_coverage_model=args.core_coverage_model,
                judge_model=args.judge_model,
                render_dpi=args.render_dpi,
                request_timeout=args.request_timeout,
                include_preference_dependent_slidetailor=args.include_preference_dependent_slidetailor,
                verbose=args.verbose,
            )
            run_command(command, cwd=PROJECT_ROOT, dry_run=args.dry_run)

    summary_command = build_summary_command(
        root_dir=args.manifest.parent,
        include_preference_dependent_slidetailor=args.include_preference_dependent_slidetailor,
    )
    run_command(summary_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
