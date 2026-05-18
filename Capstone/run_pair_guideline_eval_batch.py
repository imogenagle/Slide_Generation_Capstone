#!/usr/bin/env python3
"""Run pair-guideline win-rate evaluation for every paper in a batch manifest."""

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
from Capstone.pair_guidelines import infer_output_key_from_paper_path, output_key_from_paper_id


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def contents_name_for_manifest_item(item: dict[str, Any], outline_mode: str) -> str:
    paper_id = str(item["paper_id"])
    paper_path = Path(str(item["paper_path"]))
    return append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)


def baseline_plan_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = contents_name_for_manifest_item(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_slide_plan_baseline.json"


def pair_guided_plan_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = contents_name_for_manifest_item(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_slide_plan_pair_guidelines.json"


def baseline_pptx_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = contents_name_for_manifest_item(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"{model_name_t}_{model_name_v}_output_slides_baseline.pptx"


def pair_context_path(*, item: dict[str, Any]) -> Path:
    author_id = str(item["author_id"])
    paper_id = str(item.get("paper_id") or "").strip()
    paper_path = str(item["paper_path"])
    target_key = output_key_from_paper_id(paper_id) or infer_output_key_from_paper_path(paper_path)
    return PROJECT_ROOT / "Capstone" / "pair_guideline_contexts" / author_id / f"{target_key}.json"


def build_eval_command(
    *,
    item: dict[str, Any],
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
    judge_model: str,
    output_path: Path,
    request_timeout: float,
    prompt_path: Path | None,
    verbose: bool,
) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "Capstone.evaluate_pair_guideline_winrate",
        "--pair-context",
        str(pair_context_path(item=item)),
        "--baseline-plan",
        str(
            baseline_plan_path(
                item=item,
                outline_mode=outline_mode,
                model_name_t=model_name_t,
                model_name_v=model_name_v,
            )
        ),
        "--pair-guided-plan",
        str(
            pair_guided_plan_path(
                item=item,
                outline_mode=outline_mode,
                model_name_t=model_name_t,
                model_name_v=model_name_v,
            )
        ),
        "--model",
        judge_model,
        "--request-timeout",
        str(request_timeout),
        "--output",
        str(output_path),
    ]
    if prompt_path is not None:
        command.extend(["--prompt-path", str(prompt_path)])
    if verbose:
        command.append("--verbose")
    return command


def build_baseline_generation_command(
    *,
    item: dict[str, Any],
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
    outline_mode: str,
) -> list[str]:
    return [
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


def build_summary_command(*, eval_dir: Path) -> list[str]:
    return [
        sys.executable,
        "-m",
        "Capstone.summarize_pair_guideline_winrate",
        "--eval-dir",
        str(eval_dir),
    ]


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    rendered = " ".join(command)
    print(rendered)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate a pair-guideline batch manifest and summarize the win-rate.")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=PROJECT_ROOT / "Capstone" / "batch_runs" / "pair_guideline_runs" / "pair_guided_12_run" / "manifest.json",
        help="Manifest JSON produced by Capstone.run_pair_guideline_batch.",
    )
    parser.add_argument(
        "--eval-dir",
        type=Path,
        default=None,
        help="Directory for *.pairwin.json reports. Defaults to <manifest_dir>/eval.",
    )
    parser.add_argument("--judge-model", default="gpt-5.4-nano", help="Judge model for pair-guideline win-rate eval.")
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--prompt-path", type=Path, default=None)
    parser.add_argument(
        "--generate-missing-baselines",
        action="store_true",
        help="If a baseline plan is missing, run the baseline SlideGen pipeline first using cached parse artifacts.",
    )
    parser.add_argument("--include-existing", action="store_true", help="Re-evaluate reports even if they already exist.")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    args.manifest = resolve_cli_path(args.manifest)
    args.eval_dir = resolve_cli_path(args.eval_dir)
    args.prompt_path = resolve_cli_path(args.prompt_path)

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

    eval_dir = args.eval_dir or (args.manifest.parent / "eval")
    eval_dir.mkdir(parents=True, exist_ok=True)

    for index, item in enumerate(selected, start=1):
        paper_id = str(item["paper_id"])
        output_name = paper_id.replace(":", "_")
        output_path = eval_dir / f"{output_name}.pairwin.json"
        baseline_plan = baseline_plan_path(
            item=item,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        )

        if output_path.exists() and not args.include_existing and not args.dry_run:
            print(f"[{index}/{len(selected)}] {paper_id} -> skipping existing {output_path.name}")
            continue

        print(f"[{index}/{len(selected)}] {paper_id}")
        if not baseline_plan.exists():
            if not args.generate_missing_baselines:
                raise SystemExit(
                    f"Missing baseline plan for {paper_id}: {baseline_plan}\n"
                    "Rerun with --generate-missing-baselines to create it first."
                )
            baseline_command = build_baseline_generation_command(
                item=item,
                model_name_t=model_name_t,
                model_name_v=model_name_v,
                formula_mode=formula_mode,
                outline_mode=outline_mode,
            )
            run_command(baseline_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)

        command = build_eval_command(
            item=item,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
            judge_model=args.judge_model,
            output_path=output_path,
            request_timeout=args.request_timeout,
            prompt_path=args.prompt_path,
            verbose=args.verbose,
        )
        run_command(command, cwd=PROJECT_ROOT, dry_run=args.dry_run)

    summary_command = build_summary_command(eval_dir=eval_dir)
    run_command(summary_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
