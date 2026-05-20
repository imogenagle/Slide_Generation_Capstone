#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence


REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_PYTHON = REPO_ROOT / ".venv" / "bin" / "python3"


@dataclass(frozen=True)
class UserSpec:
    user_id: str
    paper_path: Path
    profile_path: Path
    paper_name_base: str


USER_SPECS: dict[str, UserSpec] = {
    "user1": UserSpec(
        user_id="user1",
        paper_path=REPO_ROOT / "data_raw" / "user1" / "target_paper" / "60_paper.pdf",
        profile_path=REPO_ROOT / "Capstone" / "tmp_user_profiles" / "user1.json",
        paper_name_base="target60_user1_eval",
    ),
    "user2": UserSpec(
        user_id="user2",
        paper_path=REPO_ROOT / "data_raw" / "user2" / "target_paper" / "60_paper.pdf",
        profile_path=REPO_ROOT / "Capstone" / "tmp_user_profiles" / "user2.json",
        paper_name_base="target60_user2_eval",
    ),
}


def resolve_python() -> Path:
    return DEFAULT_PYTHON if DEFAULT_PYTHON.exists() else Path(sys.executable)


def log(message: str) -> None:
    print(f"[user-personalization-eval] {message}", flush=True)


def quoted(parts: Sequence[object]) -> str:
    return " ".join(str(part) for part in parts)


def run_command(command: Sequence[object], *, dry_run: bool) -> None:
    command_list = [str(part) for part in command]
    log(f"Running: {quoted(command_list)}")
    if dry_run:
        return
    subprocess.run(command_list, cwd=REPO_ROOT, check=True)


def content_folder_name(spec: UserSpec, outline_mode: str) -> str:
    return f"{spec.paper_name_base}_{outline_mode}"


def plan_prefix(model_name_t: str, model_name_v: str) -> str:
    return f"<{model_name_t}_{model_name_v}>"


def build_pipeline_command(
    *,
    python_bin: Path,
    spec: UserSpec,
    model_name_t: str,
    model_name_v: str,
    outline_mode: str,
    formula_mode: int,
    personalized: bool,
) -> list[object]:
    command: list[object] = [
        python_bin,
        "-m",
        "SlidesAgent.new_pipeline",
        "--paper_path",
        spec.paper_path,
        "--paper_name",
        content_folder_name(spec, outline_mode),
        "--model_name_t",
        model_name_t,
        "--model_name_v",
        model_name_v,
        "--outline_mode",
        outline_mode,
        "--formula_mode",
        formula_mode,
    ]
    if personalized:
        command.extend(
            [
                "--use_author_preferences",
                "--author_profile_path",
                spec.profile_path,
            ]
        )
    return command


def build_eval_commands(
    *,
    python_bin: Path,
    spec: UserSpec,
    model_name_t: str,
    model_name_v: str,
    outline_mode: str,
    judge_model: str,
    core_coverage_model: str,
    skip_bundle_eval: bool,
) -> tuple[list[object], list[list[object]], dict[str, Path]]:
    folder_base = content_folder_name(spec, outline_mode)
    baseline_folder = REPO_ROOT / "contents" / folder_base
    personalized_folder = REPO_ROOT / "contents" / f"{folder_base}_personalized"
    prefix = plan_prefix(model_name_t, model_name_v)

    baseline_plan = baseline_folder / f"{prefix}_slide_plan_baseline.json"
    personalized_plan = personalized_folder / f"{prefix}_slide_plan_personalized.json"
    baseline_pptx = baseline_folder / f"{model_name_t}_{model_name_v}_output_slides_baseline.pptx"
    personalized_pptx = personalized_folder / f"{model_name_t}_{model_name_v}_output_slides_personalized.pptx"

    eval_dir = REPO_ROOT / "Capstone" / "evaluations" / "user_personalization" / spec.user_id
    eval_dir.mkdir(parents=True, exist_ok=True)
    pairwise_output = eval_dir / "personalization_alignment_pairwise.json"

    pairwise_command: list[object] = [
        python_bin,
        "Capstone/evaluate_personalization_alignment_pairwise.py",
        "--profile",
        spec.profile_path,
        "--baseline-plan",
        baseline_plan,
        "--personalized-plan",
        personalized_plan,
        "--model",
        judge_model,
        "--output",
        pairwise_output,
    ]

    bundle_commands: list[list[object]] = []
    if not skip_bundle_eval:
        baseline_bundle_dir = eval_dir / "bundle_eval_baseline"
        personalized_bundle_dir = eval_dir / "bundle_eval_personalized"
        bundle_commands = [
            [
                python_bin,
                "Capstone/evaluate_pptx_bundle.py",
                "--generated-pptx",
                baseline_pptx,
                "--source-document",
                spec.paper_path,
                "--paper-name",
                folder_base,
                "--judge-model",
                judge_model,
                "--core-coverage-model",
                core_coverage_model,
                "--output-dir",
                baseline_bundle_dir,
                "--skip-core-coverage",
            ],
            [
                python_bin,
                "Capstone/evaluate_pptx_bundle.py",
                "--generated-pptx",
                personalized_pptx,
                "--source-document",
                spec.paper_path,
                "--paper-name",
                f"{folder_base}_personalized",
                "--judge-model",
                judge_model,
                "--core-coverage-model",
                core_coverage_model,
                "--output-dir",
                personalized_bundle_dir,
                "--skip-core-coverage",
            ],
        ]

    outputs = {
        "baseline_plan": baseline_plan,
        "personalized_plan": personalized_plan,
        "pairwise_output": pairwise_output,
        "baseline_pptx": baseline_pptx,
        "personalized_pptx": personalized_pptx,
    }
    return pairwise_command, bundle_commands, outputs


def print_eval_summary(spec: UserSpec, eval_dir: Path) -> None:
    pairwise_path = eval_dir / "personalization_alignment_pairwise.json"
    if not pairwise_path.exists():
        log(f"{spec.user_id}: evaluation outputs missing under {eval_dir}")
        return

    pairwise = json.loads(pairwise_path.read_text(encoding="utf-8"))
    log(
        (
            f"{spec.user_id}: pairwise winner={pairwise.get('summary', {}).get('winner')} "
            f"overall_lift={pairwise.get('lift', {}).get('overall_style_alignment')} "
            f"wins={pairwise.get('summary', {}).get('win_counts')}"
        )
    )


def run_for_user(
    *,
    python_bin: Path,
    spec: UserSpec,
    model_name_t: str,
    model_name_v: str,
    outline_mode: str,
    formula_mode: int,
    judge_model: str,
    core_coverage_model: str,
    skip_bundle_eval: bool,
    dry_run: bool,
) -> None:
    log(f"Starting {spec.user_id}")
    if not spec.paper_path.exists():
        raise FileNotFoundError(f"Paper not found: {spec.paper_path}")
    if not spec.profile_path.exists():
        raise FileNotFoundError(f"Profile not found: {spec.profile_path}")

    baseline_command = build_pipeline_command(
        python_bin=python_bin,
        spec=spec,
        model_name_t=model_name_t,
        model_name_v=model_name_v,
        outline_mode=outline_mode,
        formula_mode=formula_mode,
        personalized=False,
    )
    personalized_command = build_pipeline_command(
        python_bin=python_bin,
        spec=spec,
        model_name_t=model_name_t,
        model_name_v=model_name_v,
        outline_mode=outline_mode,
        formula_mode=formula_mode,
        personalized=True,
    )

    run_command(baseline_command, dry_run=dry_run)
    run_command(personalized_command, dry_run=dry_run)

    pairwise_command, bundle_commands, outputs = build_eval_commands(
        python_bin=python_bin,
        spec=spec,
        model_name_t=model_name_t,
        model_name_v=model_name_v,
        outline_mode=outline_mode,
        judge_model=judge_model,
        core_coverage_model=core_coverage_model,
        skip_bundle_eval=skip_bundle_eval,
    )

    run_command(pairwise_command, dry_run=dry_run)
    for command in bundle_commands:
        run_command(command, dry_run=dry_run)

    if dry_run:
        return

    log(f"{spec.user_id}: baseline plan -> {outputs['baseline_plan']}")
    log(f"{spec.user_id}: personalized plan -> {outputs['personalized_plan']}")
    print_eval_summary(spec, outputs["alignment_output"].parent)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate and evaluate baseline vs personalized decks for user1/user2.")
    parser.add_argument("--users", nargs="+", choices=sorted(USER_SPECS), default=sorted(USER_SPECS))
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--judge-model", default="gpt-5")
    parser.add_argument("--core-coverage-model", default="4o-mini")
    parser.add_argument("--skip-bundle-eval", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    python_bin = resolve_python()
    log(f"Using Python: {python_bin}")
    for user_id in args.users:
        run_for_user(
            python_bin=python_bin,
            spec=USER_SPECS[user_id],
            model_name_t=args.model_name_t,
            model_name_v=args.model_name_v,
            outline_mode=args.outline_mode,
            formula_mode=args.formula_mode,
            judge_model=args.judge_model,
            core_coverage_model=args.core_coverage_model,
            skip_bundle_eval=args.skip_bundle_eval,
            dry_run=args.dry_run,
        )
    log("Done")


if __name__ == "__main__":
    main()
