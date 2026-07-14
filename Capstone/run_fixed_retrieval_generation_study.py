#!/usr/bin/env python3
"""Generate a fixed paper cohort across baseline, personalized, and original pipelines.

This script:
1. Builds a manifest of papers whose eligible author/coauthor has enough history.
2. Forces specific paper ids to be included (defaults to the existing 5-paper set).
3. Generates retrieval-conditioned profiles for every selected paper.
4. Generates SlideGen baseline + personalized decks for the full cohort.
5. Generates SlideGen Original decks for the same cohort.

It intentionally does not run evaluation.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
PYTHON = REPO_ROOT / ".venv" / "bin" / "python3"
ORIGINAL_REPO_ROOT = REPO_ROOT / "SlideGen_Original"
DEFAULT_FIXED_PAPER_IDS = [
    "acl18:74",
    "acl20:317",
    "cvpr20:1183",
    "eccv20:47",
    "icml20:398",
]


def resolve_python() -> Path:
    return PYTHON if PYTHON.exists() else Path(sys.executable)


def sanitize_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in value.strip())
    return cleaned.strip("_") or "generation_study"


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    print("Running:", " ".join(command), flush=True)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def try_run_command(
    command: list[str],
    *,
    cwd: Path,
    dry_run: bool,
    stage: str,
    paper_id: str | None,
    failures: list[dict[str, Any]],
) -> bool:
    print("Running:", " ".join(command), flush=True)
    if dry_run:
        return True
    try:
        subprocess.run(command, cwd=cwd, check=True)
        return True
    except subprocess.CalledProcessError as exc:
        failures.append(
            {
                "stage": stage,
                "paper_id": paper_id,
                "cwd": str(cwd),
                "returncode": exc.returncode,
                "cmd": exc.cmd,
                "error_type": "CalledProcessError",
            }
        )
        print(f"[warn] {stage} failed for {paper_id or 'study'}; continuing.", flush=True)
        return False
    except Exception as exc:
        failures.append(
            {
                "stage": stage,
                "paper_id": paper_id,
                "cwd": str(cwd),
                "error_type": exc.__class__.__name__,
                "message": str(exc),
            }
        )
        print(f"[warn] {stage} failed for {paper_id or 'study'}; continuing.", flush=True)
        return False


def load_manifest(manifest_path: Path) -> dict:
    with manifest_path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def build_manifest_command(
    *,
    python_bin: Path,
    count: int,
    seed: int,
    min_author_paper_count: int,
    output_path: Path,
    include_paper_ids: list[str],
) -> list[str]:
    command = [
        str(python_bin),
        "Capstone/build_fixed_retrieval_manifest.py",
        "--count",
        str(count),
        "--seed",
        str(seed),
        "--min-author-paper-count",
        str(min_author_paper_count),
        "--output",
        str(output_path),
    ]
    for paper_id in include_paper_ids:
        command.extend(["--include-paper-id", paper_id])
    return command


def build_profile_command(
    *,
    python_bin: Path,
    author_id: str,
    paper_id: str,
    output_dir: Path,
    max_retrieved: int,
    model: str,
    retrieval_ranker: str,
    force_refresh: bool,
) -> list[str]:
    command = [
        str(python_bin),
        "Capstone/retrieval_profile_pilot.py",
        "--author-id",
        author_id,
        "--target-paper-id",
        paper_id,
        "--output-dir",
        str(output_dir),
        "--max-retrieved",
        str(max_retrieved),
        "--model",
        model,
        "--retrieval-ranker",
        retrieval_ranker,
    ]
    if force_refresh:
        command.append("--force-refresh")
    return command


def build_slidegen_batch_command(
    *,
    python_bin: Path,
    manifest_path: Path,
    experiment_name: str,
    output_root: Path,
    retrieval_profile_dir: Path,
    model_t: str,
    model_v: str,
    formula_mode: int,
    outline_mode: str,
) -> list[str]:
    return [
        str(python_bin),
        "Capstone/run_batch_experiment.py",
        "--use-manifest",
        str(manifest_path),
        "--experiment-name",
        experiment_name,
        "--model-name-t",
        model_t,
        "--model-name-v",
        model_v,
        "--formula-mode",
        str(formula_mode),
        "--outline-mode",
        outline_mode,
        "--personalization-mode",
        "retrieval",
        "--retrieval-profile-dir",
        str(retrieval_profile_dir),
        "--output-dir",
        str(output_root),
        "--skip-bundle-eval",
        "--skip-personalization-eval",
    ]


def build_original_command(
    *,
    python_bin: Path,
    original_repo_root: Path,
    paper_path: str,
    paper_key: str,
    model_t: str,
    model_v: str,
    formula_mode: int,
    output_root: Path,
) -> list[str]:
    return [
        str(python_bin),
        "-m",
        "SlidesAgent.new_pipeline",
        "--paper_path",
        paper_path,
        "--paper_name",
        f"{paper_key}_original",
        "--model_name_t",
        model_t,
        "--model_name_v",
        model_v,
        "--formula_mode",
        str(formula_mode),
        "--output_root",
        str(output_root),
    ]


def paper_key(paper_id: str) -> str:
    return paper_id.replace(":", "_")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run a fixed 50-paper generation-only study.")
    parser.add_argument("--count", type=int, default=50)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--study-name", default="fixed50_generation_study")
    parser.add_argument("--output-root", type=Path, default=REPO_ROOT / "outputs")
    parser.add_argument("--original-repo-root", type=Path, default=ORIGINAL_REPO_ROOT)
    parser.add_argument("--manifest-path", type=Path, default=None)
    parser.add_argument("--min-history-pairs", type=int, default=3)
    parser.add_argument("--include-paper-id", action="append", default=None)
    parser.add_argument("--profile-history-count", type=int, default=3)
    parser.add_argument("--profile-model", default="gpt-5.4-nano")
    parser.add_argument(
        "--retrieval-ranker",
        choices=["title_similarity", "llm_title_abstract"],
        default="llm_title_abstract",
    )
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--force-refresh-profiles", action="store_true")
    parser.add_argument("--skip-manifest-build", action="store_true")
    parser.add_argument("--skip-profile-generation", action="store_true")
    parser.add_argument("--skip-slidegen-generation", action="store_true")
    parser.add_argument("--skip-original-generation", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    if args.count <= 0:
        raise SystemExit("--count must be positive.")
    if args.min_history_pairs <= 0:
        raise SystemExit("--min-history-pairs must be positive.")
    if args.profile_history_count <= 0:
        raise SystemExit("--profile-history-count must be positive.")

    python_bin = resolve_python()
    study_name = sanitize_name(args.study_name)
    args.output_root = args.output_root.resolve()
    args.original_repo_root = args.original_repo_root.resolve()

    study_dir = args.output_root / "generation_studies" / study_name
    study_dir.mkdir(parents=True, exist_ok=True)
    failures_path = study_dir / "failures.json"
    failures: list[dict[str, Any]] = []

    manifest_path = (args.manifest_path.resolve() if args.manifest_path else study_dir / "manifest.json")
    profile_dir = study_dir / "retrieval_profiles"
    slidegen_output_root = study_dir / "slidegen_outputs"
    original_output_root = study_dir / "original_outputs"
    include_paper_ids = list(dict.fromkeys(args.include_paper_id or DEFAULT_FIXED_PAPER_IDS))

    if not args.skip_manifest_build:
        manifest_command = build_manifest_command(
            python_bin=python_bin,
            count=args.count,
            seed=args.seed,
            min_author_paper_count=args.min_history_pairs,
            output_path=manifest_path,
            include_paper_ids=include_paper_ids,
        )
        manifest_ok = try_run_command(
            manifest_command,
            cwd=REPO_ROOT,
            dry_run=args.dry_run,
            stage="manifest_build",
            paper_id=None,
            failures=failures,
        )
        if not manifest_ok and not args.dry_run:
            failures_path.write_text(json.dumps(failures, indent=2), encoding="utf-8")
            raise SystemExit(f"Manifest build failed; see {failures_path}")

    if not args.dry_run and not manifest_path.exists():
        raise SystemExit(f"Manifest not found: {manifest_path}")

    manifest = {"selected_papers": []} if args.dry_run else load_manifest(manifest_path)
    selected_papers = list(manifest.get("selected_papers") or [])
    if not args.dry_run and not selected_papers:
        raise SystemExit(f"Manifest has no selected_papers: {manifest_path}")

    if not args.skip_profile_generation:
        profile_dir.mkdir(parents=True, exist_ok=True)
        for item in selected_papers:
            command = build_profile_command(
                python_bin=python_bin,
                author_id=str(item["author_id"]),
                paper_id=str(item["paper_id"]),
                output_dir=profile_dir,
                max_retrieved=args.profile_history_count,
                model=args.profile_model,
                retrieval_ranker=args.retrieval_ranker,
                force_refresh=args.force_refresh_profiles,
            )
            try_run_command(
                command,
                cwd=REPO_ROOT,
                dry_run=args.dry_run,
                stage="profile_generation",
                paper_id=str(item["paper_id"]),
                failures=failures,
            )

    if not args.skip_slidegen_generation:
        batch_command = build_slidegen_batch_command(
            python_bin=python_bin,
            manifest_path=manifest_path,
            experiment_name=study_name,
            output_root=slidegen_output_root,
            retrieval_profile_dir=profile_dir,
            model_t=args.model_name_t,
            model_v=args.model_name_v,
            formula_mode=args.formula_mode,
            outline_mode=args.outline_mode,
        )
        try_run_command(
            batch_command,
            cwd=REPO_ROOT,
            dry_run=args.dry_run,
            stage="slidegen_batch_generation",
            paper_id=None,
            failures=failures,
        )

    if not args.skip_original_generation:
        original_output_root.mkdir(parents=True, exist_ok=True)
        for item in selected_papers:
            command = build_original_command(
                python_bin=python_bin,
                original_repo_root=args.original_repo_root,
                paper_path=str(item["paper_path"]),
                paper_key=paper_key(str(item["paper_id"])),
                model_t=args.model_name_t,
                model_v=args.model_name_v,
                formula_mode=args.formula_mode,
                output_root=original_output_root,
            )
            try_run_command(
                command,
                cwd=args.original_repo_root,
                dry_run=args.dry_run,
                stage="original_generation",
                paper_id=str(item["paper_id"]),
                failures=failures,
            )

    output_summary = {
        "study_dir": str(study_dir),
        "manifest_path": str(manifest_path),
        "profile_dir": str(profile_dir),
        "slidegen_output_root": str(slidegen_output_root),
        "original_output_root": str(original_output_root),
        "count": len(selected_papers) if not args.dry_run else args.count,
        "included_paper_ids": include_paper_ids,
        "failure_count": len(failures),
        "failures_path": str(failures_path),
    }
    if not args.dry_run:
        failures_path.write_text(json.dumps(failures, indent=2), encoding="utf-8")
    print(json.dumps(output_summary, indent=2))


if __name__ == "__main__":
    main()
