#!/usr/bin/env python3
"""Run pair-guided generation plus general and personalization evaluation for one cohort."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_BATCH_DIR = PROJECT_ROOT / "Capstone" / "batch_runs" / "pair_guideline_runs"
DEFAULT_CONTENTS_DIR = PROJECT_ROOT / "contents"

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from Capstone.generate_random_decks import append_outline_mode_suffix, output_dir_key


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    rendered = " ".join(command)
    print(rendered)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def generated_contents_dir_for_item(item: dict, outline_mode: str) -> Path:
    paper_id = str(item["paper_id"])
    paper_path = Path(str(item["paper_path"]))
    folder_name = append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)
    return DEFAULT_CONTENTS_DIR / folder_name


def reset_path(path: Path) -> None:
    if not path.exists() and not path.is_symlink():
        return
    if path.is_symlink() or path.is_file():
        path.unlink()
        return
    shutil.rmtree(path)


def link_path(src: Path, dst: Path) -> None:
    if not src.exists():
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    reset_path(dst)
    dst.symlink_to(src.resolve(), target_is_directory=src.is_dir())


def write_collection_index(collection_dir: Path, *, experiment_name: str, run_dir: Path, manifest: dict) -> None:
    payload = {
        "experiment_name": experiment_name,
        "run_dir": str(run_dir),
        "manifest_path": str(run_dir / "manifest.json"),
        "selected_count": len(manifest.get("selected_papers") or []),
    }
    (collection_dir / "RUN_INFO.json").write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def sync_contents_collection(
    *,
    experiment_name: str,
    collection_name: str,
    manifest_path: Path,
    dry_run: bool,
) -> None:
    run_dir = manifest_path.parent
    collection_dir = DEFAULT_CONTENTS_DIR / collection_name

    if dry_run:
        print(f"[dry-run] would sync run outputs into {collection_dir}")
        return

    manifest = load_json(manifest_path)
    outline_mode = str(manifest.get("outline_mode") or "high_level")
    collection_dir.mkdir(parents=True, exist_ok=True)

    link_path(manifest_path, collection_dir / "manifest.json")
    write_collection_index(collection_dir, experiment_name=experiment_name, run_dir=run_dir, manifest=manifest)

    for summary_name in ("summary.json", "summary.partial.json", "BUNDLE_EVAL_SUMMARY.json", "BUNDLE_EVAL_DETAILS.csv"):
        link_path(run_dir / summary_name, collection_dir / summary_name)

    link_path(run_dir / "eval", collection_dir / "eval")

    generated_root = collection_dir / "generated"
    bundle_eval_root = collection_dir / "bundle_eval"
    generated_root.mkdir(parents=True, exist_ok=True)
    bundle_eval_root.mkdir(parents=True, exist_ok=True)

    for item in manifest.get("selected_papers") or []:
        paper_id = str(item["paper_id"])
        paper_key = paper_id.replace(":", "_")
        generated_dir = generated_contents_dir_for_item(item, outline_mode)
        link_path(generated_dir, generated_root / paper_key)
        link_path(run_dir / paper_key / "bundle_eval", bundle_eval_root / paper_key)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Run the full pair-guided batch workflow: generation, pair-guideline eval, and bundle eval."
    )
    parser.add_argument("--count", type=int, default=50)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--experiment-name", required=True)
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--pair-guideline-model", default="gpt-5.4-nano")
    parser.add_argument("--pair-guideline-max-pairs", type=int, default=2)
    parser.add_argument("--pair-guideline-candidate-pool", type=int, default=5)
    parser.add_argument("--min-author-paper-count", type=int, default=5)
    parser.add_argument("--author-role-scope", choices=["any", "primary"], default="any")
    parser.add_argument(
        "--contents-collection-dir",
        default=None,
        help="Optional folder name to create under contents/ with links to generated outputs and eval summaries.",
    )
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--include-preference-dependent-slidetailor", action="store_true")
    parser.add_argument("--force-refresh-pair-guidelines", action="store_true")
    parser.add_argument("--include-existing", action="store_true")
    parser.add_argument("--skip-generation", action="store_true")
    parser.add_argument("--skip-personalization-eval", action="store_true")
    parser.add_argument("--skip-general-eval", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    manifest_path = DEFAULT_BATCH_DIR / args.experiment_name / "manifest.json"

    if not args.skip_generation:
        generation_command = [
            sys.executable,
            "Capstone/run_pair_guideline_batch.py",
            "--experiment-name",
            args.experiment_name,
            "--model-name-t",
            args.model_name_t,
            "--model-name-v",
            args.model_name_v,
            "--formula-mode",
            str(args.formula_mode),
            "--outline-mode",
            args.outline_mode,
            "--pair-guideline-model",
            args.pair_guideline_model,
            "--pair-guideline-max-pairs",
            str(args.pair_guideline_max_pairs),
            "--pair-guideline-candidate-pool",
            str(args.pair_guideline_candidate_pool),
            "--min-author-paper-count",
            str(args.min_author_paper_count),
            "--author-role-scope",
            args.author_role_scope,
        ]
        if manifest_path.exists():
            generation_command.extend(["--use-manifest", str(manifest_path)])
        else:
            generation_command.extend([
                "--count",
                str(args.count),
                "--seed",
                str(args.seed),
            ])
        if args.force_refresh_pair_guidelines:
            generation_command.append("--force-refresh-pair-guidelines")
        if args.include_existing:
            generation_command.append("--include-existing")
        run_command(generation_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)
        if args.contents_collection_dir:
            sync_contents_collection(
                experiment_name=args.experiment_name,
                collection_name=args.contents_collection_dir,
                manifest_path=manifest_path,
                dry_run=args.dry_run,
            )

    if not manifest_path.exists() and not args.dry_run:
        raise SystemExit(f"Manifest not found after generation stage: {manifest_path}")

    if not args.skip_personalization_eval:
        personalization_eval_command = [
            sys.executable,
            "Capstone/run_pair_guideline_eval_batch.py",
            "--manifest",
            str(manifest_path),
            "--judge-model",
            args.judge_model,
            "--request-timeout",
            str(args.request_timeout),
            "--generate-missing-baselines",
            "--generate-missing-pair-contexts",
        ]
        if args.include_existing:
            personalization_eval_command.append("--include-existing")
        if args.verbose:
            personalization_eval_command.append("--verbose")
        run_command(personalization_eval_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)
        if args.contents_collection_dir:
            sync_contents_collection(
                experiment_name=args.experiment_name,
                collection_name=args.contents_collection_dir,
                manifest_path=manifest_path,
                dry_run=args.dry_run,
            )

    if not args.skip_general_eval:
        general_eval_command = [
            sys.executable,
            "Capstone/run_pair_guideline_bundle_eval_batch.py",
            "--manifest",
            str(manifest_path),
            "--core-coverage-model",
            args.core_coverage_model,
            "--judge-model",
            args.judge_model,
            "--render-dpi",
            str(args.render_dpi),
            "--request-timeout",
            str(args.request_timeout),
            "--generate-missing-baselines",
        ]
        if args.include_preference_dependent_slidetailor:
            general_eval_command.append("--include-preference-dependent-slidetailor")
        if args.include_existing:
            general_eval_command.append("--include-existing")
        if args.verbose:
            general_eval_command.append("--verbose")
        run_command(general_eval_command, cwd=PROJECT_ROOT, dry_run=args.dry_run)
        if args.contents_collection_dir:
            sync_contents_collection(
                experiment_name=args.experiment_name,
                collection_name=args.contents_collection_dir,
                manifest_path=manifest_path,
                dry_run=args.dry_run,
            )

    if args.contents_collection_dir and (args.dry_run or manifest_path.exists()):
        sync_contents_collection(
            experiment_name=args.experiment_name,
            collection_name=args.contents_collection_dir,
            manifest_path=manifest_path,
            dry_run=args.dry_run,
        )


if __name__ == "__main__":
    main()
