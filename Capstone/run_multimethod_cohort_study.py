#!/usr/bin/env python3
"""Run a multi-method cohort study over sampled paper/PPT pairs.

This script:
1. Samples a cohort of papers whose selected author has enough historical paper-PPT pairs.
2. Builds retrieval-conditioned personalization profiles for those papers.
3. Generates SlideGen baseline (high-level), SlideGen personalized, SlideGen technical, and SlideGen Original decks.
4. Runs bundle eval for all four methods.
5. Summarizes bundle eval and retrieval-personalization eval into reusable tables.
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import subprocess
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
PYTHON = REPO_ROOT / ".venv" / "bin" / "python3"
ORIGINAL_REPO_ROOT = REPO_ROOT / "SlideGen_Original"

if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.generate_random_decks import (  # noqa: E402
    DEFAULT_PAPER_AUTHORS_CSV,
    DEFAULT_PAPERS_CSV,
    save_manifest,
)


BUNDLE_METRICS = [
    "core_coverage_topic_iou",
    "geometry_aware_density_gad_geom",
    "visual_appeal_deck_score",
    "logical_flow_deck_score",
    "paper_faithfulness_deck_score",
]


def resolve_python() -> Path:
    return PYTHON if PYTHON.exists() else Path(sys.executable)


def sanitize_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in value.strip())
    return cleaned.strip("_") or "cohort_study"


def output_dir_key(paper_id: str) -> str:
    return paper_id.replace(":", "_")


def normalize_bool(value: str) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "y"}


def maybe_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except Exception:
        return default


def load_paired_candidates(papers_csv: Path) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            paper_path_raw = (row.get("paper_pdf_path") or "").strip()
            title = (row.get("paper_title") or row.get("ppt_title") or "").strip()
            if not paper_id or not paper_path_raw:
                continue
            if not normalize_bool(row.get("paper_pdf_exists")):
                continue
            if maybe_int(row.get("slide_image_count"), default=0) <= 0:
                continue
            paper_path = Path(paper_path_raw)
            if not paper_path.is_absolute():
                paper_path = REPO_ROOT.parent / paper_path
            if not paper_path.exists():
                continue
            candidates.append(
                {
                    "paper_id": paper_id,
                    "paper_path": str(paper_path.resolve()),
                    "title": title,
                    "slide_image_count": maybe_int(row.get("slide_image_count"), default=0),
                }
            )
    return candidates


def load_authorship(paper_authors_csv: Path) -> tuple[dict[str, list[dict[str, Any]]], dict[str, set[str]]]:
    authors_by_paper: dict[str, list[dict[str, Any]]] = defaultdict(list)
    papers_by_author: dict[str, set[str]] = defaultdict(set)
    with paper_authors_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            author_id = (row.get("author_id") or "").strip()
            author_name = (row.get("author_name") or "").strip()
            author_order = maybe_int(row.get("author_order"), default=999999)
            if not paper_id or not author_id:
                continue
            authors_by_paper[paper_id].append(
                {
                    "author_id": author_id,
                    "author_name": author_name,
                    "author_order": author_order,
                }
            )
            papers_by_author[author_id].add(paper_id)
    for paper_id in authors_by_paper:
        authors_by_paper[paper_id].sort(key=lambda item: (item["author_order"], item["author_id"]))
    return authors_by_paper, papers_by_author


def sample_manifest(
    *,
    papers_csv: Path,
    paper_authors_csv: Path,
    count: int,
    seed: int,
    min_history_pairs: int,
    distinct_authors: bool,
) -> dict[str, Any]:
    candidates = load_paired_candidates(papers_csv)
    authors_by_paper, papers_by_author = load_authorship(paper_authors_csv)
    paired_paper_ids = {item["paper_id"] for item in candidates}

    eligible_entries: list[dict[str, Any]] = []
    for item in candidates:
        paper_id = item["paper_id"]
        eligible_authors: list[dict[str, Any]] = []
        for author in authors_by_paper.get(paper_id, []):
            author_id = str(author["author_id"])
            historical_pairs = len((papers_by_author.get(author_id, set()) & paired_paper_ids) - {paper_id})
            if historical_pairs >= min_history_pairs:
                enriched = dict(author)
                enriched["historical_pair_count"] = historical_pairs
                eligible_authors.append(enriched)
        if not eligible_authors:
            continue
        eligible_authors.sort(
            key=lambda author: (
                -int(author["historical_pair_count"]),
                int(author["author_order"]),
                str(author["author_id"]),
            )
        )
        selected_author = eligible_authors[0]
        eligible_entries.append(
            {
                "paper_id": paper_id,
                "paper_path": item["paper_path"],
                "title": item["title"],
                "paper_key": output_dir_key(paper_id),
                "slide_image_count": item["slide_image_count"],
                "author_id": selected_author["author_id"],
                "author_name": selected_author["author_name"],
                "historical_pair_count": selected_author["historical_pair_count"],
                "eligible_authors": eligible_authors,
            }
        )

    if distinct_authors:
        papers_by_selected_author: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for entry in eligible_entries:
            papers_by_selected_author[str(entry["author_id"])].append(entry)
        if len(papers_by_selected_author) < count:
            raise SystemExit(
                f"Requested {count} distinct authors, but only {len(papers_by_selected_author)} eligible authors were found."
            )
        rng = random.Random(seed)
        chosen_author_ids = rng.sample(sorted(papers_by_selected_author), count)
        selected = [rng.choice(papers_by_selected_author[author_id]) for author_id in chosen_author_ids]
    else:
        if len(eligible_entries) < count:
            raise SystemExit(f"Requested {count} papers, but only {len(eligible_entries)} eligible papers were found.")
        rng = random.Random(seed)
        selected = rng.sample(eligible_entries, count)

    selected_papers: list[dict[str, Any]] = []
    for entry in selected:
        selected_papers.append(
            {
                "paper_id": entry["paper_id"],
                "paper_path": entry["paper_path"],
                "title": entry["title"],
                "paper_key": entry["paper_key"],
                "author_id": entry["author_id"],
                "author_name": entry["author_name"],
                "historical_pair_count": entry["historical_pair_count"],
                "slide_image_count": entry["slide_image_count"],
                "eligible_author_ids": [author["author_id"] for author in entry["eligible_authors"]],
            }
        )

    return {
        "count": len(selected_papers),
        "seed": seed,
        "min_history_pairs": min_history_pairs,
        "distinct_authors": distinct_authors,
        "selected_papers": selected_papers,
    }


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    print("Running:", " ".join(command), flush=True)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def pick_existing(paths: list[Path]) -> Path:
    for path in paths:
        if path.exists():
            return path
    tried = "\n".join(str(path) for path in paths)
    raise FileNotFoundError(f"None of the candidate files exist:\n{tried}")


def slidegen_baseline_pptx(output_root: Path, paper_key: str, outline_mode: str, model_t: str, model_v: str) -> Path:
    base = output_root / "contents" / f"{paper_key}_{outline_mode}"
    return pick_existing(
        [
            base / f"{model_t}_{model_v}_output_slides_baseline_themed.pptx",
            base / f"{model_t}_{model_v}_output_slides_baseline.pptx",
        ]
    )


def slidegen_personalized_pptx(output_root: Path, paper_key: str, model_t: str, model_v: str) -> Path:
    base = output_root / "contents" / f"{paper_key}_high_level_personalized_retrieval"
    return pick_existing(
        [
            base / f"{model_t}_{model_v}_output_slides_personalized_retrieval_themed.pptx",
            base / f"{model_t}_{model_v}_output_slides_personalized_retrieval.pptx",
        ]
    )


def slidegen_plan_path(
    output_root: Path,
    paper_key: str,
    variant_dir: str,
    variant_suffix: str,
    model_t: str,
    model_v: str,
) -> Path:
    return output_root / "contents" / variant_dir / f"<{model_t}_{model_v}>_slide_plan_{variant_suffix}.json"


def original_pptx(output_root: Path, paper_key: str, model_t: str, model_v: str) -> Path:
    base = output_root / "contents" / f"{paper_key}_original"
    return pick_existing(
        [
            base / f"{model_t}_{model_v}_output_slides_themed.pptx",
            base / f"{model_t}_{model_v}_output_slides.pptx",
        ]
    )


def build_profile_command(
    *,
    python_bin: Path,
    paper_id: str,
    author_id: str,
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


def build_slidegen_generation_command(
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
    retrieval_eval: bool,
    include_personalized: bool,
) -> list[str]:
    command = [
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
        "--output-dir",
        str(output_root),
        "--skip-bundle-eval",
        "--skip-personalization-eval",
    ]
    if include_personalized:
        command.extend(
            [
                "--personalization-mode",
                "retrieval",
                "--retrieval-profile-dir",
                str(retrieval_profile_dir),
            ]
        )
        if retrieval_eval:
            command.append("--run-retrieval-all-eval")
    else:
        command.append("--skip-personalized-generation")
    return command


def build_original_generation_command(
    *,
    python_bin: Path,
    paper_path: Path,
    paper_key: str,
    output_root: Path,
    model_t: str,
    model_v: str,
    formula_mode: int,
) -> list[str]:
    return [
        str(python_bin),
        "-m",
        "SlidesAgent.new_pipeline",
        "--paper_path",
        str(paper_path),
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


def build_bundle_eval_command(
    *,
    python_bin: Path,
    generated_pptx: Path,
    paper_id: str,
    paper_path: Path,
    output_dir: Path,
    judge_model: str,
    core_coverage_model: str,
    render_dpi: int,
) -> list[str]:
    return [
        str(python_bin),
        "Capstone/evaluate_pptx_bundle.py",
        "--generated-pptx",
        str(generated_pptx),
        "--paper-id",
        paper_id,
        "--source-document",
        str(paper_path),
        "--judge-model",
        judge_model,
        "--core-coverage-model",
        core_coverage_model,
        "--render-dpi",
        str(render_dpi),
        "--output-dir",
        str(output_dir),
    ]


def build_bundle_compare_command(
    *,
    python_bin: Path,
    original_root: Path,
    baseline_root: Path,
    technical_root: Path,
    personalized_root: Path,
    output_dir: Path,
) -> list[str]:
    command = [
        str(python_bin),
        "Capstone/compare_bundle_eval_methods.py",
        "--method",
        f"SlideGen_Original={original_root}",
        "--method",
        f"SlideGen_Baseline={baseline_root}",
        "--method",
        f"SlideGen_Technical={technical_root}",
        "--method",
        f"SlideGen_Personalized={personalized_root}",
        "--output-dir",
        str(output_dir),
        "--metrics",
        *BUNDLE_METRICS,
    ]
    return command


def build_retrieval_summary_command(*, python_bin: Path, eval_dir: Path, output_dir: Path) -> list[str]:
    return [
        str(python_bin),
        "Capstone/summarize_retrieval_eval.py",
        "--eval-dir",
        str(eval_dir),
        "--output-dir",
        str(output_dir),
    ]


def main() -> None:
    parser = argparse.ArgumentParser(description="Run a 10-paper multi-method study with generation and eval summaries.")
    parser.add_argument("--count", type=int, default=10)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--study-name", default="ten_paper_multimethod_study")
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--min-history-pairs", type=int, default=3)
    parser.add_argument("--distinct-authors", action="store_true")
    parser.add_argument("--profile-history-count", type=int, default=3)
    parser.add_argument("--profile-model", default="gpt-5.4-nano")
    parser.add_argument("--retrieval-ranker", choices=["title_similarity", "llm_title_abstract"], default="llm_title_abstract")
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--output-root", type=Path, default=REPO_ROOT / "outputs")
    parser.add_argument("--original-repo-root", type=Path, default=ORIGINAL_REPO_ROOT)
    parser.add_argument("--force-refresh-profiles", action="store_true")
    parser.add_argument("--skip-profile-generation", action="store_true")
    parser.add_argument("--skip-slidegen-highlevel-generation", action="store_true")
    parser.add_argument("--skip-slidegen-technical-generation", action="store_true")
    parser.add_argument("--skip-original-generation", action="store_true")
    parser.add_argument("--skip-bundle-eval", action="store_true")
    parser.add_argument("--skip-retrieval-summary", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    python_bin = resolve_python()
    if args.count <= 0:
        raise SystemExit("--count must be positive")
    if args.min_history_pairs <= 0:
        raise SystemExit("--min-history-pairs must be positive")
    if args.profile_history_count <= 0:
        raise SystemExit("--profile-history-count must be positive")

    args.papers_csv = args.papers_csv.resolve()
    args.paper_authors_csv = args.paper_authors_csv.resolve()
    args.output_root = args.output_root.resolve()
    args.original_repo_root = args.original_repo_root.resolve()

    study_name = sanitize_name(args.study_name)
    study_dir = args.output_root / "cohort_studies" / study_name
    study_dir.mkdir(parents=True, exist_ok=True)

    required_history_pairs = max(args.min_history_pairs, args.profile_history_count)

    manifest = sample_manifest(
        papers_csv=args.papers_csv,
        paper_authors_csv=args.paper_authors_csv,
        count=args.count,
        seed=args.seed,
        min_history_pairs=required_history_pairs,
        distinct_authors=args.distinct_authors,
    )
    manifest.update(
        {
            "study_name": study_name,
            "requested_min_history_pairs": args.min_history_pairs,
            "effective_min_history_pairs": required_history_pairs,
            "profile_history_count": args.profile_history_count,
            "profile_model": args.profile_model,
            "retrieval_ranker": args.retrieval_ranker,
        }
    )
    manifest_path = study_dir / "manifest.json"
    save_manifest(manifest_path, manifest)
    print(f"Saved study manifest to {manifest_path}", flush=True)

    profiles_dir = study_dir / "retrieval_profiles"
    highlevel_experiment = f"{study_name}_highlevel"
    technical_experiment = f"{study_name}_technical"
    highlevel_experiment_dir = args.output_root / "Capstone" / "batch_runs" / "experiments" / highlevel_experiment
    technical_experiment_dir = args.output_root / "Capstone" / "batch_runs" / "experiments" / technical_experiment
    original_output_root = study_dir / "original_slidegen_runs"
    bundle_eval_root = study_dir / "bundle_eval"
    bundle_summary_dir = study_dir / "bundle_eval_summary"
    retrieval_summary_dir = study_dir / "retrieval_eval_summary"

    if not args.skip_profile_generation:
        for item in manifest["selected_papers"]:
            run_command(
                build_profile_command(
                    python_bin=python_bin,
                    paper_id=str(item["paper_id"]),
                    author_id=str(item["author_id"]),
                    output_dir=profiles_dir,
                    max_retrieved=args.profile_history_count,
                    model=args.profile_model,
                    retrieval_ranker=args.retrieval_ranker,
                    force_refresh=args.force_refresh_profiles,
                ),
                cwd=REPO_ROOT,
                dry_run=args.dry_run,
            )

    if not args.skip_slidegen_highlevel_generation:
        run_command(
            build_slidegen_generation_command(
                python_bin=python_bin,
                manifest_path=manifest_path,
                experiment_name=highlevel_experiment,
                output_root=args.output_root,
                retrieval_profile_dir=profiles_dir,
                model_t=args.model_name_t,
                model_v=args.model_name_v,
                formula_mode=args.formula_mode,
                outline_mode="high_level",
                retrieval_eval=True,
                include_personalized=True,
            ),
            cwd=REPO_ROOT,
            dry_run=args.dry_run,
        )

    if not args.skip_slidegen_technical_generation:
        run_command(
            build_slidegen_generation_command(
                python_bin=python_bin,
                manifest_path=manifest_path,
                experiment_name=technical_experiment,
                output_root=args.output_root,
                retrieval_profile_dir=profiles_dir,
                model_t=args.model_name_t,
                model_v=args.model_name_v,
                formula_mode=args.formula_mode,
                outline_mode="technical",
                retrieval_eval=False,
                include_personalized=False,
            ),
            cwd=REPO_ROOT,
            dry_run=args.dry_run,
        )

    if not args.skip_original_generation:
        if not args.original_repo_root.exists():
            raise FileNotFoundError(f"SlideGen_Original repo not found: {args.original_repo_root}")
        for item in manifest["selected_papers"]:
            paper_key = str(item["paper_key"])
            if not args.dry_run:
                try:
                    existing_original = original_pptx(
                        original_output_root,
                        paper_key,
                        args.model_name_t,
                        args.model_name_v,
                    )
                    print(f"Skipping existing original deck: {existing_original}", flush=True)
                    continue
                except FileNotFoundError:
                    pass
            run_command(
                build_original_generation_command(
                    python_bin=python_bin,
                    paper_path=Path(str(item["paper_path"])),
                    paper_key=paper_key,
                    output_root=original_output_root,
                    model_t=args.model_name_t,
                    model_v=args.model_name_v,
                    formula_mode=args.formula_mode,
                ),
                cwd=args.original_repo_root,
                dry_run=args.dry_run,
            )

    if not args.skip_bundle_eval:
        for item in manifest["selected_papers"]:
            paper_id = str(item["paper_id"])
            paper_key = str(item["paper_key"])
            paper_path = Path(str(item["paper_path"]))

            original_pptx_path = original_pptx(original_output_root, paper_key, args.model_name_t, args.model_name_v)
            baseline_pptx_path = slidegen_baseline_pptx(args.output_root, paper_key, "high_level", args.model_name_t, args.model_name_v)
            technical_pptx_path = slidegen_baseline_pptx(args.output_root, paper_key, "technical", args.model_name_t, args.model_name_v)
            personalized_pptx_path = slidegen_personalized_pptx(args.output_root, paper_key, args.model_name_t, args.model_name_v)

            for method_label, pptx_path in (
                ("original", original_pptx_path),
                ("baseline", baseline_pptx_path),
                ("technical", technical_pptx_path),
                ("personalized", personalized_pptx_path),
            ):
                run_command(
                    build_bundle_eval_command(
                        python_bin=python_bin,
                        generated_pptx=pptx_path,
                        paper_id=paper_id,
                        paper_path=paper_path,
                        output_dir=bundle_eval_root / method_label / paper_key,
                        judge_model=args.judge_model,
                        core_coverage_model=args.core_coverage_model,
                        render_dpi=args.render_dpi,
                    ),
                    cwd=REPO_ROOT,
                    dry_run=args.dry_run,
                )

        run_command(
            build_bundle_compare_command(
                python_bin=python_bin,
                original_root=bundle_eval_root / "original",
                baseline_root=bundle_eval_root / "baseline",
                technical_root=bundle_eval_root / "technical",
                personalized_root=bundle_eval_root / "personalized",
                output_dir=bundle_summary_dir,
            ),
            cwd=REPO_ROOT,
            dry_run=args.dry_run,
        )

    if not args.skip_retrieval_summary:
        run_command(
            build_retrieval_summary_command(
                python_bin=python_bin,
                eval_dir=highlevel_experiment_dir / "retrieval_eval",
                output_dir=retrieval_summary_dir,
            ),
            cwd=REPO_ROOT,
            dry_run=args.dry_run,
        )

    summary = {
        "study_name": study_name,
        "manifest_path": str(manifest_path),
        "profiles_dir": str(profiles_dir),
        "highlevel_experiment_dir": str(highlevel_experiment_dir),
        "technical_experiment_dir": str(technical_experiment_dir),
        "original_output_root": str(original_output_root),
        "bundle_eval_root": str(bundle_eval_root),
        "bundle_summary_dir": str(bundle_summary_dir),
        "retrieval_summary_dir": str(retrieval_summary_dir),
    }
    (study_dir / "study_summary.json").write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps(summary, indent=2), flush=True)


if __name__ == "__main__":
    main()
