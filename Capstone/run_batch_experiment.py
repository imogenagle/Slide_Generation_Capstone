#!/usr/bin/env python3
"""Run a sampled batch experiment for baseline and personalized deck generation."""

from __future__ import annotations

import argparse
import csv
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from Capstone.generate_random_decks import (
    DEFAULT_MANIFEST_DIR,
    DEFAULT_PAPER_AUTHORS_CSV,
    DEFAULT_PAPERS_CSV,
    DEFAULT_RAW_ROOT,
    append_outline_mode_suffix,
    load_candidates_from_papers_csv,
    load_primary_author_ids,
    output_dir_key,
    resolve_cli_path,
    save_manifest,
    scan_candidates_from_raw_root,
)


DEFAULT_EXPERIMENT_ROOT = PROJECT_ROOT / "Capstone" / "batch_runs" / "experiments"


def sanitize_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in value.strip())
    return cleaned.strip("_") or "batch_experiment"


def build_contents_dir_name(paper_id: str, paper_path: Path, outline_mode: str, personalized: bool) -> str:
    base = append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)
    return f"{base}_personalized" if personalized else base


def resolve_output_root(output_dir: str | None) -> Path:
    if not output_dir:
        return PROJECT_ROOT
    path = Path(output_dir)
    return path if path.is_absolute() else PROJECT_ROOT / path


def personalized_variant_suffix(personalization_mode: str) -> str:
    if personalization_mode == "retrieval":
        return "personalized_retrieval"
    return "personalized"


def personalized_folder_suffix(personalization_mode: str) -> str:
    if personalization_mode == "retrieval":
        return "_personalized_retrieval"
    return "_personalized"


def build_pptx_path(
    output_root: Path,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = (
        f"{append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)}{personalized_folder_suffix(personalization_mode)}"
        if personalized else build_contents_dir_name(paper_id, paper_path, outline_mode, personalized)
    )
    variant_suffix = personalized_variant_suffix(personalization_mode) if personalized else "baseline"
    return output_root / "contents" / contents_name / f"{model_name_t}_{model_name_v}_output_slides_{variant_suffix}.pptx"


def build_plan_path(
    output_root: Path,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = (
        f"{append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)}{personalized_folder_suffix(personalization_mode)}"
        if personalized else build_contents_dir_name(paper_id, paper_path, outline_mode, personalized)
    )
    variant_suffix = personalized_variant_suffix(personalization_mode) if personalized else "baseline"
    return output_root / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_slide_plan_{variant_suffix}.json"


def build_raw_content_path(
    output_root: Path,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = (
        f"{append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)}{personalized_folder_suffix(personalization_mode)}"
        if personalized else build_contents_dir_name(paper_id, paper_path, outline_mode, personalized)
    )
    return output_root / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_raw_content.json"


def build_figures_path(
    output_root: Path,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = (
        f"{append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)}{personalized_folder_suffix(personalization_mode)}"
        if personalized else build_contents_dir_name(paper_id, paper_path, outline_mode, personalized)
    )
    return output_root / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_figures.json"


def build_formula_match_path(
    output_root: Path,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = (
        f"{append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)}{personalized_folder_suffix(personalization_mode)}"
        if personalized else build_contents_dir_name(paper_id, paper_path, outline_mode, personalized)
    )
    return output_root / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_formula_match.json"


def is_nonempty_json_file(path: Path) -> bool:
    if not path.exists() or path.stat().st_size == 0:
        return False
    try:
        json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return False
    return True


def generation_outputs_ready(
    *,
    output_root: Path,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> bool:
    pptx_path = build_pptx_path(
        output_root, paper_id, paper_path, outline_mode, personalized, personalization_mode, model_name_t, model_name_v
    )
    if not pptx_path.exists() or pptx_path.stat().st_size == 0:
        return False

    required_jsons = [
        build_raw_content_path(
            output_root, paper_id, paper_path, outline_mode, personalized, personalization_mode, model_name_t, model_name_v
        ),
        build_plan_path(
            output_root, paper_id, paper_path, outline_mode, personalized, personalization_mode, model_name_t, model_name_v
        ),
        build_figures_path(
            output_root, paper_id, paper_path, outline_mode, personalized, personalization_mode, model_name_t, model_name_v
        ),
        build_formula_match_path(
            output_root, paper_id, paper_path, outline_mode, personalized, personalization_mode, model_name_t, model_name_v
        ),
    ]
    return all(is_nonempty_json_file(path) for path in required_jsons)


def resolve_manifest_path(experiment_name: str, manifest_path: Path | None, batch_run_root: Path) -> Path:
    if manifest_path is not None:
        return manifest_path
    return batch_run_root / f"{experiment_name}.manifest.json"


def load_candidates(raw_root: Path, papers_csv: Path) -> list[dict[str, str]]:
    candidates = load_candidates_from_papers_csv(papers_csv)
    if candidates:
        return candidates
    return scan_candidates_from_raw_root(raw_root)


def load_primary_author_paper_counts(paper_authors_csv: Path) -> dict[str, int]:
    if not paper_authors_csv.exists():
        return {}

    primary_by_paper: dict[str, tuple[int, str]] = {}
    with paper_authors_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            author_id = (row.get("author_id") or "").strip()
            try:
                author_order = int((row.get("author_order") or "999999").strip())
            except ValueError:
                author_order = 999999
            if not paper_id or not author_id:
                continue
            current = primary_by_paper.get(paper_id)
            if current is None or author_order < current[0]:
                primary_by_paper[paper_id] = (author_order, author_id)

    counts: dict[str, int] = {}
    for _, author_id in primary_by_paper.values():
        counts[author_id] = counts.get(author_id, 0) + 1
    return counts


def load_author_paper_index(paper_authors_csv: Path) -> tuple[dict[str, set[str]], dict[str, list[str]]]:
    if not paper_authors_csv.exists():
        return {}, {}

    paper_ids_by_author: dict[str, set[str]] = {}
    author_ids_by_paper: dict[str, list[str]] = {}
    with paper_authors_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            author_id = (row.get("author_id") or "").strip()
            if not paper_id or not author_id:
                continue
            paper_ids_by_author.setdefault(author_id, set()).add(paper_id)
            author_ids_by_paper.setdefault(paper_id, []).append(author_id)
    return paper_ids_by_author, author_ids_by_paper


def sample_papers(
    *,
    candidates: list[dict[str, str]],
    count: int,
    seed: int,
    split_filters: list[str] | None,
    primary_author_ids: dict[str, str],
    require_personalization: bool,
    outline_mode: str,
    include_existing: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
    min_primary_author_paper_count: int,
    primary_author_paper_counts: dict[str, int],
    distinct_primary_authors: bool,
    author_ids_by_paper: dict[str, list[str]],
    paper_ids_by_author: dict[str, set[str]],
    author_count_mode: str,
    output_root: Path,
) -> list[dict[str, str]]:
    import random

    allowed_splits = {value.strip() for value in (split_filters or []) if value and value.strip()}
    eligible: list[dict[str, str]] = []
    for item in candidates:
        paper_id = item["paper_id"]
        if allowed_splits and paper_id.split(":", 1)[0] not in allowed_splits:
            continue
        if require_personalization and paper_id not in primary_author_ids:
            continue
        if min_primary_author_paper_count > 0:
            if author_count_mode == "any":
                author_ids = author_ids_by_paper.get(paper_id, [])
                if not author_ids:
                    continue
                if not any(
                    len(paper_ids_by_author.get(author_id, set())) >= min_primary_author_paper_count
                    for author_id in author_ids
                ):
                    continue
            else:
                author_id = primary_author_ids.get(paper_id)
                if not author_id:
                    continue
                if primary_author_paper_counts.get(author_id, 0) < min_primary_author_paper_count:
                    continue
        if not include_existing:
            paper_path = Path(item["paper_path"])
            baseline_ready = generation_outputs_ready(
                output_root=output_root,
                paper_id=paper_id,
                paper_path=paper_path,
                outline_mode=outline_mode,
                personalized=False,
                personalization_mode=personalization_mode,
                model_name_t=model_name_t,
                model_name_v=model_name_v,
            )
            personalized_ready = generation_outputs_ready(
                output_root=output_root,
                paper_id=paper_id,
                paper_path=paper_path,
                outline_mode=outline_mode,
                personalized=True,
                personalization_mode=personalization_mode,
                model_name_t=model_name_t,
                model_name_v=model_name_v,
            )
            if baseline_ready or (require_personalization and personalized_ready):
                continue
        eligible.append(item)

    rng = random.Random(seed)
    if distinct_primary_authors:
        papers_by_author: dict[str, list[dict[str, str]]] = {}
        for item in eligible:
            paper_id = item["paper_id"]
            if author_count_mode == "any":
                candidate_author_ids = [
                    author_id
                    for author_id in author_ids_by_paper.get(paper_id, [])
                    if len(paper_ids_by_author.get(author_id, set())) >= min_primary_author_paper_count
                ]
            else:
                primary_author_id = primary_author_ids.get(paper_id)
                candidate_author_ids = [primary_author_id] if primary_author_id else []

            for author_id in candidate_author_ids:
                if not author_id:
                    continue
                item_with_author = dict(item)
                item_with_author["author_id"] = author_id
                papers_by_author.setdefault(author_id, []).append(item_with_author)

        if len(papers_by_author) < count:
            raise SystemExit(
                f"Requested {count} distinct primary authors, but only {len(papers_by_author)} authors are eligible "
                f"(personalization_required={require_personalization}, include_existing={include_existing}, "
                f"min_primary_author_paper_count={min_primary_author_paper_count}, author_count_mode={author_count_mode})."
            )

        selected_authors = rng.sample(sorted(papers_by_author), count)
        selected_items: list[dict[str, str]] = []
        for author_id in selected_authors:
            selected_items.append(rng.choice(papers_by_author[author_id]))
        return selected_items

    if len(eligible) < count:
        raise SystemExit(
            f"Requested {count} papers, but only {len(eligible)} are eligible "
            f"(personalization_required={require_personalization}, include_existing={include_existing})."
        )

    return rng.sample(eligible, count)


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    rendered = " ".join(str(part) for part in command)
    print(rendered)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def generation_command(
    *,
    paper_path: Path,
    author_id: str | None,
    author_profile_path: Path | None,
    personalized: bool,
    personalization_mode: str,
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
    outline_mode: str,
    preference_model: str,
    preference_max_papers: int,
    force_refresh_preferences: bool,
    output_dir: str | None,
) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "SlidesAgent.new_pipeline",
        "--paper_path",
        str(paper_path),
        "--model_name_t",
        model_name_t,
        "--model_name_v",
        model_name_v,
        "--formula_mode",
        str(formula_mode),
        "--outline_mode",
        outline_mode,
        "--personalization_mode",
        personalization_mode,
    ]
    if output_dir:
        command.extend(["--output_dir", output_dir])
    if personalized:
        if not author_id:
            raise ValueError("author_id is required for personalized generation")
        command.extend(
            [
                "--use_author_preferences",
                "--author_id",
                author_id,
            ]
        )
        if author_profile_path is not None:
            command.extend(["--author_profile_path", str(author_profile_path)])
        else:
            command.extend(
                [
                    "--preference_model",
                    preference_model,
                    "--preference_max_papers",
                    str(preference_max_papers),
                ]
            )
        if force_refresh_preferences:
            command.append("--force_refresh_preferences")
    return command


def resolve_profile_path(
    *,
    personalization_mode: str,
    author_id: str | None,
    paper_id: str,
    retrieval_profile_dir: Path | None,
) -> Path | None:
    if not author_id:
        return None
    if personalization_mode == "retrieval":
        if retrieval_profile_dir is None:
            raise ValueError(
                "--retrieval-profile-dir is required when --personalization-mode retrieval is used."
            )
        expected_name = f"{author_id}.{paper_id.replace(':', '_')}.retrieval.json"
        candidate = retrieval_profile_dir / expected_name
        if candidate.exists():
            return candidate
        matches = sorted(retrieval_profile_dir.glob(f"{author_id}.{paper_id.replace(':', '_')}*.retrieval.json"))
        if matches:
            return matches[0]
        raise FileNotFoundError(
            f"Retrieval profile not found for author_id={author_id}, paper_id={paper_id} under {retrieval_profile_dir}"
        )
    return PROJECT_ROOT / "Capstone" / "profiles" / f"{author_id}.json"


def bundle_eval_command(
    *,
    generated_pptx: Path,
    paper_id: str,
    output_dir: Path,
    core_coverage_model: str,
    judge_model: str,
    render_dpi: int,
    include_preference_dependent_slidetailor: bool,
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
    ]
    if include_preference_dependent_slidetailor:
        command.append("--include-preference-dependent-slidetailor")
    return command


def personalization_eval_command(
    *,
    profile_path: Path,
    baseline_plan: Path,
    personalized_plan: Path,
    output_path: Path,
    judge_model: str,
) -> list[str]:
    return [
        sys.executable,
        "Capstone/evaluate_personalization_alignment_pairwise.py",
        "--profile",
        str(profile_path),
        "--baseline-plan",
        str(baseline_plan),
        "--personalized-plan",
        str(personalized_plan),
        "--model",
        judge_model,
        "--output",
        str(output_path),
    ]


def retrieval_numeric_eval_command(
    *,
    profile_path: Path,
    baseline_plan: Path,
    personalized_plan: Path,
    output_path: Path,
) -> list[str]:
    return [
        sys.executable,
        "Capstone/evaluate_retrieval_alignment_numeric.py",
        "--profile",
        str(profile_path),
        "--baseline-plan",
        str(baseline_plan),
        "--personalized-plan",
        str(personalized_plan),
        "--output",
        str(output_path),
    ]


def retrieval_section_eval_command(
    *,
    profile_path: Path,
    baseline_plan: Path,
    personalized_plan: Path,
    output_path: Path,
    judge_model: str,
) -> list[str]:
    return [
        sys.executable,
        "Capstone/evaluate_retrieval_alignment_sections_llm.py",
        "--profile",
        str(profile_path),
        "--baseline-plan",
        str(baseline_plan),
        "--personalized-plan",
        str(personalized_plan),
        "--model",
        judge_model,
        "--output",
        str(output_path),
    ]


def retrieval_all_eval_command(
    *,
    profile_path: Path,
    baseline_plan: Path,
    personalized_plan: Path,
    output_path: Path,
    judge_model: str,
) -> list[str]:
    return [
        sys.executable,
        "Capstone/evaluate_retrieval_alignment_all.py",
        "--profile",
        str(profile_path),
        "--baseline-plan",
        str(baseline_plan),
        "--personalized-plan",
        str(personalized_plan),
        "--model",
        judge_model,
        "--output",
        str(output_path),
    ]


def resolve_experiment_dir(experiment_name: str) -> Path:
    return DEFAULT_EXPERIMENT_ROOT / experiment_name


def load_manifest(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_summary(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sample one cohort and run baseline/personalized generation plus evaluations."
    )
    parser.add_argument("--count", type=int, default=60, help="Number of papers to include when sampling.")
    parser.add_argument("--seed", type=int, default=42, help="Random seed used for cohort sampling.")
    parser.add_argument("--experiment-name", default=None, help="Optional stable name for the run folder.")
    parser.add_argument("--manifest-path", type=Path, default=None, help="Optional manifest JSON path.")
    parser.add_argument("--use-manifest", type=Path, default=None, help="Reuse an existing manifest instead of sampling.")
    parser.add_argument("--raw-root", type=Path, default=DEFAULT_RAW_ROOT)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--split", action="append", default=None, help="Restrict to one or more dataset splits.")
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--personalization-mode", choices=["standard", "retrieval"], default="standard")
    parser.add_argument(
        "--retrieval-profile-dir",
        type=Path,
        default=None,
        help="Directory containing retrieval profiles named like <author_id>.<paper_id_with_underscore>.retrieval.json.",
    )
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Base output directory for generated contents and Capstone batch artifacts.",
    )
    parser.add_argument("--preference-model", default="gpt-5.4-nano")
    parser.add_argument("--preference-max-papers", type=int, default=5)
    parser.add_argument(
        "--min-primary-author-paper-count",
        type=int,
        default=0,
        help="Only sample papers whose primary author appears as primary author on at least this many papers.",
    )
    parser.add_argument(
        "--distinct-primary-authors",
        action="store_true",
        help="Sample at most one paper per primary author, enforcing distinct authors across the cohort.",
    )
    parser.add_argument(
        "--author-count-mode",
        choices=["primary", "any"],
        default="primary",
        help="Whether author-history thresholds and distinct-author sampling should use only the primary author or any listed author.",
    )
    parser.add_argument("--force-refresh-preferences", action="store_true")
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--skip-baseline-generation", action="store_true")
    parser.add_argument("--skip-personalized-generation", action="store_true")
    parser.add_argument("--skip-bundle-eval", action="store_true")
    parser.add_argument("--skip-personalization-eval", action="store_true")
    parser.add_argument(
        "--run-retrieval-numeric-eval",
        action="store_true",
        help="Run the retrieval-only numeric evaluator after generation. Intended for personalization-mode retrieval.",
    )
    parser.add_argument(
        "--run-retrieval-section-eval",
        action="store_true",
        help="Run the retrieval-only LLM section evaluator after generation. Intended for personalization-mode retrieval.",
    )
    parser.add_argument(
        "--run-retrieval-all-eval",
        action="store_true",
        help="Run the combined retrieval evaluator (numeric + qualitative section) after generation, saving one JSON per paper.",
    )
    parser.add_argument(
        "--include-preference-dependent-slidetailor",
        action="store_true",
        help="Opt in to SlideTailor preference-based structure/template metrics during bundle eval.",
    )
    parser.add_argument("--include-existing", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    args.raw_root = resolve_cli_path(args.raw_root)
    args.papers_csv = resolve_cli_path(args.papers_csv)
    args.paper_authors_csv = resolve_cli_path(args.paper_authors_csv)
    args.manifest_path = resolve_cli_path(args.manifest_path)
    args.use_manifest = resolve_cli_path(args.use_manifest)
    args.retrieval_profile_dir = resolve_cli_path(args.retrieval_profile_dir)
    output_root = resolve_output_root(args.output_dir)
    batch_run_root = output_root / "Capstone" / "batch_runs"
    experiment_root = batch_run_root / "experiments"

    if args.count <= 0:
        raise SystemExit("--count must be positive")
    if args.min_primary_author_paper_count < 0:
        raise SystemExit("--min-primary-author-paper-count must be non-negative")

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    default_name = (
        f"cohort_{args.count}_seed{args.seed}_{args.outline_mode}_{args.model_name_t}_{args.model_name_v}_{timestamp}"
    )
    experiment_name = sanitize_name(args.experiment_name or default_name)
    experiment_dir = experiment_root / experiment_name
    experiment_dir.mkdir(parents=True, exist_ok=True)

    primary_author_ids = load_primary_author_ids(args.paper_authors_csv)
    primary_author_paper_counts = load_primary_author_paper_counts(args.paper_authors_csv)
    paper_ids_by_author, author_ids_by_paper = load_author_paper_index(args.paper_authors_csv)
    manifest_path = resolve_manifest_path(experiment_name, args.manifest_path, batch_run_root)

    if args.use_manifest:
        manifest = load_manifest(args.use_manifest)
        selected = list(manifest.get("selected_papers") or [])
        if not selected:
            raise SystemExit(f"Manifest has no selected_papers: {args.use_manifest}")
    else:
        candidates = load_candidates(args.raw_root, args.papers_csv)
        selected = sample_papers(
            candidates=candidates,
            count=args.count,
            seed=args.seed,
            split_filters=args.split,
            primary_author_ids=primary_author_ids,
            require_personalization=(not args.skip_personalized_generation) or (not args.skip_personalization_eval),
            outline_mode=args.outline_mode,
            include_existing=args.include_existing,
            personalization_mode=args.personalization_mode,
            model_name_t=args.model_name_t,
            model_name_v=args.model_name_v,
            min_primary_author_paper_count=args.min_primary_author_paper_count,
            primary_author_paper_counts=primary_author_paper_counts,
            distinct_primary_authors=args.distinct_primary_authors,
            author_ids_by_paper=author_ids_by_paper,
            paper_ids_by_author=paper_ids_by_author,
            author_count_mode=args.author_count_mode,
            output_root=output_root,
        )
        manifest = {
            "experiment_name": experiment_name,
            "count": len(selected),
            "seed": args.seed,
            "outline_mode": args.outline_mode,
            "model_name_t": args.model_name_t,
            "model_name_v": args.model_name_v,
            "formula_mode": args.formula_mode,
            "preference_model": args.preference_model,
            "preference_max_papers": args.preference_max_papers,
            "min_primary_author_paper_count": args.min_primary_author_paper_count,
            "distinct_primary_authors": args.distinct_primary_authors,
            "author_count_mode": args.author_count_mode,
            "selected_papers": selected,
        }
        save_manifest(manifest_path, manifest)
        print(f"Saved manifest to {manifest_path}")

    summary: dict[str, Any] = {
        "experiment_name": experiment_name,
        "experiment_dir": str(experiment_dir),
        "manifest_path": str(args.use_manifest or manifest_path),
        "selected_count": len(selected),
        "outline_mode": args.outline_mode,
        "model_name_t": args.model_name_t,
        "model_name_v": args.model_name_v,
        "papers": [],
    }

    for index, item in enumerate(selected, start=1):
        paper_id = str(item["paper_id"])
        paper_path = Path(item["paper_path"])
        author_id = item.get("author_id") or primary_author_ids.get(paper_id)
        print(f"[{index}/{len(selected)}] {paper_id}")

        baseline_pptx = build_pptx_path(
            output_root, paper_id, paper_path, args.outline_mode, False, args.personalization_mode, args.model_name_t, args.model_name_v
        )
        personalized_pptx = build_pptx_path(
            output_root, paper_id, paper_path, args.outline_mode, True, args.personalization_mode, args.model_name_t, args.model_name_v
        )
        baseline_plan = build_plan_path(
            output_root, paper_id, paper_path, args.outline_mode, False, args.personalization_mode, args.model_name_t, args.model_name_v
        )
        personalized_plan = build_plan_path(
            output_root, paper_id, paper_path, args.outline_mode, True, args.personalization_mode, args.model_name_t, args.model_name_v
        )
        profile_path = resolve_profile_path(
            personalization_mode=args.personalization_mode,
            author_id=author_id,
            paper_id=paper_id,
            retrieval_profile_dir=args.retrieval_profile_dir,
        ) if author_id else None

        paper_summary: dict[str, Any] = {
            "paper_id": paper_id,
            "paper_path": str(paper_path),
            "author_id": author_id,
            "baseline_pptx": str(baseline_pptx),
            "personalized_pptx": str(personalized_pptx),
            "statuses": {},
        }

        try:
            if not args.skip_baseline_generation:
                if generation_outputs_ready(
                    output_root=output_root,
                    paper_id=paper_id,
                    paper_path=paper_path,
                    outline_mode=args.outline_mode,
                    personalized=False,
                    personalization_mode=args.personalization_mode,
                    model_name_t=args.model_name_t,
                    model_name_v=args.model_name_v,
                ) and not args.dry_run:
                    paper_summary["statuses"]["baseline_generation"] = "skipped_existing"
                else:
                    run_command(
                        generation_command(
                            paper_path=paper_path,
                            author_id=None,
                            author_profile_path=None,
                            personalized=False,
                            personalization_mode=args.personalization_mode,
                            model_name_t=args.model_name_t,
                            model_name_v=args.model_name_v,
                            formula_mode=args.formula_mode,
                            outline_mode=args.outline_mode,
                            preference_model=args.preference_model,
                            preference_max_papers=args.preference_max_papers,
                            force_refresh_preferences=args.force_refresh_preferences,
                            output_dir=args.output_dir,
                        ),
                        cwd=PROJECT_ROOT,
                        dry_run=args.dry_run,
                    )
                    paper_summary["statuses"]["baseline_generation"] = "requested"

            if not args.skip_personalized_generation:
                if not author_id:
                    raise RuntimeError(f"Missing primary author_id for personalized run: {paper_id}")
                if generation_outputs_ready(
                    output_root=output_root,
                    paper_id=paper_id,
                    paper_path=paper_path,
                    outline_mode=args.outline_mode,
                    personalized=True,
                    personalization_mode=args.personalization_mode,
                    model_name_t=args.model_name_t,
                    model_name_v=args.model_name_v,
                ) and not args.dry_run:
                    paper_summary["statuses"]["personalized_generation"] = "skipped_existing"
                else:
                    run_command(
                        generation_command(
                            paper_path=paper_path,
                            author_id=author_id,
                            author_profile_path=profile_path,
                            personalized=True,
                            personalization_mode=args.personalization_mode,
                            model_name_t=args.model_name_t,
                            model_name_v=args.model_name_v,
                            formula_mode=args.formula_mode,
                            outline_mode=args.outline_mode,
                            preference_model=args.preference_model,
                            preference_max_papers=args.preference_max_papers,
                            force_refresh_preferences=args.force_refresh_preferences,
                            output_dir=args.output_dir,
                        ),
                        cwd=PROJECT_ROOT,
                        dry_run=args.dry_run,
                    )
                    paper_summary["statuses"]["personalized_generation"] = "requested"

            if not args.skip_bundle_eval:
                baseline_eval_dir = experiment_dir / "bundle_eval" / "baseline" / paper_id.replace(":", "_")
                personalized_eval_dir = experiment_dir / "bundle_eval" / "personalized" / paper_id.replace(":", "_")
                run_command(
                    bundle_eval_command(
                        generated_pptx=baseline_pptx,
                        paper_id=paper_id,
                        output_dir=baseline_eval_dir,
                        core_coverage_model=args.core_coverage_model,
                        judge_model=args.judge_model,
                        render_dpi=args.render_dpi,
                        include_preference_dependent_slidetailor=args.include_preference_dependent_slidetailor,
                    ),
                    cwd=PROJECT_ROOT,
                    dry_run=args.dry_run,
                )
                paper_summary["statuses"]["baseline_bundle_eval"] = "requested"
                should_run_personalized_bundle = (
                    author_id is not None
                    and (not args.skip_personalized_generation or personalized_pptx.exists() or args.dry_run)
                )
                if should_run_personalized_bundle:
                    run_command(
                        bundle_eval_command(
                            generated_pptx=personalized_pptx,
                            paper_id=paper_id,
                            output_dir=personalized_eval_dir,
                            core_coverage_model=args.core_coverage_model,
                            judge_model=args.judge_model,
                            render_dpi=args.render_dpi,
                            include_preference_dependent_slidetailor=args.include_preference_dependent_slidetailor,
                        ),
                        cwd=PROJECT_ROOT,
                        dry_run=args.dry_run,
                    )
                    paper_summary["statuses"]["personalized_bundle_eval"] = "requested"

            if not args.skip_personalization_eval:
                if not author_id or profile_path is None:
                    raise RuntimeError(f"Missing author profile context for personalization eval: {paper_id}")
                output_path = experiment_dir / "personalization_eval" / f"{paper_id.replace(':', '_')}.alignment.json"
                run_command(
                    personalization_eval_command(
                        profile_path=profile_path,
                        baseline_plan=baseline_plan,
                        personalized_plan=personalized_plan,
                        output_path=output_path,
                        judge_model=args.judge_model,
                    ),
                    cwd=PROJECT_ROOT,
                    dry_run=args.dry_run,
                )
                paper_summary["statuses"]["personalization_eval"] = "requested"

            if args.run_retrieval_numeric_eval:
                if args.personalization_mode != "retrieval":
                    raise RuntimeError("--run-retrieval-numeric-eval requires --personalization-mode retrieval")
                if not author_id or profile_path is None:
                    raise RuntimeError(f"Missing retrieval profile context for retrieval numeric eval: {paper_id}")
                output_path = experiment_dir / "retrieval_numeric_eval" / f"{paper_id.replace(':', '_')}.json"
                run_command(
                    retrieval_numeric_eval_command(
                        profile_path=profile_path,
                        baseline_plan=baseline_plan,
                        personalized_plan=personalized_plan,
                        output_path=output_path,
                    ),
                    cwd=PROJECT_ROOT,
                    dry_run=args.dry_run,
                )
                paper_summary["statuses"]["retrieval_numeric_eval"] = "requested"

            if args.run_retrieval_section_eval:
                if args.personalization_mode != "retrieval":
                    raise RuntimeError("--run-retrieval-section-eval requires --personalization-mode retrieval")
                if not author_id or profile_path is None:
                    raise RuntimeError(f"Missing retrieval profile context for retrieval section eval: {paper_id}")
                output_path = experiment_dir / "retrieval_section_eval" / f"{paper_id.replace(':', '_')}.json"
                run_command(
                    retrieval_section_eval_command(
                        profile_path=profile_path,
                        baseline_plan=baseline_plan,
                        personalized_plan=personalized_plan,
                        output_path=output_path,
                        judge_model=args.judge_model,
                    ),
                    cwd=PROJECT_ROOT,
                    dry_run=args.dry_run,
                )
                paper_summary["statuses"]["retrieval_section_eval"] = "requested"

            if args.run_retrieval_all_eval:
                if args.personalization_mode != "retrieval":
                    raise RuntimeError("--run-retrieval-all-eval requires --personalization-mode retrieval")
                if not author_id or profile_path is None:
                    raise RuntimeError(f"Missing retrieval profile context for retrieval combined eval: {paper_id}")
                output_path = experiment_dir / "retrieval_eval" / f"{paper_id.replace(':', '_')}.json"
                run_command(
                    retrieval_all_eval_command(
                        profile_path=profile_path,
                        baseline_plan=baseline_plan,
                        personalized_plan=personalized_plan,
                        output_path=output_path,
                        judge_model=args.judge_model,
                    ),
                    cwd=PROJECT_ROOT,
                    dry_run=args.dry_run,
                )
                paper_summary["statuses"]["retrieval_all_eval"] = "requested"

            paper_summary["status"] = "ok"
        except subprocess.CalledProcessError as exc:
            paper_summary["status"] = "failed"
            paper_summary["error"] = {
                "type": "CalledProcessError",
                "returncode": exc.returncode,
                "cmd": exc.cmd,
            }
        except Exception as exc:
            paper_summary["status"] = "failed"
            paper_summary["error"] = {
                "type": exc.__class__.__name__,
                "message": str(exc),
            }

        summary["papers"].append(paper_summary)
        write_summary(experiment_dir / "summary.json", summary)

    failures = [item for item in summary["papers"] if item.get("status") != "ok"]
    summary["failure_count"] = len(failures)
    summary["success_count"] = len(summary["papers"]) - len(failures)
    write_summary(experiment_dir / "summary.json", summary)
    print(
        json.dumps(
            {
                "experiment_name": experiment_name,
                "experiment_dir": str(experiment_dir),
                "selected_count": len(summary["papers"]),
                "success_count": summary["success_count"],
                "failure_count": summary["failure_count"],
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
