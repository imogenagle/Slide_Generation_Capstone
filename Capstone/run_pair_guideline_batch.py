#!/usr/bin/env python3
"""Sample and run a pair-guided SlideGen batch experiment."""

from __future__ import annotations

import argparse
import csv
import json
import random
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_PAPERS_CSV = PROJECT_ROOT / "Capstone" / "author_tables" / "papers.csv"
DEFAULT_PAPER_AUTHORS_CSV = PROJECT_ROOT / "Capstone" / "author_tables" / "paper_authors.csv"
DEFAULT_BATCH_DIR = PROJECT_ROOT / "Capstone" / "batch_runs" / "pair_guideline_runs"

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from Capstone.generate_random_decks import append_outline_mode_suffix, output_dir_key, resolve_cli_path


def sanitize_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in value.strip())
    return cleaned.strip("_") or "pair_guideline_batch"


def resolve_repo_path(raw_path_value: str) -> Path:
    path = Path(raw_path_value)
    if path.exists():
        return path
    if path.is_absolute():
        return path
    if path.parts and path.parts[0] == "SlideGen":
        candidate = PROJECT_ROOT / Path(*path.parts[1:])
        if candidate.exists():
            return candidate
    return PROJECT_ROOT / path


def load_paired_paper_rows(papers_csv: Path) -> list[dict[str, Any]]:
    paired_rows: list[dict[str, Any]] = []
    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            paper_pdf_path_raw = (row.get("paper_pdf_path") or "").strip()
            raw_dir_raw = (row.get("raw_dir") or "").strip()
            if not paper_id or not paper_pdf_path_raw or not raw_dir_raw:
                continue

            paper_pdf_path = resolve_repo_path(paper_pdf_path_raw)
            raw_dir = resolve_repo_path(raw_dir_raw)
            if not paper_pdf_path.exists() or not raw_dir.exists():
                continue

            paired_rows.append(
                {
                    "paper_id": paper_id,
                    "paper_path": str(paper_pdf_path.resolve()),
                    "title": (row.get("paper_title") or row.get("ppt_title") or "").strip(),
                    "split": (row.get("split") or "").strip(),
                    "record_id": (row.get("record_id") or "").strip(),
                    "raw_dir": str(raw_dir.resolve()),
                    "slide_image_count": int((row.get("slide_image_count") or "0").strip() or 0),
                }
            )
    return paired_rows


def load_paper_authors_by_paper(paper_authors_csv: Path) -> dict[str, list[tuple[int, str]]]:
    authors_by_paper: dict[str, list[tuple[int, str]]] = {}
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
            authors_by_paper.setdefault(paper_id, []).append((author_order, author_id))
    for paper_id, author_rows in authors_by_paper.items():
        authors_by_paper[paper_id] = sorted(author_rows, key=lambda item: (item[0], item[1]))
    return authors_by_paper


def load_author_pair_counts(
    paper_authors_csv: Path,
    paired_paper_ids: set[str],
    *,
    author_role_scope: str,
) -> dict[str, int]:
    authors_by_paper = load_paper_authors_by_paper(paper_authors_csv)
    counts: Counter[str] = Counter()
    for paper_id in paired_paper_ids:
        author_rows = authors_by_paper.get(paper_id, [])
        if author_role_scope == "primary":
            author_rows = author_rows[:1]
        for _author_order, author_id in author_rows:
            counts[author_id] += 1
    return dict(counts)


def pair_guided_contents_name(paper_id: str, paper_path: Path, outline_mode: str) -> str:
    return append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)


def pair_guided_pptx_path(
    *,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = pair_guided_contents_name(paper_id, paper_path, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"{model_name_t}_{model_name_v}_output_slides_pair_guidelines.pptx"


def pair_guided_plan_path(
    *,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = pair_guided_contents_name(paper_id, paper_path, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_slide_plan_pair_guidelines.json"


def pair_guided_raw_content_path(
    *,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = pair_guided_contents_name(paper_id, paper_path, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_raw_content.json"


def pair_guided_figures_path(
    *,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = pair_guided_contents_name(paper_id, paper_path, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_figures.json"


def pair_guided_formula_path(
    *,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> Path:
    contents_name = pair_guided_contents_name(paper_id, paper_path, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_formula_match.json"


def is_nonempty_json_file(path: Path) -> bool:
    if not path.exists() or path.stat().st_size == 0:
        return False
    try:
        json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return False
    return True


def pair_guided_outputs_ready(
    *,
    paper_id: str,
    paper_path: Path,
    outline_mode: str,
    model_name_t: str,
    model_name_v: str,
) -> bool:
    pptx_path = pair_guided_pptx_path(
        paper_id=paper_id,
        paper_path=paper_path,
        outline_mode=outline_mode,
        model_name_t=model_name_t,
        model_name_v=model_name_v,
    )
    if not pptx_path.exists() or pptx_path.stat().st_size == 0:
        return False

    required_jsons = [
        pair_guided_raw_content_path(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        ),
        pair_guided_plan_path(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        ),
        pair_guided_figures_path(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        ),
        pair_guided_formula_path(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        ),
    ]
    return all(is_nonempty_json_file(path) for path in required_jsons)


def sample_papers(
    *,
    paired_papers: list[dict[str, Any]],
    count: int,
    seed: int,
    split_filters: list[str] | None,
    authors_by_paper: dict[str, list[tuple[int, str]]],
    author_pair_counts: dict[str, int],
    min_author_paper_count: int,
    distinct_authors: bool,
    author_role_scope: str,
    include_existing: bool,
    model_name_t: str,
    model_name_v: str,
    outline_mode: str,
) -> list[dict[str, Any]]:
    allowed_splits = {value.strip() for value in (split_filters or []) if value and value.strip()}
    eligible: list[dict[str, Any]] = []
    for item in paired_papers:
        paper_id = item["paper_id"]
        if allowed_splits and item["split"] not in allowed_splits:
            continue

        author_rows = list(authors_by_paper.get(paper_id, []))
        if author_role_scope == "primary":
            author_rows = author_rows[:1]
        eligible_author_ids = [
            author_id
            for _author_order, author_id in author_rows
            if author_pair_counts.get(author_id, 0) >= min_author_paper_count
        ]
        if not eligible_author_ids:
            continue

        paper_path = Path(item["paper_path"])
        if not include_existing and pair_guided_outputs_ready(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        ):
            continue

        candidate = dict(item)
        candidate["eligible_author_ids"] = eligible_author_ids
        candidate["max_author_pair_count"] = max(author_pair_counts.get(author_id, 0) for author_id in eligible_author_ids)
        eligible.append(candidate)

    if not eligible:
        raise SystemExit("No eligible papers found for the requested pair-guided batch.")

    rng = random.Random(seed)
    if distinct_authors:
        eligible_author_ids = sorted({author_id for item in eligible for author_id in item["eligible_author_ids"]})
        if len(eligible_author_ids) < count:
            raise SystemExit(
                f"Requested {count} distinct authors, but only {len(eligible_author_ids)} are eligible "
                f"(min_author_paper_count={min_author_paper_count}, author_role_scope={author_role_scope}, "
                f"include_existing={include_existing})."
            )

        best_selection: list[dict[str, Any]] = []
        for _attempt in range(256):
            paper_order = list(eligible)
            rng.shuffle(paper_order)
            paper_order.sort(key=lambda item: (len(item["eligible_author_ids"]), -int(item["max_author_pair_count"])))

            used_authors: set[str] = set()
            selected: list[dict[str, Any]] = []
            for item in paper_order:
                available_authors = [author_id for author_id in item["eligible_author_ids"] if author_id not in used_authors]
                if not available_authors:
                    continue
                chosen_author = rng.choice(available_authors)
                chosen_item = dict(item)
                chosen_item["author_id"] = chosen_author
                chosen_item["author_pair_count"] = author_pair_counts.get(chosen_author, 0)
                selected.append(chosen_item)
                used_authors.add(chosen_author)
                if len(selected) >= count:
                    return selected
            if len(selected) > len(best_selection):
                best_selection = selected

        raise SystemExit(
            f"Requested {count} distinct authors, but only found a feasible assignment for {len(best_selection)} papers "
            f"(min_author_paper_count={min_author_paper_count}, author_role_scope={author_role_scope}, "
            f"include_existing={include_existing})."
        )

    if len(eligible) < count:
        raise SystemExit(
            f"Requested {count} papers, but only {len(eligible)} are eligible "
            f"(min_author_paper_count={min_author_paper_count}, author_role_scope={author_role_scope}, "
            f"include_existing={include_existing})."
        )
    selected = rng.sample(eligible, count)
    finalized: list[dict[str, Any]] = []
    for item in selected:
        chosen_author = rng.choice(item["eligible_author_ids"])
        chosen_item = dict(item)
        chosen_item["author_id"] = chosen_author
        chosen_item["author_pair_count"] = author_pair_counts.get(chosen_author, 0)
        finalized.append(chosen_item)
    return finalized


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    rendered = " ".join(str(part) for part in command)
    print(rendered)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def generation_command(
    *,
    paper_path: Path,
    author_id: str,
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
    outline_mode: str,
    pair_guideline_model: str,
    pair_guideline_max_pairs: int,
    pair_guideline_candidate_pool: int,
    force_refresh_pair_guidelines: bool,
) -> list[str]:
    command = [
        sys.executable,
        "-m",
        "SlidesAgent.new_pipeline_pair_guidelines",
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
        "--author_id",
        author_id,
        "--pair_guideline_model",
        pair_guideline_model,
        "--pair_guideline_max_pairs",
        str(pair_guideline_max_pairs),
        "--pair_guideline_candidate_pool",
        str(pair_guideline_candidate_pool),
    ]
    if force_refresh_pair_guidelines:
        command.append("--force_refresh_pair_guidelines")
    return command


def save_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sample papers and run SlideGen's pair-guided personalization pipeline."
    )
    parser.add_argument("--count", type=int, default=12, help="Number of papers to generate.")
    parser.add_argument("--seed", type=int, default=42, help="Random seed for reproducible sampling.")
    parser.add_argument("--experiment-name", default=None, help="Optional stable name for the run folder.")
    parser.add_argument("--manifest-path", type=Path, default=None, help="Optional manifest JSON path.")
    parser.add_argument("--use-manifest", type=Path, default=None, help="Reuse an existing manifest instead of sampling.")
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--split", action="append", default=None, help="Restrict to one or more dataset splits.")
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--pair-guideline-model", default="gpt-5.4-nano")
    parser.add_argument("--pair-guideline-max-pairs", type=int, default=2)
    parser.add_argument("--pair-guideline-candidate-pool", type=int, default=5)
    parser.add_argument(
        "--min-author-paper-count",
        "--min-primary-author-paper-count",
        dest="min_author_paper_count",
        type=int,
        default=5,
        help="Only sample papers whose chosen personalization author has at least this many paired paper-PPT examples.",
    )
    parser.add_argument(
        "--distinct-authors",
        "--distinct-primary-authors",
        dest="distinct_authors",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Sample at most one paper per chosen personalization author. Default: true.",
    )
    parser.add_argument(
        "--author-role-scope",
        choices=["any", "primary"],
        default="any",
        help="Which paper authors are eligible as the personalization source author.",
    )
    parser.add_argument("--force-refresh-pair-guidelines", action="store_true")
    parser.add_argument(
        "--include-existing",
        action="store_true",
        help="Include papers that already have pair-guided outputs.",
    )
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    args.papers_csv = resolve_cli_path(args.papers_csv)
    args.paper_authors_csv = resolve_cli_path(args.paper_authors_csv)
    args.manifest_path = resolve_cli_path(args.manifest_path)
    args.use_manifest = resolve_cli_path(args.use_manifest)

    if args.count <= 0:
        raise SystemExit("--count must be positive")
    if args.min_author_paper_count < 1:
        raise SystemExit("--min-author-paper-count must be at least 1")

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    default_name = (
        f"pair_guided_{args.count}_seed{args.seed}_{args.outline_mode}_"
        f"{args.model_name_t}_{args.model_name_v}_{timestamp}"
    )
    experiment_name = sanitize_name(args.experiment_name or default_name)
    experiment_dir = DEFAULT_BATCH_DIR / experiment_name
    experiment_dir.mkdir(parents=True, exist_ok=True)

    paired_papers = load_paired_paper_rows(args.papers_csv)
    authors_by_paper = load_paper_authors_by_paper(args.paper_authors_csv)
    author_pair_counts = load_author_pair_counts(
        args.paper_authors_csv,
        {item["paper_id"] for item in paired_papers},
        author_role_scope=args.author_role_scope,
    )

    if args.use_manifest:
        manifest = json.loads(args.use_manifest.read_text(encoding="utf-8"))
        selected = list(manifest.get("selected_papers") or [])
        if not selected:
            raise SystemExit(f"Manifest has no selected_papers: {args.use_manifest}")
        manifest_path = args.use_manifest
    else:
        selected = sample_papers(
            paired_papers=paired_papers,
            count=args.count,
            seed=args.seed,
            split_filters=args.split,
            authors_by_paper=authors_by_paper,
            author_pair_counts=author_pair_counts,
            min_author_paper_count=args.min_author_paper_count,
            distinct_authors=args.distinct_authors,
            author_role_scope=args.author_role_scope,
            include_existing=args.include_existing,
            model_name_t=args.model_name_t,
            model_name_v=args.model_name_v,
            outline_mode=args.outline_mode,
        )
        manifest = {
            "experiment_name": experiment_name,
            "count": len(selected),
            "seed": args.seed,
            "outline_mode": args.outline_mode,
            "model_name_t": args.model_name_t,
            "model_name_v": args.model_name_v,
            "formula_mode": args.formula_mode,
            "pair_guideline_model": args.pair_guideline_model,
            "pair_guideline_max_pairs": args.pair_guideline_max_pairs,
            "pair_guideline_candidate_pool": args.pair_guideline_candidate_pool,
            "min_author_paper_count": args.min_author_paper_count,
            "distinct_authors": args.distinct_authors,
            "author_role_scope": args.author_role_scope,
            "selected_papers": selected,
        }
        manifest_path = args.manifest_path or (experiment_dir / "manifest.json")
        save_json(manifest_path, manifest)
        print(f"Saved manifest to {manifest_path}")

    summary: dict[str, Any] = {
        "experiment_name": experiment_name,
        "experiment_dir": str(experiment_dir),
        "manifest_path": str(manifest_path),
        "selected_count": len(selected),
        "outline_mode": args.outline_mode,
        "model_name_t": args.model_name_t,
        "model_name_v": args.model_name_v,
        "papers": [],
    }

    for index, item in enumerate(selected, start=1):
        paper_id = str(item["paper_id"])
        paper_path = Path(str(item["paper_path"]))
        author_id = str(item["author_id"])
        print(f"[{index}/{len(selected)}] {paper_id} ({author_id})")

        output_pptx = pair_guided_pptx_path(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=args.outline_mode,
            model_name_t=args.model_name_t,
            model_name_v=args.model_name_v,
        )
        output_plan = pair_guided_plan_path(
            paper_id=paper_id,
            paper_path=paper_path,
            outline_mode=args.outline_mode,
            model_name_t=args.model_name_t,
            model_name_v=args.model_name_v,
        )

        paper_summary: dict[str, Any] = {
            "paper_id": paper_id,
            "paper_path": str(paper_path),
            "author_id": author_id,
            "author_pair_count": int(item.get("author_pair_count", 0)),
            "output_pptx": str(output_pptx),
            "output_plan": str(output_plan),
            "status": "pending",
        }

        try:
            if pair_guided_outputs_ready(
                paper_id=paper_id,
                paper_path=paper_path,
                outline_mode=args.outline_mode,
                model_name_t=args.model_name_t,
                model_name_v=args.model_name_v,
            ) and not args.include_existing and not args.dry_run:
                paper_summary["status"] = "skipped_existing"
            else:
                run_command(
                    generation_command(
                        paper_path=paper_path,
                        author_id=author_id,
                        model_name_t=args.model_name_t,
                        model_name_v=args.model_name_v,
                        formula_mode=args.formula_mode,
                        outline_mode=args.outline_mode,
                        pair_guideline_model=args.pair_guideline_model,
                        pair_guideline_max_pairs=args.pair_guideline_max_pairs,
                        pair_guideline_candidate_pool=args.pair_guideline_candidate_pool,
                        force_refresh_pair_guidelines=args.force_refresh_pair_guidelines,
                    ),
                    cwd=PROJECT_ROOT,
                    dry_run=args.dry_run,
                )
                paper_summary["status"] = "generated" if not args.dry_run else "dry_run"
        except subprocess.CalledProcessError as exc:
            paper_summary["status"] = "failed"
            paper_summary["returncode"] = exc.returncode
        except Exception as exc:
            paper_summary["status"] = "failed"
            paper_summary["error"] = str(exc)

        summary["papers"].append(paper_summary)
        save_json(experiment_dir / "summary.partial.json", summary)

    save_json(experiment_dir / "summary.json", summary)
    print(json.dumps({"experiment_dir": str(experiment_dir), "manifest_path": str(manifest_path), "selected_count": len(selected)}, indent=2))


if __name__ == "__main__":
    main()
