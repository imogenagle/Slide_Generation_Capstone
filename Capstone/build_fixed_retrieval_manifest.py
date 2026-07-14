#!/usr/bin/env python3
"""Build a manifest with fixed included papers plus sampled eligible papers.

Eligibility rule:
- At least one author on the paper (primary or coauthor) must have at least
  `--min-author-paper-count` paper-powerpoint pairs in history.

The output manifest is compatible with `Capstone/run_batch_experiment.py`.
"""

from __future__ import annotations

import argparse
import csv
import json
import random
import sys
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from Capstone.generate_random_decks import (  # type: ignore
    DEFAULT_PAPER_AUTHORS_CSV,
    DEFAULT_PAPERS_CSV,
    DEFAULT_RAW_ROOT,
    load_candidates_from_papers_csv,
    resolve_cli_path,
    save_manifest,
    scan_candidates_from_raw_root,
)


DEFAULT_FIXED_PAPER_IDS = [
    "acl18:74",
    "acl20:317",
    "cvpr20:1183",
    "eccv20:47",
    "icml20:398",
]


def load_candidates(raw_root: Path, papers_csv: Path) -> list[dict[str, str]]:
    candidates = load_candidates_from_papers_csv(papers_csv)
    if candidates:
        return candidates
    return scan_candidates_from_raw_root(raw_root)


def load_author_paper_index(paper_authors_csv: Path) -> tuple[dict[str, set[str]], dict[str, list[str]]]:
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


def qualifying_authors_for_paper(
    paper_id: str,
    *,
    author_ids_by_paper: dict[str, list[str]],
    paper_ids_by_author: dict[str, set[str]],
    min_author_paper_count: int,
) -> list[str]:
    author_ids = author_ids_by_paper.get(paper_id, [])
    qualified = [
        author_id
        for author_id in author_ids
        if len(paper_ids_by_author.get(author_id, set())) >= min_author_paper_count
    ]
    return sorted(set(qualified))


def choose_author_id(
    paper_id: str,
    *,
    author_ids_by_paper: dict[str, list[str]],
    paper_ids_by_author: dict[str, set[str]],
    min_author_paper_count: int,
) -> str:
    qualified = qualifying_authors_for_paper(
        paper_id,
        author_ids_by_paper=author_ids_by_paper,
        paper_ids_by_author=paper_ids_by_author,
        min_author_paper_count=min_author_paper_count,
    )
    if not qualified:
        raise SystemExit(
            f"Paper {paper_id} does not have any author/coauthor with at least "
            f"{min_author_paper_count} paper-powerpoint pairs."
        )
    # Prefer the most experienced qualifying author; break ties deterministically.
    return max(
        qualified,
        key=lambda author_id: (len(paper_ids_by_author.get(author_id, set())), author_id),
    )


def build_manifest_item(candidate: dict[str, str], author_id: str) -> dict[str, str]:
    return {
        "paper_id": str(candidate["paper_id"]),
        "paper_path": str(candidate["paper_path"]),
        "title": str(candidate.get("title") or ""),
        "author_id": author_id,
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build a 50-paper-style manifest with fixed included papers and any-coauthor eligibility."
    )
    parser.add_argument("--count", type=int, default=50, help="Total number of papers to include.")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--raw-root", type=Path, default=DEFAULT_RAW_ROOT)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--output", type=Path, required=True, help="Output manifest JSON path.")
    parser.add_argument(
        "--include-paper-id",
        action="append",
        default=None,
        help="Paper id to force include. May be provided multiple times.",
    )
    parser.add_argument(
        "--min-author-paper-count",
        type=int,
        default=3,
        help="Minimum paper-powerpoint pair count for any author/coauthor to qualify.",
    )
    args = parser.parse_args()

    args.raw_root = resolve_cli_path(args.raw_root)
    args.papers_csv = resolve_cli_path(args.papers_csv)
    args.paper_authors_csv = resolve_cli_path(args.paper_authors_csv)
    args.output = resolve_cli_path(args.output)

    include_ids = list(args.include_paper_id or DEFAULT_FIXED_PAPER_IDS)
    include_ids = list(dict.fromkeys(include_ids))
    if args.count < len(include_ids):
        raise SystemExit(f"--count={args.count} is smaller than the {len(include_ids)} fixed included papers.")

    candidates = load_candidates(args.raw_root, args.papers_csv)
    by_paper_id = {str(item["paper_id"]): item for item in candidates}
    paper_ids_by_author, author_ids_by_paper = load_author_paper_index(args.paper_authors_csv)

    selected: list[dict[str, str]] = []
    selected_ids: set[str] = set()

    for paper_id in include_ids:
        candidate = by_paper_id.get(paper_id)
        if candidate is None:
            raise SystemExit(f"Included paper_id not found in candidates: {paper_id}")
        author_id = choose_author_id(
            paper_id,
            author_ids_by_paper=author_ids_by_paper,
            paper_ids_by_author=paper_ids_by_author,
            min_author_paper_count=args.min_author_paper_count,
        )
        selected.append(build_manifest_item(candidate, author_id))
        selected_ids.add(paper_id)

    eligible_pool: list[dict[str, str]] = []
    for candidate in candidates:
        paper_id = str(candidate["paper_id"])
        if paper_id in selected_ids:
            continue
        qualified = qualifying_authors_for_paper(
            paper_id,
            author_ids_by_paper=author_ids_by_paper,
            paper_ids_by_author=paper_ids_by_author,
            min_author_paper_count=args.min_author_paper_count,
        )
        if not qualified:
            continue
        author_id = max(
            qualified,
            key=lambda aid: (len(paper_ids_by_author.get(aid, set())), aid),
        )
        eligible_pool.append(build_manifest_item(candidate, author_id))

    needed = args.count - len(selected)
    if len(eligible_pool) < needed:
        raise SystemExit(
            f"Only {len(eligible_pool)} additional eligible papers found, but need {needed} "
            f"to reach count={args.count}."
        )

    rng = random.Random(args.seed)
    sampled = rng.sample(eligible_pool, needed)
    selected.extend(sampled)

    manifest: dict[str, Any] = {
        "experiment_name": args.output.stem,
        "count": len(selected),
        "seed": args.seed,
        "min_author_paper_count": args.min_author_paper_count,
        "author_count_mode": "any",
        "included_paper_ids": include_ids,
        "selected_papers": selected,
    }

    args.output.parent.mkdir(parents=True, exist_ok=True)
    save_manifest(args.output, manifest)
    print(json.dumps({"manifest_path": str(args.output), "count": len(selected)}, indent=2))


if __name__ == "__main__":
    main()
