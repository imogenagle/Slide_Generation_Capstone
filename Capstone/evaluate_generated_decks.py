#!/usr/bin/env python3
"""Evaluate generated decks found under contents/ against reference slide decks."""

from __future__ import annotations

import argparse
import csv
import json
from pathlib import Path
from typing import Any

from evaluate_core_coverage import DEFAULT_OUTPUT_DIR, REPO_ROOT, evaluate_core_coverage

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False


DEFAULT_CONTENTS_DIR = REPO_ROOT / "contents"
DEFAULT_PAPERS_CSV = REPO_ROOT / "Capstone" / "author_tables" / "papers.csv"


def normalize_paper_name_from_pdf_path(pdf_path: str) -> str:
    return Path(pdf_path).stem.replace(" ", "_")


def normalize_paper_name_from_paper_id(paper_id: str) -> str:
    sanitized = paper_id.strip().replace(":", "_")
    sanitized = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in sanitized)
    sanitized = sanitized.strip("_")
    return sanitized


def load_paper_metadata(papers_csv: Path) -> dict[str, dict[str, Any]]:
    metadata_by_paper_name: dict[str, dict[str, Any]] = {}
    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            title = (row.get("paper_title") or row.get("ppt_title") or "").strip()
            pdf_path = (row.get("paper_pdf_path") or "").strip()
            raw_dir = (row.get("raw_dir") or "").strip()
            if not paper_id or not pdf_path or not raw_dir:
                continue
            metadata = {
                "paper_id": paper_id,
                "title": title,
                "original_slide_dir": (REPO_ROOT.parent / raw_dir).resolve()
                if not Path(raw_dir).is_absolute()
                else Path(raw_dir),
            }
            metadata_by_paper_name[normalize_paper_name_from_paper_id(paper_id)] = metadata
            metadata_by_paper_name[normalize_paper_name_from_pdf_path(pdf_path)] = metadata
    return metadata_by_paper_name


def find_generated_pptx_dirs(contents_dir: Path, model_name_t: str, model_name_v: str) -> list[tuple[str, Path]]:
    prefix = f"{model_name_t}_{model_name_v}_output_slides"
    results: list[tuple[str, Path]] = []
    for paper_dir in sorted(path for path in contents_dir.iterdir() if path.is_dir()):
        candidates = sorted(
            path
            for path in paper_dir.glob(f"{prefix}*.pptx")
            if not path.name.endswith("_themed.pptx") and not path.name.startswith("~$")
        )
        if not candidates:
            continue

        preferred = None
        for candidate in candidates:
            if candidate.name == f"{prefix}_baseline.pptx":
                preferred = candidate
                break
        if preferred is None:
            for candidate in candidates:
                if candidate.name == f"{prefix}.pptx":
                    preferred = candidate
                    break
        if preferred is None:
            preferred = candidates[0]

        results.append((paper_dir.name, preferred))
    return results


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate generated PPTX decks found under contents/.")
    parser.add_argument("--contents-dir", type=Path, default=DEFAULT_CONTENTS_DIR)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--model-name-t", default="4o-mini")
    parser.add_argument("--model-name-v", default="4o-mini")
    parser.add_argument("--eval-model", default="4o-mini", help="Model used for the evaluation LLM call.")
    parser.add_argument("--max-original-slides", type=int, default=0)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument(
        "--save-in-contents",
        action="store_true",
        help="Save each evaluation JSON inside the corresponding contents/<paper_name>/ folder instead of --output-dir.",
    )
    parser.add_argument("--limit", type=int, default=0, help="Optional cap on number of decks to evaluate.")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")

    if not args.contents_dir.exists():
        raise SystemExit(f"Contents directory not found: {args.contents_dir}")
    if not args.papers_csv.exists():
        raise SystemExit(f"papers.csv not found: {args.papers_csv}")

    metadata_by_paper_name = load_paper_metadata(args.papers_csv)
    generated = find_generated_pptx_dirs(args.contents_dir, args.model_name_t, args.model_name_v)
    if args.limit > 0:
        generated = generated[: args.limit]

    if not generated:
        raise SystemExit("No generated PPTX files found for the requested model names.")

    if not args.save_in_contents:
        args.output_dir.mkdir(parents=True, exist_ok=True)

    missing_matches: list[str] = []
    evaluated = 0

    for paper_name, generated_pptx in generated:
        metadata = metadata_by_paper_name.get(paper_name)
        if metadata is None:
            missing_matches.append(paper_name)
            print(f"[skip] No papers.csv match for contents/{paper_name}")
            continue

        paper_id = str(metadata["paper_id"])
        title = str(metadata["title"])
        original_slide_dir = Path(metadata["original_slide_dir"])
        if not original_slide_dir.exists():
            print(f"[skip] Missing original slide dir for {paper_id}: {original_slide_dir}")
            continue

        print(f"[eval] {paper_id} -> {generated_pptx}")
        if args.dry_run:
            continue

        result = evaluate_core_coverage(
            paper_id=paper_id,
            title=title,
            original_slide_dir=original_slide_dir,
            generated_pptx=generated_pptx,
            model=args.eval_model,
            max_original_slides=args.max_original_slides,
        )
        if args.save_in_contents:
            output_path = generated_pptx.parent / f"{paper_id.replace(':', '_')}.core_coverage.json"
        else:
            output_path = args.output_dir / f"{paper_id.replace(':', '_')}.core_coverage.json"
        output_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        evaluated += 1

    print(f"Matched generated decks: {len(generated) - len(missing_matches)} / {len(generated)}")
    print(f"Evaluated decks: {evaluated}")
    if missing_matches:
        print("Unmatched contents folders:")
        for paper_name in missing_matches:
            print(f"- {paper_name}")


if __name__ == "__main__":
    main()
