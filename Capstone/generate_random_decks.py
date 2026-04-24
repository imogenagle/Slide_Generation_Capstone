#!/usr/bin/env python3
"""Generate a random batch of slide decks from data_raw."""

from __future__ import annotations

import argparse
import csv
import json
import random
import subprocess
import sys
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RAW_ROOT = PROJECT_ROOT / "data_raw"
DEFAULT_PAPERS_CSV = PROJECT_ROOT / "Capstone" / "author_tables" / "papers.csv"
DEFAULT_PAPER_AUTHORS_CSV = PROJECT_ROOT / "Capstone" / "author_tables" / "paper_authors.csv"
DEFAULT_MANIFEST_DIR = PROJECT_ROOT / "Capstone" / "batch_runs"


def load_candidates_from_papers_csv(papers_csv: Path) -> list[dict[str, str]]:
    if not papers_csv.exists():
        return []

    candidates: list[dict[str, str]] = []
    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            paper_path = (row.get("paper_pdf_path") or "").strip()
            title = (row.get("paper_title") or row.get("ppt_title") or "").strip()
            if not paper_id or not paper_path:
                continue
            pdf_path = Path(paper_path)
            if not pdf_path.is_absolute():
                pdf_path = PROJECT_ROOT.parent / pdf_path
            if not pdf_path.exists():
                continue
            candidates.append(
                {
                    "paper_id": paper_id,
                    "paper_path": str(pdf_path.resolve()),
                    "title": title,
                }
            )
    return candidates


def scan_candidates_from_raw_root(raw_root: Path) -> list[dict[str, str]]:
    candidates: list[dict[str, str]] = []
    for split_dir in sorted(path for path in raw_root.iterdir() if path.is_dir()):
        for record_dir in sorted(path for path in split_dir.iterdir() if path.is_dir()):
            preferred_pdf = record_dir / f"{record_dir.name}_paper.pdf"
            if preferred_pdf.exists():
                pdf_path = preferred_pdf
            else:
                fallback_pdfs = sorted(record_dir.glob("*_paper.pdf"))
                if not fallback_pdfs:
                    continue
                pdf_path = fallback_pdfs[0]
            candidates.append(
                {
                    "paper_id": f"{split_dir.name}:{record_dir.name}",
                    "paper_path": str(pdf_path.resolve()),
                    "title": "",
                }
            )
    return candidates


def output_dir_key(paper_id: str, paper_path: Path) -> str:
    sanitized = paper_id.strip().replace(":", "_")
    sanitized = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in sanitized)
    sanitized = sanitized.strip("_")
    if sanitized:
        return sanitized
    return paper_path.stem.replace(" ", "_")


def append_outline_mode_suffix(paper_name: str, outline_mode: str) -> str:
    base = paper_name.strip().replace(" ", "_")
    if base.endswith("_high_level") or base.endswith("_technical"):
        return base
    return f"{base}_{outline_mode}"


def output_pptx_path(
    paper_id: str,
    paper_path: Path,
    model_name_t: str,
    model_name_v: str,
    outline_mode: str,
) -> Path:
    paper_name = append_outline_mode_suffix(output_dir_key(paper_id, paper_path), outline_mode)
    return PROJECT_ROOT / "contents" / paper_name / f"{model_name_t}_{model_name_v}_output_slides.pptx"


def load_primary_author_ids(paper_authors_csv: Path) -> dict[str, str]:
    if not paper_authors_csv.exists():
        return {}

    primary_author_ids: dict[str, tuple[int, str]] = {}
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
            current = primary_author_ids.get(paper_id)
            if current is None or author_order < current[0]:
                primary_author_ids[paper_id] = (author_order, author_id)
    return {paper_id: author_id for paper_id, (_, author_id) in primary_author_ids.items()}


def build_command(
    *,
    paper_path: Path,
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
    outline_mode: str,
    use_author_preferences: bool,
    author_id: str | None,
    preference_model: str,
    preference_max_papers: int,
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
    ]
    if use_author_preferences:
        command.extend(
            [
                "--use_author_preferences",
                "--author_id",
                str(author_id),
                "--preference_model",
                preference_model,
                "--preference_max_papers",
                str(preference_max_papers),
            ]
        )
    return command


def save_manifest(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def resolve_cli_path(path: Path | None) -> Path | None:
    if path is None or path.is_absolute():
        return path
    return (Path.cwd() / path).resolve()


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a random batch of slide decks from data_raw.")
    parser.add_argument("--count", type=int, default=100, help="Number of decks to generate.")
    parser.add_argument("--seed", type=int, default=42, help="Random seed for reproducible sampling.")
    parser.add_argument("--raw-root", type=Path, default=DEFAULT_RAW_ROOT, help="Root of data_raw.")
    parser.add_argument(
        "--papers-csv",
        type=Path,
        default=DEFAULT_PAPERS_CSV,
        help="Optional papers.csv with resolved paper_pdf_path metadata.",
    )
    parser.add_argument(
        "--paper-authors-csv",
        type=Path,
        default=DEFAULT_PAPER_AUTHORS_CSV,
        help="paper_authors.csv used to resolve author_id when --use-author-preferences is enabled.",
    )
    parser.add_argument("--split", action="append", default=None, help="Restrict sampling to one or more splits.")
    parser.add_argument("--model-name-t", default="4o-mini", help="Text model alias for the generator.")
    parser.add_argument("--model-name-v", default="4o-mini", help="Vision model alias for the generator.")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--use-author-preferences", action="store_true")
    parser.add_argument("--preference-model", default="4o-mini")
    parser.add_argument("--preference-max-papers", type=int, default=5)
    parser.add_argument(
        "--include-existing",
        action="store_true",
        help="Include papers whose output PPTX already exists.",
    )
    parser.add_argument(
        "--manifest-path",
        type=Path,
        default=None,
        help="Optional path for the sampled-run manifest JSON.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Sample papers and write the manifest without generating.")
    args = parser.parse_args()
    args.raw_root = resolve_cli_path(args.raw_root)
    args.papers_csv = resolve_cli_path(args.papers_csv)
    args.paper_authors_csv = resolve_cli_path(args.paper_authors_csv)
    args.manifest_path = resolve_cli_path(args.manifest_path)

    if args.count <= 0:
        raise SystemExit("--count must be positive")
    if not args.raw_root.exists():
        raise SystemExit(f"Raw root not found: {args.raw_root}")

    candidates = load_candidates_from_papers_csv(args.papers_csv)
    if not candidates:
        candidates = scan_candidates_from_raw_root(args.raw_root)
    primary_author_ids = load_primary_author_ids(args.paper_authors_csv)

    if args.split:
        allowed_splits = {value.strip() for value in args.split if value and value.strip()}
        candidates = [item for item in candidates if item["paper_id"].split(":", 1)[0] in allowed_splits]

    eligible: list[dict[str, str]] = []
    skipped_existing = 0
    for item in candidates:
        pptx_path = output_pptx_path(
            item["paper_id"],
            Path(item["paper_path"]),
            args.model_name_t,
            args.model_name_v,
            args.outline_mode,
        )
        if pptx_path.exists() and not args.include_existing:
            skipped_existing += 1
            continue
        if args.use_author_preferences and item["paper_id"] not in primary_author_ids:
            continue
        eligible.append(item)

    if len(eligible) < args.count:
        raise SystemExit(
            f"Requested {args.count} decks, but only {len(eligible)} eligible papers were found "
            f"(skipped_existing={skipped_existing})."
        )

    rng = random.Random(args.seed)
    selected = rng.sample(eligible, args.count)

    manifest_path = args.manifest_path
    if manifest_path is None:
        manifest_name = f"random_{args.count}_seed{args.seed}_{args.model_name_t}_{args.model_name_v}.json"
        manifest_path = DEFAULT_MANIFEST_DIR / manifest_name

    manifest = {
        "count": args.count,
        "seed": args.seed,
        "raw_root": str(args.raw_root),
        "model_name_t": args.model_name_t,
        "model_name_v": args.model_name_v,
        "formula_mode": args.formula_mode,
        "include_existing": args.include_existing,
        "use_author_preferences": args.use_author_preferences,
        "selected_papers": selected,
    }
    save_manifest(manifest_path, manifest)
    print(f"Saved manifest to {manifest_path}")

    if args.dry_run:
        print("Dry run only; no decks generated.")
        return

    total = len(selected)
    failures: list[dict[str, str]] = []
    for index, item in enumerate(selected, start=1):
        paper_path = Path(item["paper_path"])
        paper_id = item["paper_id"]
        print(f"[{index}/{total}] Generating {paper_id} from {paper_path}")
        command = build_command(
            paper_path=paper_path,
            model_name_t=args.model_name_t,
            model_name_v=args.model_name_v,
            formula_mode=args.formula_mode,
            outline_mode=args.outline_mode,
            use_author_preferences=args.use_author_preferences,
            author_id=primary_author_ids.get(paper_id),
            preference_model=args.preference_model,
            preference_max_papers=args.preference_max_papers,
        )
        try:
            subprocess.run(command, cwd=PROJECT_ROOT, check=True)
        except subprocess.CalledProcessError as exc:
            failures.append(
                {
                    "paper_id": paper_id,
                    "paper_path": str(paper_path),
                    "returncode": str(exc.returncode),
                }
            )
            print(f"[failed] {paper_id} (exit {exc.returncode})")

    if failures:
        failure_path = manifest_path.with_suffix(".failures.json")
        save_manifest(failure_path, {"failures": failures})
        raise SystemExit(f"{len(failures)} deck generations failed. See {failure_path}")

    print(f"Generated {total} decks successfully.")


if __name__ == "__main__":
    main()
