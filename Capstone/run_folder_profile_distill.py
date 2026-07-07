#!/usr/bin/env python3
"""Build temporary metadata tables from raw user folders and run preference distillation."""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.preference_distill import DEFAULT_OUTPUT_DIR, distill_author_profile
from Capstone.slidetailor_eval.common import render_pptx_to_images


SLIDE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
TABLE_FILENAMES = {
    "authors": ["author_id", "display_name", "normalized_name", "paper_count"],
    "paper_authors": ["paper_id", "record_id", "split", "author_id", "author_name", "author_order"],
    "papers": [
        "paper_id",
        "split",
        "record_id",
        "source_file",
        "paper_title",
        "ppt_title",
        "paper_authors_text",
        "ppt_authors_text",
        "raw_dir",
        "raw_dir_exists",
        "paper_pdf_path",
        "paper_pdf_exists",
        "slide_image_count",
    ],
}


def sanitize_path_component(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", value).strip("._") or "item"


def normalize_name(value: str) -> str:
    lowered = value.strip().lower()
    collapsed = re.sub(r"[^a-z0-9]+", "_", lowered)
    return collapsed.strip("_")


def repo_relative_string(path: Path) -> str:
    resolved = path.resolve()
    return "SlideGen/" + str(resolved.relative_to(REPO_ROOT.resolve()))


def display_title_from_pdf(pdf_path: Path) -> str:
    text = pdf_path.stem.replace("_", " ").replace("-", " ").strip()
    return re.sub(r"\s+", " ", text) or pdf_path.stem


def find_paper_folders(author_folder: Path) -> list[tuple[Path, Path, list[Path], Path | None]]:
    discovered: list[tuple[Path, Path, list[Path], Path | None]] = []
    for folder in sorted(path for path in author_folder.rglob("*") if path.is_dir()):
        pdfs = sorted(path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".pdf")
        slides = sorted(path for path in folder.iterdir() if path.is_file() and path.suffix.lower() in SLIDE_EXTENSIONS)
        pptxs = sorted(path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".pptx")
        if pdfs and (slides or pptxs):
            discovered.append((folder, pdfs[0], slides, pptxs[0] if pptxs else None))
    root_pdfs = sorted(path for path in author_folder.iterdir() if path.is_file() and path.suffix.lower() == ".pdf")
    root_slides = sorted(path for path in author_folder.iterdir() if path.is_file() and path.suffix.lower() in SLIDE_EXTENSIONS)
    root_pptxs = sorted(path for path in author_folder.iterdir() if path.is_file() and path.suffix.lower() == ".pptx")
    if root_pdfs and (root_slides or root_pptxs):
        root_tuple = (author_folder, root_pdfs[0], root_slides, root_pptxs[0] if root_pptxs else None)
        if root_tuple not in discovered:
            discovered.append(root_tuple)
    return discovered


def parse_specs(raw_specs: list[str]) -> list[tuple[str, Path]]:
    specs: list[tuple[str, Path]] = []
    for raw_spec in raw_specs:
        if "=" not in raw_spec:
            raise SystemExit(f"Invalid --spec '{raw_spec}'. Expected format author_id=path/to/folder")
        author_id, raw_path = raw_spec.split("=", 1)
        author_id = author_id.strip()
        folder = Path(raw_path.strip())
        if not author_id:
            raise SystemExit(f"Missing author_id in --spec '{raw_spec}'")
        if not folder.is_absolute():
            folder = REPO_ROOT / folder
        specs.append((author_id, folder))
    if not specs:
        raise SystemExit("At least one --spec author_id=folder is required.")
    return specs


def build_tables(
    specs: list[tuple[str, Path]],
    out_dir: Path,
    *,
    render_pptx_slides: bool,
    render_dpi: int,
) -> tuple[Path, Path, Path, dict[str, list[str]]]:
    out_dir.mkdir(parents=True, exist_ok=True)
    rendered_root = out_dir / "rendered_slides"

    authors_rows: list[dict[str, str]] = []
    paper_authors_rows: list[dict[str, str]] = []
    papers_rows: list[dict[str, str]] = []
    paper_ids_by_author: dict[str, list[str]] = {}

    for author_id, folder in specs:
        if not folder.exists():
            raise SystemExit(f"Missing folder for {author_id}: {folder}")

        paper_folders = find_paper_folders(folder)
        if not paper_folders:
            raise SystemExit(f"No paper/deck folders found under {folder} for {author_id}")

        authors_rows.append(
            {
                "author_id": author_id,
                "display_name": author_id,
                "normalized_name": normalize_name(author_id),
                "paper_count": str(len(paper_folders)),
            }
        )

        ids_for_author: list[str] = []
        for index, (paper_dir, pdf_path, slides, pptx_path) in enumerate(paper_folders, start=1):
            paper_id = f"{author_id}:{index:03d}"
            record_id = f"{index:03d}"
            title = display_title_from_pdf(pdf_path)
            ids_for_author.append(paper_id)

            effective_raw_dir = paper_dir
            effective_slides = slides
            if not effective_slides and pptx_path is not None:
                if not render_pptx_slides:
                    raise SystemExit(
                        f"{author_id} folder {paper_dir} has a PPTX but no slide images. "
                        "Rerun without --dry-run-metadata-only so the wrapper can render them."
                    )
                render_dir = (
                    rendered_root
                    / sanitize_path_component(author_id)
                    / sanitize_path_component(record_id)
                    / sanitize_path_component(pptx_path.stem)
                )
                render_dir.mkdir(parents=True, exist_ok=True)
                effective_slides = render_pptx_to_images(pptx_path, render_dir, dpi=render_dpi)
                effective_raw_dir = render_dir

            papers_rows.append(
                {
                    "paper_id": paper_id,
                    "split": author_id,
                    "record_id": record_id,
                    "source_file": f"{author_id}.folder",
                    "paper_title": title,
                    "ppt_title": title,
                    "paper_authors_text": author_id,
                    "ppt_authors_text": author_id,
                    "raw_dir": repo_relative_string(effective_raw_dir),
                    "raw_dir_exists": "True",
                    "paper_pdf_path": repo_relative_string(pdf_path),
                    "paper_pdf_exists": "True",
                    "slide_image_count": str(len(effective_slides)),
                }
            )
            paper_authors_rows.append(
                {
                    "paper_id": paper_id,
                    "record_id": record_id,
                    "split": author_id,
                    "author_id": author_id,
                    "author_name": author_id,
                    "author_order": "1",
                }
            )

        paper_ids_by_author[author_id] = ids_for_author

    authors_csv = out_dir / "authors.csv"
    paper_authors_csv = out_dir / "paper_authors.csv"
    papers_csv = out_dir / "papers.csv"

    for path, rows, key in (
        (authors_csv, authors_rows, "authors"),
        (paper_authors_csv, paper_authors_rows, "paper_authors"),
        (papers_csv, papers_rows, "papers"),
    ):
        with path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=TABLE_FILENAMES[key])
            writer.writeheader()
            writer.writerows(rows)

    return authors_csv, paper_authors_csv, papers_csv, paper_ids_by_author


def build_comparison_payload(author_a: dict[str, Any], author_b: dict[str, Any]) -> dict[str, Any]:
    numeric_a = author_a.get("numeric_preferences") or {}
    numeric_b = author_b.get("numeric_preferences") or {}
    plan_a = author_a.get("planning_preferences") or {}
    plan_b = author_b.get("planning_preferences") or {}
    pres_a = author_a.get("presentation_preferences") or {}
    pres_b = author_b.get("presentation_preferences") or {}

    payload: dict[str, Any] = {
        "author_ids": [author_a.get("author_id"), author_b.get("author_id")],
        "planning_preferences": {},
        "numeric_preferences": {},
        "signature_choices": {
            str(author_a.get("author_id")): pres_a.get("signature_choices") or [],
            str(author_b.get("author_id")): pres_b.get("signature_choices") or [],
        },
    }

    interesting_plan_keys = [
        "section_splitting_preference",
        "bullet_density_preference",
        "text_density_preference",
        "visual_density_preference",
        "layout_bias",
    ]
    for key in interesting_plan_keys:
        payload["planning_preferences"][key] = {
            str(author_a.get("author_id")): plan_a.get(key),
            str(author_b.get("author_id")): plan_b.get(key),
            "different": plan_a.get(key) != plan_b.get(key),
        }

    all_numeric_keys = sorted(set(numeric_a) | set(numeric_b))
    for key in all_numeric_keys:
        value_a = numeric_a.get(key)
        value_b = numeric_b.get(key)
        payload["numeric_preferences"][key] = {
            str(author_a.get("author_id")): value_a,
            str(author_b.get("author_id")): value_b,
            "different": value_a != value_b,
        }

    return payload


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Distill author profiles directly from one or more raw user folders."
    )
    parser.add_argument(
        "--spec",
        action="append",
        required=True,
        help="Author/folder spec in the form author_id=relative/or/absolute/path",
    )
    parser.add_argument("--tables-dir", type=Path, default=REPO_ROOT / "Capstone" / "tmp_user_profile_tables")
    parser.add_argument("--output-dir", type=Path, default=REPO_ROOT / "Capstone" / "tmp_user_profiles")
    parser.add_argument("--max-papers", type=int, default=20)
    parser.add_argument("--model", default="gpt-5.4-nano")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--force-refresh", action="store_true")
    parser.add_argument("--dry-run-metadata-only", action="store_true")
    args = parser.parse_args()

    specs = parse_specs(args.spec)
    authors_csv, paper_authors_csv, papers_csv, paper_ids_by_author = build_tables(
        specs,
        args.tables_dir,
        render_pptx_slides=not args.dry_run_metadata_only,
        render_dpi=args.render_dpi,
    )

    print(f"Wrote temp tables to {args.tables_dir}")
    for author_id, _folder in specs:
        print(f"{author_id}: {len(paper_ids_by_author[author_id])} paper/deck folders")

    if args.dry_run_metadata_only:
        return

    args.output_dir.mkdir(parents=True, exist_ok=True)
    profiles: list[dict[str, Any]] = []
    for author_id, _folder in specs:
        print(f"[distill] {author_id}")
        profile = distill_author_profile(
            author_id,
            authors_csv=authors_csv,
            paper_authors_csv=paper_authors_csv,
            papers_csv=papers_csv,
            output_dir=args.output_dir,
            max_papers=args.max_papers,
            model=args.model,
            force_refresh=args.force_refresh,
        )
        profiles.append(profile)
        print(f"Saved profile to {args.output_dir / f'{author_id}.json'}")

    if len(profiles) == 2:
        comparison = build_comparison_payload(profiles[0], profiles[1])
        comparison_path = args.output_dir / "comparison.json"
        comparison_path.write_text(json.dumps(comparison, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"Saved comparison to {comparison_path}")


if __name__ == "__main__":
    main()
