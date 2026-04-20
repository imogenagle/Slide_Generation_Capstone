#!/usr/bin/env python3
"""Build normalized paper/author tables from Capstone/authors.csv.

This script treats each row in authors.csv as a unique paper/deck example from
the processed v1.0 dataset. It uses the `(source_file, id)` pair as the primary
join key into `data_raw/<split>/<id>/`.
"""

from __future__ import annotations

import argparse
import csv
import re
import unicodedata
from collections import Counter
from pathlib import Path


DEFAULT_INPUT_CSV = Path("SlideGen/Capstone/authors.csv")
DEFAULT_RAW_ROOT = Path("SlideGen/data_raw")
DEFAULT_OUTPUT_DIR = Path("SlideGen/Capstone/author_tables")


def slugify(text: str) -> str:
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^a-z0-9]+", "_", ascii_text.lower()).strip("_")
    return slug or "unknown"


def clean_author_name(name: str) -> str:
    """Normalize author strings while staying conservative.

    We use `paper_authors` as the source of truth because it is much cleaner
    than the OCR-derived PPT author extraction.
    """
    text = unicodedata.normalize("NFKC", name or "").strip()
    text = text.replace("−", "-")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\[\s*0{4}-?0{4}-?0{4}-?[0-9Xx]{3,4}\s*\]", "", text)
    text = re.sub(r"\(\s*[A-Za-z]\s*\)$", "", text)
    text = re.sub(r"^[^A-Za-z]+", "", text)
    text = re.sub(r"\s*[*†‡]+$", "", text)
    text = re.sub(r"\s*\d+(?:,\d+)*$", "", text)
    text = re.sub(r"(?<=[A-Za-z])\d+[A-Za-z]*$", "", text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip(" ,;")
    words = re.findall(r"[A-Za-z][A-Za-z.'-]*", text)
    if len(words) < 2:
        return ""
    return " ".join(words)


def split_authors(author_text: str) -> list[str]:
    text = (author_text or "").strip()
    if not text:
        return []

    if ";" in text:
        parts = [part.strip() for part in text.split(";")]
    else:
        parts = [part.strip() for part in re.split(r"\s+(?:and|&)\s+", text)]

    cleaned: list[str] = []
    for part in parts:
        name = clean_author_name(part)
        if name:
            cleaned.append(name)
    return cleaned


def source_file_to_split(source_file: str) -> str:
    return Path(source_file).stem.strip()


def find_paper_pdf(raw_dir: Path, record_id: str) -> Path | None:
    preferred = raw_dir / f"{record_id}_paper.pdf"
    if preferred.exists():
        return preferred

    fallback = raw_dir / "0_paper.pdf"
    if fallback.exists():
        return fallback

    pdfs = sorted(raw_dir.glob("*_paper.pdf"))
    return pdfs[0] if pdfs else None


def count_slide_images(raw_dir: Path) -> int:
    return len(list(raw_dir.glob("*.jpg")))


def read_rows(input_csv: Path) -> list[dict[str, str]]:
    with input_csv.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def build_tables(rows: list[dict[str, str]], raw_root: Path) -> tuple[list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    papers: list[dict[str, object]] = []
    paper_authors: list[dict[str, object]] = []

    author_name_counts: Counter[str] = Counter()
    author_display_names: dict[str, list[str]] = {}
    seen_papers: set[str] = set()

    for row in rows:
        source_file = row["source_file"].strip()
        split = source_file_to_split(source_file)
        record_id = row["id"].strip()
        paper_id = f"{split}:{record_id}"
        raw_dir = raw_root / split / record_id
        paper_pdf = find_paper_pdf(raw_dir, record_id) if raw_dir.exists() else None
        paper_author_text = row.get("paper_authors", "").strip()
        ppt_author_text = row.get("ppt_authors", "").strip()
        author_names = split_authors(paper_author_text) or split_authors(ppt_author_text)

        if paper_id not in seen_papers:
            seen_papers.add(paper_id)
            papers.append(
                {
                    "paper_id": paper_id,
                    "split": split,
                    "record_id": record_id,
                    "source_file": source_file,
                    "paper_title": row.get("paper_title", "").strip(),
                    "ppt_title": row.get("ppt_title", "").strip(),
                    "paper_authors_text": paper_author_text,
                    "ppt_authors_text": ppt_author_text,
                    "raw_dir": str(raw_dir),
                    "raw_dir_exists": raw_dir.exists(),
                    "paper_pdf_path": str(paper_pdf) if paper_pdf else "",
                    "paper_pdf_exists": bool(paper_pdf and paper_pdf.exists()),
                    "slide_image_count": count_slide_images(raw_dir) if raw_dir.exists() else 0,
                }
            )

        seen_author_ids_for_paper: set[str] = set()
        for order, author_name in enumerate(author_names, start=1):
            author_id = slugify(author_name)
            if author_id in seen_author_ids_for_paper:
                continue

            seen_author_ids_for_paper.add(author_id)
            author_name_counts[author_id] += 1
            author_display_names.setdefault(author_id, []).append(author_name)
            paper_authors.append(
                {
                    "paper_id": paper_id,
                    "author_id": author_id,
                    "author_name": author_name,
                    "author_order": order,
                    "split": split,
                    "record_id": record_id,
                }
            )

    authors: list[dict[str, object]] = []
    for author_id in sorted(author_display_names):
        name_options = author_display_names[author_id]
        # Prefer the most common surface form, then the longest as a stable tie-breaker.
        display_name = sorted(
            Counter(name_options).items(),
            key=lambda item: (-item[1], -len(item[0]), item[0]),
        )[0][0]
        authors.append(
            {
                "author_id": author_id,
                "display_name": display_name,
                "normalized_name": slugify(display_name).replace("_", " "),
                "paper_count": author_name_counts[author_id],
            }
        )

    papers.sort(key=lambda row: str(row["paper_id"]))
    authors.sort(key=lambda row: str(row["author_id"]))
    paper_authors.sort(key=lambda row: (str(row["paper_id"]), int(row["author_order"]), str(row["author_id"])))
    return papers, authors, paper_authors


def main() -> None:
    parser = argparse.ArgumentParser(description="Build papers/authors/paper_authors tables.")
    parser.add_argument("--input-csv", type=Path, default=DEFAULT_INPUT_CSV)
    parser.add_argument("--raw-root", type=Path, default=DEFAULT_RAW_ROOT)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    args = parser.parse_args()

    if not args.input_csv.exists():
        raise SystemExit(f"Input CSV not found: {args.input_csv}")

    rows = read_rows(args.input_csv)
    papers, authors, paper_authors = build_tables(rows, args.raw_root)

    write_csv(
        args.output_dir / "papers.csv",
        [
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
        papers,
    )
    write_csv(
        args.output_dir / "authors.csv",
        ["author_id", "display_name", "normalized_name", "paper_count"],
        authors,
    )
    write_csv(
        args.output_dir / "paper_authors.csv",
        ["paper_id", "author_id", "author_name", "author_order", "split", "record_id"],
        paper_authors,
    )

    print(f"Wrote {len(papers)} papers to {args.output_dir / 'papers.csv'}")
    print(f"Wrote {len(authors)} authors to {args.output_dir / 'authors.csv'}")
    print(f"Wrote {len(paper_authors)} paper-author links to {args.output_dir / 'paper_authors.csv'}")


if __name__ == "__main__":
    main()
