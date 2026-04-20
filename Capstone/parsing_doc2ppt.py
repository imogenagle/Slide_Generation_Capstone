#!/usr/bin/env python3
"""Build a dataframe of paper/PowerPoint titles and authors from the v1.0 dataset."""

from __future__ import annotations

import argparse
import json
import pickle
import re
import unicodedata
from pathlib import Path
import pandas as pd
from typing import Any


DEFAULT_DATA_DIR = Path("/Users/imogennagle/Desktop/UChicago/v1.0")
DEFAULT_OUTPUT_CSV = Path("/Users/imogennagle/Desktop/UChicago/Capstone/authors.csv")
# These keywords help us ignore cover-slide text that is likely affiliation or footer text,
# not actual author names or titles.
AFFILIATION_HINTS = (
    "university",
    "college",
    "institute",
    "school",
    "department",
    "research",
    "laboratory",
    "lab",
    "csail",
    "berkeley",
    "mit",
    "adobe",
    "inc",
    "corp",
    "this video",
)


def normalize_text(text: str) -> str:
    # Normalize to ASCII and lowercase so fuzzy token matching is more reliable.
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", ascii_text.lower()).strip()


def split_authors(author_text: str) -> list[str]:
    # The paper metadata stores authors as a comma-separated string.
    if not author_text:
        return []
    parts = [part.strip() for part in author_text.split(",")]
    return [part for part in parts if part]


def join_authors(author_text: str) -> str:
    # Reformat authors for cleaner CSV/DataFrame output.
    return "; ".join(split_authors(author_text))


def looks_like_affiliation(text: str) -> bool:
    # Heuristic filter for institution labels, company names, and footer text.
    lowered = text.lower()
    return any(hint in lowered for hint in AFFILIATION_HINTS)


def looks_like_author_line(text: str) -> bool:
    # Keep only lines that structurally resemble a list of people.
    stripped = text.strip()
    if not stripped or looks_like_affiliation(stripped):
        return False
    if "," not in stripped and " and " not in stripped.lower():
        return False

    words = stripped.replace(",", " ").split()
    capitalized = sum(1 for word in words if word[:1].isupper())
    return capitalized >= 2


def token_matches(candidate_token: str, author_token: str) -> bool:
    # Allow small OCR variations by accepting exact or prefix-style matches.
    if len(candidate_token) < 3 or len(author_token) < 3:
        return candidate_token == author_token
    return (
        candidate_token == author_token
        or candidate_token.startswith(author_token)
        or author_token.startswith(candidate_token)
    )


def score_author_line(text: str, paper_author_text: str) -> int:
    # Score each slide line by how much it overlaps with the paper's author metadata.
    candidate_tokens = normalize_text(text).split()
    author_tokens = normalize_text(paper_author_text).split()

    score = 0
    for candidate_token in candidate_tokens:
        if any(token_matches(candidate_token, author_token) for author_token in author_tokens):
            score += 1
    return score


def score_title_line(text: str, paper_title_text: str) -> int:
    # Title extraction uses token overlap with the paper title as a simple ranking signal.
    candidate_tokens = normalize_text(text).split()
    title_tokens = normalize_text(paper_title_text).split()

    score = 0
    for candidate_token in candidate_tokens:
        if candidate_token in title_tokens:
            score += 1
    return score


def extract_ppt_title_text(record: dict[str, Any], paper_title_text: str) -> str:
    # Search the PowerPoint cover slide for the line that most likely represents the title.
    slide = record.get("slide") or {}
    cover = slide.get("cover") or {}
    sentences = cover.get("sentences") or []

    candidates: list[tuple[int, float, str]] = []
    for sentence in sentences:
        text = str(sentence.get("text", "")).strip()
        if not text or looks_like_affiliation(text):
            continue
        bbox = sentence.get("bbox") or [0, 0, 0, 0]
        # If scores tie, prefer wider text boxes since titles are usually visually prominent.
        score = score_title_line(text, paper_title_text)
        candidates.append((score, -float(bbox[2]), text))

    if candidates:
        candidates.sort(reverse=True)
        best_title = candidates[0][2]
        if candidates[0][0] > 0:
            return best_title

    return paper_title_text


def extract_ppt_author_text(record: dict[str, Any], paper_author_text: str) -> str:
    # Search the cover slide for author lines and keep them in top-to-bottom reading order.
    slide = record.get("slide") or {}
    cover = slide.get("cover") or {}
    sentences = cover.get("sentences") or []

    candidate_lines: list[tuple[float, str]] = []
    for sentence in sentences:
        text = str(sentence.get("text", "")).strip()
        if not looks_like_author_line(text):
            continue
        score = score_author_line(text, paper_author_text)
        if score > 0:
            bbox = sentence.get("bbox") or [0, 0, 0, 0]
            candidate_lines.append((float(bbox[1]), text))

    if candidate_lines:
        candidate_lines.sort(key=lambda item: item[0])
        return " ".join(text for _, text in candidate_lines)

    return paper_author_text


def extract_record_authors(record: dict[str, Any]) -> dict[str, Any]:
    # Combine paper metadata with best-effort PowerPoint title/author extraction.
    paper = record.get("paper") or {}
    paper_title = str(((paper.get("title") or {}).get("text")) or "").strip()
    paper_author_text = str(((paper.get("author") or {}).get("text")) or "").strip()
    ppt_title = extract_ppt_title_text(record, paper_title)
    ppt_author_text = extract_ppt_author_text(record, paper_author_text)

    return {
        "id": record.get("idd"),
        "paper_title": paper_title,
        "ppt_title": ppt_title,
        "paper_authors_text": paper_author_text,
        "paper_authors": join_authors(paper_author_text),
        "ppt_authors_text": ppt_author_text,
        "ppt_authors": join_authors(ppt_author_text),
    }


def load_pickle(path: Path) -> Any:
    # Each conference file is stored as a pickle containing a list of records.
    with path.open("rb") as handle:
        return pickle.load(handle)


def collect_authors(data_dir: Path) -> dict[str, list[dict[str, Any]]]:
    # Walk through all conference pickle files and extract the fields we care about.
    results: dict[str, list[dict[str, Any]]] = {}

    for path in sorted(data_dir.glob("*.pkl")):
        payload = load_pickle(path)
        if not isinstance(payload, list):
            continue

        entries: list[dict[str, Any]] = []
        for record in payload:
            if isinstance(record, dict):
                entries.append(extract_record_authors(record))
        results[path.name] = entries

    return results


def build_rows(results: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    # Flatten the per-file extraction results into row dictionaries for tabular output.
    rows: list[dict[str, Any]] = []
    for filename, entries in results.items():
        for entry in entries:
            rows.append(
                {
                    "source_file": filename,
                    "id": entry["id"],
                    "paper_title": entry["paper_title"],
                    "ppt_title": entry["ppt_title"],
                    "paper_authors": entry["paper_authors"],
                    "ppt_authors": entry["ppt_authors"],
                }
            )
    return rows


def build_dataframe(rows: list[dict[str, Any]]):
    # Convert extracted rows into a pandas DataFrame for inspection and CSV export.
    if pd is None:
        raise SystemExit(
            "pandas is required to build the dataframe. Install it first, then rerun "
            "`python3 Capstone/doc2ppt.py`."
        )
    return pd.DataFrame(rows)


def main() -> None:
    # Command-line entry point: parse arguments, extract rows, then print/save results.
    parser = argparse.ArgumentParser(
        description="Build a dataframe of paper and PowerPoint titles/authors."
    )
    parser.add_argument(
        "--data-dir",
        type=Path,
        default=DEFAULT_DATA_DIR,
        help=f"Folder containing .pkl files (default: {DEFAULT_DATA_DIR})",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print the extracted rows as JSON instead of a DataFrame preview.",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_OUTPUT_CSV,
        help=f"Path to save the dataframe as a CSV file (default: {DEFAULT_OUTPUT_CSV}).",
    )
    args = parser.parse_args()

    if not args.data_dir.exists():
        raise SystemExit(f"Data directory not found: {args.data_dir}")

    # Extraction happens in two stages: first collect structured data, then flatten it.
    results = collect_authors(args.data_dir)
    rows = build_rows(results)
    if args.json:
        print(json.dumps(rows, indent=2, ensure_ascii=False))
        return

    dataframe = build_dataframe(rows)
    args.csv.parent.mkdir(parents=True, exist_ok=True)
    # Save a reusable CSV copy and also print the DataFrame preview to the terminal.
    dataframe.to_csv(args.csv, index=False)
    print(dataframe)
    print(f"\nSaved dataframe to: {args.csv}")


if __name__ == "__main__":
    main()
