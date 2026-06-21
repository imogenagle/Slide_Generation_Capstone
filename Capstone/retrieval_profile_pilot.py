#!/usr/bin/env python3
"""Build a target-conditioned profile using retrieved historical papers.

This is a parallel pilot path. It does not replace the existing
`distill_author_profile` flow.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any

import yaml
from dotenv import load_dotenv
from jinja2 import Environment, StrictUndefined


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.preference_distill import (
    DEFAULT_AUTHORS_CSV,
    DEFAULT_PAPER_AUTHORS_CSV,
    DEFAULT_PAPERS_CSV,
    SLIDE_EXTENSIONS,
    encode_image_data_uri,
    extract_json_object,
    load_csv_rows,
    normalize_path_for_compare,
    numeric_slide_sort_key,
    resolve_repo_path,
)
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "preference_distiller_target_conditioned.yaml"
DEFAULT_OUTPUT_DIR_RETRIEVAL = REPO_ROOT / "Capstone" / "profiles_retrieval"
CORE_NUMERIC_TARGET_KEYS = (
    "target_avg_words_per_slide",
    "target_image_slide_count",
    "target_table_slide_count",
    "target_formula_slide_count",
)
VISUAL_NUMERIC_TARGET_KEYS = (
    "target_image_slide_count",
    "target_table_slide_count",
    "target_formula_slide_count",
)
TESSERACT_CMD = "/opt/homebrew/bin/tesseract"


def is_content_policy_error(exc: Exception) -> bool:
    message = str(exc).lower()
    return "content_policy_violation" in message or "content safety system" in message


def sanitize_path_component(value: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", value).strip("._") or "item"


def normalize_text(value: str) -> str:
    value = (value or "").strip().lower()
    value = re.sub(r"[^a-z0-9\s]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def tokenize(value: str) -> list[str]:
    return [token for token in normalize_text(value).split() if token]


def token_jaccard(a: str, b: str) -> float:
    tokens_a = set(tokenize(a))
    tokens_b = set(tokenize(b))
    if not tokens_a or not tokens_b:
        return 0.0
    return len(tokens_a & tokens_b) / len(tokens_a | tokens_b)


def token_cosine(a: str, b: str) -> float:
    counts_a = Counter(tokenize(a))
    counts_b = Counter(tokenize(b))
    if not counts_a or not counts_b:
        return 0.0
    common = set(counts_a) & set(counts_b)
    dot = sum(counts_a[token] * counts_b[token] for token in common)
    norm_a = sum(value * value for value in counts_a.values()) ** 0.5
    norm_b = sum(value * value for value in counts_b.values()) ** 0.5
    if norm_a <= 0.0 or norm_b <= 0.0:
        return 0.0
    return dot / (norm_a * norm_b)


def char_ngrams(value: str, n: int = 3) -> set[str]:
    normalized = f" {normalize_text(value)} "
    if len(normalized) < n:
        return {normalized} if normalized.strip() else set()
    return {normalized[i : i + n] for i in range(len(normalized) - n + 1)}


def char_ngram_jaccard(a: str, b: str, n: int = 3) -> float:
    grams_a = char_ngrams(a, n=n)
    grams_b = char_ngrams(b, n=n)
    if not grams_a or not grams_b:
        return 0.0
    return len(grams_a & grams_b) / len(grams_a | grams_b)


def title_similarity_breakdown(target_title: str, candidate_title: str) -> dict[str, float]:
    exact_match = 1.0 if normalize_text(target_title) == normalize_text(candidate_title) and normalize_text(target_title) else 0.0
    jaccard = token_jaccard(target_title, candidate_title)
    cosine = token_cosine(target_title, candidate_title)
    char_score = char_ngram_jaccard(target_title, candidate_title)
    combined = exact_match if exact_match > 0.0 else (0.35 * jaccard + 0.4 * cosine + 0.25 * char_score)
    return {
        "exact_match": round(exact_match, 4),
        "token_jaccard": round(jaccard, 4),
        "token_cosine": round(cosine, 4),
        "char_ngram_jaccard": round(char_score, 4),
        "combined": round(min(1.0, max(0.0, combined)), 4),
    }


def infer_year_from_split(split_value: str) -> int:
    split = (split_value or "").strip().lower()
    match = re.search(r"(\d{2,4})", split)
    if not match:
        return -1
    raw = match.group(1)
    year = int(raw)
    if len(raw) == 2:
        year += 2000 if year <= 30 else 1900
    return year


def recency_sort_key(candidate: dict[str, Any]) -> tuple[int, int, str]:
    year = infer_year_from_split(str(candidate.get("split", "")))
    try:
        record_id = int(str(candidate.get("record_id", "0") or "0"))
    except ValueError:
        record_id = 0
    return (year, record_id, str(candidate.get("paper_id", "")))


def resolve_target_row(
    *,
    target_paper_id: str | None,
    target_paper_path: Path | None,
    paper_rows: list[dict[str, str]],
) -> dict[str, Any]:
    if target_paper_id:
        row = next((item for item in paper_rows if (item.get("paper_id") or "").strip() == target_paper_id.strip()), None)
        if row is None:
            raise ValueError(f"Unknown target paper_id: {target_paper_id}")
        resolved_pdf = resolve_repo_path(row["paper_pdf_path"])
        return {
            "paper_id": row["paper_id"],
            "paper_title": (row.get("paper_title") or row.get("ppt_title") or "").strip(),
            "paper_pdf_path": str(resolved_pdf),
            "record_id": row.get("record_id", ""),
            "split": row.get("split", ""),
        }

    if target_paper_path is None:
        raise ValueError("One of --target-paper-id or --target-paper-path is required.")

    normalized_target = normalize_path_for_compare(target_paper_path)
    row = next(
        (
            item
            for item in paper_rows
            if normalize_path_for_compare(item.get("paper_pdf_path", "")) == normalized_target
        ),
        None,
    )
    if row is not None:
        resolved_pdf = resolve_repo_path(row["paper_pdf_path"])
        return {
            "paper_id": row["paper_id"],
            "paper_title": (row.get("paper_title") or row.get("ppt_title") or "").strip(),
            "paper_pdf_path": str(resolved_pdf),
            "record_id": row.get("record_id", ""),
            "split": row.get("split", ""),
        }

    title = target_paper_path.stem.replace("_", " ").replace("-", " ").strip()
    return {
        "paper_id": sanitize_path_component(target_paper_path.stem),
        "paper_title": title,
        "paper_pdf_path": str(target_paper_path.resolve()),
        "record_id": "",
        "split": "",
    }


def select_candidate_papers(
    *,
    author_id: str,
    target_row: dict[str, Any],
    paper_author_rows: list[dict[str, str]],
    paper_rows: list[dict[str, str]],
) -> list[dict[str, Any]]:
    paper_ids_for_author = {
        row["paper_id"]
        for row in paper_author_rows
        if (row.get("author_id") or "").strip() == author_id
    }
    target_paper_id = str(target_row.get("paper_id") or "").strip()
    target_pdf = normalize_path_for_compare(target_row.get("paper_pdf_path"))

    candidates: list[dict[str, Any]] = []
    for row in paper_rows:
        paper_id = (row.get("paper_id") or "").strip()
        if paper_id not in paper_ids_for_author:
            continue
        if paper_id == target_paper_id:
            continue
        paper_pdf_path = row.get("paper_pdf_path", "")
        if normalize_path_for_compare(paper_pdf_path) == target_pdf:
            continue
        raw_dir = resolve_repo_path(row["raw_dir"])
        if not raw_dir.exists():
            continue
        title = (row.get("paper_title") or row.get("ppt_title") or "").strip()
        candidates.append(
            {
                "paper_id": paper_id,
                "paper_title": title,
                "record_id": row.get("record_id", ""),
                "split": row.get("split", ""),
                "raw_dir": raw_dir,
                "paper_pdf_path": str(resolve_repo_path(paper_pdf_path)),
                "slide_image_count": int(str(row.get("slide_image_count", "0") or "0")),
            }
        )
    return candidates


def rank_candidates(target_title: str, candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ranked: list[dict[str, Any]] = []
    for candidate in candidates:
        breakdown = title_similarity_breakdown(target_title, candidate.get("paper_title", ""))
        ranked.append(
            {
                **candidate,
                "similarity": breakdown,
            }
        )
    ranked.sort(
        key=lambda item: (
            -float((item.get("similarity") or {}).get("combined", 0.0)),
            -int(item.get("slide_image_count", 0)),
            str(item.get("paper_id", "")),
        )
    )
    return ranked


def select_retrieved_papers(
    ranked_candidates: list[dict[str, Any]],
    *,
    max_retrieved: int,
    obvious_similarity_threshold: float = 0.32,
) -> tuple[list[dict[str, Any]], str]:
    if not ranked_candidates:
        return [], "no_candidates"

    best_score = float((ranked_candidates[0].get("similarity") or {}).get("combined", 0.0))
    if best_score >= obvious_similarity_threshold:
        return ranked_candidates[:max_retrieved], "title_similarity"

    most_recent = sorted(
        ranked_candidates,
        key=recency_sort_key,
        reverse=True,
    )
    return most_recent[:max_retrieved], "most_recent_fallback"


def _coerce_float(raw_value: Any, *, low: float, high: float) -> float | None:
    if raw_value is None:
        return None
    try:
        value = float(raw_value)
    except Exception:
        return None
    if value != value or value in (float("inf"), float("-inf")):
        return None
    return round(max(low, min(high, value)), 4)


def extract_ocr_text(image_path: Path, *, tesseract_cmd: str = TESSERACT_CMD) -> str:
    try:
        completed = subprocess.run(
            [tesseract_cmd, str(image_path), "stdout", "--psm", "6"],
            check=True,
            capture_output=True,
            text=True,
        )
    except Exception:
        return ""
    return completed.stdout or ""


def compute_ocr_word_count(text: str) -> int:
    return len(re.findall(r"\b[\w'-]+\b", text or ""))


def build_ocr_text_density_summary(selected_papers: list[dict[str, Any]]) -> dict[str, Any]:
    per_slide_word_counts: list[int] = []
    slide_records: list[dict[str, Any]] = []

    for paper in selected_papers:
        slide_paths = sorted(
            [path for path in paper["raw_dir"].iterdir() if path.suffix.lower() in SLIDE_EXTENSIONS],
            key=numeric_slide_sort_key,
        )
        for slide_index, slide_path in enumerate(slide_paths):
            ocr_text = extract_ocr_text(slide_path)
            word_count = compute_ocr_word_count(ocr_text)
            per_slide_word_counts.append(word_count)
            slide_records.append(
                {
                    "paper_id": paper["paper_id"],
                    "slide_index": slide_index,
                    "slide_path": str(slide_path),
                    "ocr_word_count": word_count,
                }
            )

    avg_words = round(sum(per_slide_word_counts) / len(per_slide_word_counts), 4) if per_slide_word_counts else None
    return {
        "target_avg_words_per_slide": avg_words,
        "slide_count_for_ocr": len(per_slide_word_counts),
        "per_slide_word_counts": slide_records,
    }


def sanitize_numeric_only_profile(
    *,
    raw_profile: dict[str, Any],
    author_id: str,
    selected_papers: list[dict[str, Any]],
    max_retrieved: int,
    target_metadata: dict[str, Any],
    selection_strategy: str,
    retrieval_matches: list[dict[str, Any]],
    deterministic_numeric_preferences: dict[str, Any] | None = None,
) -> dict[str, Any]:
    raw_numeric = dict(raw_profile.get("numeric_preferences") or {})
    deterministic_numeric_preferences = deterministic_numeric_preferences or {}
    specs = {
        "target_avg_words_per_slide": (0.0, 200.0),
        "target_image_slide_count": (0.0, 100.0),
        "target_table_slide_count": (0.0, 100.0),
        "target_formula_slide_count": (0.0, 100.0),
    }

    numeric_preferences: dict[str, Any] = {}
    for key in CORE_NUMERIC_TARGET_KEYS:
        low, high = specs[key]
        value = _coerce_float(raw_numeric.get(key), low=low, high=high)
        if value is not None:
            numeric_preferences[key] = value
    deterministic_words = _coerce_float(
        deterministic_numeric_preferences.get("target_avg_words_per_slide"),
        low=0.0,
        high=200.0,
    )
    if deterministic_words is not None:
        numeric_preferences["target_avg_words_per_slide"] = deterministic_words

    notes = str(((raw_profile.get("evidence_summary") or {}).get("notes")) or "").strip()
    if not notes:
        notes = "Target-conditioned numeric profile distilled from the retrieved historical slide deck."

    return {
        "author_id": author_id,
        "profile_version": 5,
        "profile_method": "retrieval_conditioned_numeric_only_pilot",
        "distilled_from": {
            "paper_count": len(selected_papers),
            "paper_ids": [paper["paper_id"] for paper in selected_papers],
            "deck_sample_policy": {
                "max_papers": max_retrieved,
                "slides_per_deck": "full_retrieved_deck",
                "selection_strategy": "title_similarity_then_recency_fallback",
            },
        },
        "planning_preferences": {},
        "numeric_preferences": numeric_preferences,
        "evidence_summary": {
            "notes": notes,
        },
        "retrieval_context": {
            "target_paper": target_metadata,
            "selection_strategy": selection_strategy,
            "max_retrieved": max_retrieved,
            "retrieval_matches": retrieval_matches,
            "numeric_target_keys": list(CORE_NUMERIC_TARGET_KEYS),
            "visual_numeric_target_keys": list(VISUAL_NUMERIC_TARGET_KEYS),
            "deterministic_numeric_preferences": deterministic_numeric_preferences,
        },
    }


def build_full_deck_image_evidence(selected_papers: list[dict[str, Any]]) -> list[dict[str, Any]]:
    deck_evidence: list[dict[str, Any]] = []
    for paper in selected_papers:
        slide_paths = sorted(
            [path for path in paper["raw_dir"].iterdir() if path.suffix.lower() in SLIDE_EXTENSIONS],
            key=numeric_slide_sort_key,
        )
        deck_evidence.append(
            {
                "paper_id": paper["paper_id"],
                "paper_title": paper["paper_title"],
                "slide_count": len(slide_paths),
                "raw_dir": str(paper["raw_dir"]),
                "all_slide_paths": [str(path) for path in slide_paths],
            }
        )
    return deck_evidence


def build_deck_multimodal_content(
    *,
    user_prompt: str,
    deck_evidence: list[dict[str, Any]],
    excluded_slide_keys: set[str] | None = None,
    max_slides_per_deck: int | None = None,
) -> list[dict[str, Any]]:
    excluded_slide_keys = excluded_slide_keys or set()
    content: list[dict[str, Any]] = [{"type": "text", "text": user_prompt}]
    for deck in deck_evidence:
        all_slide_paths = list(deck.get("all_slide_paths") or [])
        if max_slides_per_deck is not None and len(all_slide_paths) > max_slides_per_deck:
            chosen_indices = {
                int(round(i * (len(all_slide_paths) - 1) / max(1, max_slides_per_deck - 1)))
                for i in range(max_slides_per_deck)
            }
            sampled_slide_paths = [
                slide_path for idx, slide_path in enumerate(all_slide_paths) if idx in chosen_indices
            ]
            slide_count_note = (
                f"Slide count: {deck['slide_count']} (sending {len(sampled_slide_paths)} sampled slides due to safety fallback)"
            )
        else:
            sampled_slide_paths = all_slide_paths
            slide_count_note = f"Slide count: {deck['slide_count']}"
        content.append(
            {
                "type": "text",
                "text": (
                    f"Retrieved historical deck: {deck['paper_id']}\n"
                    f"Title: {deck['paper_title']}\n"
                    f"{slide_count_note}"
                ),
            }
        )
        for slide_index, slide_path_str in enumerate(sampled_slide_paths):
            slide_key = f"{deck['paper_id']}::{slide_index}"
            if slide_key in excluded_slide_keys:
                continue
            slide_path = Path(slide_path_str)
            content.append(
                {
                    "type": "text",
                    "text": f"Historical deck {deck['paper_id']} / slide {slide_index}",
                }
            )
            content.append(
                {
                    "type": "image_url",
                    "image_url": {
                        "url": encode_image_data_uri(slide_path),
                    },
                }
            )
    return content


def find_flagged_slide_keys(
    *,
    client: Any,
    resolved_model_name: str,
    system_prompt: str,
    deck_evidence: list[dict[str, Any]],
) -> set[str]:
    flagged: set[str] = set()
    for deck in deck_evidence:
        for slide_index, slide_path_str in enumerate(deck.get("all_slide_paths") or []):
            slide_key = f"{deck['paper_id']}::{slide_index}"
            slide_path = Path(slide_path_str)
            probe_kwargs = {
                "model": resolved_model_name,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": (
                                    f"Safety probe for retrieved historical deck {deck['paper_id']} "
                                    f"slide {slide_index}. Return any short text response."
                                ),
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": encode_image_data_uri(slide_path),
                                },
                            },
                        ],
                    },
                ],
                "temperature": 0.0,
            }
            if "gpt-5" in resolved_model_name.lower():
                probe_kwargs["max_completion_tokens"] = 50
            else:
                probe_kwargs["max_tokens"] = 50
            try:
                client.chat.completions.create(**probe_kwargs)
            except Exception as exc:
                if is_content_policy_error(exc):
                    flagged.add(slide_key)
                    print(
                        f"[retrieval-profile] Skipping flagged historical slide {slide_key} due to content policy.",
                        flush=True,
                    )
                else:
                    raise
    return flagged


def call_retrieval_distiller_model(
    model_name: str,
    *,
    system_prompt: str,
    user_prompt: str,
    deck_evidence: list[dict[str, Any]],
) -> dict[str, Any]:
    client = build_openai_client()
    resolved_model_name = resolve_direct_model_name(model_name)
    excluded_slide_keys: set[str] = set()
    max_slides_per_deck: int | None = None

    print(
        f"[retrieval-profile] Sending retrieval distiller request to model={resolved_model_name} "
        f"with {len(deck_evidence)} retrieved deck(s)",
        flush=True,
    )
    while True:
        content = build_deck_multimodal_content(
            user_prompt=user_prompt,
            deck_evidence=deck_evidence,
            excluded_slide_keys=excluded_slide_keys,
            max_slides_per_deck=max_slides_per_deck,
        )
        request_kwargs = {
            "model": resolved_model_name,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": content},
            ],
            "temperature": 0.2,
        }
        if "gpt-5" in resolved_model_name.lower():
            request_kwargs["max_completion_tokens"] = 1400
        else:
            request_kwargs["max_tokens"] = 1400
        try:
            response = client.chat.completions.create(**request_kwargs)
            break
        except Exception as exc:
            if not is_content_policy_error(exc):
                raise
            print(
                "[retrieval-profile] Full multimodal request hit content policy. "
                "Probing retrieved slides individually and retrying without flagged slides.",
                flush=True,
            )
            newly_flagged = find_flagged_slide_keys(
                client=client,
                resolved_model_name=resolved_model_name,
                system_prompt=system_prompt,
                deck_evidence=deck_evidence,
            )
            if newly_flagged and not newly_flagged.issubset(excluded_slide_keys):
                excluded_slide_keys.update(newly_flagged)
                print(
                    f"[retrieval-profile] Retrying retrieval distillation with {len(excluded_slide_keys)} slide(s) excluded.",
                    flush=True,
                )
                continue
            fallback_schedule = [8, 6, 4, 2, 1]
            next_cap = None
            for candidate_cap in fallback_schedule:
                if max_slides_per_deck is None or candidate_cap < max_slides_per_deck:
                    next_cap = candidate_cap
                    break
            if next_cap is None:
                raise
            max_slides_per_deck = next_cap
            print(
                f"[retrieval-profile] No individually flagged slide identified. "
                f"Retrying with at most {max_slides_per_deck} sampled slide(s) per deck.",
                flush=True,
            )
    print("[retrieval-profile] Distiller response received", flush=True)
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def render_target_prompt(
    prompt_path: Path,
    *,
    author_metadata: dict[str, Any],
    target_metadata: dict[str, Any],
    retrieval_matches: list[dict[str, Any]],
    deck_evidence: list[dict[str, Any]],
) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    rendered = template.render(
        author_metadata=author_metadata,
        target_metadata=target_metadata,
        retrieval_matches=retrieval_matches,
        deck_evidence=deck_evidence,
        numeric_target_keys=list(VISUAL_NUMERIC_TARGET_KEYS),
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": rendered,
    }


def build_author_metadata(
    *,
    author_id: str,
    authors_rows: list[dict[str, str]],
    selected_papers: list[dict[str, Any]],
    max_papers: int,
    target_metadata: dict[str, Any],
    retrieval_matches: list[dict[str, Any]],
) -> dict[str, Any]:
    author_row = next((row for row in authors_rows if (row.get("author_id") or "").strip() == author_id), None)
    display_name = (author_row or {}).get("display_name") or author_id
    return {
        "author_id": author_id,
        "display_name": display_name,
        "paper_count": len(selected_papers),
        "paper_ids": [paper["paper_id"] for paper in selected_papers],
        "deck_sample_policy": {
            "max_papers": max_papers,
            "slides_per_deck": "representative",
            "selection_strategy": "title_similarity_then_recency_fallback",
        },
        "target_paper": target_metadata,
        "retrieval_matches": retrieval_matches,
    }


def compact_retrieval_matches(ranked_candidates: list[dict[str, Any]], limit: int = 5) -> list[dict[str, Any]]:
    compact: list[dict[str, Any]] = []
    for candidate in ranked_candidates[:limit]:
        compact.append(
            {
                "paper_id": candidate["paper_id"],
                "paper_title": candidate["paper_title"],
                "paper_pdf_path": candidate["paper_pdf_path"],
                "slide_image_count": candidate["slide_image_count"],
                "similarity": candidate["similarity"],
            }
        )
    return compact


def main() -> None:
    parser = argparse.ArgumentParser(description="Pilot retrieval-conditioned profile builder.")
    parser.add_argument("--author-id", required=True)
    parser.add_argument("--target-paper-id", default=None)
    parser.add_argument("--target-paper-path", type=Path, default=None)
    parser.add_argument("--authors-csv", type=Path, default=DEFAULT_AUTHORS_CSV)
    parser.add_argument("--paper-authors-csv", type=Path, default=DEFAULT_PAPER_AUTHORS_CSV)
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR_RETRIEVAL)
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--max-retrieved", type=int, default=1)
    parser.add_argument("--model", default="gpt-5.4-nano")
    parser.add_argument("--force-refresh", action="store_true")
    parser.add_argument("--dry-run-metadata-only", action="store_true")
    args = parser.parse_args()

    if not args.target_paper_id and args.target_paper_path is None:
        raise SystemExit("One of --target-paper-id or --target-paper-path is required.")
    if args.max_retrieved <= 0:
        raise SystemExit("--max-retrieved must be positive.")

    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    load_dotenv(REPO_ROOT / ".env")
    authors_rows = load_csv_rows(args.authors_csv)
    paper_author_rows = load_csv_rows(args.paper_authors_csv)
    paper_rows = load_csv_rows(args.papers_csv)

    target_row = resolve_target_row(
        target_paper_id=args.target_paper_id,
        target_paper_path=args.target_paper_path,
        paper_rows=paper_rows,
    )
    target_key = sanitize_path_component(target_row["paper_id"])
    profile_path = output_dir / f"{args.author_id}.{target_key}.retrieval.json"
    bundle_path = output_dir / f"{args.author_id}.{target_key}.retrieval.input.json"

    if profile_path.exists() and not args.force_refresh and not args.dry_run_metadata_only:
        cached = json.loads(profile_path.read_text(encoding="utf-8"))
        print(f"Reusing cached retrieval profile: {profile_path}")
        print(json.dumps(cached.get("retrieval_context", {}), indent=2, ensure_ascii=False))
        return

    candidates = select_candidate_papers(
        author_id=args.author_id,
        target_row=target_row,
        paper_author_rows=paper_author_rows,
        paper_rows=paper_rows,
    )
    if not candidates:
        raise SystemExit(f"No eligible historical papers found for author_id={args.author_id}")

    ranked = rank_candidates(target_row["paper_title"], candidates)
    selected_papers, selection_strategy = select_retrieved_papers(
        ranked,
        max_retrieved=args.max_retrieved,
    )
    retrieval_matches = compact_retrieval_matches(ranked, limit=min(5, len(ranked)))
    deck_evidence = build_full_deck_image_evidence(selected_papers)
    ocr_text_density_summary = build_ocr_text_density_summary(selected_papers)
    author_metadata = build_author_metadata(
        author_id=args.author_id,
        authors_rows=authors_rows,
        selected_papers=selected_papers,
        max_papers=args.max_retrieved,
        target_metadata=target_row,
        retrieval_matches=retrieval_matches,
    )
    rendered_prompt = render_target_prompt(
        args.prompt_path,
        author_metadata=author_metadata,
        target_metadata=target_row,
        retrieval_matches=retrieval_matches,
        deck_evidence=deck_evidence,
    )

    bundle_payload = {
        "target_paper": target_row,
        "author_metadata": author_metadata,
        "retrieval_matches": retrieval_matches,
        "selected_source_papers": [
            {
                "paper_id": item["paper_id"],
                "paper_title": item["paper_title"],
                "paper_pdf_path": item["paper_pdf_path"],
                "slide_image_count": item["slide_image_count"],
                "similarity": item["similarity"],
                "selection_strategy": selection_strategy,
            }
            for item in selected_papers
        ],
        "deck_evidence": deck_evidence,
        "ocr_text_density_summary": ocr_text_density_summary,
        "prompt_preview": rendered_prompt["user_prompt"],
    }
    bundle_path.write_text(json.dumps(bundle_payload, indent=2, ensure_ascii=False), encoding="utf-8")

    if args.dry_run_metadata_only:
        print(f"Saved retrieval evidence bundle to {bundle_path}")
        print(
            f"Top match ({selection_strategy}): "
            f"{selected_papers[0]['paper_id']} | {selected_papers[0]['paper_title']}"
        )
        return

    raw_profile = call_retrieval_distiller_model(
        args.model,
        system_prompt=rendered_prompt["system_prompt"],
        user_prompt=rendered_prompt["user_prompt"],
        deck_evidence=deck_evidence,
    )
    profile = sanitize_numeric_only_profile(
        raw_profile=raw_profile,
        author_id=args.author_id,
        max_retrieved=args.max_retrieved,
        selected_papers=selected_papers,
        target_metadata=target_row,
        selection_strategy=selection_strategy,
        retrieval_matches=retrieval_matches,
        deterministic_numeric_preferences=ocr_text_density_summary,
    )
    profile_path.write_text(json.dumps(profile, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Saved retrieval-conditioned profile to {profile_path}")
    print(f"Saved retrieval evidence bundle to {bundle_path}")


if __name__ == "__main__":
    main()
