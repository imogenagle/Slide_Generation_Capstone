#!/usr/bin/env python3
"""Summarize per-paper evaluation summaries into a compact meta-summary."""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from pathlib import Path
from typing import Any

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False


REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


DEFAULT_EVAL_DIR = REPO_ROOT / "Capstone" / "evaluations"
DEFAULT_OUTPUT_PATH = DEFAULT_EVAL_DIR / "core_coverage_meta_summary.json"


def load_eval(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def extract_json_object(raw_text: str) -> dict[str, Any]:
    text = (raw_text or "").strip()
    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, flags=re.DOTALL)
    if fenced:
        text = fenced.group(1).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        first = text.find("{")
        last = text.rfind("}")
        if first != -1 and last != -1 and last > first:
            return json.loads(text[first:last + 1])
        raise


def chunked(items: list[dict[str, Any]], size: int) -> list[list[dict[str, Any]]]:
    return [items[index:index + size] for index in range(0, len(items), size)]


def summarize_chunk(*, client: Any, model: str, chunk_index: int, total_chunks: int, items: list[dict[str, Any]]) -> dict[str, Any]:
    lines = []
    for item in items:
        lines.append(
            f"- {item['paper_id']} | topic_iou={item['topic_iou']} | title={item['title']}\n"
            f"  overall_summary: {item['overall_summary']}\n"
            f"  missing_topics: {', '.join(item['reference_only_topics']) or 'none'}\n"
            f"  extra_topics: {', '.join(item['generated_only_topics']) or 'none'}"
        )

    system_prompt = (
        "You are summarizing evaluation summaries for generated scientific slide decks.\n"
        "Identify recurring patterns in what generated decks preserve, omit, overemphasize, or restructure.\n"
        "Return only valid JSON."
    )
    user_prompt = f"""
Chunk {chunk_index} of {total_chunks}

Below are per-paper evaluation summaries.

{chr(10).join(lines)}

Return exactly this JSON schema:
{{
  "chunk_index": {chunk_index},
  "num_papers": {len(items)},
  "common_strengths": [],
  "common_weaknesses": [],
  "recurring_missing_topics": [],
  "recurring_extra_topics": [],
  "chunk_summary": ""
}}
""".strip()

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.1,
        max_tokens=1200,
    )
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def summarize_across_chunks(*, client: Any, model: str, chunk_summaries: list[dict[str, Any]], num_papers: int) -> dict[str, Any]:
    chunk_text = json.dumps(chunk_summaries, indent=2, ensure_ascii=False)
    system_prompt = (
        "You are synthesizing chunk-level summaries of slide-deck evaluations.\n"
        "Focus on the highest-level recurring strengths, weaknesses, and topic patterns.\n"
        "Return only valid JSON."
    )
    user_prompt = f"""
There are {num_papers} evaluated papers total.

Chunk summaries:
{chunk_text}

Return exactly this JSON schema:
{{
  "num_papers": {num_papers},
  "overall_takeaways": [],
  "most_common_strengths": [],
  "most_common_weaknesses": [],
  "most_common_missing_topics": [],
  "most_common_extra_topics": [],
  "meta_summary": ""
}}
""".strip()

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.1,
        max_tokens=1200,
    )
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def main() -> None:
    parser = argparse.ArgumentParser(description="Summarize evaluation overall_summary fields in chunks.")
    parser.add_argument("--eval-dir", type=Path, default=DEFAULT_EVAL_DIR)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_PATH)
    parser.add_argument("--model", default="4o-mini", help="Model used for the meta-summary.")
    parser.add_argument("--chunk-size", type=int, default=20, help="How many evaluations to summarize per chunk.")
    parser.add_argument("--limit", type=int, default=0, help="Optional cap on number of evaluation files.")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")
    eval_paths = sorted(args.eval_dir.glob("*.core_coverage.json"))
    if args.limit > 0:
        eval_paths = eval_paths[: args.limit]
    if not eval_paths:
        raise SystemExit(f"No evaluation files found in {args.eval_dir}")
    if args.chunk_size <= 0:
        raise SystemExit("--chunk-size must be positive")

    items = []
    for path in eval_paths:
        payload = load_eval(path)
        items.append(
            {
                "paper_id": str(payload.get("paper_id", "")).strip(),
                "title": str(payload.get("title", "")).strip(),
                "topic_iou": round(safe_float(payload.get("topic_iou", payload.get("coverage_ratio"))), 3),
                "reference_only_topics": [str(x) for x in payload.get("reference_only_topics", [])],
                "generated_only_topics": [str(x) for x in payload.get("generated_only_topics", [])],
                "overall_summary": str(((payload.get("difference_summary") or {}).get("overall_summary")) or "").strip(),
            }
        )

    client = build_openai_client()
    resolved_model_name = resolve_direct_model_name(args.model)
    chunks = chunked(items, args.chunk_size)
    chunk_summaries = []

    for index, chunk in enumerate(chunks, start=1):
        print(f"Summarizing chunk {index}/{len(chunks)} ({len(chunk)} papers)")
        chunk_summaries.append(
            summarize_chunk(
                client=client,
                model=resolved_model_name,
                chunk_index=index,
                total_chunks=len(chunks),
                items=chunk,
            )
        )

    print("Synthesizing chunk summaries")
    final_summary = summarize_across_chunks(
        client=client,
        model=resolved_model_name,
        chunk_summaries=chunk_summaries,
        num_papers=len(items),
    )

    output_payload = {
        "num_papers": len(items),
        "chunk_size": args.chunk_size,
        "num_chunks": len(chunks),
        "chunk_summaries": chunk_summaries,
        "final_summary": final_summary,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output_payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Saved meta-summary to {args.output}")


if __name__ == "__main__":
    main()
