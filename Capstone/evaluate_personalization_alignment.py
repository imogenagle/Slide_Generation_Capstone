#!/usr/bin/env python3
"""Score baseline vs personalized slide plans for alignment with an author profile."""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path
from typing import Any

import yaml
from jinja2 import Environment, StrictUndefined

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "personalization_alignment_evaluator.yaml"

SCORE_KEYS = [
    "section_structure_alignment",
    "bullet_density_alignment",
    "text_density_alignment",
    "figure_usage_alignment",
    "table_usage_alignment",
    "formula_usage_alignment",
    "layout_bias_alignment",
    "overall_style_alignment",
]


def log(message: str, *, verbose: bool) -> None:
    if verbose:
        print(f"[evaluate_personalization_alignment] {message}", file=sys.stderr, flush=True)


def extract_json_object(raw_text: str) -> dict[str, Any]:
    text = (raw_text or "").strip()
    if not text:
        raise ValueError("Model returned an empty response.")

    fenced = re.search(r"```(?:json)?\s*(\{.*\})\s*```", text, flags=re.DOTALL)
    if fenced:
        text = fenced.group(1)

    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError(f"Could not find JSON object in response: {raw_text[:400]}")

    return json.loads(text[start : end + 1])


def classify_level(value: float, low: float, high: float) -> str:
    if value < low:
        return "low"
    if value < high:
        return "medium"
    return "high"


def classify_structure(value: float) -> str:
    if value < 1.5:
        return "coarse"
    if value > 2.5:
        return "fine_grained"
    return "balanced"


def mean(values: list[float]) -> float:
    return sum(values) / len(values) if values else 0.0


def stdev(values: list[float]) -> float:
    if len(values) < 2:
        return 0.0
    avg = mean(values)
    variance = sum((v - avg) ** 2 for v in values) / len(values)
    return variance ** 0.5


def summarize_slide_plan(plan: dict[str, Any]) -> dict[str, Any]:
    slides = list(plan.get("slides") or [])
    slide_count = len(slides)

    section_order: list[str] = []
    section_counts: dict[str, int] = {}
    bullets_per_slide: list[int] = []
    sub_bullets_per_slide: list[int] = []
    words_per_slide: list[int] = []
    figure_flags: list[int] = []
    table_flags: list[int] = []
    formula_flags: list[int] = []
    layout_bias_counts = {
        "text_only": 0,
        "image_right": 0,
        "image_left": 0,
        "image_top": 0,
        "multi_visual": 0,
        "formula_capable": 0,
    }

    for slide in slides:
        section = str(slide.get("section") or "").strip() or "UNKNOWN"
        if section not in section_counts:
            section_order.append(section)
            section_counts[section] = 0
        section_counts[section] += 1
        bullets = list(slide.get("bullets") or [])
        bullet_count = len(bullets)
        sub_bullet_count = sum(len(b.get("sub") or []) for b in bullets if isinstance(b, dict))
        bullets_per_slide.append(bullet_count)
        sub_bullets_per_slide.append(sub_bullet_count)

        words = 0
        for bullet in bullets:
            if not isinstance(bullet, dict):
                continue
            words += len(str(bullet.get("text") or "").split())
            for sub in bullet.get("sub") or []:
                words += len(str(sub).split())
        words_per_slide.append(words)

        image_count = len(slide.get("images") or [])
        table_count = len(slide.get("tables") or [])
        formula_count = len(slide.get("formulas") or [])
        figure_flags.append(1 if image_count > 0 else 0)
        table_flags.append(1 if table_count > 0 else 0)
        formula_flags.append(1 if formula_count > 0 else 0)

        template_id = str(slide.get("template_id") or "")
        if template_id == "T1_TextOnly":
            layout_bias_counts["text_only"] += 1
        if "ImageRight" in template_id:
            layout_bias_counts["image_right"] += 1
        if "ImageLeft" in template_id:
            layout_bias_counts["image_left"] += 1
        if "ImageTop" in template_id:
            layout_bias_counts["image_top"] += 1
        if image_count + table_count + formula_count >= 2 or "2Text" in template_id:
            layout_bias_counts["multi_visual"] += 1
        if formula_count > 0:
            layout_bias_counts["formula_capable"] += 1

    section_slide_counts = [section_counts[name] for name in section_order]
    section_count = len(section_order)
    slides_per_section = (slide_count / section_count) if section_count else 0.0

    summary = {
        "slide_count": slide_count,
        "section_count": section_count,
        "section_titles": section_order,
        "section_slide_counts": section_counts,
        "avg_slides_per_section": round(slides_per_section, 3),
        "section_splitting_estimate": classify_structure(slides_per_section) if section_count else "unknown",
        "avg_bullets_per_slide": round(mean(bullets_per_slide), 3),
        "avg_sub_bullets_per_slide": round(mean(sub_bullets_per_slide), 3),
        "avg_words_per_slide": round(mean(words_per_slide), 3),
        "bullet_density_estimate": classify_level(mean(bullets_per_slide), 2.0, 4.0) if slides else "unknown",
        "text_density_estimate": classify_level(mean(words_per_slide), 18.0, 38.0) if slides else "unknown",
        "figure_usage_estimate": classify_level(sum(figure_flags) / slide_count, 0.2, 0.55) if slide_count else "unknown",
        "table_usage_estimate": classify_level(sum(table_flags) / slide_count, 0.08, 0.22) if slide_count else "unknown",
        "formula_usage_estimate": classify_level(sum(formula_flags) / slide_count, 0.08, 0.22) if slide_count else "unknown",
        "layout_bias_observed": [key for key, count in layout_bias_counts.items() if count > 0],
        "layout_bias_counts": layout_bias_counts,
        "prefers_takeaway_like_close": any(
            "conclusion" in title.lower() or "takeaway" in title.lower() or "future" in title.lower()
            for title in section_order
        ),
        "multi_slide_method_like_sections": [
            title for title, count in section_counts.items()
            if count >= 2 and any(token in title.lower() for token in ("method", "approach", "system", "model", "core idea"))
        ],
        "multi_slide_results_like_sections": [
            title for title, count in section_counts.items()
            if count >= 2 and any(token in title.lower() for token in ("result", "evaluation", "experiment", "analysis", "comparison"))
        ],
        "section_slide_count_std": round(stdev(section_slide_counts), 3),
    }
    return summary


def _count_visual_refs(value: Any, prefix: str) -> int:
    count = 0
    if isinstance(value, dict):
        for key, item in value.items():
            if isinstance(key, str) and key.startswith(prefix):
                count += 1
            count += _count_visual_refs(item, prefix)
    elif isinstance(value, list):
        for item in value:
            count += _count_visual_refs(item, prefix)
    return count


def summarize_extracted_assets(plan_path: Path) -> dict[str, Any]:
    parent = plan_path.parent
    name = plan_path.name
    model_match = re.match(r"(<[^>]+>)_slide_plan(?:_[^.]+)?\.json$", name)
    if not model_match:
        return {}

    prefix = model_match.group(1)
    figures_path = parent / f"{prefix}_figures.json"
    formula_path = parent / f"{prefix}_formula_match.json"

    summary = {
        "figures_path": str(figures_path),
        "formula_path": str(formula_path),
        "image_refs_extracted": 0,
        "table_refs_extracted": 0,
        "formula_sections_extracted": 0,
        "table_penalty_applicable": True,
        "formula_penalty_applicable": True,
    }

    if figures_path.exists():
        figures_json = json.loads(figures_path.read_text(encoding="utf-8"))
        summary["image_refs_extracted"] = _count_visual_refs(figures_json, "image")
        summary["table_refs_extracted"] = _count_visual_refs(figures_json, "table")

    if formula_path.exists():
        formula_json = json.loads(formula_path.read_text(encoding="utf-8"))
        sections = formula_json.get("sections", []) if isinstance(formula_json, dict) else []
        if isinstance(sections, list):
            summary["formula_sections_extracted"] = len(sections)

    summary["table_penalty_applicable"] = summary["table_refs_extracted"] > 0
    summary["formula_penalty_applicable"] = summary["formula_sections_extracted"] > 0
    return summary


def build_numeric_target_summary(author_profile: dict[str, Any]) -> dict[str, Any]:
    numeric = dict(author_profile.get("numeric_preferences") or {})
    if not numeric:
        return {}

    summary = {key: value for key, value in numeric.items() if value not in (None, [], {}, "")}

    avg_text_density_proxy = numeric.get("target_avg_text_density_proxy", numeric.get("avg_text_density_proxy"))
    text_density_proxy_std = numeric.get("text_density_proxy_std")
    if avg_text_density_proxy is not None:
        avg_text_density_proxy = float(avg_text_density_proxy or 0.0)
        summary["target_avg_text_density_proxy"] = avg_text_density_proxy
    if text_density_proxy_std is not None:
        text_density_proxy_std = float(text_density_proxy_std or 0.0)
        summary["text_density_proxy_std"] = text_density_proxy_std
    if avg_text_density_proxy is not None and text_density_proxy_std is not None:
        summary["text_density_proxy_target_range"] = [
            round(max(0.0, avg_text_density_proxy - text_density_proxy_std), 4),
            round(avg_text_density_proxy + text_density_proxy_std, 4),
        ]

    return summary


def build_numeric_comparison(
    target_summary: dict[str, Any],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
) -> dict[str, Any]:
    if not target_summary:
        return {}

    return {}


def render_prompt(
    prompt_path: Path,
    author_profile: dict[str, Any],
    baseline_summary: dict[str, Any],
    personalized_summary: dict[str, Any],
    extracted_asset_summary: dict[str, Any],
) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    user_prompt = template.render(
        author_profile_json=author_profile,
        baseline_plan_summary_json=baseline_summary,
        personalized_plan_summary_json=personalized_summary,
        extracted_asset_summary_json=extracted_asset_summary,
    )
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": user_prompt,
    }


def call_alignment_judge(
    model: str,
    system_prompt: str,
    user_prompt: str,
    *,
    request_timeout: float,
    verbose: bool,
) -> dict[str, Any]:
    client = build_openai_client()
    resolved_model = resolve_direct_model_name(model)

    request_kwargs: dict[str, Any] = {
        "model": resolved_model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.1,
        "timeout": request_timeout,
    }
    if "gpt-5" in resolved_model.lower():
        request_kwargs["max_completion_tokens"] = 2200
    else:
        request_kwargs["max_tokens"] = 2200

    log(
        f"Sending judge request to model={resolved_model!r} with timeout={request_timeout:.1f}s",
        verbose=verbose,
    )
    started_at = time.time()
    response = client.chat.completions.create(**request_kwargs)
    elapsed = time.time() - started_at
    log(f"Judge response received in {elapsed:.2f}s", verbose=verbose)
    raw_text = response.choices[0].message.content or ""
    return extract_json_object(raw_text)


def coerce_scores(report: dict[str, Any]) -> dict[str, Any]:
    for bucket in ("baseline", "personalized"):
        section = report.setdefault(bucket, {})
        scores = section.setdefault("scores", {})
        for key in SCORE_KEYS:
            value = scores.get(key, 0.0)
            try:
                value = float(value)
            except Exception:
                value = 0.0
            scores[key] = max(0.0, min(1.0, value))

    lift = report.setdefault("lift", {})
    for key in SCORE_KEYS:
        base_val = report["baseline"]["scores"][key]
        pers_val = report["personalized"]["scores"][key]
        lift[key] = round(pers_val - base_val, 4)
    return report


def determine_applicable_dimensions(extracted_asset_summary: dict[str, Any]) -> tuple[list[str], list[str]]:
    applicable = {
        "section_structure_alignment",
        "bullet_density_alignment",
        "text_density_alignment",
        "figure_usage_alignment",
        "layout_bias_alignment",
    }
    skipped: list[str] = []

    baseline_assets = dict(extracted_asset_summary.get("baseline") or {})
    personalized_assets = dict(extracted_asset_summary.get("personalized") or {})

    table_applicable = bool(
        baseline_assets.get("table_penalty_applicable") or personalized_assets.get("table_penalty_applicable")
    )
    formula_applicable = bool(
        baseline_assets.get("formula_penalty_applicable") or personalized_assets.get("formula_penalty_applicable")
    )

    if table_applicable:
        applicable.add("table_usage_alignment")
    else:
        skipped.append("table_usage_alignment")

    if formula_applicable:
        applicable.add("formula_usage_alignment")
    else:
        skipped.append("formula_usage_alignment")

    return [key for key in SCORE_KEYS if key != "overall_style_alignment" and key in applicable], skipped


def apply_dimension_applicability(
    report: dict[str, Any],
    extracted_asset_summary: dict[str, Any],
) -> dict[str, Any]:
    applicable_dims, skipped_dims = determine_applicable_dimensions(extracted_asset_summary)

    for metric_key in skipped_dims:
        for bucket in ("baseline", "personalized"):
            report[bucket]["scores"][metric_key] = 0.5
            rationale = report[bucket].setdefault("rationale", {})
            rationale[metric_key] = "Skipped from comparison because the target paper did not expose supporting extracted assets for this dimension."

    for bucket in ("baseline", "personalized"):
        component_values = [float(report[bucket]["scores"][key]) for key in applicable_dims]
        report[bucket]["scores"]["overall_style_alignment"] = round(mean(component_values), 4) if component_values else 0.5

    lift = report.setdefault("lift", {})
    for key in SCORE_KEYS:
        base_val = report["baseline"]["scores"][key]
        pers_val = report["personalized"]["scores"][key]
        lift[key] = round(pers_val - base_val, 4)

    overall_delta = float(lift.get("overall_style_alignment", 0.0))
    winner = "tie"
    if overall_delta > 0.03:
        winner = "personalized"
    elif overall_delta < -0.03:
        winner = "baseline"

    summary = report.setdefault("summary", {})
    summary["winner"] = winner
    if skipped_dims:
        summary["headline"] = (
            summary.get("headline")
            or "Comparison excludes unsupported asset-dependent dimensions for this target paper."
        )
    report["applicable_dimensions"] = applicable_dims
    report["skipped_dimensions"] = skipped_dims
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description="Evaluate personalization alignment lift for baseline vs personalized slide plans.")
    parser.add_argument("--profile", type=Path, required=True, help="Path to distilled author profile JSON.")
    parser.add_argument("--baseline-plan", type=Path, required=True, help="Path to baseline slide_plan JSON.")
    parser.add_argument("--personalized-plan", type=Path, required=True, help="Path to personalized slide_plan JSON.")
    parser.add_argument("--model", default="gpt-5", help="Judge model identifier.")
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--output", type=Path, default=None, help="Optional output JSON path.")
    parser.add_argument(
        "--request-timeout",
        type=float,
        default=180.0,
        help="Timeout in seconds for the judge model request.",
    )
    parser.add_argument("--print-prompt", action="store_true")
    parser.add_argument("--verbose", action="store_true", help="Print progress messages to stderr.")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")

    for path_arg in (args.profile, args.baseline_plan, args.personalized_plan, args.prompt_path):
        if not path_arg.exists():
            raise FileNotFoundError(f"Required input not found: {path_arg}")

    log("Loading profile and slide plans", verbose=args.verbose)
    author_profile = json.loads(args.profile.read_text(encoding="utf-8"))
    baseline_plan = json.loads(args.baseline_plan.read_text(encoding="utf-8"))
    personalized_plan = json.loads(args.personalized_plan.read_text(encoding="utf-8"))

    log("Summarizing baseline and personalized plans", verbose=args.verbose)
    baseline_summary = summarize_slide_plan(baseline_plan)
    personalized_summary = summarize_slide_plan(personalized_plan)
    extracted_asset_summary = {
        "baseline": summarize_extracted_assets(args.baseline_plan),
        "personalized": summarize_extracted_assets(args.personalized_plan),
    }
    log("Rendering judge prompt", verbose=args.verbose)
    prompt = render_prompt(
        args.prompt_path,
        author_profile,
        baseline_summary,
        personalized_summary,
        extracted_asset_summary,
    )

    if args.print_prompt:
        print(prompt["user_prompt"])
        return

    report = call_alignment_judge(
        args.model,
        prompt["system_prompt"],
        prompt["user_prompt"],
        request_timeout=args.request_timeout,
        verbose=args.verbose,
    )
    report = coerce_scores(report)
    report = apply_dimension_applicability(report, extracted_asset_summary)
    numeric_target_summary = build_numeric_target_summary(author_profile)
    numeric_comparison = build_numeric_comparison(
        numeric_target_summary,
        baseline_summary,
        personalized_summary,
    )
    report["inputs"] = {
        "profile": str(args.profile),
        "baseline_plan": str(args.baseline_plan),
        "personalized_plan": str(args.personalized_plan),
        "judge_model": args.model,
    }
    report["derived_plan_summaries"] = {
        "baseline": baseline_summary,
        "personalized": personalized_summary,
    }
    report["extracted_asset_summary"] = extracted_asset_summary
    if numeric_target_summary:
        report["numeric_target_summary"] = numeric_target_summary
        report["numeric_comparison"] = numeric_comparison

    output_text = json.dumps(report, indent=2, ensure_ascii=False)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output_text, encoding="utf-8")
        log(f"Wrote report to {args.output}", verbose=args.verbose)
    print(output_text)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
