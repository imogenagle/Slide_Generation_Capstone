#!/usr/bin/env python3
"""Offline GEPA pilot for the layout planner system prompt.

This script intentionally keeps GEPA out of the normal slide-generation path.
It optimizes only ``utils/prompt_templates/layout_agent_xin.yaml``'s
``system_prompt`` and evaluates candidate prompts by generating slide plans,
rendering PPTX decks, and running the existing deck evaluation bundle.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
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

from camel.agents import ChatAgent
from camel.models import ModelFactory
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name, should_use_direct_openai_client
from SlidesAgent.layout_filler import generate_pptx_from_plan
from utils.src.utils import get_json_from_response
from utils.wei_utils import account_token, chat_via_vllm, get_agent_config, openai_chat_text


DEFAULT_PROMPT_PATH = REPO_ROOT / "utils" / "prompt_templates" / "layout_agent_xin.yaml"
DEFAULT_PAPERS_CSV = REPO_ROOT / "Capstone" / "author_tables" / "papers.csv"
DEFAULT_RUN_ROOT = REPO_ROOT / "Capstone" / "gepa_runs" / "layout_agent_xin"
SCORE_KEYS = {
    "core_coverage": ("topic_iou",),
    "geometry_aware_density": ("gad_geom",),
    "slidetailor_aesthetic_quality": ("deck_score",),
    "slidetailor_content_informativeness": ("deck_score",),
    "slidetailor_structure_similarity": (
        "content_structure_similarity",
        "coverage_iou",
        "flow_ngld",
    ),
    "slidetailor_template_similarity": ("score",),
}


@dataclass(frozen=True)
class PilotPaper:
    paper_id: str
    paper_name: str
    paper_path: Path | None
    title: str


def sanitize_key(value: str) -> str:
    key = value.strip().replace(":", "_")
    key = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in key)
    return key.strip("_") or "paper"


def append_outline_mode_suffix(paper_name: str, outline_mode: str) -> str:
    base = paper_name.strip().replace(" ", "_")
    if base.endswith("_high_level") or base.endswith("_technical"):
        return base
    return f"{base}_{outline_mode}"


def model_prefix(model_name_t: str, model_name_v: str) -> str:
    return f"<{model_name_t}_{model_name_v}>"


def load_pilot_papers(
    *,
    papers_csv: Path,
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
    outline_mode: str,
    limit: int,
    include_missing_artifacts: bool,
    include_paper_ids: set[str] | None = None,
    exclude_paper_ids: set[str] | None = None,
) -> list[PilotPaper]:
    if not papers_csv.exists():
        raise FileNotFoundError(f"papers.csv not found: {papers_csv}")

    papers: list[PilotPaper] = []
    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            if not paper_id:
                continue
            if include_paper_ids is not None and paper_id not in include_paper_ids:
                continue
            if exclude_paper_ids is not None and paper_id in exclude_paper_ids:
                continue
            raw_name = sanitize_key(paper_id)
            paper_name = append_outline_mode_suffix(raw_name, outline_mode)
            paper_path_raw = (row.get("paper_pdf_path") or "").strip()
            paper_path: Path | None = None
            if paper_path_raw:
                candidate = Path(paper_path_raw)
                paper_path = candidate if candidate.is_absolute() else (REPO_ROOT.parent / candidate).resolve()
            title = (row.get("paper_title") or row.get("ppt_title") or "").strip()

            if not include_missing_artifacts:
                missing = required_artifact_paths(
                    paper_name=paper_name,
                    model_name_t=model_name_t,
                    model_name_v=model_name_v,
                    formula_mode=formula_mode,
                )
                if any(not path.exists() for path in missing):
                    continue

            papers.append(
                PilotPaper(
                    paper_id=paper_id,
                    paper_name=paper_name,
                    paper_path=paper_path,
                    title=title,
                )
            )
            if len(papers) >= limit:
                break

    if not papers:
        raise RuntimeError(
            "No pilot papers were found with the required cached layout artifacts. "
            "Run the normal pipeline first, or pass --include-missing-artifacts to see missing-file errors."
        )
    return papers


def parse_paper_id_set(values: list[str] | None) -> set[str] | None:
    if not values:
        return None
    paper_ids: set[str] = set()
    for value in values:
        for item in str(value).split(","):
            item = item.strip()
            if item:
                paper_ids.add(item)
    return paper_ids or None


def required_artifact_paths(
    *,
    paper_name: str,
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
) -> list[Path]:
    prefix = model_prefix(model_name_t, model_name_v)
    image_root = REPO_ROOT / f"{prefix}_images_and_tables"
    paths = [
        REPO_ROOT / "contents" / paper_name / f"{prefix}_raw_content.json",
        REPO_ROOT / "contents" / paper_name / f"{prefix}_figures.json",
        image_root / paper_name / "images_filtered.json",
        image_root / paper_name / "tables_filtered.json",
    ]
    if formula_mode in {1, 2}:
        paths.append(REPO_ROOT / "contents" / paper_name / f"{prefix}_formula_match.json")
    else:
        paths.append(REPO_ROOT / "contents" / paper_name / "formula_index_formula_mode3.json")
    return paths


def ensure_link_or_copy(src: Path, dst: Path) -> None:
    if dst.exists() or dst.is_symlink():
        return
    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.symlink(src, dst, target_is_directory=src.is_dir())
    except OSError:
        if src.is_dir():
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)


def stage_candidate_paper(
    *,
    original_paper_name: str,
    staged_paper_name: str,
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
) -> None:
    prefix = model_prefix(model_name_t, model_name_v)
    src_content = REPO_ROOT / "contents" / original_paper_name
    dst_content = REPO_ROOT / "contents" / staged_paper_name
    dst_content.mkdir(parents=True, exist_ok=True)

    file_names = [
        f"{prefix}_raw_content.json",
        f"{prefix}_figures.json",
    ]
    if formula_mode in {1, 2}:
        file_names.append(f"{prefix}_formula_match.json")
    else:
        file_names.append("formula_index_formula_mode3.json")

    for name in file_names:
        src = src_content / name
        if not src.exists():
            raise FileNotFoundError(f"Missing cached artifact: {src}")
        shutil.copy2(src, dst_content / name)

    src_formula_images = src_content / "formula_images"
    if src_formula_images.exists():
        ensure_link_or_copy(src_formula_images, dst_content / "formula_images")

    src_image_root = REPO_ROOT / f"{prefix}_images_and_tables"
    src_image_dir = src_image_root / original_paper_name
    dst_image_dir = src_image_root / staged_paper_name
    if not src_image_dir.exists():
        raise FileNotFoundError(f"Missing cached image/table artifact dir: {src_image_dir}")
    ensure_link_or_copy(src_image_dir, dst_image_dir)

    for suffix in ("images", "tables"):
        src_json = src_image_root / f"{original_paper_name}_{suffix}.json"
        if src_json.exists():
            shutil.copy2(src_json, src_image_root / f"{staged_paper_name}_{suffix}.json")


def cleanup_candidate_paper(
    *,
    staged_paper_name: str,
    model_name_t: str,
    model_name_v: str,
) -> None:
    shutil.rmtree(REPO_ROOT / "contents" / staged_paper_name, ignore_errors=True)
    image_root = REPO_ROOT / f"{model_prefix(model_name_t, model_name_v)}_images_and_tables"
    staged_image_dir = image_root / staged_paper_name
    if staged_image_dir.is_symlink() or staged_image_dir.is_file():
        staged_image_dir.unlink(missing_ok=True)
    elif staged_image_dir.exists():
        shutil.rmtree(staged_image_dir, ignore_errors=True)
    for suffix in ("images", "tables"):
        (image_root / f"{staged_paper_name}_{suffix}.json").unlink(missing_ok=True)


def load_prompt_config(prompt_path: Path) -> dict[str, Any]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    if not isinstance(prompt_cfg, dict):
        raise ValueError(f"Prompt YAML must be a mapping: {prompt_path}")
    if "system_prompt" not in prompt_cfg or "template" not in prompt_cfg:
        raise ValueError(f"Prompt YAML must contain system_prompt and template: {prompt_path}")
    return prompt_cfg


def render_layout_prompt(
    *,
    prompt_cfg: dict[str, Any],
    args: SimpleNamespace,
) -> str:
    prefix = model_prefix(args.model_name_t, args.model_name_v)
    paper_outline_json = REPO_ROOT / "contents" / args.paper_name / f"{prefix}_raw_content.json"
    figures_path = REPO_ROOT / "contents" / args.paper_name / f"{prefix}_figures.json"
    formulas_path = (
        REPO_ROOT / "contents" / args.paper_name / f"{prefix}_formula_match.json"
        if args.formula_mode in {1, 2}
        else REPO_ROOT / "contents" / args.paper_name / "formula_index_formula_mode3.json"
    )
    image_root = REPO_ROOT / f"{prefix}_images_and_tables"
    images_path = image_root / args.paper_name / "images_filtered.json"
    tables_path = image_root / args.paper_name / "tables_filtered.json"

    jinja_args = {
        "raw_result_json": json.loads(paper_outline_json.read_text(encoding="utf-8")),
        "figures_json": json.loads(figures_path.read_text(encoding="utf-8")),
        "formulas_json": json.loads(formulas_path.read_text(encoding="utf-8")),
        "image_informations_json": json.loads(images_path.read_text(encoding="utf-8")),
        "table_informations_json": json.loads(tables_path.read_text(encoding="utf-8")),
        "use_author_preferences": False,
        "author_preference_profile_json": None,
    }
    env = Environment(undefined=StrictUndefined)
    return env.from_string(prompt_cfg["template"]).render(**jinja_args)


def generate_candidate_slide_plan(
    *,
    prompt_cfg: dict[str, Any],
    candidate_system_prompt: str,
    args: SimpleNamespace,
) -> tuple[dict[str, Any], dict[str, int]]:
    planner_prompt = render_layout_prompt(prompt_cfg=prompt_cfg, args=args)
    cfg = get_agent_config(args.model_name_v)
    use_gpt5_responses = should_use_direct_openai_client(args.model_name_t)

    if use_gpt5_responses:
        client = build_openai_client()
        raw_text, in_tok, out_tok = openai_chat_text(
            client=client,
            model=resolve_direct_model_name(args.model_name_v),
            user_prompt=planner_prompt,
            system_prompt=candidate_system_prompt,
            prefer_responses=True,
        )
    elif args.model_name_t.startswith("vllm_qwen"):
        model = ModelFactory.create(
            model_platform=cfg["model_platform"],
            model_type=cfg["model_type"],
            model_config_dict=cfg["model_config"],
            url=cfg["url"],
        )
        response = chat_via_vllm(planner_prompt, cfg, model, candidate_system_prompt)
        raw_text = response.choices[0].message.content
        in_tok = response.usage.prompt_tokens
        out_tok = response.usage.completion_tokens
    else:
        model = ModelFactory.create(
            model_platform=cfg["model_platform"],
            model_type=cfg["model_type"],
            model_config_dict=cfg["model_config"],
            url=cfg.get("url"),
        )
        agent = ChatAgent(
            system_message=candidate_system_prompt,
            model=model,
            message_window_size=5,
        )
        response = agent.step(planner_prompt)
        raw_text = response.msgs[0].content
        in_tok, out_tok = account_token(response)

    slide_plan = get_json_from_response(raw_text)
    validate_slide_plan(slide_plan)
    return slide_plan, {"input_tokens": int(in_tok or 0), "output_tokens": int(out_tok or 0)}


def validate_slide_plan(slide_plan: dict[str, Any]) -> None:
    if not isinstance(slide_plan, dict):
        raise ValueError("Slide plan must be a JSON object.")
    slides = slide_plan.get("slides")
    if not isinstance(slides, list) or not slides:
        raise ValueError("Slide plan must contain a non-empty slides array.")
    required_keys = {"section", "subsection", "template_id", "bullets", "images", "tables", "formulas"}
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            raise ValueError(f"Slide {index} must be an object.")
        missing = sorted(required_keys - set(slide))
        if missing:
            raise ValueError(f"Slide {index} is missing required keys: {missing}")
        if not isinstance(slide["bullets"], list):
            raise ValueError(f"Slide {index} bullets must be a list.")
        for visual_key in ("images", "tables", "formulas"):
            if not isinstance(slide[visual_key], list):
                raise ValueError(f"Slide {index} {visual_key} must be a list.")


def score_plan_only(slide_plan: dict[str, Any]) -> tuple[float, dict[str, float], list[str]]:
    validate_slide_plan(slide_plan)
    slides = slide_plan["slides"]
    diagnostics: list[str] = []
    metric_values: dict[str, float] = {"schema_valid": 1.0}

    valid_template_count = 0
    visual_template_count = 0
    zero_visual_correct_count = 0
    formula_rule_count = 0
    formula_rule_total = 0
    bullet_count_scores: list[float] = []
    bullet_length_scores: list[float] = []
    directness_scores: list[float] = []
    supported_asset_scores: list[float] = []

    formula_templates = {"T14_ImageRight_1Formula", "T15_ImageLeft_1Formula", "T16_1Img_2formula_TopTextBottom", "T17_2Img_1formula_TopTextBottom", "T18_2formula_TopTextBottom"}
    meta_phrases = (
        "this section",
        "the paper",
        "the authors",
        "describes",
        "discusses",
        "presents",
        "highlights",
        "summarizes",
    )

    for index, slide in enumerate(slides, start=1):
        template_id = str(slide.get("template_id") or "")
        visuals = list(slide.get("images") or []) + list(slide.get("tables") or []) + list(slide.get("formulas") or [])
        formulas = list(slide.get("formulas") or [])
        bullets = list(slide.get("bullets") or [])

        if template_id.startswith("T") and "_" in template_id:
            valid_template_count += 1
        else:
            diagnostics.append(f"slide {index}: suspicious template_id={template_id!r}")

        if visuals:
            visual_template_count += 0 if template_id == "T1_TextOnly" else 1
            if template_id == "T1_TextOnly":
                diagnostics.append(f"slide {index}: has visuals but uses T1_TextOnly")
        else:
            zero_visual_correct_count += 1 if template_id == "T1_TextOnly" else 0
            if template_id != "T1_TextOnly":
                diagnostics.append(f"slide {index}: no visuals but uses {template_id}")

        if formulas:
            formula_rule_total += 1
            if template_id in formula_templates or (len(formulas) == 1 and len(visuals) == 1 and template_id in {"T2_ImageRight", "T3_ImageLeft", "T4_ImageTop"}):
                formula_rule_count += 1
            else:
                diagnostics.append(f"slide {index}: formula present with non-formula-capable template {template_id}")

        top_count = len(bullets)
        bullet_count_scores.append(1.0 if 1 <= top_count <= 6 else 0.4)

        length_ok = 0
        length_total = 0
        direct_ok = 0
        for bullet in bullets:
            text = str(bullet.get("text") or "") if isinstance(bullet, dict) else str(bullet)
            words = text.split()
            length_total += 1
            if len(words) <= 20:
                length_ok += 1
            lowered = text.lower()
            if not any(phrase in lowered for phrase in meta_phrases):
                direct_ok += 1
            for sub in (bullet.get("sub") or []) if isinstance(bullet, dict) else []:
                sub_words = str(sub).split()
                length_total += 1
                if len(sub_words) <= 25:
                    length_ok += 1
                if not any(phrase in str(sub).lower() for phrase in meta_phrases):
                    direct_ok += 1
        bullet_length_scores.append(length_ok / length_total if length_total else 0.6)
        directness_scores.append(direct_ok / length_total if length_total else 0.6)

        unique_visuals = {str(item) for item in visuals}
        supported_asset_scores.append(1.0 if len(unique_visuals) == len(visuals) else 0.7)

    slide_count = len(slides)
    visual_slide_total = sum(1 for slide in slides if (slide.get("images") or slide.get("tables") or slide.get("formulas")))
    zero_visual_total = slide_count - visual_slide_total

    metric_values["valid_template_ids"] = valid_template_count / slide_count
    metric_values["visual_template_consistency"] = visual_template_count / visual_slide_total if visual_slide_total else 1.0
    metric_values["text_only_consistency"] = zero_visual_correct_count / zero_visual_total if zero_visual_total else 1.0
    metric_values["formula_template_consistency"] = formula_rule_count / formula_rule_total if formula_rule_total else 1.0
    metric_values["bullet_count_bounds"] = sum(bullet_count_scores) / len(bullet_count_scores)
    metric_values["bullet_length_bounds"] = sum(bullet_length_scores) / len(bullet_length_scores)
    metric_values["content_directness"] = sum(directness_scores) / len(directness_scores)
    metric_values["asset_deduplication"] = sum(supported_asset_scores) / len(supported_asset_scores)

    if slide_count < 4:
        metric_values["slide_count_reasonable"] = 0.5
        diagnostics.append(f"slide count is very low: {slide_count}")
    elif slide_count > 40:
        metric_values["slide_count_reasonable"] = 0.5
        diagnostics.append(f"slide count is very high: {slide_count}")
    else:
        metric_values["slide_count_reasonable"] = 1.0

    score = aggregate_score(metric_values)
    return score, metric_values, diagnostics[:20]


def write_candidate_plan(
    *,
    slide_plan: dict[str, Any],
    args: SimpleNamespace,
) -> Path:
    prefix = model_prefix(args.model_name_t, args.model_name_v)
    path = REPO_ROOT / "contents" / args.paper_name / f"{prefix}_slide_plan_baseline.json"
    path.write_text(json.dumps(slide_plan, indent=2, ensure_ascii=False), encoding="utf-8")
    return path


def run_evaluation_bundle(
    *,
    generated_pptx: Path,
    output_dir: Path,
    paper: PilotPaper,
    papers_csv: Path,
    core_coverage_model: str,
    judge_model: str,
    request_timeout: float,
    verbose: bool,
) -> dict[str, Any]:
    command = [
        sys.executable,
        str(REPO_ROOT / "Capstone" / "evaluate_pptx_bundle.py"),
        "--generated-pptx",
        str(generated_pptx),
        "--papers-csv",
        str(papers_csv),
        "--paper-id",
        paper.paper_id,
        "--output-dir",
        str(output_dir),
        "--core-coverage-model",
        core_coverage_model,
        "--judge-model",
        judge_model,
        "--request-timeout",
        str(request_timeout),
    ]
    if verbose:
        command.append("--verbose")
    completed = subprocess.run(
        command,
        cwd=REPO_ROOT,
        check=True,
        text=True,
        capture_output=True,
    )
    summary_path = output_dir / "summary.json"
    if not summary_path.exists():
        raise FileNotFoundError(
            f"Evaluation bundle did not write summary.json. stdout={completed.stdout[:1000]} stderr={completed.stderr[:1000]}"
        )
    return json.loads(summary_path.read_text(encoding="utf-8"))


def subprocess_error_summary(exc: subprocess.CalledProcessError) -> dict[str, Any]:
    return {
        "returncode": exc.returncode,
        "cmd": [str(part) for part in (exc.cmd or [])],
        "stdout_tail": (exc.stdout or "")[-4000:] if isinstance(exc.stdout, str) else "",
        "stderr_tail": (exc.stderr or "")[-4000:] if isinstance(exc.stderr, str) else "",
    }


def metric_values_from_summary(summary: dict[str, Any]) -> tuple[dict[str, float], list[str]]:
    values: dict[str, float] = {}
    diagnostics: list[str] = []
    metrics = summary.get("metrics") or {}
    skipped = summary.get("skipped") or {}
    for metric_name, keys in SCORE_KEYS.items():
        metric_payload = metrics.get(metric_name)
        if not isinstance(metric_payload, dict):
            if metric_name in skipped:
                diagnostics.append(f"{metric_name} skipped: {skipped[metric_name]}")
            continue
        for key in keys:
            raw_value = metric_payload.get(key)
            if raw_value is None:
                continue
            try:
                value = float(raw_value)
            except Exception:
                continue
            if key == "flow_ngld":
                value = 1.0 - value
            values[f"{metric_name}.{key}"] = max(0.0, min(1.0, value))
            break
    return values, diagnostics


def aggregate_score(values: dict[str, float]) -> float:
    if not values:
        return 0.0
    return round(sum(values.values()) / len(values), 6)


def candidate_id(candidate: str) -> str:
    return hashlib.sha1(candidate.encode("utf-8")).hexdigest()[:12]


def make_eval_args(args: argparse.Namespace, staged_paper_name: str) -> SimpleNamespace:
    return SimpleNamespace(
        paper_name=staged_paper_name,
        model_name_t=args.model_name_t,
        model_name_v=args.model_name_v,
        formula_mode=args.formula_mode,
        use_author_preferences=False,
    )


def save_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def mean_score(records: list[dict[str, Any]]) -> float:
    if not records:
        return 0.0
    return round(sum(float(record.get("score") or 0.0) for record in records) / len(records), 6)


def mean_metric_values(records: list[dict[str, Any]]) -> dict[str, float]:
    totals: dict[str, float] = {}
    counts: dict[str, int] = {}
    for record in records:
        for key, value in dict(record.get("metric_values") or {}).items():
            try:
                numeric = float(value)
            except Exception:
                continue
            totals[key] = totals.get(key, 0.0) + numeric
            counts[key] = counts.get(key, 0) + 1
    return {key: round(totals[key] / counts[key], 6) for key in sorted(totals)}


def write_records_csv(path: Path, records: list[dict[str, Any]]) -> None:
    metric_keys = sorted({key for record in records for key in dict(record.get("metric_values") or {})})
    fieldnames = [
        "paper_id",
        "paper_name",
        "score",
        "eval_mode",
        "slide_count",
        "input_tokens",
        "output_tokens",
        "eval_dir",
        "error",
    ] + metric_keys
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for record in records:
            token_usage = dict(record.get("token_usage") or {})
            row = {
                "paper_id": record.get("paper_id"),
                "paper_name": record.get("paper_name"),
                "score": record.get("score"),
                "eval_mode": record.get("eval_mode"),
                "slide_count": record.get("slide_count"),
                "input_tokens": token_usage.get("input_tokens"),
                "output_tokens": token_usage.get("output_tokens"),
                "eval_dir": record.get("eval_dir"),
                "error": record.get("error"),
            }
            for key in metric_keys:
                row[key] = dict(record.get("metric_values") or {}).get(key)
            writer.writerow(row)


def evaluate_prompt_on_papers(
    *,
    evaluator,
    candidate: str,
    papers: list[PilotPaper],
    run_dir: Path,
    label: str,
) -> dict[str, Any]:
    records: list[dict[str, Any]] = []
    for paper in papers:
        _score, side_info = evaluator(candidate, paper)
        records.append(side_info)

    summary = {
        "label": label,
        "paper_count": len(records),
        "average_score": mean_score(records),
        "average_metric_values": mean_metric_values(records),
        "records": records,
    }
    output_dir = run_dir / "aggregate"
    save_json(output_dir / f"{label}.json", summary)
    write_records_csv(output_dir / f"{label}.csv", records)
    return summary


def compare_aggregate_summaries(
    *,
    run_dir: Path,
    baseline: dict[str, Any],
    candidate: dict[str, Any],
    label: str,
) -> dict[str, Any]:
    baseline_score = float(baseline.get("average_score") or 0.0)
    candidate_score = float(candidate.get("average_score") or 0.0)
    absolute_lift = round(candidate_score - baseline_score, 6)
    relative_lift_pct = round((absolute_lift / baseline_score) * 100.0, 3) if baseline_score else None
    metric_lift = _metric_lift(
        dict(baseline.get("average_metric_values") or {}),
        dict(candidate.get("average_metric_values") or {}),
    )
    comparison = {
        "label": label,
        "baseline_label": baseline.get("label"),
        "candidate_label": candidate.get("label"),
        "paper_count": candidate.get("paper_count"),
        "baseline_average_score": baseline_score,
        "candidate_average_score": candidate_score,
        "absolute_lift": absolute_lift,
        "relative_lift_pct": relative_lift_pct,
        "metric_lift": metric_lift,
    }
    save_json(run_dir / "aggregate" / f"{label}.comparison.json", comparison)
    return comparison


def build_evaluator(
    *,
    args: argparse.Namespace,
    prompt_cfg: dict[str, Any],
    run_dir: Path,
    papers: list[PilotPaper],
):
    eval_counter = {"count": 0}

    def evaluate(
        candidate: str,
        example: PilotPaper | dict[str, Any] | None = None,
        **_kwargs: Any,
    ) -> tuple[float, dict[str, Any]]:
        eval_counter["count"] += 1
        if isinstance(example, dict):
            paper = PilotPaper(
                paper_id=example["paper_id"],
                paper_name=example["paper_name"],
                paper_path=Path(example["paper_path"]) if example.get("paper_path") else None,
                title=example.get("title", ""),
            )
        elif isinstance(example, PilotPaper):
            paper = example
        else:
            paper = papers[(eval_counter["count"] - 1) % len(papers)]

        cid = candidate_id(candidate)
        staged_paper_name = f"{paper.paper_name}__gepa_{run_dir.name}_{cid}_{eval_counter['count']:04d}"
        eval_dir = run_dir / "evaluations" / f"{eval_counter['count']:04d}_{sanitize_key(paper.paper_id)}_{cid}"
        eval_dir.mkdir(parents=True, exist_ok=True)
        (eval_dir / "candidate_system_prompt.txt").write_text(candidate, encoding="utf-8")

        try:
            stage_candidate_paper(
                original_paper_name=paper.paper_name,
                staged_paper_name=staged_paper_name,
                model_name_t=args.model_name_t,
                model_name_v=args.model_name_v,
                formula_mode=args.formula_mode,
            )
            eval_args = make_eval_args(args, staged_paper_name)
            slide_plan, token_usage = generate_candidate_slide_plan(
                prompt_cfg=prompt_cfg,
                candidate_system_prompt=candidate,
                args=eval_args,
            )
            plan_path = write_candidate_plan(slide_plan=slide_plan, args=eval_args)
            shutil.copy2(plan_path, eval_dir / plan_path.name)
            if args.eval_mode == "plan-only":
                score, values, diagnostics = score_plan_only(slide_plan)
                copied_pptx = None
            else:
                generate_pptx_from_plan(eval_args, args.template_id)
                pptx_path = (
                    REPO_ROOT
                    / "contents"
                    / staged_paper_name
                    / f"{args.model_name_t}_{args.model_name_v}_output_slides_baseline.pptx"
                )
                if not pptx_path.exists():
                    raise FileNotFoundError(f"Renderer did not produce expected PPTX: {pptx_path}")

                copied_pptx = eval_dir / pptx_path.name
                shutil.copy2(pptx_path, copied_pptx)
                bundle_dir = eval_dir / "bundle"
                summary = run_evaluation_bundle(
                    generated_pptx=copied_pptx,
                    output_dir=bundle_dir,
                    paper=paper,
                    papers_csv=args.papers_csv,
                    core_coverage_model=args.core_coverage_model,
                    judge_model=args.judge_model,
                    request_timeout=args.request_timeout,
                    verbose=args.verbose,
                )
                values, diagnostics = metric_values_from_summary(summary)
                score = aggregate_score(values)
            side_info = {
                "paper_id": paper.paper_id,
                "paper_name": paper.paper_name,
                "staged_paper_name": staged_paper_name,
                "eval_mode": args.eval_mode,
                "score": score,
                "metric_values": values,
                "diagnostics": diagnostics,
                "slide_count": len(slide_plan.get("slides") or []),
                "token_usage": token_usage,
                "eval_dir": str(eval_dir),
            }
            if copied_pptx is not None:
                side_info["pptx_path"] = str(copied_pptx)
            save_json(eval_dir / "score.json", side_info)
            return score, side_info
        except Exception as exc:
            side_info = {
                "paper_id": paper.paper_id,
                "paper_name": paper.paper_name,
                "score": 0.0,
                "error": f"{type(exc).__name__}: {exc}",
                "eval_dir": str(eval_dir),
            }
            if isinstance(exc, subprocess.CalledProcessError):
                side_info["subprocess"] = subprocess_error_summary(exc)
            save_json(eval_dir / "score.json", side_info)
            return 0.0, side_info
        finally:
            if not args.keep_staged_artifacts:
                cleanup_candidate_paper(
                    staged_paper_name=staged_paper_name,
                    model_name_t=args.model_name_t,
                    model_name_v=args.model_name_v,
                )

    return evaluate


def import_gepa():
    try:
        import gepa.optimize_anything as oa
        from gepa.optimize_anything import EngineConfig, GEPAConfig, ReflectionConfig, optimize_anything
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "GEPA is not installed. Install it only for offline optimization with:\n"
            "  python -m pip install gepa\n"
            "Normal slide generation does not require GEPA."
        ) from exc
    return oa, optimize_anything, GEPAConfig, EngineConfig, ReflectionConfig


def message_text_from_reflection_messages(messages: Any) -> tuple[str | None, str]:
    if isinstance(messages, str):
        return None, messages
    if not isinstance(messages, list):
        return None, str(messages)

    system_parts: list[str] = []
    user_parts: list[str] = []
    for message in messages:
        if isinstance(message, dict):
            role = str(message.get("role") or "user")
            content = message.get("content")
        else:
            role = str(getattr(message, "role", "user") or "user")
            content = getattr(message, "content", "")

        if isinstance(content, list):
            text = "\n".join(
                str(part.get("text", part)) if isinstance(part, dict) else str(part)
                for part in content
            )
        else:
            text = str(content)

        if role == "system":
            system_parts.append(text)
        else:
            user_parts.append(text)

    return "\n".join(system_parts) or None, "\n\n".join(user_parts)


def build_reflection_lm(reflection_lm: str, model_name_t: str):
    if reflection_lm != "direct":
        return reflection_lm

    def direct_reflection_lm(messages: Any, **_kwargs: Any) -> str:
        system_prompt, user_prompt = message_text_from_reflection_messages(messages)
        client = build_openai_client()
        text, _in_tok, _out_tok = openai_chat_text(
            client=client,
            model=resolve_direct_model_name(model_name_t),
            user_prompt=user_prompt,
            system_prompt=system_prompt,
            prefer_responses=False,
        )
        return text

    return direct_reflection_lm


def _score_json_files(run_dir: Path) -> list[dict[str, Any]]:
    score_files = sorted((run_dir / "evaluations").glob("*/score.json"))
    records: list[dict[str, Any]] = []
    for path in score_files:
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            continue
        payload["score_path"] = str(path)
        records.append(payload)
    return records


def _metric_lift(seed_metrics: dict[str, Any], best_metrics: dict[str, Any]) -> dict[str, float]:
    keys = sorted(set(seed_metrics) | set(best_metrics))
    lift: dict[str, float] = {}
    for key in keys:
        try:
            seed_value = float(seed_metrics.get(key, 0.0))
            best_value = float(best_metrics.get(key, 0.0))
        except Exception:
            continue
        lift[key] = round(best_value - seed_value, 6)
    return lift


def write_lift_summary(
    *,
    run_dir: Path,
    seed_side_info: dict[str, Any] | None = None,
    best_candidate_id: str | None = None,
) -> dict[str, Any]:
    records = _score_json_files(run_dir)
    if not records:
        summary = {"error": "No score.json files found.", "run_dir": str(run_dir)}
        save_json(run_dir / "lift_summary.json", summary)
        return summary

    if seed_side_info is None:
        seed_record = records[0]
    else:
        seed_eval_dir = seed_side_info.get("eval_dir")
        seed_record = next((record for record in records if record.get("eval_dir") == seed_eval_dir), seed_side_info)

    if best_candidate_id:
        best_record = next(
            (
                record
                for record in records
                if best_candidate_id in str(record.get("eval_dir", ""))
                or best_candidate_id in str(record.get("staged_paper_name", ""))
            ),
            None,
        )
    else:
        best_record = None
    if best_record is None:
        best_record = max(records, key=lambda record: float(record.get("score") or 0.0))

    seed_score = float(seed_record.get("score") or 0.0)
    best_score = float(best_record.get("score") or 0.0)
    absolute_lift = round(best_score - seed_score, 6)
    relative_lift_pct = round((absolute_lift / seed_score) * 100.0, 3) if seed_score else None
    metric_lift = _metric_lift(
        dict(seed_record.get("metric_values") or {}),
        dict(best_record.get("metric_values") or {}),
    )
    summary = {
        "run_dir": str(run_dir),
        "seed_score": seed_score,
        "best_score": best_score,
        "absolute_lift": absolute_lift,
        "relative_lift_pct": relative_lift_pct,
        "metric_lift": metric_lift,
        "seed_eval_dir": seed_record.get("eval_dir"),
        "best_eval_dir": best_record.get("eval_dir"),
        "seed_score_path": seed_record.get("score_path"),
        "best_score_path": best_record.get("score_path"),
        "best_candidate_id": best_candidate_id,
        "eval_count": len(records),
    }
    save_json(run_dir / "lift_summary.json", summary)
    return summary


def write_comparison_only(run_dir: Path) -> dict[str, Any]:
    aggregate_dir = run_dir / "aggregate"
    seed_path = aggregate_dir / "seed.json"
    candidate_path = aggregate_dir / "candidate.json"
    if not seed_path.exists() or not candidate_path.exists():
        raise FileNotFoundError(f"Expected {seed_path} and {candidate_path}")
    return compare_aggregate_summaries(
        run_dir=run_dir,
        baseline=json.loads(seed_path.read_text(encoding="utf-8")),
        candidate=json.loads(candidate_path.read_text(encoding="utf-8")),
        label="seed_vs_candidate",
    )


def promote_best_prompt(best_prompt_path: Path, prompt_path: Path) -> None:
    if not best_prompt_path.exists():
        raise FileNotFoundError(f"Best prompt file not found: {best_prompt_path}")
    best_prompt = best_prompt_path.read_text(encoding="utf-8").rstrip() + "\n"
    original = prompt_path.read_text(encoding="utf-8")
    marker = "template:"
    template_start = original.find(marker)
    system_start = original.find("system_prompt:")
    if system_start == -1 or template_start == -1 or template_start <= system_start:
        prompt_cfg = load_prompt_config(prompt_path)
        prompt_cfg["system_prompt"] = best_prompt.rstrip("\n")
        prompt_path.write_text(yaml.safe_dump(prompt_cfg, sort_keys=False, allow_unicode=True), encoding="utf-8")
        return

    replacement = "system_prompt: |\n" + "".join(f"  {line}\n" for line in best_prompt.splitlines())
    prompt_path.write_text(replacement + "\n" + original[template_start:], encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run an offline GEPA pilot for the layout planner system prompt.")
    parser.add_argument("--papers-csv", type=Path, default=DEFAULT_PAPERS_CSV)
    parser.add_argument("--prompt-path", type=Path, default=DEFAULT_PROMPT_PATH)
    parser.add_argument("--candidate-system-prompt", type=Path, default=None, help="Use a prompt text file as the candidate system prompt without editing the production YAML.")
    parser.add_argument("--run-root", type=Path, default=DEFAULT_RUN_ROOT)
    parser.add_argument("--run-id", default=None)
    parser.add_argument("--paper-limit", type=int, default=3)
    parser.add_argument("--paper-id", action="append", default=None, help="Restrict selected papers to one or more paper ids. Can be repeated or comma-separated.")
    parser.add_argument("--exclude-paper-id", action="append", default=None, help="Exclude one or more paper ids. Can be repeated or comma-separated.")
    parser.add_argument("--max-metric-calls", type=int, default=20)
    parser.add_argument("--model-name-t", default="4o-mini")
    parser.add_argument("--model-name-v", default="4o-mini")
    parser.add_argument("--formula-mode", type=int, choices=[1, 2, 3], default=1)
    parser.add_argument("--outline-mode", choices=["high_level", "technical"], default="high_level")
    parser.add_argument("--template-id", type=int, default=3)
    parser.add_argument("--reflection-lm", default="openai/gpt-5")
    parser.add_argument("--core-coverage-model", default=None)
    parser.add_argument("--judge-model", default=None)
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--seed", type=int, default=0)
    parser.add_argument("--eval-mode", choices=["plan-only", "full-bundle"], default="plan-only")
    parser.add_argument("--include-missing-artifacts", action="store_true")
    parser.add_argument("--keep-staged-artifacts", action="store_true")
    parser.add_argument("--evaluate-seed-only", action="store_true", help="Run one seed-prompt evaluation without importing GEPA.")
    parser.add_argument("--evaluate-all-selected", action="store_true", help="Evaluate the current seed/candidate prompt across every selected paper and write aggregate JSON/CSV.")
    parser.add_argument("--compare-with-seed", action="store_true", help="With --candidate-system-prompt and --evaluate-all-selected, also evaluate the production seed prompt and write a comparison report.")
    parser.add_argument("--write-lift-summary-only", type=Path, default=None, help="Recompute lift_summary.json for an existing GEPA run directory.")
    parser.add_argument("--write-comparison-only", type=Path, default=None, help="Recompute aggregate/seed_vs_candidate.comparison.json for an existing run with aggregate seed/candidate JSON.")
    parser.add_argument("--promote-best", type=Path, default=None, help="Replace prompt YAML system_prompt from a reviewed best prompt file.")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    load_dotenv(REPO_ROOT / ".env")
    args.papers_csv = args.papers_csv.resolve()
    args.prompt_path = args.prompt_path.resolve()
    if args.candidate_system_prompt is not None:
        args.candidate_system_prompt = args.candidate_system_prompt.resolve()
    args.run_root = args.run_root.resolve()

    if args.write_lift_summary_only:
        summary = write_lift_summary(run_dir=args.write_lift_summary_only.resolve())
        print(json.dumps(summary, indent=2, ensure_ascii=False))
        return
    if args.write_comparison_only:
        comparison = write_comparison_only(args.write_comparison_only.resolve())
        print(json.dumps(comparison, indent=2, ensure_ascii=False))
        return

    if args.promote_best:
        promote_best_prompt(args.promote_best.resolve(), args.prompt_path)
        print(f"Promoted reviewed prompt from {args.promote_best} into {args.prompt_path}")
        return

    if args.paper_limit <= 0:
        raise SystemExit("--paper-limit must be positive")
    if args.max_metric_calls <= 0:
        raise SystemExit("--max-metric-calls must be positive")
    if args.core_coverage_model is None:
        args.core_coverage_model = args.model_name_t
    if args.judge_model is None:
        args.judge_model = args.model_name_t

    prompt_cfg = load_prompt_config(args.prompt_path)
    production_seed_prompt = str(prompt_cfg["system_prompt"])
    run_id = args.run_id or time.strftime("%Y%m%d_%H%M%S")
    run_dir = args.run_root / sanitize_key(run_id)
    run_dir.mkdir(parents=True, exist_ok=True)
    if args.candidate_system_prompt is not None:
        if not args.candidate_system_prompt.exists():
            raise FileNotFoundError(f"Candidate system prompt not found: {args.candidate_system_prompt}")
        seed_prompt = args.candidate_system_prompt.read_text(encoding="utf-8")
    else:
        seed_prompt = production_seed_prompt
    (run_dir / "seed_system_prompt.txt").write_text(seed_prompt, encoding="utf-8")

    include_paper_ids = parse_paper_id_set(args.paper_id)
    exclude_paper_ids = parse_paper_id_set(args.exclude_paper_id)
    papers = load_pilot_papers(
        papers_csv=args.papers_csv,
        model_name_t=args.model_name_t,
        model_name_v=args.model_name_v,
        formula_mode=args.formula_mode,
        outline_mode=args.outline_mode,
        limit=args.paper_limit,
        include_missing_artifacts=args.include_missing_artifacts,
        include_paper_ids=include_paper_ids,
        exclude_paper_ids=exclude_paper_ids,
    )
    save_json(
        run_dir / "run_config.json",
        {
            "prompt_path": str(args.prompt_path),
            "candidate_system_prompt": str(args.candidate_system_prompt) if args.candidate_system_prompt else None,
            "papers_csv": str(args.papers_csv),
            "papers": [paper.__dict__ | {"paper_path": str(paper.paper_path) if paper.paper_path else None} for paper in papers],
            "paper_ids": sorted(include_paper_ids) if include_paper_ids else None,
            "exclude_paper_ids": sorted(exclude_paper_ids) if exclude_paper_ids else None,
            "model_name_t": args.model_name_t,
            "model_name_v": args.model_name_v,
            "formula_mode": args.formula_mode,
            "outline_mode": args.outline_mode,
            "max_metric_calls": args.max_metric_calls,
            "eval_mode": args.eval_mode,
            "reflection_lm": args.reflection_lm,
            "judge_model": args.judge_model,
        },
    )

    evaluator = build_evaluator(args=args, prompt_cfg=prompt_cfg, run_dir=run_dir, papers=papers)
    if args.evaluate_all_selected:
        candidate_summary = evaluate_prompt_on_papers(
            evaluator=evaluator,
            candidate=seed_prompt,
            papers=papers,
            run_dir=run_dir,
            label="candidate" if args.candidate_system_prompt else "seed",
        )
        output: dict[str, Any] = {
            "run_dir": str(run_dir),
            "candidate_summary": candidate_summary,
        }
        if args.compare_with_seed and args.candidate_system_prompt is not None:
            seed_summary = evaluate_prompt_on_papers(
                evaluator=evaluator,
                candidate=production_seed_prompt,
                papers=papers,
                run_dir=run_dir,
                label="seed",
            )
            output["seed_summary"] = seed_summary
            output["comparison"] = compare_aggregate_summaries(
                run_dir=run_dir,
                baseline=seed_summary,
                candidate=candidate_summary,
                label="seed_vs_candidate",
            )
        print(json.dumps(output, indent=2, ensure_ascii=False))
        return

    seed_score, seed_side_info = evaluator(seed_prompt, papers[0])

    if args.evaluate_seed_only:
        lift_summary = write_lift_summary(run_dir=run_dir, seed_side_info=seed_side_info)
        print(json.dumps({"score": seed_score, "side_info": seed_side_info, "lift_summary": lift_summary, "run_dir": str(run_dir)}, indent=2, ensure_ascii=False))
        return

    _oa, optimize_anything, GEPAConfig, EngineConfig, ReflectionConfig = import_gepa()
    dataset = [
        {
            "paper_id": paper.paper_id,
            "paper_name": paper.paper_name,
            "paper_path": str(paper.paper_path) if paper.paper_path else None,
            "title": paper.title,
        }
        for paper in papers
    ]
    result = optimize_anything(
        seed_candidate=seed_prompt,
        evaluator=evaluator,
        dataset=dataset,
        objective=(
            "Optimize the SlideGen layout planner system prompt so generated slide-plan JSON "
            "produces high-quality scientific presentation decks under the existing renderer "
            "and full evaluation bundle."
        ),
        background=(
            "Only the system prompt may change. The Jinja user template, JSON schema, template ids, "
            "asset filenames, and downstream renderer contracts are fixed. Candidates must return "
            "valid JSON only and must not invent unsupported paper content or visual assets."
        ),
        config=GEPAConfig(
            engine=EngineConfig(
                run_dir=str(run_dir / "gepa_state"),
                seed=args.seed,
                max_metric_calls=args.max_metric_calls,
                display_progress_bar=True,
                parallel=False,
                cache_evaluation=True,
                capture_stdio=False,
            ),
            reflection=ReflectionConfig(
                reflection_lm=build_reflection_lm(args.reflection_lm, args.model_name_t)
            ),
        ),
    )

    best_candidate = getattr(result, "best_candidate", None)
    if best_candidate is None and hasattr(result, "candidates") and hasattr(result, "best_idx"):
        best_candidate = result.candidates[result.best_idx]
    if not isinstance(best_candidate, str):
        best_candidate = str(best_candidate)
    best_prompt_path = run_dir / "best_system_prompt.txt"
    best_prompt_path.write_text(best_candidate, encoding="utf-8")
    best_id = candidate_id(best_candidate)
    lift_summary = write_lift_summary(
        run_dir=run_dir,
        seed_side_info=seed_side_info,
        best_candidate_id=best_id,
    )

    summary = {
        "run_dir": str(run_dir),
        "best_prompt_path": str(best_prompt_path),
        "lift_summary_path": str(run_dir / "lift_summary.json"),
        "seed_score": seed_score,
        "best_score": lift_summary.get("best_score"),
        "absolute_lift": lift_summary.get("absolute_lift"),
        "relative_lift_pct": lift_summary.get("relative_lift_pct"),
        "total_metric_calls": getattr(result, "total_metric_calls", None),
        "best_idx": getattr(result, "best_idx", None),
        "val_aggregate_scores": getattr(result, "val_aggregate_scores", None),
    }
    save_json(run_dir / "result_summary.json", summary)
    print(json.dumps(summary, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
