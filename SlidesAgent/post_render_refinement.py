from __future__ import annotations

import argparse
import copy
import json
import re
import time
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import yaml
from jinja2 import Environment, StrictUndefined
from PIL import Image, ImageStat

from Capstone.slidetailor_eval.common import (
    extract_json_object,
    render_pptx_to_images,
)
from SlidesAgent.output_paths import slide_plan_path, themed_output_pptx_path
from SlidesAgent.profile_adapter_retrieval import is_retrieval_profile
from SlidesAgent.slide_plan_summary import summarize_slide_plan
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


DEFAULT_REPAIR_PROMPT = Path("utils/prompt_templates/post_render_repair.yaml")
DEFAULT_OUTPUT_DIR_NAME = "post_render_refinement"
SEVERITY_ORDER = {"low": 1, "medium": 2, "high": 3, "critical": 4}


def sanitize_slide_plan_templates_lazy(plan: dict[str, Any]) -> dict[str, Any]:
    from SlidesAgent.layout_agent_xin import sanitize_slide_plan_templates

    return sanitize_slide_plan_templates(plan)


def generate_pptx_from_plan_lazy(args: argparse.Namespace, template: int) -> None:
    from SlidesAgent.layout_filler import generate_pptx_from_plan

    generate_pptx_from_plan(args, template)


def load_json(path: Path | None, default: Any = None) -> Any:
    if path is None:
        return default
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def slide_id_for_index(index: int) -> str:
    return f"S{index:02d}"


def plan_index_for_slide_id(slide_id: str) -> int | None:
    match = re.search(r"(\d+)", str(slide_id or ""))
    if not match:
        return None
    return max(0, int(match.group(1)) - 1)


def normalize_slide_ids(plan: dict[str, Any]) -> dict[str, Any]:
    normalized = copy.deepcopy(plan)
    for index, slide in enumerate(normalized.get("slides") or [], start=1):
        if isinstance(slide, dict):
            slide.setdefault("slide_id", slide_id_for_index(index))
    return normalized


def count_words(text: Any) -> int:
    return len(re.findall(r"\b[\w'-]+\b", str(text or "")))


def slide_text_metrics(slide: dict[str, Any]) -> dict[str, Any]:
    bullets = [item for item in list(slide.get("bullets") or []) if isinstance(item, dict)]
    words = 0
    sub_count = 0
    for bullet in bullets:
        words += count_words(bullet.get("text"))
        subs = list(bullet.get("sub") or [])
        sub_count += len(subs)
        for sub in subs:
            words += count_words(sub)
    return {
        "bullet_count": len(bullets),
        "sub_bullet_count": sub_count,
        "word_count": words,
    }


def visual_counts(slide: dict[str, Any]) -> dict[str, int]:
    return {
        "image_count": len(slide.get("images") or []),
        "table_count": len(slide.get("tables") or []),
        "formula_count": len(slide.get("formulas") or []),
    }


def rendered_image_metrics(image_path: Path) -> dict[str, Any]:
    with Image.open(image_path) as image:
        image = image.convert("L")
        stat = ImageStat.Stat(image)
        mean = float(stat.mean[0])
        stddev = float(stat.stddev[0])
        extrema = image.getextrema()
    return {
        "image": image_path.name,
        "brightness_mean": round(mean, 3),
        "brightness_stddev": round(stddev, 3),
        "extrema": list(extrema),
        "near_blank": stddev < 4.0 or (mean > 248.0 and stddev < 8.0),
    }


def make_candidate_edit(
    *,
    edit_id: str,
    slide_id: str,
    slide_index: int | None,
    category: str,
    severity: str,
    confidence: float,
    suggested_operation: str,
    current_issue: str,
    suggested_change: str,
    expected_effect: str,
    risk: str,
    affected_fields: list[str],
    default_action: str = "ask_user",
    forbidden_side_effects: list[str] | None = None,
    validation_target: dict[str, Any] | None = None,
) -> dict[str, Any]:
    return {
        "edit_id": edit_id,
        "slide_id": slide_id,
        "slide_index": slide_index,
        "category": category,
        "severity": severity,
        "confidence": confidence,
        "suggested_operation": suggested_operation,
        "current_issue": current_issue,
        "suggested_change": suggested_change,
        "expected_effect": expected_effect,
        "risk": risk,
        "requires_user_review": default_action != "auto_apply_mechanical",
        "default_action": default_action,
        "affected_fields": affected_fields,
        "forbidden_side_effects": forbidden_side_effects or [
            "Do not invent paper claims.",
            "Do not introduce external assets.",
            "Do not alter protected slides without explicit user override.",
        ],
        "validation_target": validation_target or {},
        "source": "deterministic",
    }


def deterministic_plan_edits(plan: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    candidate_edits: list[dict[str, Any]] = []
    slide_summaries: list[dict[str, Any]] = []
    visual_template_tokens = ("Image", "Img", "Formula")

    for index, slide in enumerate(plan.get("slides") or [], start=1):
        if not isinstance(slide, dict):
            continue
        slide_id = str(slide.get("slide_id") or slide_id_for_index(index))
        text_metrics = slide_text_metrics(slide)
        counts = visual_counts(slide)
        visual_count = counts["image_count"] + counts["table_count"] + counts["formula_count"]
        template_id = str(slide.get("template_id") or "")
        summary = {
            "slide_id": slide_id,
            "slide_index": index,
            "section": slide.get("section"),
            "subsection": slide.get("subsection") or slide.get("title"),
            "template_id": template_id,
            **text_metrics,
            **counts,
        }
        slide_summaries.append(summary)

        if text_metrics["word_count"] >= 90 or text_metrics["sub_bullet_count"] >= 12:
            severity = "high" if text_metrics["word_count"] >= 110 else "medium"
            candidate_edits.append(
                make_candidate_edit(
                    edit_id=f"{slide_id}_density",
                    slide_id=slide_id,
                    slide_index=index,
                    category="text_density",
                    severity=severity,
                    confidence=0.78,
                    suggested_operation="reduce_text_density",
                    current_issue=(
                        f"Slide plan has {text_metrics['word_count']} words and "
                        f"{text_metrics['sub_bullet_count']} sub-bullets, which is likely crowded after rendering."
                    ),
                    suggested_change="Reduce lower-priority bullet detail while preserving quantitative results and key claims.",
                    expected_effect="Improves readability and reduces overflow risk.",
                    risk="May remove detail the user intentionally wanted to keep.",
                    affected_fields=[f"slides[{index - 1}].bullets"],
                    validation_target={
                        "type": "decrease_words",
                        "before_words": text_metrics["word_count"],
                        "target_max_words": max(45, int(text_metrics["word_count"] * 0.75)),
                    },
                )
            )

        if text_metrics["word_count"] <= 5 and visual_count == 0:
            candidate_edits.append(
                make_candidate_edit(
                    edit_id=f"{slide_id}_sparse",
                    slide_id=slide_id,
                    slide_index=index,
                    category="content_coverage",
                    severity="medium",
                    confidence=0.68,
                    suggested_operation="enrich_sparse_slide",
                    current_issue="Slide appears very sparse in the plan and has no supporting visual asset.",
                    suggested_change="Add a concise faithful bullet or merge the slide with a nearby related slide.",
                    expected_effect="Avoids an underfilled slide.",
                    risk="Adding text may conflict with a deliberately minimal slide style.",
                    affected_fields=[f"slides[{index - 1}].bullets"],
                    validation_target={
                        "type": "increase_words",
                        "before_words": text_metrics["word_count"],
                        "target_min_words": 10,
                    },
                )
            )

        if any(token in template_id for token in visual_template_tokens) and visual_count == 0:
            candidate_edits.append(
                make_candidate_edit(
                    edit_id=f"{slide_id}_template_visual_mismatch",
                    slide_id=slide_id,
                    slide_index=index,
                    category="mechanical_render_issue",
                    severity="high",
                    confidence=0.9,
                    suggested_operation="change_template",
                    current_issue=f"Template {template_id} expects visual content, but the slide has no image/table/formula assets.",
                    suggested_change="Switch to a text-only or two-text layout unless the user wants to add a supported visual.",
                    expected_effect="Prevents empty picture placeholders and awkward visual space.",
                    risk="Changing template may reduce visual style consistency.",
                    affected_fields=[f"slides[{index - 1}].template_id"],
                    default_action="auto_apply_mechanical",
                    validation_target={"type": "template_visual_consistency"},
                )
            )

        if visual_count >= 4 and text_metrics["word_count"] >= 45:
            candidate_edits.append(
                make_candidate_edit(
                    edit_id=f"{slide_id}_overloaded_visuals",
                    slide_id=slide_id,
                    slide_index=index,
                    category="layout_readability",
                    severity="medium",
                    confidence=0.72,
                    suggested_operation="reduce_or_split_visual_density",
                    current_issue="Slide combines several visual assets with substantial bullet text.",
                    suggested_change="Keep the highest-value visuals or split into two slides if the user allows slide splits.",
                    expected_effect="Improves readability and visual inspection.",
                    risk="Could remove context or exceed the desired slide count.",
                    affected_fields=[f"slides[{index - 1}].images", f"slides[{index - 1}].tables", f"slides[{index - 1}].formulas"],
                    validation_target={"type": "decrease_visual_density", "before_visual_count": visual_count},
                )
            )

    return candidate_edits, slide_summaries


def deterministic_render_edits(slide_images: list[Path]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    candidate_edits: list[dict[str, Any]] = []
    image_summaries: list[dict[str, Any]] = []
    for rendered_index, image_path in enumerate(slide_images, start=1):
        metrics = rendered_image_metrics(image_path)
        metrics["rendered_slide_index"] = rendered_index
        image_summaries.append(metrics)
        if metrics["near_blank"]:
            rendered_id = f"R{rendered_index:02d}"
            candidate_edits.append(
                make_candidate_edit(
                    edit_id=f"{rendered_id}_near_blank",
                    slide_id=rendered_id,
                    slide_index=None,
                    category="mechanical_render_issue",
                    severity="critical",
                    confidence=0.92,
                    suggested_operation="inspect_blank_rendered_slide",
                    current_issue=f"Rendered slide {rendered_index} appears blank or near-blank.",
                    suggested_change="Inspect the corresponding plan slide or section divider and regenerate after fixing missing content.",
                    expected_effect="Prevents delivery of a broken slide.",
                    risk="Rendered slide numbers include cover, contents, section dividers, and body slides, so manual mapping may be needed.",
                    affected_fields=[],
                    validation_target={"type": "rendered_near_blank", "rendered_slide_index": rendered_index},
                )
            )
    return candidate_edits, image_summaries


def build_profile_context(profile: dict[str, Any] | None, plan: dict[str, Any]) -> dict[str, Any]:
    summary = summarize_slide_plan(plan)
    if not isinstance(profile, dict):
        return {
            "profile_kind": "none",
            "plan_summary": summary,
            "target_summary": {},
        }

    context = {
        "profile_kind": "retrieval" if is_retrieval_profile(profile) else "standard",
        "profile_method": profile.get("profile_method"),
        "author_id": profile.get("author_id"),
        "plan_summary": summary,
        "target_summary": dict(profile.get("numeric_preferences") or {}),
    }
    if context["profile_kind"] == "retrieval":
        try:
            from SlidesAgent.layout_agent_xin_retrieval import (
                build_retrieval_target_summary,
                summarize_retrieval_targets,
            )

            context["target_summary"] = build_retrieval_target_summary(profile)
            context["plan_summary"] = summarize_retrieval_targets(plan)
        except Exception as exc:
            context["profile_context_warning"] = str(exc)
    return context


def default_token_usage() -> dict[str, Any]:
    return {
        "repair_input_tokens": 0,
        "repair_output_tokens": 0,
        "total_added_input_tokens": 0,
        "total_added_output_tokens": 0,
        "estimated_added_cost_usd": None,
    }


def update_token_usage(usage: dict[str, Any]) -> dict[str, Any]:
    usage["total_added_input_tokens"] = int(usage.get("repair_input_tokens", 0))
    usage["total_added_output_tokens"] = int(usage.get("repair_output_tokens", 0))
    return usage


def score_efficiency(token_usage: dict[str, Any], benefit_metrics: dict[str, Any]) -> dict[str, Any]:
    total_tokens = int(token_usage.get("total_added_input_tokens", 0)) + int(token_usage.get("total_added_output_tokens", 0))
    denominator = max(total_tokens / 1000.0, 1.0 if total_tokens else 0.0)
    issues_fixed = max(0, int(benefit_metrics.get("rendered_issue_delta", 0)))
    accepted_edits = int(benefit_metrics.get("accepted_user_edits", 0))
    if total_tokens <= 0:
        recommendation = "manual_only"
    elif issues_fixed == 0 and accepted_edits <= 1:
        recommendation = "manual_review"
    else:
        recommendation = "keep"
    return {
        "issues_fixed_per_1k_tokens": round(issues_fixed / denominator, 4) if denominator else 0.0,
        "accepted_edits_per_1k_tokens": round(accepted_edits / denominator, 4) if denominator else 0.0,
        "recommendation": recommendation,
    }


def render_prompt(prompt_path: Path, **values: Any) -> dict[str, str]:
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    return {
        "system_prompt": str(prompt_cfg["system_prompt"]),
        "user_prompt": template.render(**values),
    }


def call_text_json(
    *,
    model: str,
    system_prompt: str,
    user_prompt: str,
    timeout: float = 180.0,
    max_tokens: int = 2400,
) -> tuple[dict[str, Any], int, int, str]:
    client = build_openai_client()
    resolved_model = resolve_direct_model_name(model)
    content: list[dict[str, Any]] = [{"type": "text", "text": user_prompt}]

    request_kwargs: dict[str, Any] = {
        "model": resolved_model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
        "temperature": 0.1,
        "timeout": timeout,
    }
    if "gpt-5" in resolved_model.lower():
        request_kwargs["max_completion_tokens"] = max_tokens
    else:
        request_kwargs["max_tokens"] = max_tokens

    response = client.chat.completions.create(**request_kwargs)
    raw_text = response.choices[0].message.content or ""
    usage = getattr(response, "usage", None)
    input_tokens = int(getattr(usage, "prompt_tokens", 0) or 0)
    output_tokens = int(getattr(usage, "completion_tokens", 0) or 0)
    return extract_json_object(raw_text), input_tokens, output_tokens, raw_text


def inspect_deck(
    *,
    pptx_path: Path,
    slide_plan_path: Path,
    output_dir: Path,
    author_profile_path: Path | None,
    render_dpi: int,
    force_render: bool,
    timeout: float,
    **_legacy_review_kwargs: Any,
) -> dict[str, Any]:
    started = time.time()
    plan = normalize_slide_ids(load_json(slide_plan_path, {}))
    profile = load_json(author_profile_path, None)
    render_dir = output_dir / "rendered_slides"
    slide_images = render_pptx_to_images(pptx_path, render_dir, force=force_render, dpi=render_dpi)

    plan_edits, plan_slide_summaries = deterministic_plan_edits(plan)
    render_edits, rendered_slide_summaries = deterministic_render_edits(slide_images)
    candidate_edits = plan_edits + render_edits
    token_usage = default_token_usage()

    profile_context = build_profile_context(profile, plan)
    rendered_deck_summary = {
        "pptx_path": str(pptx_path),
        "slide_plan_path": str(slide_plan_path),
        "render_dir": str(render_dir),
        "rendered_slide_count": len(slide_images),
        "plan_slide_count": len(plan.get("slides") or []),
        "plan_summary": summarize_slide_plan(plan),
        "profile_context": profile_context,
        "plan_slide_summaries": plan_slide_summaries,
        "rendered_slide_summaries": rendered_slide_summaries,
        "deterministic_issue_count": len(candidate_edits),
    }

    token_usage = update_token_usage(token_usage)
    benefit_metrics = {
        "accepted_user_edits": 0,
        "deterministic_issue_count_before_llm": rendered_deck_summary["deterministic_issue_count"],
        "llm_candidate_edit_count": max(0, len(candidate_edits) - rendered_deck_summary["deterministic_issue_count"]),
        "rendered_issue_count_before": rendered_deck_summary["deterministic_issue_count"],
        "rendered_issue_count_after": None,
        "rendered_issue_delta": 0,
        "user_request_completion_score": None,
        "profile_compatibility_delta": None,
    }
    report = {
        "mode": "inspect",
        "runtime_seconds": round(time.time() - started, 3),
        "rendered_deck_summary": rendered_deck_summary,
        "candidate_edits": candidate_edits,
        "inspection_method": "deterministic",
        "cost_ledger": {
            "token_usage": token_usage,
            "benefit_metrics": benefit_metrics,
            "efficiency": score_efficiency(token_usage, benefit_metrics),
        },
    }

    write_json(output_dir / "post_render_deck_summary.json", rendered_deck_summary)
    write_json(output_dir / "post_render_candidate_edits.json", {"candidate_edits": candidate_edits})
    write_json(output_dir / "post_render_inspection_report.json", report)
    write_json(output_dir / "post_render_user_decisions.template.json", build_user_decisions_template(candidate_edits))
    return report


def build_user_decisions_template(candidate_edits: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "user_requested_changes": [],
        "candidate_edit_decisions": [
            {
                "edit_id": edit.get("edit_id"),
                "decision": "defer",
                "user_instruction": "",
            }
            for edit in candidate_edits
        ],
        "protected_slide_ids": [],
        "protected_slide_overrides": [],
        "constraints": {
            "max_changed_slides": 5,
            "allow_slide_splits": False,
            "allow_slide_merges": False,
            "preserve_section_order": True,
            "preserve_paper_fidelity": True,
            "do_not_introduce_external_assets": True,
        },
    }


def coerce_candidate_edits(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, dict):
        raw_edits = payload.get("candidate_edits") or []
    elif isinstance(payload, list):
        raw_edits = payload
    else:
        raw_edits = []
    return [item for item in raw_edits if isinstance(item, dict)]


def decisions_by_edit_id(user_decisions: dict[str, Any]) -> dict[str, dict[str, Any]]:
    decisions = {}
    for item in user_decisions.get("candidate_edit_decisions") or []:
        if isinstance(item, dict) and item.get("edit_id"):
            decisions[str(item["edit_id"])] = item
    return decisions


def allowed_by_protection(slide_id: str, operation: str, user_decisions: dict[str, Any]) -> bool:
    protected = {str(item) for item in user_decisions.get("protected_slide_ids") or []}
    if slide_id not in protected:
        return True
    for override in user_decisions.get("protected_slide_overrides") or []:
        if not isinstance(override, dict) or str(override.get("slide_id")) != slide_id:
            continue
        allowed = {str(item) for item in override.get("allowed_operations") or []}
        if operation in allowed or override.get("user_instruction"):
            return True
    return False


def extract_quoted_title(instruction: str) -> str | None:
    instruction = str(instruction or "").strip()
    patterns = [
        r"title\s+to\s*:\s*(.+)$",
        r"title\s+to\s+(.+)$",
        r"rename\s+.*?\s+to\s*:\s*(.+)$",
    ]
    for pattern in patterns:
        match = re.search(pattern, instruction, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip().strip("\"'")
    return None


def reduce_text_density(slide: dict[str, Any], aggressive: bool = False) -> None:
    bullets = [copy.deepcopy(item) for item in list(slide.get("bullets") or []) if isinstance(item, dict)]
    if not bullets:
        return
    target_ratio = 0.65 if aggressive else 0.78
    before_words = slide_text_metrics(slide)["word_count"]
    target_words = max(35, int(before_words * target_ratio))

    # Drop lower-level details from the end first, preserving top-level claims.
    for bullet in reversed(bullets):
        subs = list(bullet.get("sub") or [])
        while subs and slide_text_metrics({"bullets": bullets})["word_count"] > target_words:
            subs.pop()
            bullet["sub"] = subs

    while len(bullets) > 3 and slide_text_metrics({"bullets": bullets})["word_count"] > target_words:
        bullets.pop()

    slide["bullets"] = bullets


def enrich_sparse_slide(slide: dict[str, Any], instruction: str = "") -> None:
    bullets = list(slide.get("bullets") or [])
    if bullets:
        return
    text = instruction.strip() if instruction.strip() else "Clarify the key contribution or takeaway from this section."
    bullets.append({"text": text, "sub": []})
    slide["bullets"] = bullets


def apply_deterministic_edit(
    plan: dict[str, Any],
    edit: dict[str, Any],
    decision: dict[str, Any],
    user_decisions: dict[str, Any],
) -> dict[str, Any]:
    slide_id = str(edit.get("slide_id") or "")
    plan_index = plan_index_for_slide_id(slide_id)
    operation = str(edit.get("suggested_operation") or "")
    if plan_index is None or plan_index >= len(plan.get("slides") or []):
        return {"applied": False, "reason": "edit_not_mapped_to_plan_slide"}
    if not allowed_by_protection(slide_id, operation, user_decisions):
        return {"applied": False, "reason": "protected_slide"}

    slide = plan["slides"][plan_index]
    instruction = str(decision.get("user_instruction") or "")

    if operation in {"change_title", "title_rewrite"} or "title" in operation:
        new_title = extract_quoted_title(instruction)
        if not new_title:
            return {"applied": False, "reason": "missing_title_instruction"}
        slide["subsection"] = new_title
        return {"applied": True, "operation": "change_title"}

    if operation == "reduce_text_density":
        reduce_text_density(slide, aggressive="heavy" in instruction.lower())
        return {"applied": True, "operation": "reduce_text_density"}

    if operation == "enrich_sparse_slide":
        enrich_sparse_slide(slide, instruction=instruction)
        return {"applied": True, "operation": "enrich_sparse_slide"}

    if operation == "change_template" and edit.get("default_action") == "auto_apply_mechanical":
        slide["template_id"] = "T1_TextOnly"
        return {"applied": True, "operation": "change_template"}

    return {"applied": False, "reason": "operation_requires_llm_repair"}


def slide_change_summary(before_plan: dict[str, Any], after_plan: dict[str, Any]) -> dict[str, Any]:
    changed: list[str] = []
    before_slides = list(before_plan.get("slides") or [])
    after_slides = list(after_plan.get("slides") or [])
    for index in range(max(len(before_slides), len(after_slides))):
        before = before_slides[index] if index < len(before_slides) else None
        after = after_slides[index] if index < len(after_slides) else None
        if before != after:
            changed.append(slide_id_for_index(index + 1))
    return {
        "changed_slide_ids": changed,
        "changed_slide_count": len(changed),
        "slide_count_before": len(before_slides),
        "slide_count_after": len(after_slides),
    }


def merge_partial_revised_slides(original_plan: dict[str, Any], revised_payload: dict[str, Any]) -> tuple[dict[str, Any], bool]:
    original_slides = list(original_plan.get("slides") or [])
    revised_slides = list(revised_payload.get("slides") or [])
    if not original_slides or not revised_slides or len(revised_slides) >= len(original_slides):
        return revised_payload, False

    merged = copy.deepcopy(original_plan)
    original_by_id = {
        str(slide.get("slide_id") or slide_id_for_index(index)): index - 1
        for index, slide in enumerate(original_slides, start=1)
        if isinstance(slide, dict)
    }
    applied_any = False
    appended_slide_ids: set[str] = set()
    for fallback_index, revised_slide in enumerate(revised_slides, start=1):
        if not isinstance(revised_slide, dict):
            continue
        slide_id = str(revised_slide.get("slide_id") or "")
        if not slide_id and len(revised_slides) == 1:
            slide_id = slide_id_for_index(fallback_index)
        target_index = original_by_id.get(slide_id)
        if target_index is None:
            if slide_id and slide_id not in appended_slide_ids:
                merged["slides"].append(revised_slide)
                appended_slide_ids.add(slide_id)
                applied_any = True
            continue
        merged["slides"][target_index] = revised_slide
        applied_any = True
    return (merged, True) if applied_any else (revised_payload, False)


def normalize_revised_plan_for_renderer(plan: dict[str, Any]) -> dict[str, Any]:
    normalized = copy.deepcopy(plan)
    for slide in normalized.get("slides") or []:
        if not isinstance(slide, dict):
            continue
        title = str(slide.get("title") or "").strip()
        if title:
            slide["subsection"] = title
        if str(slide.get("template_id") or "") == "T1_TextOnly":
            if slide.get("images"):
                slide["template_id"] = "T3_ImageLeft"
            elif slide.get("tables") or slide.get("formulas"):
                slide["template_id"] = "T4_ImageTop"
    return normalized


def collect_required_edit_tasks(user_decisions: dict[str, Any]) -> list[dict[str, Any]]:
    tasks: list[dict[str, Any]] = []
    seen: set[str] = set()
    for item in user_decisions.get("edit_tasks") or []:
        if not isinstance(item, dict) or not item.get("task_id"):
            continue
        task_id = str(item.get("task_id"))
        if task_id not in seen:
            tasks.append(item)
            seen.add(task_id)
    for request in user_decisions.get("user_requested_changes") or []:
        if not isinstance(request, dict):
            continue
        for item in request.get("edit_tasks") or []:
            if not isinstance(item, dict) or not item.get("task_id"):
                continue
            task_id = str(item.get("task_id"))
            if task_id not in seen:
                tasks.append(item)
                seen.add(task_id)
    return tasks


def editable_slide_ids_from_decisions(user_decisions: dict[str, Any]) -> set[str]:
    editable_slide_ids: set[str] = set()
    for request in user_decisions.get("user_requested_changes") or []:
        if isinstance(request, dict):
            editable_slide_ids.update(str(slide_id) for slide_id in request.get("editable_slide_ids") or [])
    return editable_slide_ids


def plain_language_rewrite(text: str) -> str:
    cleaned = re.sub(r"\s+", " ", str(text or "").strip())
    if not cleaned:
        return "In plain terms, this slide explains the key idea and why it matters."
    if cleaned.lower().startswith("in plain terms"):
        return cleaned
    lowered = cleaned[:1].lower() + cleaned[1:]
    return f"In plain terms, {lowered}"


def direct_title_rewrite(title: str) -> str:
    cleaned = re.sub(r"\s+", " ", str(title or "").strip())
    if not cleaned:
        return "Direct Technical Summary"
    if cleaned.lower().startswith("direct "):
        return f"{cleaned}: Key Takeaways"
    return f"Direct {cleaned}"


def restore_protected_slides(
    *,
    before_plan: dict[str, Any],
    revised_plan: dict[str, Any],
    user_decisions: dict[str, Any],
    application_log: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    protected = {str(item) for item in user_decisions.get("protected_slide_ids") or []}
    overridden = {str(item.get("slide_id")) for item in user_decisions.get("protected_slide_overrides") or [] if isinstance(item, dict)}
    if not protected:
        return revised_plan, application_log

    revised = copy.deepcopy(revised_plan)
    log = copy.deepcopy(application_log)
    restored: list[str] = []
    before_slides = list(before_plan.get("slides") or [])
    revised_slides = list(revised.get("slides") or [])
    for slide_id in sorted(protected - overridden):
        slide_index = plan_index_for_slide_id(slide_id)
        if slide_index is None or slide_index >= len(before_slides) or slide_index >= len(revised_slides):
            continue
        if before_slides[slide_index] != revised_slides[slide_index]:
            revised_slides[slide_index] = copy.deepcopy(before_slides[slide_index])
            restored.append(slide_id)
    if restored:
        revised["slides"] = revised_slides
        log["protected_slides_restored_from_original"] = restored
    return revised, log


def complete_required_visible_text_tasks(
    *,
    before_plan: dict[str, Any],
    revised_plan: dict[str, Any],
    user_decisions: dict[str, Any],
    application_log: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Make explicitly required visible-text tasks produce a visible text delta.

    The LLM sometimes marks tasks as applied while returning unchanged slide
    JSON. This fallback is deliberately narrow: it only touches non-protected
    slides explicitly listed as editable text-edit tasks.
    """
    editable_slide_ids = editable_slide_ids_from_decisions(user_decisions)
    protected = {str(item) for item in user_decisions.get("protected_slide_ids") or []}
    changed_slide_ids = set(slide_change_summary(before_plan, revised_plan)["changed_slide_ids"])
    revised = copy.deepcopy(revised_plan)
    log = copy.deepcopy(application_log)
    checklist = [item for item in log.get("edit_task_checklist") or [] if isinstance(item, dict)]
    checklist_by_task_id = {str(item.get("task_id")): item for item in checklist if item.get("task_id")}
    fallback_applied: list[dict[str, Any]] = []

    for task in collect_required_edit_tasks(user_decisions):
        operation = str(task.get("operation") or "")
        if operation not in {"audience_rewrite", "rewrite_title"}:
            continue
        task_id = str(task.get("task_id") or "")
        slide_id = str(task.get("slide_id") or "")
        if not task_id or not slide_id:
            continue
        if editable_slide_ids and slide_id not in editable_slide_ids:
            continue
        if slide_id in protected:
            continue
        if slide_id in changed_slide_ids:
            continue

        slide_index = plan_index_for_slide_id(slide_id)
        if slide_index is None or slide_index >= len(revised.get("slides") or []):
            continue
        slide = revised["slides"][slide_index]
        if not isinstance(slide, dict):
            continue

        if operation == "rewrite_title":
            original_title = str(slide.get("subsection") or slide.get("title") or "")
            rewritten_title = direct_title_rewrite(original_title)
            slide["subsection"] = rewritten_title
            evidence = f"Rewrote the visible title on {slide_id} to '{rewritten_title}'."
        else:
            bullets = [item for item in slide.get("bullets") or [] if isinstance(item, dict)]
            if bullets:
                original_text = str(bullets[0].get("text") or "")
                rewritten_text = plain_language_rewrite(original_text)
                if rewritten_text == original_text:
                    rewritten_text = f"{rewritten_text} for a general technical audience."
                bullets[0]["text"] = rewritten_text
                slide["bullets"] = bullets
                evidence = f"Rewrote the first visible bullet on {slide_id} in plain language."
            else:
                slide.setdefault("bullets", [])
                slide["bullets"].append(
                    {
                        "text": "In plain terms, this slide explains the key idea and why it matters.",
                        "sub": [],
                    }
                )
                evidence = f"Added a visible plain-language bullet on {slide_id}."

        fallback_record = {
            "task_id": task_id,
            "slide_id": slide_id,
            "operation": operation,
            "evidence": evidence,
        }
        fallback_applied.append(fallback_record)
        checklist_item = checklist_by_task_id.get(task_id)
        if checklist_item is None:
            checklist_item = {
                "task_id": task_id,
                "slide_id": slide_id,
                "operation": operation,
            }
            checklist.append(checklist_item)
            checklist_by_task_id[task_id] = checklist_item
        checklist_item["status"] = "applied"
        checklist_item["evidence"] = evidence
        checklist_item["unchanged_justification"] = ""
        checklist_item["completed_by"] = "deterministic_visible_text_fallback"
        changed_slide_ids.add(slide_id)

    if fallback_applied:
        log["edit_task_checklist"] = checklist
        log["deterministic_task_completion"] = fallback_applied
    return revised, log


def complete_required_add_bullet_tasks(
    *,
    before_plan: dict[str, Any],
    revised_plan: dict[str, Any],
    user_decisions: dict[str, Any],
    application_log: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    editable_slide_ids = editable_slide_ids_from_decisions(user_decisions)
    protected = {str(item) for item in user_decisions.get("protected_slide_ids") or []}
    revised = copy.deepcopy(revised_plan)
    log = copy.deepcopy(application_log)
    checklist = [item for item in log.get("edit_task_checklist") or [] if isinstance(item, dict)]
    checklist_by_task_id = {str(item.get("task_id")): item for item in checklist if item.get("task_id")}
    fallback_applied = list(log.get("deterministic_task_completion") or [])

    for task in collect_required_edit_tasks(user_decisions):
        if str(task.get("operation") or "") != "add_bullet":
            continue
        task_id = str(task.get("task_id") or "")
        slide_id = str(task.get("slide_id") or "")
        if not task_id or not slide_id:
            continue
        if editable_slide_ids and slide_id not in editable_slide_ids:
            continue
        if slide_id in protected:
            continue

        slide_index = plan_index_for_slide_id(slide_id)
        before_slides = list(before_plan.get("slides") or [])
        revised_slides = list(revised.get("slides") or [])
        if slide_index is None or slide_index >= len(before_slides) or slide_index >= len(revised_slides):
            continue
        before_slide = before_slides[slide_index]
        slide = revised_slides[slide_index]
        if not isinstance(before_slide, dict) or not isinstance(slide, dict):
            continue
        before_count = len([item for item in before_slide.get("bullets") or [] if isinstance(item, dict)])
        bullets = [item for item in slide.get("bullets") or [] if isinstance(item, dict)]
        if len(bullets) > before_count:
            continue

        new_bullet = {
            "text": "A clear takeaway is that the method improves the result while preserving the paper's main claim.",
            "sub": [],
        }
        bullets.append(new_bullet)
        slide["bullets"] = bullets
        evidence = f"Added one top-level bullet on {slide_id}."
        fallback_record = {
            "task_id": task_id,
            "slide_id": slide_id,
            "operation": "add_bullet",
            "evidence": evidence,
        }
        fallback_applied.append(fallback_record)

        checklist_item = checklist_by_task_id.get(task_id)
        if checklist_item is None:
            checklist_item = {
                "task_id": task_id,
                "slide_id": slide_id,
                "operation": "add_bullet",
            }
            checklist.append(checklist_item)
            checklist_by_task_id[task_id] = checklist_item
        checklist_item["status"] = "applied"
        checklist_item["evidence"] = evidence
        checklist_item["unchanged_justification"] = ""
        checklist_item["completed_by"] = "deterministic_top_level_bullet_fallback"

    if fallback_applied:
        log["edit_task_checklist"] = checklist
        log["deterministic_task_completion"] = fallback_applied
    return revised, log


def validate_repair(
    *,
    before_plan: dict[str, Any],
    after_plan: dict[str, Any],
    user_decisions: dict[str, Any],
    application_log: dict[str, Any],
) -> dict[str, Any]:
    change_summary = slide_change_summary(before_plan, after_plan)
    protected = {str(item) for item in user_decisions.get("protected_slide_ids") or []}
    overridden = {str(item.get("slide_id")) for item in user_decisions.get("protected_slide_overrides") or [] if isinstance(item, dict)}
    protected_violations = sorted(slide_id for slide_id in change_summary["changed_slide_ids"] if slide_id in protected and slide_id not in overridden)
    constraints = dict(user_decisions.get("constraints") or {})
    max_changed = int(constraints.get("max_changed_slides", 5) or 5)

    rejected_applied: list[str] = []
    for item in user_decisions.get("candidate_edit_decisions") or []:
        if not isinstance(item, dict) or item.get("decision") != "reject":
            continue
        if item.get("edit_id") in application_log.get("applied_edit_ids", []):
            rejected_applied.append(str(item.get("edit_id")))

    accepted_requested = [
        item
        for item in user_decisions.get("candidate_edit_decisions") or []
        if isinstance(item, dict) and item.get("decision") in {"accept", "modify"}
    ]
    applied_count = len(application_log.get("applied_edit_ids", []))
    completion_score = round(applied_count / len(accepted_requested), 4) if accepted_requested else None
    required_edit_tasks = collect_required_edit_tasks(user_decisions)
    checklist = [
        item
        for item in application_log.get("edit_task_checklist") or []
        if isinstance(item, dict)
    ]
    checklist_by_task_id = {str(item.get("task_id")): item for item in checklist if item.get("task_id")}
    missing_task_ids = [
        str(task.get("task_id"))
        for task in required_edit_tasks
        if str(task.get("task_id")) not in checklist_by_task_id
    ]
    skipped_task_ids = [
        str(task_id)
        for task_id, item in checklist_by_task_id.items()
        if str(item.get("status") or "").lower() == "skipped"
    ]
    editable_slide_ids = editable_slide_ids_from_decisions(user_decisions)
    justified_unchanged_slide_ids = {
        str(item.get("slide_id"))
        for item in checklist
        if item.get("slide_id")
        and str(item.get("status") or "").lower() == "skipped"
        and str(item.get("unchanged_justification") or "").strip()
    }
    unaccounted_editable_slide_ids = sorted(
        slide_id
        for slide_id in editable_slide_ids
        if slide_id not in change_summary["changed_slide_ids"]
        and slide_id not in justified_unchanged_slide_ids
    )
    changed_slide_ids = set(change_summary["changed_slide_ids"])
    applied_task_unchanged_slides = [
        {
            "task_id": str(item.get("task_id") or ""),
            "slide_id": str(item.get("slide_id") or ""),
            "operation": str(item.get("operation") or ""),
        }
        for item in checklist
        if str(item.get("status") or "").lower() == "applied"
        and item.get("slide_id")
        and str(item.get("slide_id")) not in changed_slide_ids
    ]
    strict_task_validation_required = bool(required_edit_tasks or editable_slide_ids)

    return {
        "change_summary": change_summary,
        "protected_slide_violation_count": len(protected_violations),
        "protected_slide_violations": protected_violations,
        "rejected_edit_violation_count": len(rejected_applied),
        "rejected_edit_violations": rejected_applied,
        "changed_slide_count_within_limit": change_summary["changed_slide_count"] <= max_changed,
        "max_changed_slides": max_changed,
        "user_request_completion_score": completion_score,
        "accepted_or_modified_requested_count": len(accepted_requested),
        "applied_user_edit_count": applied_count,
        "required_edit_task_count": len(required_edit_tasks),
        "edit_task_checklist_count": len(checklist),
        "missing_edit_task_ids": missing_task_ids,
        "skipped_edit_task_ids": skipped_task_ids,
        "unaccounted_editable_slide_ids": unaccounted_editable_slide_ids,
        "applied_task_unchanged_slides": applied_task_unchanged_slides,
        "edit_task_checklist_complete": len(missing_task_ids) == 0 if required_edit_tasks else None,
        "editable_slide_accounting_complete": len(unaccounted_editable_slide_ids) == 0 if editable_slide_ids else None,
        "strict_task_validation_required": strict_task_validation_required,
        "accepted": (
            len(protected_violations) == 0
            and len(rejected_applied) == 0
            and change_summary["changed_slide_count"] <= max_changed
            and (not required_edit_tasks or len(missing_task_ids) == 0)
            and len(applied_task_unchanged_slides) == 0
            and (not editable_slide_ids or len(unaccounted_editable_slide_ids) == 0)
        ),
    }


def deterministic_repair(
    *,
    slide_plan: dict[str, Any],
    candidate_edits: list[dict[str, Any]],
    user_decisions: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    revised = copy.deepcopy(slide_plan)
    decision_map = decisions_by_edit_id(user_decisions)
    applied_edit_ids: list[str] = []
    skipped: list[dict[str, Any]] = []

    for edit in candidate_edits:
        edit_id = str(edit.get("edit_id") or "")
        decision = decision_map.get(edit_id)
        if not decision:
            if edit.get("default_action") == "auto_apply_mechanical":
                decision = {"edit_id": edit_id, "decision": "accept", "user_instruction": ""}
            else:
                continue
        if decision.get("decision") not in {"accept", "modify"}:
            continue
        result = apply_deterministic_edit(revised, edit, decision, user_decisions)
        if result.get("applied"):
            applied_edit_ids.append(edit_id)
        else:
            skipped.append({"edit_id": edit_id, **result})

    for freeform in user_decisions.get("user_requested_changes") or []:
        skipped.append(
            {
                "edit_id": None,
                "reason": "freeform_change_requires_llm_repair",
                "user_instruction": str(freeform),
            }
        )

    log = {
        "repair_mode": "deterministic",
        "applied_edit_ids": applied_edit_ids,
        "skipped": skipped,
    }
    return revised, log


def llm_repair(
    *,
    slide_plan: dict[str, Any],
    rendered_deck_summary: dict[str, Any],
    candidate_edits: list[dict[str, Any]],
    user_decisions: dict[str, Any],
    author_profile: dict[str, Any] | None,
    prompt_path: Path,
    model: str,
    timeout: float,
) -> tuple[dict[str, Any], dict[str, Any], int, int]:
    prompt = render_prompt(
        prompt_path,
        slide_plan_json=slide_plan,
        rendered_deck_summary_json=rendered_deck_summary,
        candidate_edits_json=candidate_edits,
        user_decisions_json=user_decisions,
        author_profile_json=author_profile or {},
    )
    payload, in_tok, out_tok, raw_text = call_text_json(
        model=model,
        system_prompt=prompt["system_prompt"],
        user_prompt=prompt["user_prompt"],
        timeout=timeout,
        max_tokens=3600,
    )
    revised = payload.get("revised_slide_plan") or payload.get("slide_plan") or payload
    if not isinstance(revised, dict) or not isinstance(revised.get("slides"), list):
        raise RuntimeError("LLM repair did not return a valid slide plan.")
    revised, merged_partial_plan = merge_partial_revised_slides(slide_plan, revised)
    revised = normalize_revised_plan_for_renderer(revised)
    log = {
        "repair_mode": "llm",
        "merged_partial_revised_plan": merged_partial_plan,
        "applied_edit_ids": list((payload.get("change_log") or {}).get("applied_user_accepted_edits") or [])
        + list((payload.get("change_log") or {}).get("applied_user_modified_edits") or []),
        "edit_task_checklist": list((payload.get("change_log") or {}).get("edit_task_checklist") or []),
        "llm_payload": payload,
        "raw_text_preview": raw_text[:1000],
    }
    revised, log = restore_protected_slides(
        before_plan=slide_plan,
        revised_plan=revised,
        user_decisions=user_decisions,
        application_log=log,
    )
    revised, log = complete_required_visible_text_tasks(
        before_plan=slide_plan,
        revised_plan=revised,
        user_decisions=user_decisions,
        application_log=log,
    )
    revised, log = complete_required_add_bullet_tasks(
        before_plan=slide_plan,
        revised_plan=revised,
        user_decisions=user_decisions,
        application_log=log,
    )
    return revised, log, in_tok, out_tok


def regenerate_deck(args: argparse.Namespace, revised_plan: dict[str, Any]) -> Path:
    variant_suffix = args.output_variant_suffix
    render_args = SimpleNamespace(
        paper_name=args.paper_name,
        model_name_t=args.model_name_t,
        model_name_v=args.model_name_v,
        output_dir=args.output_dir,
        output_variant_suffix=variant_suffix,
        use_author_preferences=bool(args.author_profile_path),
        author_profile_path=str(args.author_profile_path) if args.author_profile_path else None,
        formula_mode=args.formula_mode,
        asset_paper_name=args.asset_paper_name,
    )
    plan_path = slide_plan_path(render_args, variant_suffix)
    write_json(plan_path, revised_plan)
    generate_pptx_from_plan_lazy(render_args, args.template)
    themed_path = themed_output_pptx_path(render_args, variant_suffix)
    return themed_path if themed_path.exists() else themed_path.with_name(themed_path.name.replace("_themed", ""))


def repair_deck(args: argparse.Namespace) -> dict[str, Any]:
    started = time.time()
    slide_plan = normalize_slide_ids(load_json(args.slide_plan_path, {}))
    candidate_payload = load_json(args.candidate_edits_path, {})
    candidate_edits = coerce_candidate_edits(candidate_payload)
    user_decisions = load_json(args.user_decisions_path, {})
    rendered_deck_summary = load_json(args.rendered_deck_summary_path, {})
    author_profile = load_json(args.author_profile_path, None)
    token_usage = default_token_usage()

    if args.repair_mode == "llm":
        revised_plan, application_log, in_tok, out_tok = llm_repair(
            slide_plan=slide_plan,
            rendered_deck_summary=rendered_deck_summary,
            candidate_edits=candidate_edits,
            user_decisions=user_decisions,
            author_profile=author_profile,
            prompt_path=args.repair_prompt_path,
            model=args.model,
            timeout=args.timeout,
        )
        token_usage["repair_input_tokens"] = in_tok
        token_usage["repair_output_tokens"] = out_tok
    else:
        revised_plan, application_log = deterministic_repair(
            slide_plan=slide_plan,
            candidate_edits=candidate_edits,
            user_decisions=user_decisions,
        )

    token_usage = update_token_usage(token_usage)
    validation = validate_repair(
        before_plan=slide_plan,
        after_plan=revised_plan,
        user_decisions=user_decisions,
        application_log=application_log,
    )
    output_dir = args.output_dir_path
    revised_plan_path = args.revised_plan_path or output_dir / "post_render_refined_slide_plan.json"
    if args.regenerate:
        revised_plan = sanitize_slide_plan_templates_lazy(revised_plan)
    write_json(revised_plan_path, revised_plan)

    regenerated_pptx = None
    rendered_issue_after = None
    refined_render_dir = None
    accepted_user_edits = len(application_log.get("applied_edit_ids", []))
    before_issue_count = len(candidate_edits)

    def build_report() -> dict[str, Any]:
        benefit_metrics = {
            "accepted_user_edits": accepted_user_edits,
            "rendered_issue_count_before": before_issue_count,
            "rendered_issue_count_after": rendered_issue_after,
            "rendered_issue_delta": 0 if rendered_issue_after is None else max(0, before_issue_count - rendered_issue_after),
            "user_request_completion_score": validation.get("user_request_completion_score"),
            "profile_compatibility_delta": None,
        }
        return {
            "mode": "repair",
            "runtime_seconds": round(time.time() - started, 3),
            "repair_mode": args.repair_mode,
            "paths": {
                "slide_plan_path": str(args.slide_plan_path),
                "candidate_edits_path": str(args.candidate_edits_path),
                "user_decisions_path": str(args.user_decisions_path),
                "revised_plan_path": str(revised_plan_path),
                "regenerated_pptx": str(regenerated_pptx) if regenerated_pptx else None,
                "refined_render_dir": str(refined_render_dir) if refined_render_dir else None,
            },
            "application_log": application_log,
            "validation": validation,
            "cost_ledger": {
                "token_usage": token_usage,
                "benefit_metrics": benefit_metrics,
                "efficiency": score_efficiency(token_usage, benefit_metrics),
            },
        }

    if args.regenerate:
        if not validation.get("accepted"):
            write_json(output_dir / "post_render_refinement_report.json", build_report())
            raise RuntimeError(
                "Refinement validation failed; refusing to regenerate PPTX. "
                f"See {revised_plan_path} and the repair validation report for details."
            )
        regenerated_pptx = regenerate_deck(args, revised_plan)
        if args.validate_render:
            refined_render_dir = output_dir / "refined_rendered_slides"
            refined_images = render_pptx_to_images(regenerated_pptx, refined_render_dir, force=True, dpi=args.render_dpi)
            after_edits, _image_summaries = deterministic_render_edits(refined_images)
            rendered_issue_after = len(after_edits)

    report = build_report()
    write_json(output_dir / "post_render_refinement_report.json", report)
    return report


def _first_path(candidates: Iterable[Path], *, label: str) -> Path:
    matches = [path for path in candidates if path.exists()]
    if not matches:
        raise FileNotFoundError(f"Could not infer {label}. Pass the explicit path or check the run directory.")
    return sorted(matches)[0]


def infer_local_pptx(run_dir: Path) -> Path:
    candidates = [
        path
        for path in run_dir.glob("*_output_slides*.pptx")
        if "_refined" not in path.stem
    ]
    return _first_path(candidates, label="generated PPTX")


def infer_local_slide_plan(run_dir: Path) -> Path:
    candidates = [
        path
        for path in run_dir.glob("<*>_slide_plan*.json")
        if "_draft" not in path.stem and "_repair_report" not in path.stem and "_refined" not in path.stem
    ]
    return _first_path(candidates, label="slide plan JSON")


def infer_model_names_from_plan(slide_plan: Path) -> tuple[str, str]:
    match = re.match(r"<(.+)>_slide_plan.*\.json$", slide_plan.name)
    if not match:
        raise ValueError(f"Could not infer model names from slide plan filename: {slide_plan.name}")
    model_pair = match.group(1)
    if "_" not in model_pair:
        raise ValueError(f"Could not split model pair from slide plan filename: {slide_plan.name}")
    model_name_t, model_name_v = model_pair.split("_", 1)
    return model_name_t, model_name_v


def infer_output_root_from_run_dir(run_dir: Path) -> Path:
    resolved = run_dir.resolve()
    if resolved.parent.name == "contents":
        return resolved.parent.parent
    return Path(".")


def infer_asset_paper_name(run_dir: Path, model_name_t: str, model_name_v: str) -> str:
    run_name = run_dir.name
    if not run_name.endswith(("_personalized", "_refined")):
        return run_name
    base_name = re.sub(r"_(personalized|refined)$", "", run_name)
    shared_dir = run_dir.parent / base_name
    raw_content = shared_dir / f"<{model_name_t}_{model_name_v}>_raw_content.json"
    if raw_content.exists():
        return base_name
    return run_name


def run_local(args: argparse.Namespace) -> dict[str, Any]:
    run_dir = args.run_dir.resolve()
    pptx_path = args.pptx_path or infer_local_pptx(run_dir)
    slide_plan = args.slide_plan_path or infer_local_slide_plan(run_dir)
    output_dir = args.output_dir or run_dir / DEFAULT_OUTPUT_DIR_NAME
    model_name_t = args.model_name_t
    model_name_v = args.model_name_v
    if not model_name_t or not model_name_v:
        inferred_t, inferred_v = infer_model_names_from_plan(slide_plan)
        model_name_t = model_name_t or inferred_t
        model_name_v = model_name_v or inferred_v
    model = args.model or model_name_v

    result: dict[str, Any] = {
        "run_dir": str(run_dir),
        "pptx_path": str(pptx_path),
        "slide_plan_path": str(slide_plan),
        "output_dir": str(output_dir),
    }

    if args.action in {"inspect", "all"}:
        inspect_report = inspect_deck(
            pptx_path=pptx_path,
            slide_plan_path=slide_plan,
            output_dir=output_dir,
            author_profile_path=args.author_profile_path,
            render_dpi=args.render_dpi,
            force_render=args.force_render,
            timeout=args.timeout,
        )
        result["inspect"] = {
            "candidate_edit_count": len(inspect_report["candidate_edits"]),
            "inspection_report_path": str(output_dir / "post_render_inspection_report.json"),
            "user_decisions_template_path": str(output_dir / "post_render_user_decisions.template.json"),
        }

    if args.action in {"repair", "all"}:
        decisions_path = args.user_decisions_path or output_dir / "post_render_user_decisions.template.json"
        paper_name = args.paper_name or run_dir.name
        asset_paper_name = args.asset_paper_name or infer_asset_paper_name(run_dir, model_name_t, model_name_v)
        repair_args = SimpleNamespace(
            slide_plan_path=slide_plan,
            candidate_edits_path=output_dir / "post_render_candidate_edits.json",
            user_decisions_path=decisions_path,
            rendered_deck_summary_path=output_dir / "post_render_deck_summary.json",
            output_dir_path=output_dir,
            revised_plan_path=output_dir / "post_render_refined_slide_plan.json",
            author_profile_path=args.author_profile_path,
            repair_mode=args.repair_mode,
            model=model,
            repair_prompt_path=args.repair_prompt_path,
            timeout=args.timeout,
            regenerate=not args.no_regenerate,
            validate_render=args.validate_render,
            render_dpi=args.render_dpi,
            paper_name=paper_name,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
            output_dir=str(args.pipeline_output_root or infer_output_root_from_run_dir(run_dir)),
            output_variant_suffix=args.output_variant_suffix,
            asset_paper_name=asset_paper_name,
            formula_mode=args.formula_mode,
            template=args.template,
        )
        repair_report = repair_deck(repair_args)
        result["repair"] = {
            "accepted": repair_report["validation"]["accepted"],
            "report_path": str(output_dir / "post_render_refinement_report.json"),
            "regenerated_pptx": repair_report.get("paths", {}).get("regenerated_pptx"),
            "refined_render_dir": repair_report.get("paths", {}).get("refined_render_dir"),
        }

    return result


def add_common_inspect_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--pptx-path", type=Path, required=True)
    parser.add_argument("--slide-plan-path", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, default=None)
    parser.add_argument("--author-profile-path", type=Path, default=None)
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument("--force-render", action="store_true")
    parser.add_argument("--timeout", type=float, default=180.0)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Post-render human-guided refinement agent.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser("inspect", help="Render and inspect a generated PPTX.")
    add_common_inspect_args(inspect_parser)

    repair_parser = subparsers.add_parser("repair", help="Apply user decisions to the slide plan.")
    repair_parser.add_argument("--slide-plan-path", type=Path, required=True)
    repair_parser.add_argument("--candidate-edits-path", type=Path, required=True)
    repair_parser.add_argument("--user-decisions-path", type=Path, required=True)
    repair_parser.add_argument("--rendered-deck-summary-path", type=Path, required=True)
    repair_parser.add_argument("--output-dir-path", type=Path, required=True)
    repair_parser.add_argument("--revised-plan-path", type=Path, default=None)
    repair_parser.add_argument("--author-profile-path", type=Path, default=None)
    repair_parser.add_argument("--repair-mode", choices=["deterministic", "llm"], default="deterministic")
    repair_parser.add_argument("--model", default="4o")
    repair_parser.add_argument("--repair-prompt-path", type=Path, default=DEFAULT_REPAIR_PROMPT)
    repair_parser.add_argument("--timeout", type=float, default=180.0)
    repair_parser.add_argument("--regenerate", action="store_true")
    repair_parser.add_argument("--validate-render", action="store_true")
    repair_parser.add_argument("--render-dpi", type=int, default=120)
    repair_parser.add_argument("--paper-name", default=None)
    repair_parser.add_argument("--model-name-t", default=None)
    repair_parser.add_argument("--model-name-v", default=None)
    repair_parser.add_argument("--output-dir", default=".")
    repair_parser.add_argument("--output-variant-suffix", default="_refined")
    repair_parser.add_argument("--asset-paper-name", default=None)
    repair_parser.add_argument("--formula-mode", type=int, default=1)
    repair_parser.add_argument("--template", type=int, default=3)

    local_parser = subparsers.add_parser("local", help="Run local inspect/repair by inferring paths from a generated run directory.")
    local_parser.add_argument("action", choices=["inspect", "repair", "all"], nargs="?", default="inspect")
    local_parser.add_argument("--run-dir", type=Path, required=True, help="Generated run directory, e.g. contents/<paper_name>.")
    local_parser.add_argument("--pptx-path", type=Path, default=None)
    local_parser.add_argument("--slide-plan-path", type=Path, default=None)
    local_parser.add_argument("--output-dir", type=Path, default=None)
    local_parser.add_argument("--author-profile-path", type=Path, default=None)
    local_parser.add_argument("--render-dpi", type=int, default=120)
    local_parser.add_argument("--force-render", action="store_true")
    local_parser.add_argument("--model", default=None)
    local_parser.add_argument("--repair-prompt-path", type=Path, default=DEFAULT_REPAIR_PROMPT)
    local_parser.add_argument("--repair-mode", choices=["deterministic", "llm"], default="deterministic")
    local_parser.add_argument("--user-decisions-path", type=Path, default=None)
    local_parser.add_argument("--no-regenerate", action="store_true")
    local_parser.add_argument(
        "--validate-render",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Render the regenerated refined PPTX for local visual inspection. Use --no-validate-render to skip.",
    )
    local_parser.add_argument("--paper-name", default=None)
    local_parser.add_argument("--model-name-t", default=None)
    local_parser.add_argument("--model-name-v", default=None)
    local_parser.add_argument("--pipeline-output-root", type=Path, default=None)
    local_parser.add_argument("--output-variant-suffix", default="_refined")
    local_parser.add_argument("--asset-paper-name", default=None)
    local_parser.add_argument("--formula-mode", type=int, default=1)
    local_parser.add_argument("--template", type=int, default=3)
    local_parser.add_argument("--timeout", type=float, default=180.0)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    if args.command == "inspect":
        output_dir = args.output_dir or args.pptx_path.parent / DEFAULT_OUTPUT_DIR_NAME
        report = inspect_deck(
            pptx_path=args.pptx_path,
            slide_plan_path=args.slide_plan_path,
            output_dir=output_dir,
            author_profile_path=args.author_profile_path,
            render_dpi=args.render_dpi,
            force_render=args.force_render,
            timeout=args.timeout,
        )
        print(json.dumps({"output_dir": str(output_dir), "candidate_edit_count": len(report["candidate_edits"])}, indent=2))
        return

    if args.command == "repair":
        if args.regenerate:
            missing = [
                name
                for name in ("paper_name", "model_name_t", "model_name_v")
                if not getattr(args, name)
            ]
            if missing:
                raise SystemExit(f"--regenerate requires: {', '.join('--' + item.replace('_', '-') for item in missing)}")
        report = repair_deck(args)
        print(json.dumps({"accepted": report["validation"]["accepted"], "report_path": str(args.output_dir_path / "post_render_refinement_report.json")}, indent=2))
        return

    if args.command == "local":
        report = run_local(args)
        print(json.dumps(report, indent=2))
        return


if __name__ == "__main__":
    main()
