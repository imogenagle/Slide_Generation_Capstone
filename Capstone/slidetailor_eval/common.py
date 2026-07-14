"""Shared helpers for SlideTailor-derived evaluation scripts."""

from __future__ import annotations

import argparse
import base64
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Iterable

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:
    def load_dotenv(*_args: Any, **_kwargs: Any) -> bool:
        return False


REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from .paths import metric_output_dir


DEFAULT_PAPERS_CSV = REPO_ROOT / "Capstone" / "author_tables" / "papers.csv"
DEFAULT_MAX_REFERENCE_STRUCTURE_IMAGES = 50
DEFAULT_CATEGORY_LIST = [
    "Title and Authors",
    "Definitions and Background",
    "Motivation and Challenges",
    "Related Work",
    "Dataset Construction",
    "Problem Formulation",
    "Method Overview",
    "Experiment Setup and Results",
    "Future Directions",
    "Summary and Conclusions",
]
SLIDE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
DEFAULT_MAX_VISION_IMAGES = 24
DEFAULT_REFERENCE_OUTLINE_CHUNK_SIZE = 12
DEFAULT_STANDARDIZE_CHUNK_SIZE = 16


def add_shared_args(parser: argparse.ArgumentParser, *, metric_name: str) -> None:
    parser.add_argument("--model", default="gpt-5", help="Judge model identifier.")
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help=f"Optional output JSON path. Defaults to Capstone/evaluations/slidetailor/{metric_name}/",
    )
    parser.add_argument("--request-timeout", type=float, default=180.0)
    parser.add_argument("--verbose", action="store_true")


def log(message: str, *, verbose: bool) -> None:
    if verbose:
        print(f"[slidetailor_eval] {message}", file=sys.stderr, flush=True)


def load_runtime_env() -> None:
    load_dotenv(REPO_ROOT / ".env")


def render_prompt(prompt_path: Path, **values: Any) -> dict[str, str]:
    try:
        import yaml
    except ModuleNotFoundError as exc:
        raise RuntimeError("PyYAML is required to load SlideTailor evaluation prompts.") from exc
    try:
        from jinja2 import Environment, StrictUndefined
    except ModuleNotFoundError as exc:
        raise RuntimeError("Jinja2 is required to render SlideTailor evaluation prompts.") from exc
    prompt_cfg = yaml.safe_load(prompt_path.read_text(encoding="utf-8"))
    env = Environment(undefined=StrictUndefined)
    env.filters["tojson"] = lambda value: json.dumps(value, ensure_ascii=False, indent=2)
    template = env.from_string(prompt_cfg["template"])
    user_prompt = template.render(**values)
    return {
        "system_prompt": prompt_cfg["system_prompt"],
        "user_prompt": user_prompt,
    }


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
    json_text = text[start : end + 1]
    try:
        return json.loads(json_text)
    except json.JSONDecodeError as exc:
        # Some model replies include stray backslashes inside the free-form reason
        # field, which makes otherwise-correct JSON fail with "Invalid \\escape".
        # Repair only unsupported backslash escapes and retry.
        if "Invalid \\escape" not in str(exc):
            raise
        repaired = re.sub(r'\\(?!["\\/bfnrtu])', r"\\\\", json_text)
        return json.loads(repaired)


def extract_message_text(message: Any) -> str:
    content = getattr(message, "content", None)
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        text_parts: list[str] = []
        for part in content:
            if isinstance(part, dict):
                text_value = part.get("text")
                if isinstance(text_value, str):
                    text_parts.append(text_value)
            else:
                text_value = getattr(part, "text", None)
                if isinstance(text_value, str):
                    text_parts.append(text_value)
        return "\n".join(part for part in text_parts if part).strip()
    refusal = getattr(message, "refusal", None)
    if isinstance(refusal, str) and refusal.strip():
        return refusal.strip()
    return ""


def resolve_output_path(metric_name: str, output: Path | None, stem: str) -> Path:
    if output is not None:
        return output
    out_dir = metric_output_dir(metric_name)
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"{stem}.{metric_name}.json"


def _normalize_score(value: Any, scale: str = "zero_to_one") -> float:
    try:
        numeric = float(value)
    except Exception:
        return 0.0
    if scale == "one_to_five":
        numeric = max(1.0, min(5.0, numeric))
        return round((numeric - 1.0) / 4.0, 4)
    if scale == "zero_to_one":
        return max(0.0, min(1.0, numeric))
    raise ValueError(f"Unsupported score scale: {scale}")


def encode_image_data_uri(image_path: Path) -> str:
    mime_type = "image/jpeg"
    if image_path.suffix.lower() == ".png":
        mime_type = "image/png"
    elif image_path.suffix.lower() == ".webp":
        mime_type = "image/webp"
    data = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{mime_type};base64,{data}"


def numeric_slide_sort_key(path: Path) -> tuple[int, str]:
    match = re.search(r"(\d+)", path.stem)
    if match:
        return (int(match.group(1)), path.name)
    return (10**9, path.name)


def sample_paths(paths: list[Path], max_images: int) -> list[Path]:
    if max_images <= 0 or len(paths) <= max_images:
        return paths
    if max_images <= 4:
        indices = [0, len(paths) // 3, (2 * len(paths)) // 3, len(paths) - 1]
        chosen = sorted(set(max(0, min(i, len(paths) - 1)) for i in indices))
        return [paths[i] for i in chosen]
    selected = {0, len(paths) - 1}
    remaining = max_images - len(selected)
    if remaining > 0:
        span = len(paths) - 1
        for idx in range(remaining):
            pos = round(((idx + 1) * span) / (remaining + 1))
            selected.add(max(0, min(pos, len(paths) - 1)))
    return [paths[i] for i in sorted(selected)[:max_images]]


def render_pptx_to_images(pptx_path: Path, output_dir: Path, *, force: bool = False, dpi: int = 120) -> list[Path]:
    try:
        from pdf2image import convert_from_path
    except ModuleNotFoundError as exc:
        raise RuntimeError("pdf2image is required to render slide images.") from exc
    pptx_path = pptx_path.resolve()
    if not pptx_path.exists():
        raise FileNotFoundError(f"PPTX not found: {pptx_path}")
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    existing = sorted(output_dir.glob("slide_*.jpg"), key=numeric_slide_sort_key)
    if existing and not force:
        return existing

    for path in output_dir.glob("slide_*.jpg"):
        path.unlink()

    soffice_path = shutil.which("soffice")
    if not soffice_path:
        raise RuntimeError("LibreOffice 'soffice' is required to render PPTX files, but it was not found on PATH.")

    with tempfile.TemporaryDirectory() as temp_dir, tempfile.TemporaryDirectory() as user_install_dir:
        command_list = [
            soffice_path,
            "--headless",
            "--norestore",
            "--nolockcheck",
            f"-env:UserInstallation=file://{user_install_dir}",
            "--convert-to",
            "pdf",
            str(pptx_path),
            "--outdir",
            temp_dir,
        ]
        env = os.environ.copy()
        env["LC_ALL"] = "en_US.UTF-8"
        env["LANG"] = "en_US.UTF-8"
        completed = subprocess.run(
            command_list,
            check=False,
            capture_output=True,
            text=True,
            env=env,
        )
        if completed.returncode != 0:
            stdout = (completed.stdout or "").strip()
            stderr = (completed.stderr or "").strip()
            details = "\n".join(part for part in [stdout, stderr] if part)
            if not details:
                details = "no stdout/stderr captured"
            raise RuntimeError(
                f"LibreOffice failed while rendering PPTX to PDF (exit {completed.returncode}): {pptx_path}\n{details}"
            )

        pdf_candidates = [path for path in Path(temp_dir).iterdir() if path.suffix.lower() == ".pdf"]
        if not pdf_candidates:
            raise RuntimeError(
                f"LibreOffice completed without producing a PDF for: {pptx_path}\n"
                f"Temporary output directory: {temp_dir}"
            )
        images = convert_from_path(str(pdf_candidates[0]), dpi=dpi)
        for index, image in enumerate(images, start=1):
            image.save(output_dir / f"slide_{index:04d}.jpg", "JPEG")

    return sorted(output_dir.glob("slide_*.jpg"), key=numeric_slide_sort_key)


def collect_slide_images(slide_dir: Path) -> list[Path]:
    return sorted(
        [path for path in slide_dir.iterdir() if path.suffix.lower() in SLIDE_EXTENSIONS],
        key=numeric_slide_sort_key,
    )


def call_json_judge(
    *,
    model: str,
    system_prompt: str,
    user_prompt: str,
    image_paths: Iterable[Path] | None = None,
    image_labels: list[str] | None = None,
    request_timeout: float = 180.0,
    verbose: bool = False,
) -> dict[str, Any]:
    from slidegen_openai_utils import build_openai_client, resolve_direct_model_name

    client = build_openai_client()
    resolved_model = resolve_direct_model_name(model)
    content: list[dict[str, Any]] = [{"type": "text", "text": user_prompt}]
    image_path_list = list(image_paths or [])
    label_list = image_labels or [f"Image: {path.name}" for path in image_path_list]
    for image_path, image_label in zip(image_path_list, label_list):
        content.append({"type": "text", "text": image_label})
        content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": encode_image_data_uri(image_path),
                    "detail": "low",
                },
            }
        )
    request_kwargs: dict[str, Any] = {
        "model": resolved_model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
        "temperature": 0.1,
        "timeout": request_timeout,
    }
    if "gpt-5" in resolved_model.lower():
        request_kwargs["max_completion_tokens"] = 2200
    else:
        request_kwargs["max_tokens"] = 2200
    log(f"Calling judge model={resolved_model!r}", verbose=verbose)
    response = client.chat.completions.create(**request_kwargs)
    raw_text = extract_message_text(response.choices[0].message)
    if not raw_text:
        finish_reason = getattr(response.choices[0], "finish_reason", None)
        raise ValueError(
            f"Model returned an empty response. finish_reason={finish_reason!r}, model={resolved_model!r}"
        )
    return extract_json_object(raw_text)


def summarize_scores(per_slide: list[dict[str, Any]]) -> float:
    scores = [_normalize_score(item.get("score")) for item in per_slide]
    return round(sum(scores) / len(scores), 4) if scores else 0.0


def extract_presentation_text(pptx_path: Path) -> str:
    from utils.src.presentation import Presentation
    from utils.src.utils import Config

    presentation = Presentation.from_file(str(pptx_path), Config("/tmp", debug=False))
    return presentation.to_text(show_image=False)


def extract_outline_from_presentation_text(
    *,
    pptx_path: Path,
    categories: list[str],
    prompt_path: Path,
    model: str,
    request_timeout: float,
    verbose: bool,
) -> dict[str, Any]:
    prompt = render_prompt(
        prompt_path,
        categories_json=json.dumps(categories, ensure_ascii=False, indent=2),
        presentation_text=extract_presentation_text(pptx_path),
    )
    return call_json_judge(
        model=model,
        system_prompt=prompt["system_prompt"],
        user_prompt=prompt["user_prompt"],
        request_timeout=request_timeout,
        verbose=verbose,
    )


def extract_outline_from_reference_images(
    *,
    image_paths: list[Path],
    categories: list[str],
    prompt_path: Path,
    model: str,
    request_timeout: float,
    verbose: bool,
    max_images: int = DEFAULT_MAX_VISION_IMAGES,
    chunk_size: int = DEFAULT_REFERENCE_OUTLINE_CHUNK_SIZE,
) -> dict[str, Any]:
    sampled = sample_paths(image_paths, max_images)
    if not sampled:
        return {
            "slide_descriptions": [],
            "deck_summary": "",
            "sampled_reference_slides": [],
        }

    prompt = render_prompt(
        prompt_path,
        categories_json=json.dumps(categories, ensure_ascii=False, indent=2),
    )
    chunk_size = max(1, chunk_size)
    combined_descriptions: list[dict[str, Any]] = []
    deck_summaries: list[str] = []
    global_slide_index = 1

    for start in range(0, len(sampled), chunk_size):
        chunk = sampled[start : start + chunk_size]
        outline = call_json_judge(
            model=model,
            system_prompt=prompt["system_prompt"],
            user_prompt=prompt["user_prompt"],
            image_paths=chunk,
            request_timeout=request_timeout,
            verbose=verbose,
        )
        chunk_descriptions = list(outline.get("slide_descriptions") or [])
        for local_index, image_path in enumerate(chunk, start=1):
            source_item = chunk_descriptions[local_index - 1] if local_index - 1 < len(chunk_descriptions) else {}
            combined_descriptions.append(
                {
                    "slide_index": global_slide_index,
                    "title": str(source_item.get("title") or "").strip(),
                    "description": str(source_item.get("description") or "").strip(),
                    "category_guess": str(source_item.get("category_guess") or "").strip(),
                    "source_image": image_path.name,
                }
            )
            global_slide_index += 1
        deck_summary = str(outline.get("deck_summary") or "").strip()
        if deck_summary:
            deck_summaries.append(deck_summary)

    return {
        "slide_descriptions": combined_descriptions,
        "deck_summary": " ".join(deck_summaries).strip(),
        "sampled_reference_slides": [path.name for path in sampled],
        "chunk_size": chunk_size,
    }


def standardize_narrative_items(
    *,
    narrative_items: list[str],
    categories: list[str],
    prompt_path: Path,
    model: str,
    request_timeout: float,
    verbose: bool,
    chunk_size: int = DEFAULT_STANDARDIZE_CHUNK_SIZE,
) -> list[dict[str, Any]]:
    if not narrative_items:
        return []
    chunk_size = max(1, chunk_size)
    standardized: list[dict[str, Any]] = []
    for start in range(0, len(narrative_items), chunk_size):
        chunk = narrative_items[start : start + chunk_size]
        prompt = render_prompt(
            prompt_path,
            categories_json=json.dumps(categories, ensure_ascii=False, indent=2),
            narrative_items_json=json.dumps(chunk, ensure_ascii=False, indent=2),
        )
        result = call_json_judge(
            model=model,
            system_prompt=prompt["system_prompt"],
            user_prompt=prompt["user_prompt"],
            request_timeout=request_timeout,
            verbose=verbose,
        )
        items = list(result.get("items") or [])
        if len(items) < len(chunk):
            items.extend(
                {
                    "raw": raw_item,
                    "standard": "",
                    "confidence": "low",
                }
                for raw_item in chunk[len(items):]
            )
        standardized.extend(items[: len(chunk)])
    return standardized


def condensed_standard_flow(items: list[dict[str, Any]]) -> list[str]:
    flow: list[str] = []
    prev = None
    for item in items:
        standard = str(item.get("standard") or "").strip()
        if not standard:
            continue
        if standard != prev:
            flow.append(standard)
        prev = standard
    return flow


def coverage_iou(target_flow: list[str], reference_flow: list[str]) -> float:
    target = set(target_flow)
    reference = set(reference_flow)
    if not target and not reference:
        return 1.0
    union = target | reference
    return round(len(target & reference) / len(union), 4) if union else 0.0


def _levenshtein_distance(s1: list[str], s2: list[str]) -> int:
    if len(s1) > len(s2):
        s1, s2 = s2, s1
    distances = list(range(len(s1) + 1))
    for i2, c2 in enumerate(s2):
        updated = [i2 + 1]
        for i1, c1 in enumerate(s1):
            if c1 == c2:
                updated.append(distances[i1])
            else:
                updated.append(1 + min(distances[i1], distances[i1 + 1], updated[-1]))
        distances = updated
    return distances[-1]


def flow_ngld(target_flow: list[str], reference_flow: list[str]) -> float:
    if not target_flow and not reference_flow:
        return 1.0
    gld = _levenshtein_distance(target_flow, reference_flow)
    denominator = len(target_flow) + len(reference_flow) + gld
    if denominator == 0:
        return 0.0
    normalized = 2 * gld / denominator
    return round(1 - normalized, 4)


def evaluate_single_slide_images(
    *,
    metric_name: str,
    slide_images: list[Path],
    prompt_path: Path,
    model: str,
    request_timeout: float,
    verbose: bool,
    score_scale: str = "one_to_five",
    prompt_values: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    prompt_cfg = render_prompt(prompt_path, **(prompt_values or {}))
    results: list[dict[str, Any]] = []
    for index, image_path in enumerate(slide_images, start=1):
        log(f"Evaluating {metric_name} on {image_path.name}", verbose=verbose)
        response = call_json_judge(
            model=model,
            system_prompt=prompt_cfg["system_prompt"],
            user_prompt=prompt_cfg["user_prompt"],
            image_paths=[image_path],
            request_timeout=request_timeout,
            verbose=verbose,
        )
        results.append(
            {
                "slide_index": index,
                "slide_image": image_path.name,
                "score": _normalize_score(response.get("score"), scale=score_scale),
                "reason": str(response.get("reason", "")).strip(),
            }
        )
    return results


def extract_source_document_text(source_path: Path, *, max_chars: int = 40000) -> str:
    if not source_path.exists():
        raise FileNotFoundError(f"Source document not found: {source_path}")
    suffix = source_path.suffix.lower()
    if suffix in {".txt", ".md"}:
        text = source_path.read_text(encoding="utf-8")
        return text[:max_chars]
    if suffix == ".json":
        data = json.loads(source_path.read_text(encoding="utf-8"))
        text = json.dumps(data, ensure_ascii=False, indent=2)
        return text[:max_chars]
    if suffix == ".pdf":
        try:
            from docling.document_converter import DocumentConverter
            converter = DocumentConverter()
            result = converter.convert(source_path)
            text = result.document.export_to_markdown()
            if text and text.strip():
                return text[:max_chars]
        except Exception as exc:
            raise RuntimeError(
                f"Failed to extract text from PDF via Docling: {source_path}"
            ) from exc
    raise ValueError(f"Unsupported source document type for text extraction: {source_path}")


def load_paper_metadata(papers_csv: Path = DEFAULT_PAPERS_CSV) -> dict[str, dict[str, Any]]:
    metadata_by_paper_name: dict[str, dict[str, Any]] = {}
    with papers_csv.open(newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            paper_id = (row.get("paper_id") or "").strip()
            title = (row.get("paper_title") or row.get("ppt_title") or "").strip()
            pdf_path = (row.get("paper_pdf_path") or "").strip()
            raw_dir = (row.get("raw_dir") or "").strip()
            if not paper_id:
                continue
            normalized_id = normalize_paper_name_from_paper_id(paper_id)
            metadata = {
                "paper_id": paper_id,
                "title": title,
                "paper_pdf_path": (REPO_ROOT.parent / pdf_path).resolve() if pdf_path and not Path(pdf_path).is_absolute() else Path(pdf_path) if pdf_path else None,
                "reference_slide_dir": (REPO_ROOT.parent / raw_dir).resolve() if raw_dir and not Path(raw_dir).is_absolute() else Path(raw_dir) if raw_dir else None,
            }
            metadata_by_paper_name[normalized_id] = metadata
            if pdf_path:
                metadata_by_paper_name[Path(pdf_path).stem.replace(" ", "_")] = metadata
    return metadata_by_paper_name


def normalize_paper_name_from_paper_id(paper_id: str) -> str:
    sanitized = paper_id.strip().replace(":", "_")
    sanitized = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in sanitized)
    return sanitized.strip("_")
