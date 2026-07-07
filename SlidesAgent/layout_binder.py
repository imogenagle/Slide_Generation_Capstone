"""Layout binder + user-template renderer.

Sits between SlideGen's refiner and rendering. When --template_path is set,
SlideGen runs its full agent pipeline as normal; this module then:
  1. Renders the user template's slides to images.
  2. Has a vision LLM describe each user-template slide.
  3. Calls an LLM that picks, for each SlideGen plan slide, which user-template
     slide to clone (the binder agent).
  4. Renders the output PPTX by duplicating chosen user-template slides and
     filling them with SlideGen's content (titles, bullets, figures).

Content (titles, bullets, figure choices) comes straight from SlideGen's plan;
this module only decides which template slide hosts each piece and renders it.
"""

from __future__ import annotations

import base64
import copy
import json
import os
import re
import subprocess
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Optional

from pdf2image import convert_from_path
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from slidegen_openai_utils import build_openai_client, resolve_direct_model_name


SOFFICE_BIN = "/opt/homebrew/bin/soffice"


DESCRIBE_SYSTEM_PROMPT = """You are analyzing a single slide from a PowerPoint template.
Look at the slide image and return a JSON object describing its structure.

Output schema (return ONLY the JSON, no prose):
{
  "slide_type": "title" | "section_divider" | "content_bullets" | "image_focus" | "two_column" | "comparison" | "data_table" | "thanks" | "blank" | "instruction" | "other",
  "description": "1-2 sentences describing what this slide visually shows",
  "is_meta": true | false,
  "text_regions": [
    {"role": "title|subtitle|body|caption|footer|other", "approx_position": "top|center|bottom|left|right"}
  ],
  "image_regions": [
    {"approx_position": "top|center|bottom|left|right", "size": "small|medium|large"}
  ],
  "suitable_for": ["short tag describing what content this slide layout could host"]
}

Set "is_meta": true when the slide is a template-instruction page (e.g. "delete this note",
"how to use this template", "click here in PowerPoint"). Otherwise set "is_meta": false.
"""


BIND_SYSTEM_PROMPT = """You are a layout binder. SlideGen has produced a deck plan for a research paper. The user has uploaded a PowerPoint template. Your job is to pick, for each slide in SlideGen's plan, which slide of the user template to clone as its visual host.

You will be given:
  1. SlideGen's plan slides (each has section, subsection, template_id hint, bullet count, and counts for images / tables / formulas — tables and formulas are rendered as images on the slide, so a slide may carry multiple visuals).
  2. Deck-level metadata: the deck title and the subtitle (author string), used only to fill the cover slide. There is NO per-slide metadata for contents / section dividers / thanks — choose those indices purely from the user-template descriptions in (3).
  3. A JSON list describing each slide of the user template (slide_type, image_regions, is_meta, and `has_picture_placeholder` — a programmatically-verified flag for whether the slide has a real picture placeholder a figure can be dropped into).

SlideGen's `template_id` hints encode the structural need:
  - T1_TextOnly                       -> bullets only, no visuals
  - T2_ImageRight                     -> bullets left, visual right
  - T3_ImageLeft                      -> visual left, bullets right
  - T4_ImageTop                       -> visual top, bullets below
  - T16/T17/T18 (formula variants)    -> include a formula image + optionally an image
  - T19_2Text                         -> two text columns
  - Other T* codes                    -> similar text+visual variants

Rules:
  - NEVER pick a user-template slide where `is_meta` is true.
  - For the cover, prefer slide_type "title".
  - For the contents/agenda, prefer "content_bullets" or "two_column". DO NOT return null unless the template literally has zero non-meta slides with body text — even an imperfect match is better than skipping the agenda entirely.
  - For section dividers, prefer "section_divider"; fall back to "title". Return null ONLY if the template has nothing close.
  - For thanks, prefer "thanks"; fall back to "title". Return null ONLY if the template has nothing close.
  - For body slides:
      - HARD REQUIREMENT: if the SlideGen slide has any visuals (n_images + n_tables + n_formulas > 0), you MUST pick a user-template slide whose `has_picture_placeholder` is true. A figure dropped onto a slide with no picture placeholder lands on top of the bullet text. Only ignore this if NO non-meta template slide has `has_picture_placeholder` true.
      - Among `has_picture_placeholder` slides, prefer slide_type "image_focus", "two_column", or "comparison".
      - If it has multiple visuals (>= 2) -> strongly prefer "image_focus" or "comparison" (more room).
      - If it's text-only (total_visuals == 0) -> prefer "content_bullets" or "two_column"; `has_picture_placeholder` does not matter.
      - For data_table content (n_tables > 0 and n_bullets is low) -> prefer "data_table" if available, else an "image_focus" slide with a picture placeholder.
      - Match figure-position hint (right/left/top) to image_regions when possible.
  - It is OK to reuse the same user-template slide for many body slides (multiple bullet-style slides can all clone the same content_bullets layout).
  - Pick the SAME section_divider layout for every section transition.

Output schema (return ONLY the JSON, no prose):
{
  "cover_slide_index": <int>,
  "contents_slide_index": <int or null>,
  "section_divider_slide_index": <int or null>,
  "thanks_slide_index": <int or null>,
  "body_assignments": [<int>, <int>, ...]
}

`body_assignments` must have length equal to the number of SlideGen plan slides, in the same order. Each entry is the user-template slide index to clone for that plan slide.
"""


# ---------- Logging ---------- #


@contextmanager
def _stage(num: int, label: str):
    print(f"[layout_binder] (stage {num}) {label}...")
    import time
    t0 = time.time()
    try:
        yield
    finally:
        print(f"[layout_binder]   (stage {num} took {time.time() - t0:.1f}s)")


# ---------- Stage 1: render template to images ---------- #


def render_template_to_images(template_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    if not Path(SOFFICE_BIN).exists():
        raise FileNotFoundError(
            f"LibreOffice not found at {SOFFICE_BIN}. "
            "Install with `brew install --cask libreoffice` or update SOFFICE_BIN in layout_binder.py."
        )
    try:
        subprocess.run(
            [SOFFICE_BIN, "--headless", "--convert-to", "pdf",
             "--outdir", str(output_dir), str(template_path)],
            check=True,
            capture_output=True,
            timeout=120,
        )
    except subprocess.CalledProcessError as e:
        stderr = e.stderr.decode("utf-8", errors="replace") if e.stderr else "(no stderr)"
        raise RuntimeError(
            f"LibreOffice failed to convert template to PDF (exit {e.returncode}). "
            f"Stderr: {stderr.strip()[:500]}"
        ) from e
    except subprocess.TimeoutExpired as e:
        raise RuntimeError(
            f"LibreOffice timed out after {e.timeout}s converting {template_path.name}."
        ) from e

    pdf_path = output_dir / f"{template_path.stem}.pdf"
    if not pdf_path.exists():
        raise FileNotFoundError(
            f"LibreOffice exited cleanly but produced no PDF at {pdf_path}."
        )

    images = convert_from_path(pdf_path, dpi=120)
    image_paths: list[Path] = []
    for i, img in enumerate(images):
        path = output_dir / f"slide_{i:03d}.png"
        img.save(path, "PNG")
        image_paths.append(path)
    return image_paths


# ---------- Stage 2: describe template slides ---------- #


def _parse_json_from_response(text: str) -> dict[str, Any]:
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```", 2)[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.rsplit("```", 1)[0].strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        try:
            from json_repair import repair_json
            return json.loads(repair_json(text))
        except Exception:
            return {}


def describe_template_slides(image_paths: list[Path]) -> list[dict[str, Any]]:
    client = build_openai_client()
    deployment = resolve_direct_model_name(os.environ.get("AZURE_DEPLOYMENT_NAME", ""))
    descriptions: list[dict[str, Any]] = []
    for i, img_path in enumerate(image_paths):
        with img_path.open("rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()
        response = client.chat.completions.create(
            model=deployment,
            messages=[
                {"role": "system", "content": DESCRIBE_SYSTEM_PROMPT},
                {"role": "user", "content": [
                    {"type": "text", "text": f"This is slide {i+1} of the template."},
                    {"type": "image_url",
                     "image_url": {"url": f"data:image/png;base64,{img_b64}"}},
                ]},
            ],
            max_completion_tokens=800,
        )
        raw = response.choices[0].message.content or ""
        parsed = _parse_json_from_response(raw)
        parsed.setdefault("slide_type", "other")
        parsed.setdefault("description", "")
        parsed.setdefault("is_meta", False)
        parsed.setdefault("text_regions", [])
        parsed.setdefault("image_regions", [])
        parsed.setdefault("suitable_for", [])
        parsed["slide_index"] = i
        descriptions.append(parsed)
        meta_tag = " [META]" if parsed.get("is_meta") else ""
        print(f"[layout_binder]   slide {i+1}: {parsed.get('slide_type', '?')}{meta_tag} - {parsed.get('description', '')[:70]}")
    return descriptions


def scan_picture_placeholders(template_path: Path) -> dict[int, int]:
    """Return {slide_index: count of PICTURE placeholders} for the template.

    Programmatic ground truth — far more reliable than the vision LLM's
    slide_type guess for deciding which slides can host a figure.
    PP_PLACEHOLDER.PICTURE == 18.
    """
    prs = Presentation(str(template_path))
    counts: dict[int, int] = {}
    for i, slide in enumerate(prs.slides):
        n = 0
        for ph in slide.placeholders:
            if ph.placeholder_format.type == 18:
                n += 1
        counts[i] = n
    return counts


# ---------- Stage 3: bind SlideGen plan slides to user-template slides ---------- #


def _summarize_plan_for_binder(slidegen_plan: dict) -> list[dict]:
    """Produce a compact per-slide summary for the binder LLM."""
    summary = []
    for i, s in enumerate(slidegen_plan.get("slides", [])):
        n_images = len(s.get("images") or [])
        n_tables = len(s.get("tables") or [])
        n_formulas = len(s.get("formulas") or [])
        summary.append({
            "plan_index": i,
            "section": s.get("section", ""),
            "subsection": s.get("subsection", ""),
            "template_id": s.get("template_id", ""),
            "n_bullets": len(s.get("bullets") or []),
            "n_images": n_images,
            "n_tables": n_tables,
            "n_formulas": n_formulas,
            "total_visuals": n_images + n_tables + n_formulas,
        })
    return summary


def bind_user_template_to_plan(
    slidegen_plan: dict,
    template_descriptions: list[dict],
    deck_meta: dict,
) -> dict:
    """LLM call: pick a user-template slide for each SlideGen plan slide.

    Returns a dict with cover/contents/section_divider/thanks indices plus a
    list of body_assignments parallel to slidegen_plan['slides'].
    """
    client = build_openai_client()
    deployment = resolve_direct_model_name(os.environ.get("AZURE_DEPLOYMENT_NAME", ""))

    plan_summary = _summarize_plan_for_binder(slidegen_plan)

    user_payload = {
        "deck_meta": deck_meta,
        "plan_slides": plan_summary,
        "template_slides": template_descriptions,
    }

    response = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": BIND_SYSTEM_PROMPT},
            {"role": "user", "content": json.dumps(user_payload)},
        ],
        max_completion_tokens=2000,
    )
    raw = response.choices[0].message.content or ""
    binding = _parse_json_from_response(raw)
    if not binding or "body_assignments" not in binding:
        raise RuntimeError(
            f"Binder returned malformed response. Head: {raw[:300]}"
        )

    n_plan_slides = len(slidegen_plan.get("slides", []))
    assignments = binding.get("body_assignments") or []
    if len(assignments) != n_plan_slides:
        print(
            f"[layout_binder]   WARN: binder returned {len(assignments)} assignments "
            f"for {n_plan_slides} plan slides; padding/truncating."
        )
        if len(assignments) < n_plan_slides:
            fallback = binding.get("cover_slide_index", 0) or 0
            assignments.extend([fallback] * (n_plan_slides - len(assignments)))
        else:
            assignments = assignments[:n_plan_slides]
        binding["body_assignments"] = assignments

    # Enforce the picture-placeholder rule the prompt asks for: any plan slide
    # with visuals MUST land on a template slide that has a picture placeholder,
    # otherwise the figure renders on top of the bullet text. We don't trust the
    # LLM to honor this — verify against the programmatic flag and repair.
    pic_slides = [d["slide_index"] for d in template_descriptions
                  if d.get("has_picture_placeholder") and not d.get("is_meta")]
    if pic_slides:
        plan_slides = slidegen_plan.get("slides", [])
        repaired = 0
        for i, plan_slide in enumerate(plan_slides):
            n_visuals = (len(plan_slide.get("images") or [])
                         + len(plan_slide.get("tables") or [])
                         + len(plan_slide.get("formulas") or []))
            if n_visuals == 0:
                continue
            if assignments[i] not in pic_slides:
                # Reassign to a picture-placeholder slide. Prefer one the binder
                # already chose for another figure slide (visual consistency).
                already = next((a for a in assignments if a in pic_slides), None)
                assignments[i] = already if already is not None else pic_slides[0]
                repaired += 1
        if repaired:
            print(f"[layout_binder]   repaired {repaired} figure slide(s) "
                  f"onto picture-placeholder layouts")
    else:
        print("[layout_binder]   WARN: template has no picture-placeholder "
              "slides; figures will use beside-text fallback placement")

    return binding


# ---------- Stage 4: render output PPTX ---------- #


# Placeholder type IDs (PP_PLACEHOLDER):
#   1 TITLE, 2 BODY, 3 CENTER_TITLE, 4 SUBTITLE, 7 OBJECT
_TITLE_TYPES = {1, 3}
_SUBTITLE_TYPES = {4}
_SKIP_TYPES = {13, 14, 15, 16}  # slide number / header / footer / date


_A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _duplicate_slide(prs, src_index: int):
    """Append a deep copy of slide `src_index` to the end of `prs`."""
    source = prs.slides[src_index]
    new_slide = prs.slides.add_slide(source.slide_layout)

    new_spTree = new_slide.shapes._spTree
    for shape in list(new_slide.shapes):
        new_spTree.remove(shape.element)

    for shape in source.shapes:
        new_spTree.append(copy.deepcopy(shape.element))

    return new_slide


def _delete_slides_at(prs, indices: list[int]) -> None:
    sldIdLst = prs.slides._sldIdLst
    sldId_elems = list(sldIdLst)
    for i in sorted(indices, reverse=True):
        sldIdLst.remove(sldId_elems[i])


# Substrings (lowercased) that mark a shape as leftover template sample content.
_SAMPLE_TEXT_MARKERS = (
    "bullet level", "lorem ipsum", "helvetica", "additional content here",
    "delete this slide", "image caption", "caption text here", "caption or",
    "description text here", "presentation title here", "divider slide text",
    "slide title", "firstname lastname", "cnet@uchicago", "url.uchicago",
    "head level", "text box ma", "text level", "click here", "placeholder text",
)


def _collect_shape_text(shape) -> str:
    """Gather all text from a shape, recursing into groups. Lowercased."""
    parts: list[str] = []
    if shape.has_text_frame and shape.text_frame.text.strip():
        parts.append(shape.text_frame.text)
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            parts.append(_collect_shape_text(child))
    return " ".join(parts).lower()


def _strip_sample_shapes(slide) -> int:
    """Remove leftover template sample content from a cloned slide.

    Targets non-placeholder shapes only (placeholders are handled by
    _fill_slide_text). Removes sample charts and tables outright, and any
    shape whose text matches a known sample marker (e.g. callout groups with
    'Bullet level 1' / 'Helvetica Bold 24pt'). Brand decoration with no
    sample text is left intact. Returns the number stripped.
    """
    to_remove = []
    for shape in slide.shapes:
        # Sample charts / tables — template examples, never real content here.
        if getattr(shape, "has_chart", False) or getattr(shape, "has_table", False):
            to_remove.append(shape)
            continue
        if shape.is_placeholder:
            continue  # _fill_slide_text owns placeholders
        text = _collect_shape_text(shape)
        if text and any(marker in text for marker in _SAMPLE_TEXT_MARKERS):
            to_remove.append(shape)
    for shape in to_remove:
        shape._element.getparent().remove(shape._element)
    return len(to_remove)


def _enable_shrink_to_fit(tf) -> None:
    if tf is None:
        return
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE


def _set_text_frame(tf, text: str) -> None:
    if not tf.paragraphs:
        tf.text = text
        _enable_shrink_to_fit(tf)
        return
    p0 = tf.paragraphs[0]
    for r in list(p0.runs)[1:]:
        r._r.getparent().remove(r._r)
    if p0.runs:
        p0.runs[0].text = text
    else:
        p0.text = text
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    _enable_shrink_to_fit(tf)


_LATEX_UNICODE = {
    r"\\leq": "≤", r"\\geq": "≥", r"\\le": "≤", r"\\ge": "≥",
    r"\\neq": "≠", r"\\approx": "≈", r"\\sim": "∼",
    r"\\to": "→", r"\\rightarrow": "→", r"\\Rightarrow": "⇒",
    r"\\leftarrow": "←", r"\\Leftarrow": "⇐", r"\\leftrightarrow": "↔",
    r"\\times": "×", r"\\cdot": "·", r"\\pm": "±",
    r"\\alpha": "α", r"\\beta": "β", r"\\gamma": "γ", r"\\delta": "δ",
    r"\\epsilon": "ε", r"\\varepsilon": "ε", r"\\zeta": "ζ", r"\\eta": "η",
    r"\\theta": "θ", r"\\lambda": "λ", r"\\mu": "μ", r"\\nu": "ν",
    r"\\pi": "π", r"\\rho": "ρ", r"\\sigma": "σ", r"\\tau": "τ",
    r"\\phi": "φ", r"\\chi": "χ", r"\\psi": "ψ", r"\\omega": "ω",
    r"\\Gamma": "Γ", r"\\Delta": "Δ", r"\\Theta": "Θ", r"\\Lambda": "Λ",
    r"\\Pi": "Π", r"\\Sigma": "Σ", r"\\Phi": "Φ", r"\\Omega": "Ω",
    r"\\ell": "ℓ", r"\\infty": "∞", r"\\partial": "∂", r"\\nabla": "∇",
    r"\\in": "∈", r"\\subset": "⊂", r"\\subseteq": "⊆", r"\\cup": "∪", r"\\cap": "∩",
    r"\\forall": "∀", r"\\exists": "∃",
}
_MATH_DELIM_RE = re.compile(r"\$+([^$]*?)\$+")


def _clean_math(text: str) -> str:
    if not text or ("$" not in text and "\\" not in text):
        return text
    for cmd, repl in _LATEX_UNICODE.items():
        text = re.sub(cmd + r"\b", repl, text)
    text = _MATH_DELIM_RE.sub(lambda m: m.group(1), text)
    text = re.sub(r"\\(?:mathbf|mathit|mathrm|text|emph|bf|it)\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\[a-zA-Z]+", "", text)
    return text.strip()


def _flatten_slidegen_bullets(raw: list) -> list[str]:
    """Convert SlideGen's nested {text, sub: [...]} bullets to flat strings.

    SlideGen emits `bullets: [{"text": "...", "sub": ["...", "..."]}, ...]`.
    Sub-bullets are flattened to two-space-indented strings on the parent's
    level so they render as visually-nested bullets in PowerPoint.
    """
    out: list[str] = []
    for item in raw or []:
        if isinstance(item, dict):
            text = item.get("text") or ""
            if text:
                out.append(text)
            for sub in (item.get("sub") or []):
                if isinstance(sub, str) and sub:
                    out.append("  " + sub)
                elif isinstance(sub, dict):
                    out.extend(_flatten_slidegen_bullets([sub]))
        elif isinstance(item, str):
            out.append(item)
    return out


def _set_bullets(tf, bullets: list[str]) -> None:
    if not bullets:
        return
    _set_text_frame(tf, bullets[0])
    template_para_xml = tf.paragraphs[0]._p
    txBody = template_para_xml.getparent()

    for bullet in bullets[1:]:
        new_p = copy.deepcopy(template_para_xml)
        for r in list(new_p.findall(f"{_A_NS}r")):
            new_p.remove(r)
        for br in list(new_p.findall(f"{_A_NS}br")):
            new_p.remove(br)
        txBody.append(new_p)
        last_para = tf.paragraphs[-1]
        run = last_para.add_run()
        run.text = bullet


def _clear_text_frame(tf) -> None:
    if not tf.paragraphs:
        return
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    for extra in list(tf.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)


def _fill_slide_text(slide, content: dict) -> None:
    """Fill placeholders BY TYPE (not by name) so it works with any template."""
    title_text = _clean_math(content.get("title", "") or "")
    subtitle_text = _clean_math(content.get("subtitle", "") or "")
    bullets = [_clean_math(b) for b in (content.get("bullets") or [])]
    bullets = [b for b in bullets if b]

    used_phs: set[int] = set()

    # 1. Title
    if title_text:
        for ph in slide.placeholders:
            if ph.placeholder_format.type in _TITLE_TYPES and ph.has_text_frame:
                _set_text_frame(ph.text_frame, title_text)
                used_phs.add(ph.placeholder_format.idx)
                break

    # 2. Bullets (before subtitle, so they win the single-content slot on Section Header layouts)
    if bullets:
        chosen = None
        for ptype in (7, 2):  # OBJECT preferred, BODY fallback
            for ph in slide.placeholders:
                if (ph.placeholder_format.type == ptype
                    and ph.has_text_frame
                    and ph.placeholder_format.idx not in used_phs):
                    chosen = ph
                    break
            if chosen is not None:
                break
        if chosen is not None:
            _set_bullets(chosen.text_frame, bullets)
            used_phs.add(chosen.placeholder_format.idx)

    # 3. Subtitle
    if subtitle_text:
        chosen = None
        for ph in slide.placeholders:
            if ph.placeholder_format.type in _SUBTITLE_TYPES and ph.has_text_frame:
                chosen = ph
                break
        if chosen is None:
            for ph in slide.placeholders:
                if (ph.placeholder_format.type == 2
                    and ph.has_text_frame
                    and ph.placeholder_format.idx not in used_phs):
                    chosen = ph
                    break
        if chosen is not None:
            _set_text_frame(chosen.text_frame, subtitle_text)
            used_phs.add(chosen.placeholder_format.idx)

    # 4. Clear remaining unused content placeholders so sample text doesn't ship.
    for ph in slide.placeholders:
        ptype = ph.placeholder_format.type
        if ptype in _SKIP_TYPES:
            continue
        if not ph.has_text_frame:
            continue
        if ph.placeholder_format.idx in used_phs:
            continue
        if ptype in (2, 4, 7):
            _clear_text_frame(ph.text_frame)


def _fit_image_in_box(slide, image_path: Path, box_left: int, box_top: int,
                      box_w: int, box_h: int) -> None:
    """Place an image into a box, aspect-preserved, centered."""
    from PIL import Image
    with Image.open(str(image_path)) as img:
        fw, fh = img.size
    figure_aspect = fw / fh if fh else 1.0
    box_aspect = box_w / box_h if box_h else 1.0

    if figure_aspect > box_aspect:
        pic = slide.shapes.add_picture(str(image_path), box_left, box_top, width=box_w)
        pic.top = box_top + (box_h - pic.height) // 2
    else:
        pic = slide.shapes.add_picture(str(image_path), box_left, box_top, height=box_h)
        pic.left = box_left + (box_w - pic.width) // 2


def _find_picture_anchor(slide):
    """Find where the template designer wants a picture to live on this slide.

    Returns ((left, top, width, height), shape) for the anchor, or (None, None).
    Preference order:
      1. A picture-type placeholder (PP_PLACEHOLDER.PICTURE == 18).
      2. The largest picture shape on the slide (the designer's sample image).
    The returned shape is the element to remove before placing our new image —
    we replace what the designer put there, rather than stacking on top of it.
    """
    for ph in slide.placeholders:
        if ph.placeholder_format.type == 18:
            return (ph.left or 0, ph.top or 0, ph.width or 0, ph.height or 0), ph

    candidates = [s for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE]
    if candidates:
        largest = max(candidates, key=lambda s: (s.width or 0) * (s.height or 0))
        return (largest.left or 0, largest.top or 0, largest.width or 0, largest.height or 0), largest

    return None, None


def _strip_pictures_in_region(slide, box_left: int, box_top: int,
                              box_w: int, box_h: int, overlap_threshold: float = 0.3) -> int:
    """Remove picture shapes on the slide whose bounding box significantly
    overlaps the target visual region. Decorative pictures elsewhere
    (e.g. corner logos) survive. Returns the number stripped.

    We strip BEFORE inserting our new visual so two copies of the same area
    don't end up stacked on top of each other.
    """
    to_remove = []
    for shape in slide.shapes:
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        sl, st = shape.left or 0, shape.top or 0
        sw, sh = shape.width or 0, shape.height or 0
        ix1, iy1 = max(sl, box_left), max(st, box_top)
        ix2, iy2 = min(sl + sw, box_left + box_w), min(st + sh, box_top + box_h)
        if ix2 > ix1 and iy2 > iy1:
            inter_area = (ix2 - ix1) * (iy2 - iy1)
            shape_area = sw * sh
            if shape_area > 0 and inter_area / shape_area >= overlap_threshold:
                to_remove.append(shape)
    for shape in to_remove:
        shape._element.getparent().remove(shape._element)
    return len(to_remove)


def _insert_visuals(slide, prs, visual_paths: list[Path]) -> int:
    """Place one or more visuals (image / table / formula PNGs) on the slide.

    Anchors to the template's picture placeholder when present; otherwise
    carves a right-hand column beside the text. Single visual fills the box;
    multiple visuals are stacked in equal-height rows. Returns the count placed.
    """
    if not visual_paths:
        return 0

    # Prefer the template designer's picture spot (placeholder or sample image)
    # so the paper figure lands where the layout was designed for it. Fall back
    # to the LLM-described region only if the slide has no picture anchor.
    anchor_box, anchor_shape = _find_picture_anchor(slide)
    if anchor_box is not None:
        box_left, box_top, box_w, box_h = anchor_box
        anchor_shape._element.getparent().remove(anchor_shape._element)
        # Strip any other lingering pictures in the same region (loose threshold).
        stripped = _strip_pictures_in_region(slide, box_left, box_top, box_w, box_h,
                                             overlap_threshold=0.1)
        if stripped:
            print(f"[layout_binder]   stripped {stripped} extra picture(s) near anchor")
        print(f"[layout_binder]   anchored to template picture spot")
    else:
        # No picture spot on this layout. Carve a right-hand column for the
        # figure and shrink the body/object text placeholders to the left half,
        # so the figure sits BESIDE the bullets at a usable size rather than
        # on top of them. (The binder should rarely land here — figure slides
        # are repaired onto picture-placeholder layouts upstream.)
        slide_w = prs.slide_width
        slide_h = prs.slide_height
        text_limit = int(slide_w * 0.50)
        for ph in slide.placeholders:
            if ph.placeholder_format.type in (2, 7) and ph.has_text_frame:
                if (ph.left or 0) + (ph.width or 0) > text_limit:
                    ph.width = max(text_limit - (ph.left or 0), int(slide_w * 0.30))
        box_left = int(slide_w * 0.54)
        box_top = int(slide_h * 0.24)
        box_w = int(slide_w * 0.44)
        box_h = int(slide_h * 0.68)
        stripped = _strip_pictures_in_region(slide, box_left, box_top, box_w, box_h)
        if stripped:
            print(f"[layout_binder]   stripped {stripped} sample picture(s) from fallback region")
        print("[layout_binder]   no picture anchor; placed figure beside text")

    n = len(visual_paths)
    gap = max(int(box_h * 0.02), 0) if n > 1 else 0
    row_h = (box_h - gap * (n - 1)) // n if n > 0 else box_h

    placed = 0
    for i, path in enumerate(visual_paths):
        try:
            row_top = box_top + i * (row_h + gap)
            _fit_image_in_box(slide, path, box_left, row_top, box_w, row_h)
            placed += 1
        except Exception as e:
            print(f"[layout_binder]   WARN: insert failed for {path.name}: {e}")
    return placed


# ---------- Build the render sequence from SlideGen plan + binding ---------- #


def _pick_fallback_slide(
    descriptions: list[dict],
    preferred_types: list[str],
    avoid_idx: Optional[int] = None,
) -> Optional[int]:
    """Find the first non-meta template slide whose slide_type matches one of
    `preferred_types`, optionally avoiding a specific index (e.g. don't reuse
    the cover for everything). Returns the slide_index, or None if nothing
    matches.
    """
    for ptype in preferred_types:
        for d in descriptions:
            if d.get("is_meta"):
                continue
            idx = d.get("slide_index")
            if avoid_idx is not None and idx == avoid_idx:
                continue
            if d.get("slide_type") == ptype:
                return idx
    return None


def _build_render_sequence(
    slidegen_plan: dict,
    binding: dict,
    deck_meta: dict,
    template_descriptions: Optional[list[dict]] = None,
) -> list[dict]:
    """Build the full list of slides to render (cover, contents, dividers, body, thanks).

    Each entry: {kind, template_slide_index, content: {title, subtitle, bullets}, slidegen_slide?}

    If the binder LLM returned null for contents/section_divider/thanks, fall
    back to scanning `template_descriptions` for a usable slide_type.
    """
    seq: list[dict] = []
    descs = template_descriptions or []

    # Cover
    cover_idx = binding.get("cover_slide_index")
    if cover_idx is None:
        cover_idx = _pick_fallback_slide(descs, ["title"]) or 0
        print(f"[layout_binder]   cover fallback -> slide {cover_idx}")
    seq.append({
        "kind": "cover",
        "template_slide_index": cover_idx,
        "content": {
            "title": deck_meta.get("deck_title", ""),
            "subtitle": deck_meta.get("deck_subtitle", ""),
        },
    })

    # Contents (agenda) — fall back to a content_bullets / two_column layout
    # if the LLM punted to null.
    unique_sections = []
    seen = set()
    for s in slidegen_plan.get("slides", []):
        sec = s.get("section", "")
        if sec and sec not in seen:
            seen.add(sec)
            unique_sections.append(sec)

    contents_idx = binding.get("contents_slide_index")
    if contents_idx is None and unique_sections:
        contents_idx = _pick_fallback_slide(
            descs, ["content_bullets", "two_column", "image_focus"],
            avoid_idx=cover_idx,
        )
        if contents_idx is not None:
            print(f"[layout_binder]   contents fallback -> slide {contents_idx}")
    if contents_idx is not None and unique_sections:
        seq.append({
            "kind": "contents",
            "template_slide_index": contents_idx,
            "content": {"title": "Contents", "bullets": unique_sections},
        })

    # Body slides, with section-divider injection on section transitions.
    section_div_idx = binding.get("section_divider_slide_index")
    if section_div_idx is None:
        section_div_idx = _pick_fallback_slide(
            descs, ["section_divider", "title"], avoid_idx=cover_idx,
        )
        if section_div_idx is not None:
            print(f"[layout_binder]   section_divider fallback -> slide {section_div_idx}")
    body_assignments = binding.get("body_assignments") or []
    current_section: Optional[str] = None
    section_counter = 0
    for i, plan_slide in enumerate(slidegen_plan.get("slides", [])):
        section = plan_slide.get("section", "")
        if section != current_section:
            current_section = section
            section_counter += 1
            if section_div_idx is not None and section:
                seq.append({
                    "kind": "section_divider",
                    "template_slide_index": section_div_idx,
                    "content": {
                        "title": section,
                        "subtitle": f"PART {section_counter:02d}",
                    },
                })

        tpl_idx = body_assignments[i] if i < len(body_assignments) else (cover_idx or 0)
        bullets = _flatten_slidegen_bullets(plan_slide.get("bullets") or [])
        seq.append({
            "kind": "body",
            "template_slide_index": tpl_idx,
            "content": {
                "title": plan_slide.get("subsection", "") or plan_slide.get("section", ""),
                "bullets": bullets,
            },
            # Carry the raw SlideGen slide so the renderer can resolve visuals
            # (images + tables + formulas) via layout_filler.resolve_visual_paths.
            "slidegen_slide": plan_slide,
        })

    # Thanks
    thanks_idx = binding.get("thanks_slide_index")
    if thanks_idx is None:
        thanks_idx = _pick_fallback_slide(
            descs, ["thanks", "title"], avoid_idx=cover_idx,
        )
        if thanks_idx is not None:
            print(f"[layout_binder]   thanks fallback -> slide {thanks_idx}")
    if thanks_idx is not None:
        seq.append({
            "kind": "thanks",
            "template_slide_index": thanks_idx,
            "content": {"title": "Thank You", "subtitle": ""},
        })

    return seq


# ---------- Top-level entry point ---------- #


def _load_slidegen_artifacts(args) -> tuple[dict, dict, dict, dict]:
    """Load SlideGen's plan / raw content / figures / formulas from contents."""
    variant_suffix = "_personalized" if getattr(args, "use_author_preferences", False) else "_baseline"
    contents_dir = REPO_ROOT / "contents" / args.paper_name
    plan_path = contents_dir / f"<{args.model_name_t}_{args.model_name_v}>_slide_plan{variant_suffix}.json"
    raw_path = contents_dir / f"<{args.model_name_t}_{args.model_name_v}>_raw_content.json"
    figs_path = contents_dir / f"<{args.model_name_t}_{args.model_name_v}>_figures.json"
    formulas_path = contents_dir / f"<{args.model_name_t}_{args.model_name_v}>_formula_match.json"

    if not plan_path.exists():
        raise FileNotFoundError(f"SlideGen plan not found: {plan_path}")
    plan = json.loads(plan_path.read_text(encoding="utf-8"))

    raw = {}
    if raw_path.exists():
        raw = json.loads(raw_path.read_text(encoding="utf-8"))

    figures = {}
    if figs_path.exists():
        figures = json.loads(figs_path.read_text(encoding="utf-8"))

    formulas = {}
    if formulas_path.exists():
        formulas = json.loads(formulas_path.read_text(encoding="utf-8"))

    return plan, raw, figures, formulas


def _append_body_notes(
    slide,
    slidegen_slide: dict,
    outline_json: dict,
    figs_data: dict,
    formula_data: dict,
    *,
    get_content,
    get_image_reasons,
    get_table_reasons,
    get_formula_reasons,
) -> bool:
    """Add the same body-slide notes used by layout_filler.generate_pptx_from_plan."""
    notes_chunks: list[str] = []
    section = slidegen_slide.get("section", "")
    subsection = slidegen_slide.get("subsection", "")

    txt = get_content(section, subsection, outline_json)
    if txt:
        notes_chunks.append(txt)

    if slidegen_slide.get("images"):
        img_r = get_image_reasons(section, subsection, slidegen_slide["images"], figs_data)
        if img_r:
            notes_chunks.append(img_r)

    if slidegen_slide.get("tables"):
        tb_r = get_table_reasons(section, subsection, slidegen_slide["tables"], figs_data)
        if tb_r:
            notes_chunks.append(tb_r)

    if slidegen_slide.get("formulas"):
        fm_r = get_formula_reasons(section, subsection, slidegen_slide["formulas"], formula_data)
        if fm_r:
            notes_chunks.append(fm_r)

    if not notes_chunks:
        return False

    nframe = slide.notes_slide.notes_text_frame
    if nframe.text and not nframe.text.endswith("\n"):
        nframe.text += "\n"
    nframe.text += "\n\n".join(notes_chunks)
    return True


def _slugify_template_label(template_path: Path) -> str:
    """Derive a filesystem-safe label from a template path (e.g. 'bee template.pptx' -> 'bee_template')."""
    stem = template_path.stem
    return re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_").lower()


def bind_and_render(args, user_template_path: str, template_label: Optional[str] = None) -> Optional[Path]:
    """Run binder + render. Assumes SlideGen's full pipeline has already run.

    Reads the cached plan/raw_content/figures from contents/<paper_name>/, calls
    the binder LLM, and writes the output PPTX. Returns the output path on
    success, None on failure.

    `template_label` is a short tag (e.g. "uchicago", "bee") used in artifact
    paths and the output filename so multiple templates can share one
    paper_name without overwriting each other. If None, derived from template
    filename.
    """
    template_path = Path(user_template_path).expanduser().resolve()
    if not template_path.exists():
        print(f"[layout_binder] ERROR: template not found: {template_path}")
        return None
    if template_path.suffix.lower() != ".pptx":
        print(f"[layout_binder] ERROR: template must be .pptx: {template_path}")
        return None

    if template_label is None:
        template_label = _slugify_template_label(template_path)

    try:
        plan, raw_content, figures, formulas = _load_slidegen_artifacts(args)
    except FileNotFoundError as e:
        print(f"[layout_binder] ERROR: {e}")
        return None

    deck_meta = {
        "deck_title": (raw_content.get("metadata") or {}).get("title", args.paper_name),
        "deck_subtitle": (raw_content.get("metadata") or {}).get("author", ""),
    }

    contents_dir = REPO_ROOT / "contents" / args.paper_name
    # Template-specific subdir so different templates against the same
    # paper_name don't collide on _user_template_images / _user_template_binding.
    template_artifacts_dir = contents_dir / f"_user_template__{template_label}"
    template_images_dir = template_artifacts_dir / "images"
    descriptions_cache = template_artifacts_dir / "descriptions.json"

    with _stage(1, f"render template to images [{template_label}]"):
        # Cache: skip LibreOffice if the per-slide PNGs already exist.
        existing_pngs = sorted(template_images_dir.glob("slide_*.png")) if template_images_dir.exists() else []
        if existing_pngs:
            image_paths = existing_pngs
            print(f"[layout_binder]   reusing {len(image_paths)} cached template slide images")
        else:
            image_paths = render_template_to_images(template_path, template_images_dir)
            print(f"[layout_binder]   rendered {len(image_paths)} template slide images")

    with _stage(2, f"describe template slides (vision LLM) [{template_label}]"):
        # Cache: skip the vision LLM if descriptions are already on disk.
        if descriptions_cache.exists():
            descriptions = json.loads(descriptions_cache.read_text(encoding="utf-8"))
            print(f"[layout_binder]   reusing cached descriptions for {len(descriptions)} slides")
        else:
            descriptions = describe_template_slides(image_paths)
            descriptions_cache.parent.mkdir(parents=True, exist_ok=True)
            descriptions_cache.write_text(json.dumps(descriptions, indent=2), encoding="utf-8")

        # Merge in programmatic picture-placeholder ground truth (always — the
        # cached descriptions.json may predate this field).
        pic_counts = scan_picture_placeholders(template_path)
        for d in descriptions:
            n = pic_counts.get(d.get("slide_index"), 0)
            d["has_picture_placeholder"] = n > 0
            d["n_picture_placeholders"] = n
        n_pic = sum(1 for d in descriptions if d.get("has_picture_placeholder"))
        print(f"[layout_binder]   {n_pic}/{len(descriptions)} template slides have a picture placeholder")

    with _stage(3, f"bind plan slides to template slides [{template_label}]"):
        binding = bind_user_template_to_plan(plan, descriptions, deck_meta)
        print(f"[layout_binder]   binding: cover={binding.get('cover_slide_index')}, "
              f"contents={binding.get('contents_slide_index')}, "
              f"divider={binding.get('section_divider_slide_index')}, "
              f"thanks={binding.get('thanks_slide_index')}, "
              f"body={binding.get('body_assignments')}")
        binding_cache = template_artifacts_dir / "binding.json"
        binding_cache.parent.mkdir(parents=True, exist_ok=True)
        binding_cache.write_text(json.dumps(binding, indent=2), encoding="utf-8")

    with _stage(4, "render output PPTX"):
        # Lazy import to avoid pulling apply_color / heavy filler deps at module load.
        from SlidesAgent.layout_filler import (
            get_content,
            get_formula_reasons,
            get_image_reasons,
            get_table_reasons,
            resolve_visual_paths,
        )

        seq = _build_render_sequence(plan, binding, deck_meta, template_descriptions=descriptions)
        prs = Presentation(str(template_path))
        n_template_slides = len(prs.slides)

        rendered = 0
        visuals_inserted = 0
        notes_inserted = 0
        for entry in seq:
            src = entry["template_slide_index"]
            if src is None or not (0 <= src < n_template_slides):
                print(f"[layout_binder]   WARN: skipping {entry['kind']} (bad template_slide_index={src})")
                continue
            new_slide = _duplicate_slide(prs, src)
            # Strip leftover template sample content (sample charts/tables,
            # callout groups with "Bullet level 1" etc.) before filling.
            stripped = _strip_sample_shapes(new_slide)
            if stripped:
                print(f"[layout_binder]   stripped {stripped} sample shape(s) from {entry['kind']}")
            try:
                _fill_slide_text(new_slide, entry["content"])
            except Exception as e:
                print(f"[layout_binder]   WARN: text fill failed on {entry['kind']}: {e}")

            # Resolve and insert visuals (images + tables + formulas) for body slides.
            slidegen_slide = entry.get("slidegen_slide")
            if slidegen_slide:
                try:
                    if _append_body_notes(
                        new_slide,
                        slidegen_slide,
                        raw_content,
                        figures,
                        formulas,
                        get_content=get_content,
                        get_image_reasons=get_image_reasons,
                        get_table_reasons=get_table_reasons,
                        get_formula_reasons=get_formula_reasons,
                    ):
                        notes_inserted += 1
                except Exception as e:
                    print(f"[layout_binder]   WARN: speaker-note fill failed on {entry['kind']}: {e}")
                try:
                    visual_paths = resolve_visual_paths(slidegen_slide, args)
                except Exception as e:
                    print(f"[layout_binder]   WARN: visual resolution failed on {entry['kind']}: {e}")
                    visual_paths = []
                if visual_paths:
                    placed = _insert_visuals(new_slide, prs, visual_paths)
                    visuals_inserted += placed
            rendered += 1

        _delete_slides_at(prs, list(range(n_template_slides)))

        variant_suffix = "_personalized" if getattr(args, "use_author_preferences", False) else "_baseline"
        output_pptx = contents_dir / f"{args.model_name_t}_{args.model_name_v}_output_slides{variant_suffix}_user_template__{template_label}.pptx"
        output_pptx.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(output_pptx))
        print(
            f"[layout_binder]   rendered {rendered} slides "
            f"({visuals_inserted} visuals placed, {notes_inserted} notes added) -> {output_pptx}"
        )

    return output_pptx
