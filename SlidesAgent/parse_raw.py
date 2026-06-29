from dotenv import load_dotenv
from utils.src.utils import get_json_from_response
import json
import random
import time
from tenacity import retry, stop_after_attempt
from slidegen_openai_utils import build_openai_client, resolve_direct_model_name
from pathlib import Path
import os
from SlidesAgent.output_paths import (
    asset_paper_name,
    formula_crop_path,
    formulas_json_path,
    formula_sections_path,
    html_referenced_path,
    images_json_path,
    markdown_embedded_path,
    markdown_referenced_path,
    page_image_path,
    paper_image_tables_dir,
    paper_output_dir,
    picture_image_path,
    raw_content_path,
    table_image_path,
    tables_json_path,
)

import PIL

from jinja2 import Template
import re
import argparse    
from typing import Any

load_dotenv()
IMAGE_RESOLUTION_SCALE = 5.0


def create_model_dict(*args, **kwargs):
    # Marker is only used as a fallback when docling parsing is too sparse.
    # Import it lazily so its environment-sensitive settings do not block
    # SlideGen startup on machines with different local env defaults.
    if os.environ.get("DEBUG") == "release":
        os.environ["DEBUG"] = "false"

    try:
        from marker.models import create_model_dict as _create_model_dict
    except Exception:
        from marker.models import load_all_models as _create_model_dict

    return _create_model_dict(*args, **kwargs)


def _import_parse_pdf():
    from utils.src.model_utils import parse_pdf

    return parse_pdf


def _import_camel_runtime():
    from camel.models import ModelFactory
    from camel.agents import ChatAgent

    return ModelFactory, ChatAgent


def _import_wei_utils_helpers():
    from utils.wei_utils import account_token, chat_via_vllm, get_agent_config, openai_chat_text

    return account_token, chat_via_vllm, get_agent_config, openai_chat_text


def _import_torch():
    import torch

    return torch

def _import_docling():
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions
    from docling.document_converter import DocumentConverter, PdfFormatOption

    return InputFormat, PdfPipelineOptions, DocumentConverter, PdfFormatOption


def _import_docling_core_types():
    from docling_core.types.doc import ImageRefMode, TableItem
    from docling_core.types.doc.document import BoundingBox

    return ImageRefMode, TableItem, BoundingBox


def build_converter():
    InputFormat, PdfPipelineOptions, DocumentConverter, PdfFormatOption = _import_docling()
    opts = PdfPipelineOptions()
    opts.images_scale = IMAGE_RESOLUTION_SCALE
    opts.generate_page_images = True
    opts.generate_picture_images = True

    conv = DocumentConverter(
        format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)}
    )
    return conv



 
import fitz  # PyMuPDF
from PIL import Image as PILImage
from pathlib import Path
import json

def _page_size_from_doc(doc, page_no: int):
    """
    返回 (width, height)。doc.pages 是 1-based 字典。
    """
    pages = getattr(doc, "pages", {}) or {}
    page = pages.get(page_no)
    if page is None:
        return None, None
    size = getattr(page, "size", None)
    if size is None:
        return None, None
    return getattr(size, "width", None), getattr(size, "height", None)


def _get_prov_items(el):
    prov = getattr(el, "prov", None) or getattr(el, "provenance", None)
    if prov is None:
        return []
    if isinstance(prov, (list, tuple)):
        return list(prov)
    return [prov]


def _get_page_no_from_el(el):
    prov_items = _get_prov_items(el)
    if prov_items:
        page_no = getattr(prov_items[0], "page_no", None)
        if page_no is not None:
            return page_no
    return getattr(el, "page_no", None)


def _get_formula_text(el) -> str:
    for attr_name in ("latex", "text", "content", "orig"):
        value = getattr(el, attr_name, None)
        if value is None:
            continue
        value = str(value).strip()
        if value:
            return value
    return ""


def _looks_like_formula_screenshot(
    formula_text: str,
    crop_width: int | None,
    crop_height: int | None,
    page_width: int | None,
    page_height: int | None,
) -> bool:
    if not all(isinstance(v, (int, float)) and v > 0 for v in (crop_width, crop_height, page_width, page_height)):
        return False

    width_ratio = float(crop_width) / float(page_width)
    height_ratio = float(crop_height) / float(page_height)
    area_ratio = float(crop_width * crop_height) / float(page_width * page_height)
    aspect = float(crop_width) / float(crop_height) if crop_height else None

    normalized_text = " ".join(str(formula_text or "").split())
    text_len = len(normalized_text)
    alpha_chars = sum(ch.isalpha() for ch in normalized_text)
    digit_chars = sum(ch.isdigit() for ch in normalized_text)
    symbol_chars = sum(ch in "=+-/*^_()[]{}<>|∏∑λ∀∈≤≥≈×÷" for ch in normalized_text)
    natural_language_heavy = alpha_chars > max(80, symbol_chars * 3)

    if area_ratio >= 0.18:
        return True
    if height_ratio >= 0.34 and width_ratio >= 0.18:
        return True
    if height_ratio >= 0.28 and text_len >= 220:
        return True
    if height_ratio >= 0.22 and natural_language_heavy and text_len >= 140:
        return True
    if area_ratio >= 0.10 and natural_language_heavy and text_len >= 180:
        return True
    if aspect is not None and aspect < 0.9 and text_len >= 120:
        return True
    if digit_chars < 3 and symbol_chars < 5 and natural_language_heavy and text_len >= 120:
        return True
    return False

def _doc_bbox_bottomleft_to_xyxy(bbox: dict, page_h: float):
    
    l = float(bbox["l"]); r = float(bbox["r"])
    t = float(bbox["t"]); b = float(bbox["b"])
    # BOTTOMLEFT -> TOPLEFT:y_top = page_h - y_bottom
    y0 = page_h - b
    y1 = page_h - t
    x0, x1 = l, r
     
    if x1 < x0: x0, x1 = x1, x0
    if y1 < y0: y0, y1 = y1, y0
    return (x0, y0, x1, y1)

def export_formula_crops_from_texts(args,raw_result ):
     
    # doc_converter = build_converter() 
    # conv_res = doc_converter.convert(args.paper_path)
    conv_res=raw_result
    doc = conv_res.document
     
    pdf = fitz.open(str(args.paper_path))
    out_root = paper_image_tables_dir(args)
    out_root.mkdir(parents=True, exist_ok=True)
    out_json = formulas_json_path(args)

    formulas = {}
    skipped_formula_screenshot_like = 0
    idx = 1

    for el in getattr(doc, "texts", []):
        if str(getattr(el, "label", "")).lower() != "formula":
            continue
        text = _get_formula_text(el)
        prov = _get_prov_items(el)
        if not text or not prov:
            continue

        pno = getattr(prov[0], "page_no", None)
        bb = getattr(prov[0], "bbox", None)
         
        if bb is None:
            continue
        if not isinstance(bb, dict):
            # 对象 -> dict
            bb = {
                "l": getattr(bb, "l", None),
                "t": getattr(bb, "t", None),
                "r": getattr(bb, "r", None),
                "b": getattr(bb, "b", None),
                "coord_origin": str(getattr(bb, "coord_origin", "BOTTOMLEFT")),
            }
        if None in (bb.get("l"), bb.get("t"), bb.get("r"), bb.get("b")):
            continue

         
        w, h = _page_size_from_doc(doc, int(pno))
        if h is None:
             
            try:
                page = pdf[(pno - 1)]
                rect = page.rect
                w, h = float(rect.width), float(rect.height)
            except Exception:
                continue

        x0, y0, x1, y1 = _doc_bbox_bottomleft_to_xyxy(bb, page_h=h)
 
     
        scale = IMAGE_RESOLUTION_SCALE
 
         
        out_png = formula_crop_path(args, idx)
        try:
            page = pdf[(pno - 1)]
            pm = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=fitz.Rect(x0, y0, x1, y1))
            pm.save(str(out_png))
        except Exception as e:
            print(f"[Warn] crop failed at idx={idx}: {e}")
            idx += 1
            continue

       
        width = height = size = aspect = None
        try:
            im = PILImage.open(out_png)
            width, height = im.width, im.height
            size = width * height
            aspect = width / height if height else None
        except Exception:
            pass

        page = doc.pages.get(int(pno)) if getattr(doc, "pages", None) else None
        page_img = getattr(getattr(page, "image", None), "pil_image", None)
        page_img_w = getattr(page_img, "width", None)
        page_img_h = getattr(page_img, "height", None)

        if _looks_like_formula_screenshot(
            formula_text=text,
            crop_width=width,
            crop_height=height,
            page_width=page_img_w,
            page_height=page_img_h,
        ):
            skipped_formula_screenshot_like += 1
            print(
                f"[Formulas] Skipping screenshot-like formula crop idx={idx} page={pno} "
                f"crop={width}x{height} page={page_img_w}x{page_img_h}"
            )
            idx += 1
            continue

        formulas[str(idx)] = {
            "text": text,
            "page_no": int(pno),
            "bbox_doc": {k: float(v) if isinstance(v, (int, float)) else v for k, v in bb.items()},  # 原始 l/t/r/b
            "clip_rect_xyxy": [float(x0), float(y0), float(x1), float(y1)],  # 转换后的裁剪框
            "formula_path": str(out_png),
            "width": width, "height": height,
            "figure_size": size, "figure_aspect": aspect,
            "container_attr": "texts", "method": "crop"
        }
        idx += 1

    pdf.close()
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(formulas, f, ensure_ascii=False, indent=2)

    print(f"[Formulas] JSON: {out_json}")
    print(f"[Formulas] PNG dir: {out_root}")
    print(f"[Formulas] total: {len(formulas)}")
    if skipped_formula_screenshot_like:
        print(f"[Formulas] skipped screenshot-like crops: {skipped_formula_screenshot_like}")
    return formulas,conv_res

from pathlib import Path
import re, json

 

def export_formula_sections_grouped_json_from_texts(args, conv_res, max_page_no_exclusive: int = 12):
 
    doc = conv_res.document
    out_root = paper_image_tables_dir(args)
    out_root.mkdir(parents=True, exist_ok=True)
    out_json = formula_sections_path(args)
    debug_json = out_root / f"{asset_paper_name(args)}_formula_debug.json"

    def _get_page_no(el):
        return _get_page_no_from_el(el)

    def _get_bbox_t(el):
        prov = _get_prov_items(el)
        bb = getattr(prov[0], "bbox", None) if prov else None
        if isinstance(bb, dict):
            return bb.get("t", None)
        if bb is not None:
            return getattr(bb, "t", None)
        return None

    def _norm(s): return (s or "").strip()

    def _is_text_label(label: str) -> bool:
        label = (label or "").lower()
        return label in {"paragraph", "text", "list_item", "bullet", "caption"}

    def _clean_latex(text: str) -> str:
        text = text.strip()
        m = re.fullmatch(r"\${1,2}\s*(.+?)\s*\${1,2}", text, flags=re.S) \
            or re.fullmatch(r"\\\((.+?)\\\)", text, flags=re.S) \
            or re.fullmatch(r"\\\[(.+?)\\\]", text, flags=re.S)
        return m.group(1).strip() if m else text

    _re_heading_num = re.compile(r"^\s*(\d+(?:\.\d+)*)\b")

    def _is_section_header(txt, label: str) -> bool:
        if not txt: return False
        lab = (label or "").lower()
        if lab == "page_footer":
            return False
        if txt.strip().isdigit():
            return False
        if _re_heading_num.match(txt) and re.search(r"[A-Za-z]", txt):
            return True
        return ("section" in lab and "header" in lab)

    def _heading_level(title: str) -> int:
        m = _re_heading_num.match(title or "")
        return 1 + m.group(1).count(".") if m else 99

    # -------- linearize all content --------
    label_counts: dict[str, int] = {}
    skipped_after_page_cap = 0
    formula_missing_text_count = 0
    formula_missing_page_count = 0
    candidate_formula_examples: list[dict[str, Any]] = []
    dropped_formula_examples: list[dict[str, Any]] = []
    kept_text_examples: list[dict[str, Any]] = []
    skipped_garbage_examples: list[dict[str, Any]] = []
    header_examples: list[dict[str, Any]] = []
    linear = []
    for el in getattr(doc, "texts", []):
        label = str(getattr(el, "label", "")).lower()
        label_counts[label] = label_counts.get(label, 0) + 1
        page_no = _get_page_no(el)
        if page_no is None:
            if label == "formula":
                formula_missing_page_count += 1
                if len(dropped_formula_examples) < 25:
                    dropped_formula_examples.append({
                        "reason": "missing_page_no",
                        "text": _get_formula_text(el)[:300],
                    })
            continue
        if page_no >= max_page_no_exclusive:
            if page_no is not None and page_no >= max_page_no_exclusive:
                skipped_after_page_cap += 1
            continue
        y_top = _get_bbox_t(el)
        y_top = float(y_top) if y_top is not None else -1e9
        text = _norm(_get_formula_text(el) if label == "formula" else (getattr(el, "text", "") or ""))
        if not text:
            if label == "formula":
                formula_missing_text_count += 1
                if len(dropped_formula_examples) < 25:
                    dropped_formula_examples.append({
                        "reason": "missing_formula_text",
                        "page_no": page_no,
                    })
            continue

        if label == "formula":
            if len(candidate_formula_examples) < 25:
                candidate_formula_examples.append({
                    "page_no": page_no,
                    "label": label,
                    "text": text[:300],
                })
            linear.append({
                "kind": "formula",
                "latex": _clean_latex(text),
                "page_no": page_no,
                "y_top": y_top
            })
        elif _is_text_label(label):
            if len(kept_text_examples) < 25:
                kept_text_examples.append({
                    "page_no": page_no,
                    "label": label,
                    "text": text[:300],
                })
            linear.append({
                "kind": "text",
                "content": text,
                "page_no": page_no,
                "y_top": y_top
            })
        elif _is_section_header(text, label):
            if len(header_examples) < 25:
                header_examples.append({
                    "page_no": page_no,
                    "label": label,
                    "title": text[:300],
                })
            linear.append({
                "kind": "header",
                "title": text,
                "level": _heading_level(text),
                "page_no": page_no,
                "y_top": y_top
            })

    # -------- sort by page + y_top (PDF = bottom-left origin) --------
    linear.sort(key=lambda x: (x["page_no"], -x["y_top"]))

    # -------- section partitioning --------
    sections = []
    cur_section = {
        "section_title": None,
        "section_number": None,
        "section_pages": set(),
        "content": []
    }

    def is_garbage_text(text: str) -> bool:
        text = text.strip()
        if not text:
            return True
        if len(text) <= 3:
            return True  
        if re.fullmatch(r"[\s\.,;:()\[\]<>*/\\\"'=+\-~!?^|]{0,2}\s*[A-Za-z0-9]\s*", text):
            return True 
        if re.fullmatch(r"[A-Za-z0-9+/=\\<>|*_\"'.-]{1,8}", text):
            return True 
        if any(substr in text for substr in ["<latexi", "64=", "base64", ">A", "kS+Q", "Kr/", "\\", "†"]):
            return True
        return False

    def _flush_section(): 
        if any(x["type"] == "formula" for x in cur_section["content"]):
            out = {
                "section_title": cur_section["section_title"],
                "section_number": cur_section["section_number"],
                "section_pages": sorted(cur_section["section_pages"]),
                "content": cur_section["content"]
            }
            sections.append(out)


    for item in linear:
        if item["kind"] == "header": 
            _flush_section()
            cur_section = {
                "section_title": item["title"],
                "section_number": _re_heading_num.match(item["title"]).group(1) if _re_heading_num.match(item["title"]) else None,
                "section_pages": set(),
                "content": []
            }
        elif item["kind"] in {"text", "formula"}:
            if item["kind"] == "text":
                if is_garbage_text(item["content"]):
                    if len(skipped_garbage_examples) < 25:
                        skipped_garbage_examples.append({
                            "page_no": item["page_no"],
                            "text": str(item["content"])[:300],
                        })
                    continue
            cur_section["content"].append({
                "type": item["kind"],
                **({"content": item["content"]} if item["kind"] == "text" else {"latex": item["latex"], "page_no": item["page_no"]})
            })
            cur_section["section_pages"].add(item["page_no"])

    _flush_section()

    # -------- write output --------
    out_json.parent.mkdir(parents=True, exist_ok=True)
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(sections, f, ensure_ascii=False, indent=2)

    debug_payload = {
        "paper_name": asset_paper_name(args),
        "max_page_no_exclusive": max_page_no_exclusive,
        "raw_text_element_count": len(getattr(doc, "texts", [])),
        "label_counts": label_counts,
        "skipped_after_page_cap": skipped_after_page_cap,
        "formula_missing_page_count": formula_missing_page_count,
        "formula_missing_text_count": formula_missing_text_count,
        "linear_item_count": len(linear),
        "linear_formula_count": sum(1 for item in linear if item["kind"] == "formula"),
        "linear_text_count": sum(1 for item in linear if item["kind"] == "text"),
        "linear_header_count": sum(1 for item in linear if item["kind"] == "header"),
        "output_section_count": len(sections),
        "output_formula_count": sum(
            1
            for section in sections
            for content_item in list(section.get("content") or [])
            if isinstance(content_item, dict) and content_item.get("type") == "formula"
        ),
        "candidate_formula_examples": candidate_formula_examples,
        "dropped_formula_examples": dropped_formula_examples,
        "header_examples": header_examples,
        "kept_text_examples": kept_text_examples,
        "skipped_garbage_examples": skipped_garbage_examples,
    }
    with open(debug_json, "w", encoding="utf-8") as f:
        json.dump(debug_payload, f, ensure_ascii=False, indent=2)

    print(f"[LinearSections] Saved to {out_json}")
    print(f"[LinearSections] Debug saved to {debug_json}")
    print(f"[LinearSections] Sections: {len(sections)}")
    return sections
 
@retry(stop=stop_after_attempt(1))
def parse_raw(args, actor_config, version=1):
    raw_source = args.paper_path
    markdown_clean_pattern = re.compile(r"<!--[\s\S]*?-->")

    print(f"[parse_raw] Starting docling convert for {raw_source}", flush=True)
    raw_result = build_converter().convert(raw_source)
    print("[parse_raw] Docling convert finished", flush=True)
    input_token, output_token =0,0
    
    print("[parse_raw] Exporting markdown from docling result", flush=True)
    raw_markdown = raw_result.document.export_to_markdown()
    print("[parse_raw] Markdown export finished", flush=True)
    text_content = markdown_clean_pattern.sub("", raw_markdown)
    start_time = time.time()
    if len(text_content) < 500:
        print('\nParsing with docling failed, using marker instead\n')
        torch = _import_torch()
        parse_pdf = _import_parse_pdf()
        parser_model = create_model_dict(device='cuda', dtype=torch.float16)
        text_content, rendered = parse_pdf(raw_source, model_lst=parser_model, save_file=False)

    if version == 1:
        template = Template(open("utils/prompts/gen_poster_raw_content.txt").read())
    elif version == 2:
        template = Template(open("utils/prompts/gen_slides_raw_content_v2.txt").read())
    author_preference_profile = None
    if getattr(args, "use_author_preferences", False):
        profile_path = getattr(args, "author_profile_path", None)
        if profile_path and Path(profile_path).exists():
            print(f"[parse_raw] Loading author profile from {profile_path}", flush=True)
            author_preference_profile = json.loads(Path(profile_path).read_text(encoding="utf-8"))
    use_gpt5_responses = False

    actor_sys_msg = 'You are the author of the paper, and you will create an academic presentation (slides) to explain the paper'
    account_token, chat_via_vllm, _get_agent_config, openai_chat_text = _import_wei_utils_helpers()
 
    if "gpt-5" in args.model_name_t.lower():  
        client = build_openai_client()
        use_gpt5_responses = True
    
    actor_model = None
    actor_agent = None
    if not use_gpt5_responses:
        ModelFactory, ChatAgent = _import_camel_runtime()
        if "qwen" in str(args.model_name_t).lower():
            print("model_type=actor_config['model_type']: ", actor_config['model_type'])
            actor_model = ModelFactory.create(
                model_platform=actor_config['model_platform'],
                model_type=actor_config['model_type'],
                model_config_dict=actor_config['model_config'],
                url=actor_config['url'],
            ) 
        else:
            actor_model = ModelFactory.create(
                model_platform=actor_config['model_platform'],
                model_type=actor_config['model_type'],
                model_config_dict=actor_config['model_config'],
            )

        actor_agent = ChatAgent(
            system_message=actor_sys_msg,
            model=actor_model,
            message_window_size=10,
            token_limit=actor_config.get('token_limit', None)
        )

    while True:
        outline_mode = getattr(args, "outline_mode", "high_level")
        print(f"[parse_raw] Building prompt (outline_mode={outline_mode})", flush=True)
        prompt = template.render(
            markdown_document=text_content,
            outline_mode=outline_mode,
            use_author_preferences=getattr(args, "use_author_preferences", False),
            author_preference_profile_json=author_preference_profile,
        )
        print(
            f"[parse_raw] Sending outline request to model={args.model_name_t}",
            flush=True,
        )
        if use_gpt5_responses: 
            raw_output, input_token, output_token = openai_chat_text(
                client=client,
                model=resolve_direct_model_name(args.model_name_t),
                user_prompt=prompt,
                system_prompt=actor_sys_msg,
                prefer_responses=True,
            )
            print("[parse_raw] Outline response received", flush=True)
            
        else:
            if "qwen" in str(args.model_name_t).lower():
                response = chat_via_vllm(prompt,actor_config,actor_model,actor_sys_msg)
                raw_output = response.choices[0].message.content
                print("raw_output by qwen : ")
                print(raw_output) 
                input_token = response.usage.prompt_tokens
                output_token = response.usage.completion_tokens
                print("input_token: ",input_token)
            else:
                actor_agent.reset()
                response = actor_agent.step(prompt)
                input_token, output_token = account_token(response)
                raw_output = response.msgs[0].content
                print("[parse_raw] Outline response received", flush=True)


        content_json = get_json_from_response(raw_output)

        if len(content_json) > 0:
            break
        print('[parse_raw] Error: Empty response, retrying...', flush=True)
        if "qwen" in str(args.model_name_t).lower():
            text_content = text_content[:80000]


    print(type(content_json))
    print("content_json",content_json)
    # if len(content_json['sections']) > 9:
    #     # First 2 sections + randomly select 5 sections + last 2 sections
    #     selected_sections = content_json['sections'][:2] + random.sample(content_json['sections'][2:-2], 5) + content_json['sections'][-2:]
    #     content_json['sections'] = selected_sections

    has_title = False

    for section in content_json['sections']:
   
        if 'title' in section['title'].lower():
            has_title = True

    # if not has_title:
    #     print('Ouch! The response is invalid, the LLM is not following the format :(')
    #     raise
    end_time = time.time()
    time_taken = end_time - start_time
    output_dir = paper_output_dir(args)
    output_dir.mkdir(parents=True, exist_ok=True)

    outline_path = raw_content_path(args)
    outline_path.parent.mkdir(parents=True, exist_ok=True)
    print(f"[parse_raw] Writing raw content to {outline_path}", flush=True)
    with open(outline_path, "w", encoding="utf-8") as handle:
        json.dump(content_json, handle, indent=4)
    print("[parse_raw] Raw parsing stage complete", flush=True)
    return input_token, output_token, time_taken, raw_result

from pprint import pprint

def safe_print_element_fields(element):
    print(f"[Type] {type(element)}")
    safe_dict = {}
    for k, v in vars(element).items():
        if isinstance(v, (str, int, float, tuple, list, dict, type(None))):
            safe_dict[k] = v
        else:
            safe_dict[k] = f"<{type(v).__name__}>"
    pprint(safe_dict)
 
def convert_bbox_to_pil_coords(bbox, page_width, page_height, pad=10):
    """
    将 BOTTOMLEFT 坐标的 bbox 转为 PIL 坐标（左上角原点） 
    """
    x0 = bbox.l
    x1 = bbox.r 
    y0 = page_height - bbox.b  # PIL y0: corresponds to the bottom of the PDF bbox
    y1 = page_height - bbox.t  # PIL y1: corresponds to the top of the PDF bbox

    # Add padding (and clip to image boundaries)
    return (
        max(0, x0 - pad),
        max(0, y1 - pad),
        min(page_width, x1 + pad),
        min(page_height, y0 + pad),
    )
  
def gen_image_and_table(args, conv_res):
    ImageRefMode, TableItem, BoundingBox = _import_docling_core_types()
    input_token, output_token = 0, 0
    raw_source = args.paper_path

    output_dir = paper_image_tables_dir(args)

    output_dir.mkdir(parents=True, exist_ok=True)
    doc_filename = args.paper_name

    # Save page images
    for page_no, page in conv_res.document.pages.items():
        page_no = page.page_no
        page_image_filename = page_image_path(args, page_no)
        with page_image_filename.open("wb") as fp:
            page.image.pil_image.save(fp, format="PNG")
 
    # 修改裁剪框（扩大上下左右）  
    from PIL import Image
    
    for i, picture in enumerate(conv_res.document.pictures):
        page_no = picture.prov[0].page_no
        page = conv_res.document.pages[page_no]  
        full_img = page.image.pil_image   
        page_size_in_points = page.size  #  PDF 中的尺寸（单位 pt）
        page_width_pt, page_height_pt = page_size_in_points.width, page_size_in_points.height

        scale = full_img.width / page_width_pt
         
        bbox = picture.prov[0].bbox
        pad = 1   
        padded_bbox = BoundingBox(
            l= bbox.l - pad ,
            r=  bbox.r + pad ,
            b=   bbox.b - pad ,
            t=   bbox.t + pad ,
            coord_origin=bbox.coord_origin,
        )
        tl_bbox = padded_bbox.to_top_left_origin(page_height=page_height_pt)
        pil_box = tl_bbox.scaled(scale=scale).as_tuple() 
        left, top, right, bottom = pil_box
        cropped = full_img.crop((left, top, right, bottom))  
        cropped.save(str(picture_image_path(args, i + 1)))
    
    table_counter = 0 
    for element, _level in conv_res.document.iterate_items():
        if isinstance(element, TableItem):
            table_counter += 1
            element_image_filename = (
                table_image_path(args, table_counter)
            )
            with element_image_filename.open("wb") as fp:
                element.get_image(conv_res.document).save(fp, "PNG")
   

    # These exports are useful for inspection, but they are not required for the
    # downstream slide-generation pipeline. On some synced macOS folders, PIL/Docling
    # can time out while saving referenced image assets, so we treat them as best effort.
    try:
        md_filename = markdown_embedded_path(args)
        conv_res.document.save_as_markdown(md_filename, image_mode=ImageRefMode.EMBEDDED)
    except (TimeoutError, OSError) as exc:
        print(f"[warning] Skipping embedded markdown export: {exc}")

    try:
        md_filename = markdown_referenced_path(args)
        conv_res.document.save_as_markdown(md_filename, image_mode=ImageRefMode.REFERENCED)
    except (TimeoutError, OSError) as exc:
        print(f"[warning] Skipping referenced markdown export: {exc}")

    try:
        html_filename = html_referenced_path(args)
        conv_res.document.save_as_html(html_filename, image_mode=ImageRefMode.REFERENCED)
    except (TimeoutError, OSError) as exc:
        print(f"[warning] Skipping HTML export: {exc}")

    tables = {}

    table_index = 1
    for table in conv_res.document.tables:
        caption = table.caption_text(conv_res.document)
        if len(caption) > 0:
            table_img_path = table_image_path(args, table_index)
            table_img = PIL.Image.open(table_img_path)
            tables[str(table_index)] = {
                'caption': caption,
                'page_no': table.prov[0].page_no,
                'table_path': str(table_img_path),
                'width': table_img.width,
                'height': table_img.height,
                'figure_size': table_img.width * table_img.height,
                'figure_aspect': table_img.width / table_img.height,
            }

        table_index += 1

    images = {}
    image_index = 1
    for image in conv_res.document.pictures:
        caption = image.caption_text(conv_res.document) 
        print(f"[{i}] caption: {caption}")
        if len(caption) > 0:
            image_img_path = picture_image_path(args, image_index)
            image_img = PIL.Image.open(image_img_path)
            images[str(image_index)] = {
                'caption': caption,
                'page_no': image.prov[0].page_no,
                'image_path': str(image_img_path),
                'width': image_img.width,
                'height': image_img.height,
                'figure_size': image_img.width * image_img.height,
                'figure_aspect': image_img.width / image_img.height,
            }
        image_index += 1

    images_path = images_json_path(args)
    tables_path = tables_json_path(args)
    images_path.parent.mkdir(parents=True, exist_ok=True)
    tables_path.parent.mkdir(parents=True, exist_ok=True)
    json.dump(images, open(images_path, 'w'), indent=4)
    json.dump(tables, open(tables_path, 'w'), indent=4)

    return input_token, output_token, images, tables

def append_outline_mode_suffix(paper_name: str, outline_mode: str) -> str:
    base = paper_name.strip().replace(" ", "_")
    if base.endswith("_high_level") or base.endswith("_technical"):
        return base
    return f"{base}_{outline_mode}"


def infer_output_key_from_paper_path(paper_path: str) -> str | None:
    path = Path(paper_path)
    parts = list(path.parts)

    if len(parts) >= 3:
        record_id = parts[-2].strip()
        split = parts[-3].strip()
        if record_id.isdigit() and split:
            split_key = "".join(ch if ch.isalnum() or ch in ("_", "-") else "_" for ch in split)
            return f"{split_key}_{record_id}"

    stem = path.stem.replace(" ", "_")
    if stem.lower() == "paper" and len(parts) >= 2:
        record_id = parts[-2].strip()
        if record_id:
            return record_id
    return stem or None

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--paper_name', type=str, default=None)
    parser.add_argument('--model_name', type=str, default=None)
    parser.add_argument('--model_name_t', type=str, default=None)
    parser.add_argument('--model_name_v', type=str, default=None)
    parser.add_argument('--paper_path', type=str, required=True)
    parser.add_argument('--index', type=int, default=0)
    parser.add_argument('--version', type=int, choices=[1, 2], default=2)
    parser.add_argument('--outline_only', action='store_true')
    parser.add_argument(
        '--outline_mode',
        choices=['high_level', 'technical'],
        default='high_level',
        help='Use high_level for a compact presentation narrative, or technical to preserve major paper subsections.',
    )
    args = parser.parse_args()

    if args.model_name_t is None:
        args.model_name_t = args.model_name or '4o'
    if args.model_name_v is None:
        args.model_name_v = args.model_name_t

    _account_token, _chat_via_vllm, get_agent_config, _openai_chat_text = _import_wei_utils_helpers()
    agent_config = get_agent_config(args.model_name_t)

    if args.paper_name is None:
        paper_name = infer_output_key_from_paper_path(args.paper_path)
        args.paper_name = append_outline_mode_suffix(paper_name, args.outline_mode)
    else:
        args.paper_name = append_outline_mode_suffix(args.paper_name, args.outline_mode)

    # Parse raw content
    input_token, output_token, _, raw_result = parse_raw(args, agent_config, version=args.version)

    if not args.outline_only:
        # Generate images and tables
        _, _, _, _ = gen_image_and_table(args, raw_result)

    print(f'Token consumption: {input_token} -> {output_token}')
