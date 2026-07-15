#!/usr/bin/env python3
"""Generate clean non-personalized baselines for a batch manifest."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parents[1]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from Capstone.generate_random_decks import append_outline_mode_suffix, output_dir_key, resolve_cli_path


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def clean_baseline_paper_name(item: dict[str, Any]) -> str:
    paper_id = str(item["paper_id"])
    paper_path = Path(str(item["paper_path"]))
    return f"{output_dir_key(paper_id, paper_path)}_clean_baseline"


def clean_baseline_contents_name(item: dict[str, Any], outline_mode: str) -> str:
    return append_outline_mode_suffix(clean_baseline_paper_name(item), outline_mode)


def baseline_plan_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = clean_baseline_contents_name(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"<{model_name_t}_{model_name_v}>_slide_plan_baseline.json"


def baseline_pptx_path(*, item: dict[str, Any], outline_mode: str, model_name_t: str, model_name_v: str) -> Path:
    contents_name = clean_baseline_contents_name(item, outline_mode)
    return PROJECT_ROOT / "contents" / contents_name / f"{model_name_t}_{model_name_v}_output_slides_baseline.pptx"


def build_generation_command(
    *,
    item: dict[str, Any],
    model_name_t: str,
    model_name_v: str,
    formula_mode: int,
    outline_mode: str,
) -> list[str]:
    return [
        sys.executable,
        "-m",
        "SlidesAgent.new_pipeline",
        "--paper_path",
        str(item["paper_path"]),
        "--paper_name",
        clean_baseline_paper_name(item),
        "--model_name_t",
        model_name_t,
        "--model_name_v",
        model_name_v,
        "--formula_mode",
        str(formula_mode),
        "--outline_mode",
        outline_mode,
    ]


def run_command(command: list[str], *, cwd: Path, dry_run: bool) -> None:
    rendered = " ".join(command)
    print(rendered)
    if dry_run:
        return
    subprocess.run(command, cwd=cwd, check=True)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate clean non-personalized baseline decks for a manifest.")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=PROJECT_ROOT / "Capstone" / "batch_runs" / "manifest.json",
        help="Manifest JSON describing the selected papers for the batch run.",
    )
    parser.add_argument("--include-existing", action="store_true", help="Regenerate even if clean baseline outputs already exist.")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    args.manifest = resolve_cli_path(args.manifest)
    if not args.manifest.exists():
        raise SystemExit(f"Manifest not found: {args.manifest}")

    manifest = load_json(args.manifest)
    selected = list(manifest.get("selected_papers") or [])
    if not selected:
        raise SystemExit(f"No selected_papers found in manifest: {args.manifest}")

    outline_mode = str(manifest.get("outline_mode") or "high_level")
    model_name_t = str(manifest.get("model_name_t") or "gpt-5.4-nano")
    model_name_v = str(manifest.get("model_name_v") or "gpt-5.4-nano")
    formula_mode = int(manifest.get("formula_mode") or 1)

    summary: dict[str, Any] = {
        "manifest_path": str(args.manifest),
        "outline_mode": outline_mode,
        "model_name_t": model_name_t,
        "model_name_v": model_name_v,
        "papers": [],
    }

    for index, item in enumerate(selected, start=1):
        paper_id = str(item["paper_id"])
        output_plan = baseline_plan_path(
            item=item,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        )
        output_pptx = baseline_pptx_path(
            item=item,
            outline_mode=outline_mode,
            model_name_t=model_name_t,
            model_name_v=model_name_v,
        )

        row = {
            "paper_id": paper_id,
            "paper_path": str(item["paper_path"]),
            "paper_name": clean_baseline_paper_name(item),
            "output_plan": str(output_plan),
            "output_pptx": str(output_pptx),
            "status": "pending",
        }

        print(f"[{index}/{len(selected)}] {paper_id}")
        if output_plan.exists() and output_pptx.exists() and not args.include_existing and not args.dry_run:
            row["status"] = "skipped_existing"
        else:
            command = build_generation_command(
                item=item,
                model_name_t=model_name_t,
                model_name_v=model_name_v,
                formula_mode=formula_mode,
                outline_mode=outline_mode,
            )
            run_command(command, cwd=PROJECT_ROOT, dry_run=args.dry_run)
            row["status"] = "generated" if not args.dry_run else "dry_run"

        summary["papers"].append(row)

    summary_path = args.manifest.parent / "clean_baseline_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"summary_path": str(summary_path), "paper_count": len(summary["papers"])}, indent=2))


if __name__ == "__main__":
    main()
