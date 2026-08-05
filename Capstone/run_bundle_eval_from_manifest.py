#!/usr/bin/env python3
"""Run bundle eval for decks listed in a saved manifest and build comparison tables."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[1]
PYTHON = REPO_ROOT / ".venv" / "bin" / "python3"

if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from Capstone.compare_bundle_eval_methods import ALL_METRICS  # noqa: E402


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_python() -> Path:
    return PYTHON if PYTHON.exists() else Path(sys.executable)


def sanitize_name(value: str) -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in ("_", "-", ".") else "_" for ch in value.strip())
    return cleaned.strip("_") or "bundle_eval"


def run_command(command: list[str], *, dry_run: bool) -> None:
    print("Running:", " ".join(command), flush=True)
    if dry_run:
        return
    subprocess.run(command, cwd=REPO_ROOT, check=True)


def slidegen_baseline_pptx(slidegen_root: Path, paper_key: str, model_t: str, model_v: str) -> Path:
    return (
        slidegen_root
        / "contents"
        / f"{paper_key}_high_level"
        / f"{model_t}_{model_v}_output_slides_baseline.pptx"
    )


def slidegen_personalized_pptx(slidegen_root: Path, paper_key: str, model_t: str, model_v: str) -> Path:
    return (
        slidegen_root
        / "contents"
        / f"{paper_key}_high_level_personalized_retrieval"
        / f"{model_t}_{model_v}_output_slides_personalized_retrieval.pptx"
    )


def original_pptx(original_root: Path, paper_key: str, model_t: str, model_v: str) -> Path:
    return (
        original_root
        / "contents"
        / f"{paper_key}_original"
        / f"{model_t}_{model_v}_output_slides.pptx"
    )


def build_bundle_eval_command(
    *,
    python_bin: Path,
    generated_pptx: Path,
    paper_id: str,
    paper_path: Path,
    output_dir: Path,
    judge_model: str,
    core_coverage_model: str,
    render_dpi: int,
    metrics: list[str],
) -> list[str]:
    cmd = [
        str(python_bin),
        "Capstone/evaluate_pptx_bundle.py",
        "--generated-pptx",
        str(generated_pptx),
        "--paper-id",
        paper_id,
        "--source-document",
        str(paper_path),
        "--output-dir",
        str(output_dir),
        "--judge-model",
        judge_model,
        "--core-coverage-model",
        core_coverage_model,
        "--render-dpi",
        str(render_dpi),
    ]

    selected = set(metrics)
    if "core_coverage_topic_iou" not in selected:
        cmd.append("--skip-core-coverage")
    if "geometry_aware_density_gad_geom" not in selected:
        cmd.append("--skip-gad")
    if "visual_appeal_deck_score" not in selected:
        cmd.append("--skip-visual-appeal")
    if "layout_defects_deck_score" not in selected:
        cmd.append("--skip-layout-correctness")
    if "logical_flow_deck_score" not in selected:
        cmd.append("--skip-logical-flow")
    if "paper_faithfulness_deck_score" not in selected:
        cmd.append("--skip-faithfulness")
    return cmd


def build_compare_command(
    *,
    python_bin: Path,
    eval_root: Path,
    output_dir: Path,
    metrics: list[str],
) -> list[str]:
    cmd = [
        str(python_bin),
        "Capstone/compare_bundle_eval_methods.py",
        "--method",
        f"SlideGen_Original={eval_root / 'original'}",
        "--method",
        f"SlideGen_Baseline={eval_root / 'baseline'}",
        "--method",
        f"SlideGen_Personalized={eval_root / 'personalized'}",
        "--output-dir",
        str(output_dir),
    ]
    if metrics:
        cmd.extend(["--metrics", *metrics])
    return cmd


def main() -> None:
    parser = argparse.ArgumentParser(description="Run bundle eval for manifest-listed original/baseline/personalized decks.")
    parser.add_argument("--manifest", type=Path, required=True, help="Path to manifest.json with selected_papers.")
    parser.add_argument("--slidegen-root", type=Path, required=True, help="Root containing SlideGen contents/ outputs.")
    parser.add_argument("--original-root", type=Path, required=True, help="Root containing SlideGen_Original contents/ outputs.")
    parser.add_argument("--output-dir", type=Path, required=True, help="Where to save per-paper bundle eval outputs.")
    parser.add_argument("--summary-dir", type=Path, required=True, help="Where to save comparison tables.")
    parser.add_argument("--model-name-t", default="gpt-5.4-nano")
    parser.add_argument("--model-name-v", default="gpt-5.4-nano")
    parser.add_argument("--judge-model", default="gpt-5.4-nano")
    parser.add_argument("--core-coverage-model", default="gpt-5.4-nano")
    parser.add_argument("--render-dpi", type=int, default=120)
    parser.add_argument(
        "--metrics",
        nargs="+",
        choices=sorted(set(ALL_METRICS.keys()) | {"layout_defects_deck_score"}),
        default=[
            "core_coverage_topic_iou",
            "geometry_aware_density_gad_geom",
            "visual_appeal_deck_score",
            "logical_flow_deck_score",
            "paper_faithfulness_deck_score",
            "layout_defects_deck_score",
        ],
        help="Subset of bundle metrics to run and summarize.",
    )
    parser.add_argument("--skip-existing", action="store_true", help="Skip per-paper eval if summary.json already exists.")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    python_bin = resolve_python()
    manifest = load_json(args.manifest)
    selected_papers = list(manifest.get("selected_papers") or [])
    if not selected_papers:
        raise SystemExit(f"No selected_papers found in {args.manifest}")

    failures: list[dict[str, Any]] = []

    for item in selected_papers:
        paper_id = str(item["paper_id"])
        paper_key = paper_id.replace(":", "_")
        paper_path = Path(str(item["paper_path"]))

        methods = {
            "original": original_pptx(args.original_root, paper_key, args.model_name_t, args.model_name_v),
            "baseline": slidegen_baseline_pptx(args.slidegen_root, paper_key, args.model_name_t, args.model_name_v),
            "personalized": slidegen_personalized_pptx(args.slidegen_root, paper_key, args.model_name_t, args.model_name_v),
        }

        for method_label, pptx_path in methods.items():
            out_dir = args.output_dir / method_label / paper_key
            summary_path = out_dir / "summary.json"
            if args.skip_existing and summary_path.exists():
                print(f"Skipping existing {method_label} bundle eval: {summary_path}", flush=True)
                continue
            if not pptx_path.exists():
                failures.append(
                    {
                        "paper_id": paper_id,
                        "method": method_label,
                        "error": f"Missing PPTX: {pptx_path}",
                    }
                )
                print(f"[skip] Missing {method_label} PPTX for {paper_id}: {pptx_path}", flush=True)
                continue

            try:
                run_command(
                    build_bundle_eval_command(
                        python_bin=python_bin,
                        generated_pptx=pptx_path,
                        paper_id=paper_id,
                        paper_path=paper_path,
                        output_dir=out_dir,
                        judge_model=args.judge_model,
                        core_coverage_model=args.core_coverage_model,
                        render_dpi=args.render_dpi,
                        metrics=args.metrics,
                    ),
                    dry_run=args.dry_run,
                )
            except subprocess.CalledProcessError as exc:
                failures.append(
                    {
                        "paper_id": paper_id,
                        "method": method_label,
                        "error": "CalledProcessError",
                        "returncode": exc.returncode,
                    }
                )

    run_command(
        build_compare_command(
            python_bin=python_bin,
            eval_root=args.output_dir,
            output_dir=args.summary_dir,
            metrics=[metric for metric in args.metrics if metric in ALL_METRICS],
        ),
        dry_run=args.dry_run,
    )

    failure_path = args.output_dir / "bundle_eval_failures.json"
    failure_path.parent.mkdir(parents=True, exist_ok=True)
    failure_path.write_text(json.dumps(failures, indent=2, ensure_ascii=False), encoding="utf-8")

    print(
        json.dumps(
            {
                "manifest": str(args.manifest),
                "papers": len(selected_papers),
                "output_dir": str(args.output_dir),
                "summary_dir": str(args.summary_dir),
                "failure_count": len(failures),
                "failures_path": str(failure_path),
            },
            indent=2,
        ),
        flush=True,
    )


if __name__ == "__main__":
    main()
