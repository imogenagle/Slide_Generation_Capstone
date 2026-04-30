#!/usr/bin/env python3
"""Probe import/startup hot spots with per-step timing and timeouts."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def run_step(name: str, code: str, timeout: int) -> dict:
    cmd = [sys.executable, "-c", code]
    start = time.time()
    try:
        proc = subprocess.run(
            cmd,
            cwd=ROOT,
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        elapsed = round(time.time() - start, 3)
        return {
            "name": name,
            "status": "ok" if proc.returncode == 0 else "error",
            "seconds": elapsed,
            "returncode": proc.returncode,
            "stdout": proc.stdout.strip(),
            "stderr": proc.stderr.strip(),
        }
    except subprocess.TimeoutExpired as exc:
        elapsed = round(time.time() - start, 3)
        return {
            "name": name,
            "status": "timeout",
            "seconds": elapsed,
            "returncode": None,
            "stdout": (exc.stdout or "").strip() if exc.stdout else "",
            "stderr": (exc.stderr or "").strip() if exc.stderr else "",
        }


def main() -> int:
    parser = argparse.ArgumentParser(description="Probe SlideGen startup stages")
    parser.add_argument("--timeout", type=int, default=20, help="Per-step timeout in seconds")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON")
    args = parser.parse_args()

    steps = [
        (
            "python_identity",
            "import sys; print(sys.executable); print(sys.version.split()[0]); print(sys.prefix)",
        ),
        (
            "import_utils_wei_utils",
            "import utils.wei_utils; print('imported utils.wei_utils')",
        ),
        (
            "import_parse_raw",
            "import SlidesAgent.parse_raw; print('imported SlidesAgent.parse_raw')",
        ),
        (
            "import_new_pipeline_logtime",
            "import SlidesAgent.new_pipeline_logtime; print('imported SlidesAgent.new_pipeline_logtime')",
        ),
        (
            "import_docling_modules",
            (
                "from docling.datamodel.base_models import InputFormat; "
                "from docling.datamodel.pipeline_options import PdfPipelineOptions; "
                "from docling.document_converter import DocumentConverter, PdfFormatOption; "
                "print('imported docling modules')"
            ),
        ),
        (
            "construct_pdf_pipeline_options",
            (
                "from docling.datamodel.pipeline_options import PdfPipelineOptions; "
                "PdfPipelineOptions(); "
                "print('constructed PdfPipelineOptions')"
            ),
        ),
        (
            "construct_docling_converter_minimal",
            (
                "from docling.datamodel.base_models import InputFormat; "
                "from docling.datamodel.pipeline_options import PdfPipelineOptions; "
                "from docling.document_converter import DocumentConverter, PdfFormatOption; "
                "opts = PdfPipelineOptions(); "
                "DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=opts)}); "
                "print('constructed minimal DocumentConverter')"
            ),
        ),
        (
            "construct_parse_raw_converter",
            "from SlidesAgent.parse_raw import build_converter; build_converter(); print('constructed parse_raw converter')",
        ),
    ]

    results = [run_step(name, code, args.timeout) for name, code in steps]

    if args.json:
        print(json.dumps(results, indent=2))
        return 0

    print(f"Startup probe using {sys.executable}")
    for result in results:
        print(f"[{result['status']}] {result['name']} in {result['seconds']}s")
        if result["stdout"]:
            print(f"stdout: {result['stdout']}")
        if result["stderr"]:
            print(f"stderr: {result['stderr']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
