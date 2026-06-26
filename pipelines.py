#!/usr/bin/env python3
"""
Google Colab entry point for the HOYER + SAMAT rate pipeline.

Usage in Colab (after uploading/cloning code to /content/Ecolab-SAMAT-HOYER):

    from google.colab import drive
    drive.mount("/content/drive")

    import os
    exec(open("/content/Ecolab-SAMAT-HOYER/pipelines.py").read())
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

CODE_DIR = Path("/content/Ecolab-SAMAT-HOYER")
if str(CODE_DIR) not in sys.path:
    sys.path.insert(0, str(CODE_DIR))

subprocess.run(
    [sys.executable, "-m", "pip", "install", "-q", "pandas>=2.0.0", "openpyxl>=3.1.0"],
    check=False,
)

from project_paths import (  # noqa: E402
    INPUT_DIR,
    OUTPUT_DIR,
    PROCESSING_DIR,
    ensure_workspace_dirs,
    is_colab_environment,
)
from pipeline import print_summary, run_pipeline  # noqa: E402

if not is_colab_environment():
    print("Warning: Colab paths were not detected; using configured project_paths.")

ensure_workspace_dirs()

print("HOYER + SAMAT pipeline")
print(f"  Input:      {INPUT_DIR}")
print(f"  Processing: {PROCESSING_DIR}")
print(f"  Output:     {OUTPUT_DIR}")

result = run_pipeline()
print_summary(result)
