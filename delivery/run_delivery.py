"""Canonical business entrypoint for thesis delivery iteration.

Run from anywhere:

    python E:/My Project/thesis-format-checker/delivery/run_delivery.py

This is the only user-facing command for producing the next thesis DOCX/PDF
delivery. Implementation details live under delivery/.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"


def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    os.environ["PYTHONIOENCODING"] = "utf-8"
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))

    from delivery import build_lgp_docx

    build_lgp_docx.main()


if __name__ == "__main__":
    main()
