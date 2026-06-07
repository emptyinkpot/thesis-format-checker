"""Single full-test entrypoint for thesis-format-checker.

Run from anywhere:

    python E:/My Project/thesis-format-checker/run_full_tests.py

What it covers:
- Python syntax/bytecode compilation
- pytest unit tests when pytest is installed
- v012 DOCX regeneration through the canonical formatter
- NCWU checker pass on v012
- color-consistency regression check against v011
- v012 visual audit and blank-scan sanity checks
"""

from __future__ import annotations

import importlib.util
import contextlib
import io
import json
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SRC_DIR = ROOT / "src"
DOWNLOADS = Path(r"C:/Users/ASUS-KL/Downloads")
V011 = DOWNLOADS / "202213210刘高朋修改迭代版_v011_格式统一交付版.docx"
V012 = DOWNLOADS / "202213210刘高朋修改迭代版_v012_全篇黑色字体统一版.docx"
V012_BLANK_REPORT = DOWNLOADS / "202213210刘高朋修改迭代版_v012_留白扫描.json"


class SkipStep(RuntimeError):
    pass


@dataclass
class StepResult:
    name: str
    status: str
    detail: str = ""


def run_command(args: list[str], *, cwd: Path = ROOT) -> None:
    subprocess.run(args, cwd=str(cwd), check=True)


def step_compileall() -> str:
    run_command([
        sys.executable,
        "-m",
        "compileall",
        str(ROOT / "src" / "thesis_format_checker"),
        str(ROOT / "format_lgp_v012.py"),
        str(ROOT / "run_full_tests.py"),
    ])
    return "compileall passed"


def step_pytest() -> str:
    if importlib.util.find_spec("pytest") is None:
        raise SkipStep("pytest is not installed")
    run_command([sys.executable, "-m", "pytest"], cwd=ROOT)
    return "pytest passed"


def step_regenerate_v012() -> str:
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    import format_lgp_v012

    output = io.StringIO()
    try:
        with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
            format_lgp_v012.main()
    except Exception:
        print(output.getvalue())
        raise
    if not V012.exists():
        raise RuntimeError(f"v012 DOCX missing after generation: {V012}")
    return "generated v012 DOCX/PDF/report/blank-scan"


def step_check_v012() -> str:
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    from thesis_format_checker.checker import check, load_preset

    preset = load_preset("ncwu")
    _docx, _content, findings = check(V012, preset)
    if findings:
        detail = "; ".join(f"{f.rule_id}: {f.message}" for f in findings[:5])
        raise RuntimeError(f"v012 checker findings={len(findings)} {detail}")
    return "v012 checker findings=0"


def step_regression_v011_color_rule() -> str:
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    from thesis_format_checker.checker import check, load_preset

    preset = load_preset("ncwu")
    _docx, _content, findings = check(V011, preset)
    color_findings = [f for f in findings if f.rule_id == "text-color-consistency"]
    if not color_findings:
        raise RuntimeError("v011 no longer triggers text-color-consistency regression warning")
    return "v011 triggers text-color-consistency as expected"


def step_visual_audit() -> str:
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    import format_lgp_v012

    audit = format_lgp_v012.audit_visual_format()
    if audit["non_black_runs"] or audit["style_non_black"]:
        raise RuntimeError(f"v012 color audit failed: {audit}")
    return f"color audit passed: {audit}"


def step_blank_scan_sanity() -> str:
    if not V012_BLANK_REPORT.exists():
        raise RuntimeError(f"blank scan report missing: {V012_BLANK_REPORT}")
    suspects = json.loads(V012_BLANK_REPORT.read_text(encoding="utf-8"))
    pages = {item.get("page") for item in suspects}
    if 90 in pages:
        raise RuntimeError("reference orphan tail page returned at page 90")
    if len(suspects) > 7:
        raise RuntimeError(f"blank suspects increased: {len(suspects)}")
    return f"blank scan sanity passed: suspects={len(suspects)} pages={sorted(pages)}"


def run_step(name: str, fn) -> StepResult:
    print(f"\n== {name} ==")
    try:
        detail = fn()
        print(f"PASS {detail}")
        return StepResult(name, "PASS", detail)
    except SkipStep as exc:
        print(f"SKIP {exc}")
        return StepResult(name, "SKIP", str(exc))
    except Exception as exc:
        print(f"FAIL {exc}")
        return StepResult(name, "FAIL", str(exc))


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    os.environ["PYTHONIOENCODING"] = "utf-8"
    os.chdir(ROOT)

    steps = [
        ("compileall", step_compileall),
        ("pytest", step_pytest),
        ("regenerate-v012", step_regenerate_v012),
        ("check-v012", step_check_v012),
        ("regression-v011-color-rule", step_regression_v011_color_rule),
        ("visual-audit", step_visual_audit),
        ("blank-scan-sanity", step_blank_scan_sanity),
    ]
    results = [run_step(name, fn) for name, fn in steps]

    print("\n== summary ==")
    for result in results:
        print(f"{result.status:4} {result.name} {result.detail}")

    failed = [result for result in results if result.status == "FAIL"]
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
