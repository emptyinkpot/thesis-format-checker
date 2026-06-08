"""Single full-test entrypoint for thesis-format-checker.

Run from anywhere:

    python E:/My Project/thesis-format-checker/scripts/run_full_tests.py

What it covers:
- Python syntax/bytecode compilation
- built-in unit contracts for rules and preset loading
- v012 DOCX regeneration through the canonical formatter
- NCWU checker pass on v012
- color-consistency regression check against v011
- v012 visual audit and blank-scan sanity checks
"""

from __future__ import annotations

import contextlib
import io
import json
import os
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = ROOT / "src"
DOWNLOADS = Path(r"C:/Users/ASUS-KL/Downloads")
ORIGINAL = DOWNLOADS / "202213210刘高朋修改迭代版.docx"
V011 = DOWNLOADS / "202213210刘高朋修改迭代版_v011_格式统一交付版.docx"
V012 = DOWNLOADS / "202213210刘高朋修改迭代版_v012_全篇黑色字体统一版.docx"
V012_PDF = DOWNLOADS / "202213210刘高朋修改迭代版_v012_全篇黑色字体统一版.pdf"
V012_REPORT = DOWNLOADS / "202213210刘高朋修改迭代版_v012_格式检测报告.md"
V012_BLANK_REPORT = DOWNLOADS / "202213210刘高朋修改迭代版_v012_留白扫描.json"
VERSION_LOG = DOWNLOADS / "202213210刘高朋修改迭代版_版本记录.md"
EXPECTED_HEADER = "华北水利水电大学毕业设计"
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

FORBIDDEN_TERMS = [
    "代码截图",
    "本实验",
    "实验",
    "本研究",
    "本论文",
    "软件算法",
    "算法",
    "预测",
    "补偿",
    "图3.6",
    "7.2 不足与展望",
    "2.2.1 系统需求分析",
    "文献启发",
    "本文按照",
]

ALLOWED_BLANK_PAGES = {2, 3, 23, 52, 53, 69, 76}
ALLOWED_EAST_ASIA_FONTS = {"宋体", "黑体", "仿宋_GB2312", "隶书", "Consolas"}
ALLOWED_LATIN_FONTS = {"Times New Roman", "Consolas", "宋体"}


@dataclass
class StepResult:
    name: str
    status: str
    detail: str = ""


def run_command(args: list[str], *, cwd: Path = ROOT) -> None:
    subprocess.run(args, cwd=str(cwd), check=True)


def document_text(path: Path) -> str:
    doc = Document(str(path))
    parts: list[str] = [paragraph.text for paragraph in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                parts.extend(paragraph.text for paragraph in cell.paragraphs)
    return "\n".join(parts)


def iter_direct_run_fonts(path: Path):
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            root = ET.fromstring(archive.read(name))
            for run in root.findall(f".//{W_NS}r"):
                text = "".join(node.text or "" for node in run.findall(f".//{W_NS}t")).strip()
                if not text:
                    continue
                rpr = run.find(f"{W_NS}rPr")
                if rpr is None:
                    continue
                fonts = rpr.find(f"{W_NS}rFonts")
                if fonts is None:
                    continue
                yield {
                    "source": name,
                    "text": text[:80],
                    "east_asia": fonts.get(f"{W_NS}eastAsia"),
                    "ascii": fonts.get(f"{W_NS}ascii"),
                    "hansi": fonts.get(f"{W_NS}hAnsi"),
                }


def step_compileall() -> str:
    run_command([
        sys.executable,
        "-m",
        "compileall",
        str(ROOT / "src" / "thesis_format_checker"),
        str(ROOT / "format_lgp_v012.py"),
        str(ROOT / "scripts" / "run_full_tests.py"),
    ])
    return "compileall passed"


def step_unit_contracts() -> str:
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    from thesis_format_checker.rules import RULES
    from thesis_format_checker.checker import load_preset

    expected_rules = {
        "page-margins", "header-text-match", "header-on-all-sections",
        "body-font-size", "body-east-asia-font", "text-color-consistency",
        "body-line-spacing", "heading1-style", "heading2-style",
        "heading3-style", "heading-style-applied", "chapter-page-break",
        "abstract-zh-length", "abstract-en-length", "foreign-translation-length",
        "toc-present", "cover-fields",
    }
    missing = expected_rules - set(RULES.keys())
    if missing:
        raise RuntimeError(f"missing rule registrations: {sorted(missing)}")

    preset = load_preset("ncwu")
    if preset["preset_id"] != "ncwu":
        raise RuntimeError(f"unexpected preset_id: {preset['preset_id']}")
    if preset["page"]["margin_cm"]["left"] != 3.0:
        raise RuntimeError("unexpected left margin in preset")
    if preset["styles"]["body"]["size_pt"] != 12:
        raise RuntimeError("unexpected body font size in preset")
    if preset["content"]["abstract_zh"]["min_chars"] != 500:
        raise RuntimeError("unexpected zh abstract threshold in preset")

    return "rule registry and preset contracts passed"


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


def step_delivery_contract() -> str:
    required = [ORIGINAL, V011, V012, V012_PDF, V012_REPORT, V012_BLANK_REPORT, VERSION_LOG]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise RuntimeError(f"missing delivery artifacts: {missing}")
    if V012.suffix.lower() != ".docx":
        raise RuntimeError(f"final artifact is not a DOCX: {V012}")
    if "_v012_" not in V012.name:
        raise RuntimeError(f"final DOCX is not versioned as v012: {V012.name}")
    if ORIGINAL.name == V012.name:
        raise RuntimeError("final DOCX overwrote the original filename")

    log_text = VERSION_LOG.read_text(encoding="utf-8")
    if V012.name not in log_text or "可疑页 7 页" not in log_text:
        raise RuntimeError("version log is missing current v012 delivery facts")
    return "v012 DOCX/PDF/report/version-log artifacts present"


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


def step_content_integrity_contract() -> str:
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    from thesis_format_checker.checker import check, load_preset

    preset = load_preset("ncwu")
    _docx, content, _findings = check(V012, preset)
    if content.abstract_zh_chars < 500:
        raise RuntimeError(f"zh abstract too short: {content.abstract_zh_chars}")
    if content.abstract_en_words < 300:
        raise RuntimeError(f"en abstract too short: {content.abstract_en_words}")
    if content.foreign_translation_chars < 2000:
        raise RuntimeError(f"foreign translation too short: {content.foreign_translation_chars}")

    expected_chapters = {f"第{i}章" for i in range(1, 8)}
    chapter_text = "\n".join(content.chapters)
    missing_chapters = [chapter for chapter in sorted(expected_chapters) if chapter not in chapter_text]
    if missing_chapters:
        raise RuntimeError(f"missing chapters: {missing_chapters}")

    required_markers = ["摘 要", "ABSTRACT", "参考文献", "附录一", "附录二", "附录三"]
    missing_markers = [marker for marker in required_markers if marker not in content.full_text]
    if missing_markers:
        raise RuntimeError(f"missing structural markers: {missing_markers}")
    return "abstracts/translation/chapters/appendices present"


def step_header_contract() -> str:
    doc = Document(str(V012))
    headers = []
    for section in doc.sections:
        text = " ".join(paragraph.text.strip() for paragraph in section.header.paragraphs if paragraph.text.strip()).strip()
        if text:
            headers.append(text)
    if not headers:
        raise RuntimeError("no body headers found")
    bad_headers = [header for header in headers if header != EXPECTED_HEADER]
    if bad_headers:
        raise RuntimeError(f"unexpected headers: {bad_headers}")
    if any(("(" in header or ")" in header or "（" in header or "）" in header or "论文" in header) for header in headers):
        raise RuntimeError(f"header still contains parentheses or thesis suffix: {headers}")
    return f"headers exact: {EXPECTED_HEADER}"


def step_font_contract() -> str:
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    from thesis_format_checker.docx_inspector import inspect

    docx = inspect(V012)
    style_expectations = {
        "Normal": ("宋体", "Times New Roman"),
        "Body Text": ("宋体", "Times New Roman"),
    }
    for style_name, (expected_ea, expected_latin) in style_expectations.items():
        style = docx.styles.get(style_name)
        if style is None:
            raise RuntimeError(f"missing style: {style_name}")
        if style.east_asia and style.east_asia != expected_ea:
            raise RuntimeError(f"{style_name} eastAsia={style.east_asia}, expected {expected_ea}")
        if style.latin and style.latin != expected_latin:
            raise RuntimeError(f"{style_name} latin={style.latin}, expected {expected_latin}")

    bad_fonts = []
    for item in iter_direct_run_fonts(V012):
        ea = item["east_asia"]
        ascii_font = item["ascii"]
        hansi = item["hansi"]
        if ea and ea not in ALLOWED_EAST_ASIA_FONTS:
            bad_fonts.append(item)
        if ascii_font and ascii_font not in ALLOWED_LATIN_FONTS:
            bad_fonts.append(item)
        if hansi and hansi not in ALLOWED_LATIN_FONTS:
            bad_fonts.append(item)
        if len(bad_fonts) >= 5:
            break
    if bad_fonts:
        raise RuntimeError(f"unexpected direct fonts: {bad_fonts}")
    return "style and direct font families are within allowed thesis set"


def step_forbidden_terms_contract() -> str:
    text = document_text(V012)
    hits = {term: text.count(term) for term in FORBIDDEN_TERMS if text.count(term)}
    if hits:
        raise RuntimeError(f"forbidden terms present: {hits}")
    return "forbidden prose terms absent"


def step_image_table_contract() -> str:
    before = Document(str(V011))
    after = Document(str(V012))
    if len(after.inline_shapes) < len(before.inline_shapes):
        raise RuntimeError(f"inline image count decreased: v011={len(before.inline_shapes)} v012={len(after.inline_shapes)}")
    if len(after.tables) < len(before.tables):
        raise RuntimeError(f"table count decreased: v011={len(before.tables)} v012={len(after.tables)}")
    return f"images/tables preserved: images={len(after.inline_shapes)}, tables={len(after.tables)}"


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
    unexpected = pages - ALLOWED_BLANK_PAGES
    if unexpected:
        raise RuntimeError(f"unexpected blank-scan pages: {sorted(unexpected)}")
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
        ("unit-contracts", step_unit_contracts),
        ("regenerate-v012", step_regenerate_v012),
        ("delivery-contract", step_delivery_contract),
        ("check-v012", step_check_v012),
        ("content-integrity", step_content_integrity_contract),
        ("header-contract", step_header_contract),
        ("font-contract", step_font_contract),
        ("forbidden-terms", step_forbidden_terms_contract),
        ("image-table-contract", step_image_table_contract),
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
