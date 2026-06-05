"""Rule definitions and evaluation engine.

Each rule is a function decorated with @rule() that takes
(docx_result, content_result, preset) and returns findings.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Callable


@dataclass
class Finding:
    rule_id: str
    message: str
    severity: str = "error"  # error | warning | info
    expected: Any = None
    actual: Any = None
    location: str = ""
    fixable: bool = False


@dataclass
class RuleSpec:
    id: str
    fn: Callable
    default_severity: str = "error"
    default_enabled: bool = True


RULES: dict[str, RuleSpec] = {}


def rule(id: str, default_severity: str = "error", default_enabled: bool = True):
    def decorator(fn):
        RULES[id] = RuleSpec(id=id, fn=fn, default_severity=default_severity, default_enabled=default_enabled)
        return fn
    return decorator


def _get_rule_config(preset: dict, rule_id: str) -> tuple[str, bool]:
    """Get severity and enabled status for a rule from preset overrides."""
    for r in preset.get("rules", []):
        if r.get("id") == rule_id:
            return r.get("severity", "error"), r.get("enabled", True)
    return "error", True


def evaluate_all(docx_result, content_result, preset: dict) -> list[Finding]:
    """Run all enabled rules and collect findings."""
    findings = []
    for rule_id, spec in RULES.items():
        severity, enabled = _get_rule_config(preset, rule_id)
        if not enabled:
            continue
        results = spec.fn(docx_result, content_result, preset)
        for f in results:
            f.severity = severity
        findings.extend(results)
    return findings


# --- Rule implementations ---

@rule("page-margins")
def check_page_margins(docx, content, preset) -> list[Finding]:
    findings = []
    expected = preset.get("page", {}).get("margin_cm", {})
    if not expected:
        return findings
    for sec in docx.sections:
        for side in ("top", "bottom", "left", "right"):
            exp = expected.get(side)
            if exp is None:
                continue
            actual = getattr(sec, f"{side}_margin_cm")
            if abs(actual - exp) > 0.15:
                findings.append(Finding(
                    rule_id="page-margins",
                    message=f"Section {sec.index}: {side} margin {actual:.2f}cm, expected {exp}cm",
                    expected=exp, actual=round(actual, 2),
                    location=f"Section {sec.index}",
                    fixable=True,
                ))
    return findings


@rule("header-text-match", default_severity="warning")
def check_header_text(docx, content, preset) -> list[Finding]:
    expected = preset.get("page", {}).get("header_text", "")
    if not expected:
        return []
    findings = []
    for sec in docx.sections:
        if sec.header_text and expected not in sec.header_text:
            findings.append(Finding(
                rule_id="header-text-match",
                message=f"Section {sec.index} 页眉为 {sec.header_text!r}，应包含 {expected!r}",
                expected=expected, actual=sec.header_text,
                location=f"Section {sec.index}", fixable=True,
            ))
    return findings


@rule("header-on-all-sections", default_severity="warning")
def check_header_on_all_sections(docx, content, preset) -> list[Finding]:
    findings = []
    has_any_header = any(s.header_text for s in docx.sections)
    if not has_any_header:
        return findings
    for sec in docx.sections:
        if not sec.header_text and not sec.header_linked_to_previous:
            findings.append(Finding(
                rule_id="header-on-all-sections",
                message=f"Section {sec.index} 缺少页眉且未链接到前一节",
                location=f"Section {sec.index}", fixable=False,
            ))
    return findings


@rule("body-font-size")
def check_body_font_size(docx, content, preset) -> list[Finding]:
    expected = preset.get("styles", {}).get("body", {}).get("size_pt")
    if expected is None:
        return []
    normal = docx.styles.get("Normal")
    if normal is None:
        return []
    actual = normal.size_pt
    if actual is not None and abs(actual - expected) > 0.1:
        return [Finding(
            rule_id="body-font-size",
            message=f"正文字号应为 {expected}pt（小四），Normal 样式实际 {actual}pt",
            expected=expected, actual=actual,
            location="Style: Normal", fixable=True,
        )]
    return []


@rule("body-east-asia-font")
def check_body_east_asia_font(docx, content, preset) -> list[Finding]:
    expected = preset.get("styles", {}).get("body", {}).get("east_asia")
    if not expected:
        return []
    normal = docx.styles.get("Normal")
    if normal is None or normal.east_asia is None:
        return []
    if normal.east_asia != expected:
        return [Finding(
            rule_id="body-east-asia-font",
            message=f"正文中文字体应为 {expected}，Normal 样式实际 {normal.east_asia}",
            expected=expected, actual=normal.east_asia,
            location="Style: Normal", fixable=True,
        )]
    return []


@rule("body-line-spacing", default_severity="warning")
def check_body_line_spacing(docx, content, preset) -> list[Finding]:
    expected = preset.get("styles", {}).get("body", {}).get("line_spacing")
    if expected is None:
        return []
    normal = docx.styles.get("Normal")
    if normal is None or normal.line_spacing is None:
        return []
    if abs(normal.line_spacing - expected) > 0.05:
        return [Finding(
            rule_id="body-line-spacing",
            message=f"正文行距应为 {expected}，Normal 样式实际 {normal.line_spacing}",
            expected=expected, actual=normal.line_spacing,
            location="Style: Normal", fixable=True,
        )]
    return []


def _check_heading(docx, preset, level: int, style_name: str) -> list[Finding]:
    cfg = preset.get("styles", {}).get(f"heading{level}")
    if not cfg:
        return []
    style = docx.styles.get(style_name)
    if style is None:
        return [Finding(
            rule_id=f"heading{level}-style",
            message=f"样式 {style_name} 不存在",
            location=f"Style: {style_name}", fixable=False,
        )]
    findings = []
    if cfg.get("east_asia") and style.east_asia and style.east_asia != cfg["east_asia"]:
        findings.append(Finding(
            rule_id=f"heading{level}-style",
            message=f"{style_name} 中文字体应为 {cfg['east_asia']}，实际 {style.east_asia}",
            expected=cfg["east_asia"], actual=style.east_asia,
            location=f"Style: {style_name}", fixable=True,
        ))
    if cfg.get("size_pt") and style.size_pt is not None and abs(style.size_pt - cfg["size_pt"]) > 0.1:
        findings.append(Finding(
            rule_id=f"heading{level}-style",
            message=f"{style_name} 字号应为 {cfg['size_pt']}pt，实际 {style.size_pt}pt",
            expected=cfg["size_pt"], actual=style.size_pt,
            location=f"Style: {style_name}", fixable=True,
        ))
    return findings


@rule("heading1-style")
def check_heading1(docx, content, preset) -> list[Finding]:
    return _check_heading(docx, preset, 1, "Heading 1")


@rule("heading2-style")
def check_heading2(docx, content, preset) -> list[Finding]:
    return _check_heading(docx, preset, 2, "Heading 2")


@rule("heading3-style", default_severity="warning")
def check_heading3(docx, content, preset) -> list[Finding]:
    return _check_heading(docx, preset, 3, "Heading 3")


CHAPTER_TEXT_RE = re.compile(r"^第\d+章\s+\S+")
SECTION_TEXT_RE = re.compile(r"^\d+\.\d+\s+\S+")
SUBSECTION_TEXT_RE = re.compile(r"^\d+\.\d+\.\d+\s+\S+")


@rule("heading-style-applied")
def check_heading_style_applied(docx, content, preset) -> list[Finding]:
    findings = []
    for p in docx.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if CHAPTER_TEXT_RE.match(text) and p.style_name != "Heading 1":
            findings.append(Finding(
                rule_id="heading-style-applied",
                message=f"段落 {p.index} 章标题 {text[:30]!r} 应使用 Heading 1，实际 {p.style_name}",
                expected="Heading 1", actual=p.style_name,
                location=f"Paragraph {p.index}", fixable=True,
            ))
        elif SUBSECTION_TEXT_RE.match(text) and p.style_name != "Heading 3":
            findings.append(Finding(
                rule_id="heading-style-applied",
                message=f"段落 {p.index} 三级标题 {text[:30]!r} 应使用 Heading 3，实际 {p.style_name}",
                expected="Heading 3", actual=p.style_name,
                location=f"Paragraph {p.index}", fixable=True,
            ))
        elif SECTION_TEXT_RE.match(text) and p.style_name != "Heading 2" and not SUBSECTION_TEXT_RE.match(text):
            findings.append(Finding(
                rule_id="heading-style-applied",
                message=f"段落 {p.index} 二级标题 {text[:30]!r} 应使用 Heading 2，实际 {p.style_name}",
                expected="Heading 2", actual=p.style_name,
                location=f"Paragraph {p.index}", fixable=True,
            ))
    return findings


@rule("chapter-page-break", default_severity="warning")
def check_chapter_page_break(docx, content, preset) -> list[Finding]:
    if not preset.get("structure", {}).get("chapter_page_break", True):
        return []
    findings = []
    chapter_paras = [p for p in docx.paragraphs if CHAPTER_TEXT_RE.match(p.text.strip())]
    for i, p in enumerate(chapter_paras):
        if i == 0:
            continue
        if not (p.page_break_before or p.has_page_break_run):
            prev_idx = p.index - 1
            prev_has_break = False
            if prev_idx >= 0:
                prev = docx.paragraphs[prev_idx]
                prev_has_break = prev.has_page_break_run
            if not prev_has_break:
                findings.append(Finding(
                    rule_id="chapter-page-break",
                    message=f"章 {p.text.strip()[:20]!r} 前缺少分页符",
                    location=f"Paragraph {p.index}", fixable=False,
                ))
    return findings


@rule("abstract-zh-length")
def check_abstract_zh_length(docx, content, preset) -> list[Finding]:
    min_chars = preset.get("content", {}).get("abstract_zh", {}).get("min_chars", 0)
    if min_chars == 0:
        return []
    if content.abstract_zh_chars < min_chars:
        return [Finding(
            rule_id="abstract-zh-length",
            message=f"中文摘要 {content.abstract_zh_chars} 个汉字，应不少于 {min_chars}",
            expected=min_chars, actual=content.abstract_zh_chars,
            location="中文摘要", fixable=False,
        )]
    return []


@rule("abstract-en-length", default_severity="warning")
def check_abstract_en_length(docx, content, preset) -> list[Finding]:
    min_words = preset.get("content", {}).get("abstract_en", {}).get("min_words", 0)
    if min_words == 0:
        return []
    if content.abstract_en_words < min_words:
        return [Finding(
            rule_id="abstract-en-length",
            message=f"英文摘要 {content.abstract_en_words} 词，应不少于 {min_words}",
            expected=min_words, actual=content.abstract_en_words,
            location="ABSTRACT", fixable=False,
        )]
    return []


@rule("foreign-translation-length")
def check_foreign_translation_length(docx, content, preset) -> list[Finding]:
    min_chars = preset.get("content", {}).get("foreign_translation", {}).get("min_chars", 0)
    if min_chars == 0:
        return []
    if content.foreign_translation_chars < min_chars:
        return [Finding(
            rule_id="foreign-translation-length",
            message=f"外文译文 {content.foreign_translation_chars} 个汉字，应不少于 {min_chars}",
            expected=min_chars, actual=content.foreign_translation_chars,
            location="附录二 外文译文", fixable=False,
        )]
    return []


@rule("toc-present")
def check_toc_present(docx, content, preset) -> list[Finding]:
    if not preset.get("content", {}).get("toc", {}).get("required", True):
        return []
    if not content.toc_present:
        return [Finding(
            rule_id="toc-present",
            message="未检测到目录页",
            location="目录", fixable=False,
        )]
    return []


@rule("cover-fields")
def check_cover_fields(docx, content, preset) -> list[Finding]:
    required = preset.get("content", {}).get("cover_fields", {}).get("required", [])
    if not required:
        return []
    missing = [f for f in required if f not in content.cover_fields_found]
    if missing:
        return [Finding(
            rule_id="cover-fields",
            message=f"封面缺少字段: {'、'.join(missing)}",
            expected=required, actual=content.cover_fields_found,
            location="封面", fixable=False,
        )]
    return []
