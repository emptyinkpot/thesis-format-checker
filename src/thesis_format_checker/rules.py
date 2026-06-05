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


# --- NEW RULES: from 9-thesis analysis (2026-06-05) ---


@rule("body-first-line-indent", default_severity="warning")
def check_body_first_line_indent(docx, content, preset) -> list[Finding]:
    expected_chars = preset.get("styles", {}).get("body", {}).get("first_line_indent_chars")
    if expected_chars is None:
        return []
    normal = docx.styles.get("Normal")
    if normal is None:
        return []
    # python-docx stores first_line_indent in EMU; 2 chars ≈ 480000 EMU (at 12pt)
    # We check style-level only; run-level is too noisy
    # For now just check if indent is defined at all via line_spacing proxy
    # TODO: enhance docx_inspector to extract first_line_indent from pPr
    return []


@rule("page-number-format", default_severity="warning")
def check_page_number_format(docx, content, preset) -> list[Finding]:
    """正文页码应为阿拉伯数字居中底部。"""
    expected = preset.get("page", {}).get("page_number", {})
    if not expected:
        return []
    # Check footer for page numbers in body sections
    findings = []
    body_format = expected.get("body_format", "arabic")
    body_position = expected.get("body_position", "center_bottom")
    # Inspect last section (body) for footer content
    if docx.sections:
        last_sec = docx.sections[-1]
        # We'd need footer inspection - for now flag if no footer text detected
        # TODO: enhance docx_inspector to extract footer paragraphs
    return findings


@rule("page-number-roman-frontmatter", default_severity="warning")
def check_page_number_roman_frontmatter(docx, content, preset) -> list[Finding]:
    """前置部分（摘要/目录）应使用罗马数字页码。"""
    expected = preset.get("page", {}).get("page_number", {})
    if not expected.get("frontmatter_roman", False):
        return []
    # Detect via section pgNumType in early sections
    findings = []
    # TODO: enhance docx_inspector to read pgNumType from sectPr
    return findings


@rule("header-underline", default_severity="warning")
def check_header_underline(docx, content, preset) -> list[Finding]:
    """页眉下方应有细横线。"""
    if not preset.get("page", {}).get("header_underline", False):
        return []
    # python-docx can detect paragraph border on header paragraph
    # TODO: implement via header paragraph bottom border check
    return []


@rule("header-frontmatter-excluded", default_severity="warning")
def check_header_frontmatter_excluded(docx, content, preset) -> list[Finding]:
    """摘要/目录页不应有页眉（仅正文起有）。"""
    if not preset.get("page", {}).get("header_frontmatter_excluded", True):
        return []
    findings = []
    # First sections (cover + abstract + toc) should NOT have header text
    # Last section (body) should have header text
    if len(docx.sections) >= 2:
        for sec in docx.sections[:-1]:
            if sec.header_text and not sec.header_linked_to_previous:
                findings.append(Finding(
                    rule_id="header-frontmatter-excluded",
                    message=f"Section {sec.index}（前置部分）不应有页眉，实际有 {sec.header_text!r}",
                    location=f"Section {sec.index}", fixable=False,
                ))
    return findings


@rule("abstract-page-title", default_severity="warning")
def check_abstract_page_title(docx, content, preset) -> list[Finding]:
    """摘要页顶部应有论文中文题目（居中）。"""
    if not preset.get("content", {}).get("abstract_zh", {}).get("title_above", False):
        return []
    # Check paragraphs before the "摘  要" heading for a title-like paragraph
    for p in docx.paragraphs[:50]:
        text = p.text.strip()
        if "摘" in text and "要" in text and len(text) < 10:
            # Found abstract title - check if preceding para has thesis title
            idx = p.index
            if idx > 0:
                prev = docx.paragraphs[idx - 1]
                if len(prev.text.strip()) > 5:
                    return []  # There's content above abstract heading - likely title
            return [Finding(
                rule_id="abstract-page-title",
                message="摘要页上方缺少论文中文题目",
                location="摘要页", fixable=False,
            )]
    return []


@rule("abstract-title-spacing", default_severity="info")
def check_abstract_title_spacing(docx, content, preset) -> list[Finding]:
    """'摘  要'标题字间应空2格。"""
    if not preset.get("content", {}).get("abstract_zh", {}).get("title_spacing", False):
        return []
    for p in docx.paragraphs[:50]:
        text = p.text.strip()
        if text in ("摘要", "摘 要"):
            return [Finding(
                rule_id="abstract-title-spacing",
                message=f"摘要标题 {text!r} 字间距不足，应为 '摘  要'（中间空2格）",
                expected="摘  要", actual=text,
                location=f"Paragraph {p.index}", fixable=True,
            )]
        if "摘" in text and "要" in text and len(text) <= 5:
            break
    return []


@rule("keywords-bold", default_severity="info")
def check_keywords_bold(docx, content, preset) -> list[Finding]:
    """'关键词'三字应加粗。"""
    if not preset.get("content", {}).get("abstract_zh", {}).get("keywords_bold", False):
        return []
    for p in docx.paragraphs[:80]:
        if "关键词" in p.text:
            if not p.first_run_bold:
                return [Finding(
                    rule_id="keywords-bold",
                    message="'关键词'三字应加粗",
                    location=f"Paragraph {p.index}", fixable=True,
                )]
            break
    return []


@rule("classification-number", default_severity="info")
def check_classification_number(docx, content, preset) -> list[Finding]:
    """摘要末尾应有中图分类号。"""
    if not preset.get("content", {}).get("abstract_zh", {}).get("classification_number", False):
        return []
    # Search in pandoc text around abstract area
    if "中图分类号" not in content.full_text[:10000]:
        return [Finding(
            rule_id="classification-number",
            message="摘要区域未发现中图分类号（应有 '中图分类号：TPxxx'）",
            location="中文摘要", fixable=False,
        )]
    return []


@rule("toc-title-spacing", default_severity="info")
def check_toc_title_spacing(docx, content, preset) -> list[Finding]:
    """'目  录'标题字间应双空格。"""
    if not preset.get("content", {}).get("toc", {}).get("title_spacing", False):
        return []
    for p in docx.paragraphs[:60]:
        text = p.text.strip()
        if text in ("目录", "目 录"):
            return [Finding(
                rule_id="toc-title-spacing",
                message=f"目录标题 {text!r} 字间距不足，应为 '目  录'（中间双空格）",
                expected="目  录", actual=text,
                location=f"Paragraph {p.index}", fixable=True,
            )]
        if "目" in text and "录" in text and len(text) <= 5:
            break
    return []


@rule("toc-dotline", default_severity="info")
def check_toc_dotline(docx, content, preset) -> list[Finding]:
    """目录条目应有点线连接页码。"""
    # This requires inspecting tab leaders in TOC paragraphs
    # python-docx can read tab stops with leader attributes
    # For now, presence check only via content
    if not preset.get("content", {}).get("toc", {}).get("dotline", False):
        return []
    # SDT TOC usually has tab leaders - hard to verify without deeper XML parse
    # TODO: implement via SDT body paragraph tab leader inspection
    return []


FIGURE_CAPTION_RE_RULE = re.compile(r"^图\s*(\d+)[\.\-](\d+)\s+")
TABLE_CAPTION_RE_RULE = re.compile(r"^表\s*(\d+)[\.\-](\d+)\s+")


@rule("figure-caption-position", default_severity="warning")
def check_figure_caption_position(docx, content, preset) -> list[Finding]:
    """图题应在图下方（即图片段落之后）。"""
    if not preset.get("structure", {}).get("figure_caption_below", True):
        return []
    # Check that figure captions come AFTER image paragraphs
    # A figure caption appearing before any image in a sequence suggests wrong position
    findings = []
    for i, p in enumerate(docx.paragraphs):
        if FIGURE_CAPTION_RE_RULE.match(p.text.strip()):
            # Check next paragraph - if it's an image, caption is ABOVE (wrong)
            if i + 1 < len(docx.paragraphs):
                next_p = docx.paragraphs[i + 1]
                # Image paragraphs often have empty text with inline shapes
                # This is a heuristic; real check needs inline shape detection
    return findings


@rule("table-caption-position", default_severity="warning")
def check_table_caption_position(docx, content, preset) -> list[Finding]:
    """表题应在表上方（即表格之前）。"""
    if not preset.get("structure", {}).get("table_caption_above", True):
        return []
    # Heuristic: table captions should appear before table elements
    # TODO: correlate table captions with actual table positions
    return []


@rule("figure-table-numbering", default_severity="warning")
def check_figure_table_numbering(docx, content, preset) -> list[Finding]:
    """图表编号应按章编号（图X.Y / 表X.Y 点号分隔）。"""
    numbering = preset.get("structure", {}).get("figure_table_numbering", "chapter_dot")
    if numbering != "chapter_dot":
        return []
    findings = []
    bad_patterns = []
    for p in docx.paragraphs:
        text = p.text.strip()
        # Check for dash-separated numbering like 图3-1 instead of 图3.1
        if re.match(r"^图\s*\d+\-\d+", text):
            bad_patterns.append((p.index, text[:20]))
        if re.match(r"^表\s*\d+\-\d+", text):
            bad_patterns.append((p.index, text[:20]))
        # Check for non-chapter numbering like 图1, 图2 (no dot)
        if re.match(r"^图\s*\d+\s+[^\.]", text) and not re.match(r"^图\s*\d+\.\d+", text):
            bad_patterns.append((p.index, text[:20]))
    if bad_patterns:
        for idx, txt in bad_patterns[:5]:
            findings.append(Finding(
                rule_id="figure-table-numbering",
                message=f"图表编号应为章号制点号分隔（图X.Y），实际 {txt!r}",
                expected="图X.Y / 表X.Y", actual=txt,
                location=f"Paragraph {idx}", fixable=False,
            ))
    return findings


@rule("reference-superscript", default_severity="warning")
def check_reference_superscript(docx, content, preset) -> list[Finding]:
    """正文引用应为上角标 [n] 格式。"""
    if not preset.get("structure", {}).get("reference_superscript", True):
        return []
    # Look for inline [n] references that are NOT superscript
    inline_ref_re = re.compile(r"\[\d+\]")
    non_super_refs = []
    for p in docx.paragraphs:
        text = p.text.strip()
        if not text or p.style_name in ("Heading 1", "Heading 2", "Heading 3"):
            continue
        # Skip paragraphs that ARE the reference list
        if text.startswith("[") and re.match(r"^\[\d+\]", text):
            continue
        matches = inline_ref_re.findall(text)
        if matches and p.first_run_size_pt and p.first_run_size_pt >= 10:
            # References in normal-sized text = likely not superscript
            non_super_refs.append((p.index, text[:40]))
    if len(non_super_refs) > 3:
        return [Finding(
            rule_id="reference-superscript",
            message=f"发现 {len(non_super_refs)} 处正文引用 [n] 未使用上角标格式",
            expected="上角标 [n]", actual=f"{len(non_super_refs)} 处内嵌引用",
            location="正文", fixable=False,
        )]
    return []


@rule("reference-format", default_severity="info")
def check_reference_format(docx, content, preset) -> list[Finding]:
    """参考文献列表应使用 [n] 顶格 GB/T 7714 格式。"""
    if not preset.get("structure", {}).get("reference_gbt7714", True):
        return []
    # Check reference section formatting
    in_refs = False
    bad_format = []
    for p in docx.paragraphs:
        text = p.text.strip()
        if "参考文献" in text and len(text) < 10:
            in_refs = True
            continue
        if in_refs:
            if not text:
                continue
            if re.match(r"^(附录|致\s*谢)", text):
                break
            # Each reference should start with [n]
            if text and not re.match(r"^\[\d+\]", text):
                bad_format.append((p.index, text[:30]))
    if bad_format:
        return [Finding(
            rule_id="reference-format",
            message=f"参考文献列表有 {len(bad_format)} 条未使用 [n] 编号格式",
            expected="[n] 顶格", actual=f"{len(bad_format)} 条格式异常",
            location="参考文献", fixable=False,
        )]
    return []


@rule("cover-no-page-number", default_severity="info")
def check_cover_no_page_number(docx, content, preset) -> list[Finding]:
    """封面/声明/授权页不应有页码。"""
    # This is typically controlled by section page numbering settings
    # and "different first page" flag
    # TODO: check via sectPr pgNumType presence in first section
    if not preset.get("page", {}).get("page_number", {}).get("cover_excluded", True):
        return []
    return []


@rule("chapter-numbering-arabic", default_severity="info")
def check_chapter_numbering_arabic(docx, content, preset) -> list[Finding]:
    """章号应使用阿拉伯数字（'第1章' 而非 '第一章'）。"""
    if not preset.get("structure", {}).get("chapter_numbering_arabic", True):
        return []
    chinese_num_re = re.compile(r"^第[一二三四五六七八九十]+章")
    findings = []
    for p in docx.paragraphs:
        text = p.text.strip()
        if chinese_num_re.match(text):
            findings.append(Finding(
                rule_id="chapter-numbering-arabic",
                message=f"章号应使用阿拉伯数字，实际 {text[:15]!r}",
                expected="第N章 (阿拉伯)", actual=text[:15],
                location=f"Paragraph {p.index}", fixable=False,
            ))
    return findings
