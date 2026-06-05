"""Test rule registration."""

from thesis_format_checker.rules import RULES


def test_rules_registered():
    expected = {
        "page-margins", "header-text-match", "header-on-all-sections",
        "body-font-size", "body-east-asia-font", "body-line-spacing",
        "heading1-style", "heading2-style", "heading3-style",
        "heading-style-applied", "chapter-page-break",
        "abstract-zh-length", "abstract-en-length",
        "foreign-translation-length", "toc-present", "cover-fields",
    }
    assert expected.issubset(set(RULES.keys())), f"Missing: {expected - set(RULES.keys())}"


def test_preset_loads():
    from thesis_format_checker.checker import load_preset
    preset = load_preset("ncwu")
    assert preset["preset_id"] == "ncwu"
    assert preset["page"]["margin_cm"]["left"] == 3.0
    assert preset["styles"]["body"]["size_pt"] == 12
    assert preset["content"]["abstract_zh"]["min_chars"] == 500
