"""Tests for DocumentMapper markup-tolerant anchor matching (1.5.2-thesuite.3).

Word cross-reference fields render in clean-view ``full_text`` as CriticMarkup
(``[~2.1~](#_Ref123)``) and bookmark anchors as ``{#_Ref123}``. An LLM-driven
caller anchors on the *visible* text (``2.1``; bookmarks render nothing), so the
raw substring/fuzzy ladder never matches and the edit is rejected. These tests
cover the field-visible projection fallback that lets those anchors resolve back
to real ``full_text`` offsets (markup-inclusive span) without mutating the doc.

Drop this file in the fork's test tree (e.g. ``tests/redline/``).
"""

from docx import Document
from docx.oxml import parse_xml

from adeu.redline.mapper import DocumentMapper

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _bare_mapper() -> DocumentMapper:
    """A DocumentMapper instance without running __init__ — for pure-helper unit tests."""
    return DocumentMapper.__new__(DocumentMapper)


def _doc_with_cross_reference() -> Document:
    """Build a real .docx: a bookmarked heading + a paragraph that REF-references it.

    Renders in clean-view full_text as:
        '{#_Ref123}2.1 License Grant.\\n\\nExcept as expressly granted in '
        'Section [~2.1~](#_Ref123), nothing else is licensed hereunder.'
    The cross-ref points at a *different* paragraph, mirroring real documents, so
    the structural appendix never echoes the referencing sentence verbatim.
    """
    heading = (
        f'<w:p xmlns:w="{W}">'
        '<w:bookmarkStart w:id="1" w:name="_Ref123"/>'
        "<w:r><w:t>2.1 License Grant.</w:t></w:r>"
        '<w:bookmarkEnd w:id="1"/></w:p>'
    )
    body = (
        f'<w:p xmlns:w="{W}">'
        "<w:r><w:t>Except as expressly granted in Section </w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText> REF _Ref123 \\r \\h </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        "<w:r><w:t>2.1</w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        "<w:r><w:t>, nothing else is licensed hereunder.</w:t></w:r></w:p>"
    )
    doc = Document()
    doc.element.body.insert(0, parse_xml(body))
    doc.element.body.insert(0, parse_xml(heading))
    return doc


# --- Pure projection helpers (no docx parsing) ----------------------------------


def test_visible_projection_collapses_xref_and_drops_bookmark():
    m = _bare_mapper()
    text = "see {#_Ref9}Section [~10.16~](#_Ref9) and [~3~](#_Ref8) below"
    visible, index_map = m._build_field_visible_projection(text)

    assert visible == "see Section 10.16 and 3 below"
    assert "[~" not in visible and "{#" not in visible
    # index_map has one entry per visible char plus the trailing sentinel.
    assert len(index_map) == len(visible) + 1
    assert index_map[-1] == len(text)


def test_visible_projection_index_map_roundtrips_to_source():
    m = _bare_mapper()
    text = "alpha [~2.1~](#_RefA) beta {#_RefB} gamma"
    visible, index_map = m._build_field_visible_projection(text)
    # Every visible char points back at the identical char in the source text.
    for i, ch in enumerate(visible):
        assert text[index_map[i]] == ch


def test_visible_projection_noop_when_no_field_markup():
    m = _bare_mapper()
    text = "plain text with **bold** but no fields"
    visible, index_map = m._build_field_visible_projection(text)
    assert visible == text  # unchanged; the fast path short-circuits the fallback
    assert len(visible) == len(text)


def test_map_visible_span_to_full_spans_interior_markup():
    m = _bare_mapper()
    text = "Section [~2.1~](#_Ref1) here"
    visible, index_map = m._build_field_visible_projection(text)
    # 'Section 2.1' in the projection -> the real span includes the interior markup.
    v_start = visible.index("Section 2.1")
    start, length = m._map_visible_span_to_full(v_start, len("Section 2.1"), index_map)
    assert text[start : start + length] == "Section [~2.1"


# --- Integration through a real cross-reference document ------------------------


def test_find_match_index_resolves_anchor_spanning_xref_field():
    m = DocumentMapper(_doc_with_cross_reference(), clean_view=True)
    anchor = "expressly granted in Section 2.1, nothing else is licensed"

    start, length = m.find_match_index(anchor)

    assert start != -1, "field-spanning anchor should resolve via the projection fallback"
    matched = m.full_text[start : start + length]
    assert "[~2.1~](#_Ref123)" in matched  # mapped back to the real markup-laden span


def test_find_target_runs_includes_the_field_display_run():
    m = DocumentMapper(_doc_with_cross_reference(), clean_view=True)
    runs = m.find_target_runs("expressly granted in Section 2.1, nothing else is licensed")

    texts = [r.text for r in runs]
    assert texts == ["expressly granted in Section ", "2.1", ", nothing else is licensed"]


def test_find_all_match_indices_resolves_field_anchor():
    m = DocumentMapper(_doc_with_cross_reference(), clean_view=True)
    matches = m.find_all_match_indices("expressly granted in Section 2.1, nothing else is licensed")

    assert len(matches) == 1
    start, length = matches[0]
    assert "[~2.1~](#_Ref123)" in m.full_text[start : start + length]


def test_raw_match_unaffected_for_anchor_without_fields():
    """Regression: an anchor with no field markup still resolves via the raw fast path."""
    m = DocumentMapper(_doc_with_cross_reference(), clean_view=True)
    anchor = "nothing else is licensed hereunder"

    start, length = m.find_match_index(anchor)
    assert start != -1
    assert m.full_text[start : start + length] == anchor  # exact, no markup involved
