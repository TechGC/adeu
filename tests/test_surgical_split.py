"""Tests for _try_surgical_split — the per-opcode fan-out helper that splits
high-similarity ModifyText edits into surgical sub-edits.

See `adeu.redline.engine._try_surgical_split` for the function under test and
the rationale behind the default thresholds.
"""

import io

from docx import Document

from adeu.models import EditOperationType, ModifyText
from adeu.redline.engine import (
    DEFAULT_SURGICAL_SPLIT_MIN_SIMILARITY,
    DEFAULT_SURGICAL_SPLIT_MIN_TARGET_LEN,
    RedlineEngine,
    _try_surgical_split,
)


# ---------------------------------------------------------------------------
# Unit tests: _try_surgical_split returns None when the input does not qualify.
# ---------------------------------------------------------------------------


def _stub_edit(comment="rationale"):
    """Build a minimal ModifyText carrying just a comment for the helper to copy."""
    return ModifyText(type="modify", target_text="x", new_text="y", comment=comment)


def test_returns_none_when_new_text_empty():
    """Pure deletion: nothing to diff against, helper bails out."""
    long_target = "x" * (DEFAULT_SURGICAL_SPLIT_MIN_TARGET_LEN + 50)
    result = _try_surgical_split(
        _stub_edit(), long_target, "", effective_start_idx=0, active_mapper=None
    )
    assert result is None


def test_returns_none_when_target_below_min_len():
    """Short residual: even unsplit it renders as a compact mark, splitting adds noise."""
    short_target = "Short text"
    result = _try_surgical_split(
        _stub_edit(),
        short_target,
        "Shorter text",
        effective_start_idx=0,
        active_mapper=None,
    )
    assert result is None


def test_returns_none_when_similarity_below_threshold():
    """Real rewrite (target and new differ substantially): splitting would fragment intent."""
    target = "The quick brown fox jumps over the lazy dog. " * 3
    new = "Wholly different text replacing the entire clause from scratch. " * 3
    result = _try_surgical_split(
        _stub_edit(), target, new, effective_start_idx=0, active_mapper=None
    )
    assert result is None


def test_returns_none_when_single_opcode():
    """Only one non-equal opcode (pure tail insertion): nothing to fan out — caller handles it."""
    target = "Section header. " + ("a" * 100)
    new = target + " Extra appended text."
    result = _try_surgical_split(
        _stub_edit(), target, new, effective_start_idx=0, active_mapper=None
    )
    assert result is None


# ---------------------------------------------------------------------------
# Unit tests: the splitter fires and returns sub-edits with expected shape.
# ---------------------------------------------------------------------------


def test_splits_high_similarity_long_target_into_sub_edits():
    """The canonical bug shape AFTER trim_common_context: a long post-trim
    residual where target and new differ at multiple scattered points but the
    bulk of the content is identical. The splitter should fan out into one
    sub-edit per non-equal opcode."""
    # Mimics what the engine sees post-trim when the LLM emitted a verbose
    # paragraph-wide target with several small embedded edits.
    target = (
        "AAA "
        + ("common run of unchanged text " * 3)
        + "BBB "
        + ("another stretch of equal content " * 2)
        + "CCC end"
    )
    new = (
        "XXX "
        + ("common run of unchanged text " * 3)
        + "YYY "
        + ("another stretch of equal content " * 2)
        + "ZZZ end"
    )
    assert len(target) > DEFAULT_SURGICAL_SPLIT_MIN_TARGET_LEN
    result = _try_surgical_split(
        _stub_edit("rationale-A"),
        target,
        new,
        effective_start_idx=0,
        active_mapper=None,
    )
    assert result is not None, "splitter should have fired on this multi-opcode input"
    assert len(result) >= 2
    # Each sub-edit must be much smaller than the original — that is the whole point.
    for sub in result:
        assert sub.target_text != target  # no sub re-quotes the entire input
        assert sub.new_text != new
        assert sub._match_start_index >= 0
        assert len(sub.target_text or "") < 20  # surgical: small chunks only


def test_sub_edits_use_correct_mapper_coordinates():
    """`_match_start_index` of each sub-edit must equal effective_start_idx + i1."""
    target = "X" * 50 + "ABCDEF" + "Y" * 60  # 116 chars
    new = "X" * 50 + "AbCdEf" + "Y" * 60
    base_idx = 4242
    result = _try_surgical_split(
        _stub_edit(), target, new, effective_start_idx=base_idx, active_mapper=None
    )
    assert result is not None
    # All sub-edits must land inside [base_idx, base_idx + len(target)).
    for sub in result:
        assert base_idx <= sub._match_start_index < base_idx + len(target)


def test_comment_attached_to_first_sub_edit_only():
    """Word renders one balloon per non-empty comment. The original rationale
    must travel with the first sub-edit; the rest must be silent."""
    target = ("Customer's payment obligations. " * 5) + " (iii) extra " + ("Foo bar. " * 5)
    new = ("MLB's payment obligations. " * 5) + " " + ("Foo bar. " * 5)
    result = _try_surgical_split(
        _stub_edit("only-rationale-on-first"),
        target,
        new,
        effective_start_idx=0,
        active_mapper=None,
    )
    assert result is not None
    assert result[0].comment == "only-rationale-on-first"
    for sub in result[1:]:
        assert sub.comment is None or sub.comment == ""


def test_internal_op_set_per_opcode():
    """Each sub-edit must carry _internal_op matching its diff opcode kind.
    Use a fixture with both a replace AND a pure-deletion region to exercise
    multiple opcode kinds in one input."""
    target = "Header AAA " + ("padding word " * 20) + " Footer DEL_ME BBB end."
    new = "Header XXX " + ("padding word " * 20) + " Footer BBB end."
    result = _try_surgical_split(
        _stub_edit(), target, new, effective_start_idx=0, active_mapper=None
    )
    assert result is not None
    valid_ops = {
        EditOperationType.DELETION,
        EditOperationType.INSERTION,
        EditOperationType.MODIFICATION,
    }
    for sub in result:
        assert sub._internal_op in valid_ops


# ---------------------------------------------------------------------------
# Unit tests: custom thresholds via keyword arguments.
# ---------------------------------------------------------------------------


def test_custom_min_target_len_allows_shorter_splits():
    """Caller can lower the length floor to attack shorter restate-shaped edits."""
    target = "ABCDEF GHIJKLM" + " filler" * 5  # ~55 chars
    new = "AbCdEf GHIJKLM" + " filler" * 5
    # Default threshold is 100; with 30 the splitter should fire.
    result = _try_surgical_split(
        _stub_edit(),
        target,
        new,
        effective_start_idx=0,
        active_mapper=None,
        min_target_len=30,
    )
    assert result is not None


def test_custom_min_similarity_can_disable_splitter():
    """Setting similarity threshold above 1.0 makes the splitter never fire."""
    target = ("Foo bar baz qux. " * 10)  # 170 chars, near-identical to new
    new = target.replace("Foo", "FOO", 1)  # one tiny change, very high similarity
    result = _try_surgical_split(
        _stub_edit(),
        target,
        new,
        effective_start_idx=0,
        active_mapper=None,
        min_similarity=1.5,
    )
    assert result is None


# ---------------------------------------------------------------------------
# End-to-end: drive a fixture through RedlineEngine.apply_edits and assert
# the output XML shows surgical (small) tracked changes, not paragraph-wide
# delete-and-restate.
# ---------------------------------------------------------------------------


def _build_long_paragraph_docx() -> io.BytesIO:
    """Fixture: one paragraph long enough to clear the splitter threshold,
    so that an over-anchored edit triggers the splitter."""
    long_clause = (
        "Subcontractors. Company shall indemnify, defend, and hold harmless "
        "the MLB Entities from and against any and all third party claims, "
        "demands, liabilities, damages, costs, and expenses, including, but "
        "not limited to, defense costs and legal fees, arising out of "
        "(i) Company personnel actions, (ii) breach of warranty, "
        "(iii) misuse of confidential information, and the willful misconduct "
        "of any subcontractor engaged by Company."
    )
    doc = Document()
    doc.add_paragraph(long_clause)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream


def test_end_to_end_surgical_split_via_apply_edits():
    """Without the splitter the engine would emit one big <w:del> wrapping
    the original clause plus one big <w:ins> for the near-identical replacement.
    With it, the deletion should be a small, focused mark on `(iii) misuse of
    confidential information, and `."""
    stream = _build_long_paragraph_docx()
    original_clause = (
        "Subcontractors. Company shall indemnify, defend, and hold harmless "
        "the MLB Entities from and against any and all third party claims, "
        "demands, liabilities, damages, costs, and expenses, including, but "
        "not limited to, defense costs and legal fees, arising out of "
        "(i) Company personnel actions, (ii) breach of warranty, "
        "(iii) misuse of confidential information, and the willful misconduct "
        "of any subcontractor engaged by Company."
    )
    cleaned_clause = original_clause.replace(
        "(iii) misuse of confidential information, and ", ""
    )

    edit = ModifyText(
        target_text=original_clause,
        new_text=cleaned_clause,
        comment="Remove (iii); narrow to gross negligence/willful misconduct",
    )

    engine = RedlineEngine(stream)
    engine.apply_edits([edit])

    out = engine.save_to_stream()
    doc = Document(out)
    xml = doc.element.xml

    # The original 500+-char clause must NOT appear inside a single <w:delText>.
    # i.e., we should never see the whole-clause delete-and-restate shape.
    assert original_clause not in xml.replace("</w:delText>", "").replace("<w:delText>", "")
    # And the changed phrase must show up as a strikethrough somewhere.
    assert "misuse of confidential information" in xml
    assert "w:del" in xml
