# FILE: tests/test_edit_results_identity.py
from io import BytesIO

from docx import Document

from adeu.models import ModifyText
from adeu.redline.engine import RedlineEngine


def _engine_from_paragraphs(*paragraphs: str) -> RedlineEngine:
    """Builds an engine over a throwaway document containing the given paragraphs."""
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)

    stream = BytesIO()
    doc.save(stream)
    stream.seek(0)
    return RedlineEngine(stream, author="Test AI")


def test_edit_results_has_one_entry_per_caller_edit_in_input_order():
    """Every submitted edit gets exactly one result, addressed by its input index."""
    engine = _engine_from_paragraphs(
        "The Company shall pay the annual bonus in March.",
        "The Executive may terminate this agreement at will.",
        "Confidential information survives for five years.",
    )

    stats = engine.process_batch(
        [
            ModifyText(target_text="in March", new_text="in January"),
            ModifyText(target_text="at will", new_text="for Good Reason"),
            ModifyText(target_text="five years", new_text="three years"),
        ]
    )

    results = stats["edit_results"]
    assert [r["index"] for r in results] == [0, 1, 2]
    assert [r["status"] for r in results] == ["applied", "applied", "applied"]
    assert all(r["reason"] is None for r in results)


def test_edit_results_names_which_edit_was_skipped():
    """
    The bug this exists to prevent: counts alone cannot say *which* edit failed.

    Two edits resolve to overlapping spans. The engine applies the later one and drops the
    earlier, so `edits_applied` is 1 and `edits_skipped` is 1 -- identical to the result you
    would get if the *other* edit had been the casualty. Only the index tells them apart.
    """
    engine = _engine_from_paragraphs("The termination fee is twenty five thousand dollars.")

    stats = engine.process_batch(
        [
            ModifyText(target_text="twenty five thousand", new_text="fifty thousand"),
            ModifyText(target_text="twenty five thousand", new_text="ten thousand"),
        ]
    )

    results = stats["edit_results"]
    assert len(results) == 2
    assert stats["edits_applied"] + stats["edits_skipped"] == 2

    by_index = {r["index"]: r for r in results}
    statuses = {by_index[0]["status"], by_index[1]["status"]}
    assert statuses == {"applied", "skipped"}

    skipped = next(r for r in results if r["status"] == "skipped")
    assert skipped["reason"] == "overlap"
    assert skipped["parts_applied"] == 0


def test_edit_results_keeps_one_entry_when_the_surgical_split_fans_an_edit_out():
    """
    A single caller edit that splits into several tracked changes still has one outcome.

    This is why `edits_applied` cannot be used as an identity: it counts resolved spans, so it
    can exceed the number of edits submitted.
    """
    engine = _engine_from_paragraphs(
        "The Company may provide written notice to the Executive within thirty days of such event."
    )

    stats = engine.process_batch(
        [
            ModifyText(
                target_text=(
                    "The Company may provide written notice to the Executive within thirty days of such event."
                ),
                new_text=("The Company shall provide written notice to the Executive within sixty days of such event."),
            )
        ]
    )

    results = stats["edit_results"]
    assert len(results) == 1, "one submitted edit must stay one result however many spans it became"
    assert results[0]["index"] == 0
    assert results[0]["status"] == "applied"
    assert results[0]["parts_total"] >= 2, "this target is expected to fan out; the split is what makes counts useless"
    assert results[0]["parts_applied"] == results[0]["parts_total"]
    assert stats["edits_applied"] > len(results), (
        "counts exceed the submitted edits -- the reason identity cannot come from them"
    )


def test_edit_results_length_tracks_input_not_applied_count():
    """`edit_results` is indexed by the caller's list, so its length never drifts from it."""
    engine = _engine_from_paragraphs(
        "The Company may provide written notice to the Executive within thirty days of such event.",
        "Payment is due in March of each calendar year.",
    )

    edits = [
        ModifyText(
            target_text=("The Company may provide written notice to the Executive within thirty days of such event."),
            new_text=("The Company shall provide written notice to the Executive within sixty days of such event."),
        ),
        ModifyText(target_text="in March", new_text="in January"),
    ]
    stats = engine.process_batch(edits)

    assert len(stats["edit_results"]) == len(edits)
    assert [r["index"] for r in stats["edit_results"]] == [0, 1]


def test_edit_results_reports_the_resolved_offset():
    """The offset the engine resolved is returned, so a caller can check placement positionally."""
    engine = _engine_from_paragraphs("Confidential information survives for five years.")

    stats = engine.process_batch([ModifyText(target_text="five years", new_text="three years")])

    result = stats["edit_results"][0]
    assert result["status"] == "applied"
    assert isinstance(result["match_start_index"], int)
    assert result["match_start_index"] >= 0


def test_edit_results_is_reset_between_batches():
    """A second batch does not inherit the first batch's outcomes."""
    engine = _engine_from_paragraphs(
        "The Company shall pay the annual bonus in March.",
        "Confidential information survives for five years.",
        "Governing law is the State of Delaware.",
    )

    engine.process_batch(
        [
            ModifyText(target_text="in March", new_text="in January"),
            ModifyText(target_text="five years", new_text="three years"),
        ]
    )
    stats = engine.process_batch([ModifyText(target_text="State of Delaware", new_text="State of New York")])

    assert len(stats["edit_results"]) == 1
    assert stats["edit_results"][0]["index"] == 0
