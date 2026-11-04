import io
import re
import zipfile

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from adeu.models import EditComment, ModifyText
from adeu.redline.comments import CommentsManager, w14_ns, w15_ns
from adeu.redline.engine import RedlineEngine


def test_native_comment_creation_and_linking():
    doc = Document()
    doc.add_paragraph("The quick brown fox.")

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)

    edit = ModifyText(target_text="quick", new_text="slow", comment="Foxes are not always quick.")

    engine = RedlineEngine(stream)
    engine.apply_edits([edit])

    result_stream = engine.save_to_stream()
    doc = Document(result_stream)

    doc_xml = doc.element.xml
    assert "w:commentRangeStart" in doc_xml
    assert "w:commentRangeEnd" in doc_xml
    assert "w:commentReference" in doc_xml

    id_match = re.search(r'<w:commentRangeStart[^>]*w:id="(\d+)"', doc_xml)
    assert id_match
    comment_id = id_match.group(1)

    comments_part = None
    for rel in doc.part.rels.values():
        if rel.reltype == RT.COMMENTS:
            comments_part = rel.target_part
            break

    assert comments_part is not None
    comments_xml = comments_part.blob.decode("utf-8")
    assert "Foxes are not always quick." in comments_xml
    assert f'w:id="{comment_id}"' in comments_xml


def test_multiple_comments_ids():
    doc = Document()
    doc.add_paragraph("First sentence.")
    doc.add_paragraph("Second sentence.")

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)

    edit1 = ModifyText(target_text="First", new_text="First The ", comment="Comment One")
    edit2 = ModifyText(target_text="Second", new_text="Second The ", comment="Comment Two")

    engine = RedlineEngine(stream)
    engine.apply_edits([edit1, edit2])

    result_stream = engine.save_to_stream()
    doc = Document(result_stream)

    comments_part = None
    for rel in doc.part.rels.values():
        if rel.reltype == RT.COMMENTS:
            comments_part = rel.target_part
            break

    comments_xml = comments_part.blob.decode("utf-8")

    assert "Comment One" in comments_xml
    assert "Comment Two" in comments_xml

    ids = re.findall(r'<w:comment\s+[^>]*w:id="(\d+)"', comments_xml)
    assert len(set(ids)) == 2


def _apply_and_read(paragraph: str, edit: ModifyText) -> tuple[str, str]:
    doc = Document()
    doc.add_paragraph(paragraph)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)

    engine = RedlineEngine(stream, author="Jane Doe")
    engine.apply_edits([edit])
    result = Document(engine.save_to_stream())

    comments_part = next(rel.target_part for rel in result.part.rels.values() if rel.reltype == RT.COMMENTS)
    return result.element.xml, comments_part.blob.decode("utf-8")


def _authors_by_text(comments_xml: str) -> dict[str, str]:
    pairs = re.findall(r'<w:comment\b[^>]*w:author="([^"]*)"[^>]*>.*?<w:t[^>]*>([^<]*)</w:t>', comments_xml, re.S)
    return {text: author for author, text in pairs}


def test_comment_list_writes_one_comment_per_entry_with_its_author():
    edit = ModifyText(
        target_text="quick",
        new_text="slow",
        comment=[
            EditComment(text="Internal: foxes vary.", author="Jane Doe (Internal)"),
            EditComment(text="Vendor: please confirm.", author="Jane Doe (Vendor)"),
        ],
    )

    _doc_xml, comments_xml = _apply_and_read("The quick brown fox.", edit)

    assert _authors_by_text(comments_xml) == {
        "Internal: foxes vary.": "Jane Doe (Internal)",
        "Vendor: please confirm.": "Jane Doe (Vendor)",
    }


def test_comment_list_entries_share_the_same_anchor():
    edit = ModifyText(
        target_text="quick",
        new_text="slow",
        comment=[EditComment(text="One"), EditComment(text="Two")],
    )

    doc_xml, _comments_xml = _apply_and_read("The quick brown fox.", edit)

    starts = re.findall(r'<w:commentRangeStart[^>]*w:id="(\d+)"', doc_xml)
    ends = re.findall(r'<w:commentRangeEnd[^>]*w:id="(\d+)"', doc_xml)
    assert len(starts) == 2
    assert ends == list(reversed(starts))
    between_starts = doc_xml.split(f'w:id="{starts[0]}"')[1].split(f'w:id="{starts[1]}"')[0]
    assert "<w:t" not in between_starts


def test_comment_list_entry_without_author_uses_the_engine_author():
    edit = ModifyText(target_text="quick", new_text="slow", comment=[EditComment(text="No author given")])

    _doc_xml, comments_xml = _apply_and_read("The quick brown fox.", edit)

    assert _authors_by_text(comments_xml) == {"No author given": "Jane Doe"}


def test_string_comment_still_writes_a_single_comment_under_the_engine_author():
    edit = ModifyText(target_text="quick", new_text="slow", comment="Plain comment")

    _doc_xml, comments_xml = _apply_and_read("The quick brown fox.", edit)

    assert _authors_by_text(comments_xml) == {"Plain comment": "Jane Doe"}


def test_surgically_split_edit_keeps_every_comment_together():
    edit = ModifyText(
        target_text="The Supplier shall deliver all of the ordered goods to the Customer within thirty days.",
        new_text="The Vendor shall deliver all of the ordered goods to the Customer within sixty days.",
        comment=[EditComment(text="Internal note", author="A"), EditComment(text="Vendor note", author="B")],
    )

    doc_xml, comments_xml = _apply_and_read(
        "The Supplier shall deliver all of the ordered goods to the Customer within thirty days.", edit
    )

    assert len(re.findall(r"<w:ins\b", doc_xml)) > 1
    assert _authors_by_text(comments_xml) == {"Internal note": "A", "Vendor note": "B"}
    starts = re.findall(r'<w:commentRangeStart[^>]*w:id="(\d+)"', doc_xml)
    assert len(starts) == 2
    between_starts = doc_xml.split(f'w:id="{starts[0]}"')[1].split(f'w:id="{starts[1]}"')[0]
    assert "<w:t" not in between_starts


def test_comment_text_joins_a_comment_list():
    edit = ModifyText(target_text="a", new_text="b", comment=[EditComment(text="One"), EditComment(text="Two")])

    assert edit.comment_text == "One\n\nTwo"


def test_paragraphs_added_after_the_target_keep_their_comment():
    sentence = "This Agreement is governed by the laws of New York."
    edit = ModifyText(
        target_text=sentence,
        new_text=(
            f"{sentence}\nConfidentiality: Firm keeps Client information confidential."
            "\nNon-Solicitation: Firm will not recruit Client employees."
        ),
        comment="Adds the missing protections.",
    )

    doc_xml, comments_xml = _apply_and_read(sentence, edit)

    assert _authors_by_text(comments_xml) == {"Adds the missing protections.": "Jane Doe"}
    start_id = re.search(r'<w:commentRangeStart[^>]*w:id="(\d+)"', doc_xml).group(1)
    anchored = doc_xml.split("<w:commentRangeStart")[1].split(f'<w:commentRangeEnd w:id="{start_id}"')[0]
    assert "Confidentiality:" in anchored
    assert "Non-Solicitation:" in anchored


_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_W16SE_NS = "http://schemas.microsoft.com/office/word/2015/wordml/symex"


def _docx_with_existing_comments(paragraph: str, comments_root_attrs: str, comment_body: str) -> io.BytesIO:
    doc = Document()
    doc.add_paragraph(paragraph)
    source = io.BytesIO()
    doc.save(source)

    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments {comments_root_attrs}><w:comment w:id="0" w:author="Counterparty">{comment_body}</w:comment>'
        "</w:comments>"
    )
    result = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(source.getvalue())) as src, zipfile.ZipFile(result, "w") as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == "[Content_Types].xml":
                data = data.replace(
                    b"</Types>",
                    b'<Override PartName="/word/comments.xml" ContentType="application/'
                    b'vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
                )
            if item.filename == "word/_rels/document.xml.rels":
                data = data.replace(
                    b"</Relationships>",
                    f'<Relationship Id="rIdComments" Type="{RT.COMMENTS}" Target="comments.xml"/>'
                    f'<Relationship Id="rIdLink" Type="{RT.HYPERLINK}" Target="https://example.com" '
                    'TargetMode="External"/></Relationships>'.encode(),
                )
            out.writestr(item, data)
        out.writestr("word/comments.xml", comments_xml)
    result.seek(0)
    return result


def _apply_to_existing_comments(comments_root_attrs: str, comment_body: str) -> str:
    sentence = "The quick brown fox."
    engine = RedlineEngine(_docx_with_existing_comments(sentence, comments_root_attrs, comment_body), author="Jane Doe")
    engine.apply_edits([ModifyText(target_text="quick", new_text="slow", comment="Foxes are not always quick.")])
    result = Document(engine.save_to_stream())
    return next(rel.target_part for rel in result.part.rels.values() if rel.reltype == RT.COMMENTS).blob.decode()


def test_comment_added_beside_an_existing_hyperlinked_comment_keeps_its_namespace():
    hyperlinked_comment = '<w:p><w:hyperlink r:id="rIdLink"><w:r><w:t>see policy</w:t></w:r></w:hyperlink></w:p>'

    comments_xml = _apply_to_existing_comments(f'xmlns:w="{_W_NS}" xmlns:r="{_R_NS}"', hyperlinked_comment)

    assert f'xmlns:r="{_R_NS}"' in comments_xml
    assert 'r:id="rIdLink"' in comments_xml
    assert _authors_by_text(comments_xml) == {"see policy": "Counterparty", "Foxes are not always quick.": "Jane Doe"}
    assert "w14:paraId" in comments_xml
    assert "ns0:" not in comments_xml


def test_existing_ignorable_prefixes_survive_alongside_the_added_ones():
    root_attrs = f'xmlns:w="{_W_NS}" xmlns:mc="{_MC_NS}" xmlns:w16se="{_W16SE_NS}" mc:Ignorable="w16se"'

    comments_xml = _apply_to_existing_comments(root_attrs, "<w:p><w:r><w:t>Earlier note</w:t></w:r></w:p>")

    ignorable = re.search(r'mc:Ignorable="([^"]*)"', comments_xml).group(1).split()
    assert ignorable == ["w16se", "w14", "w15", "w16cid", "w16cex"]
    assert f'xmlns:w16se="{_W16SE_NS}"' in comments_xml


def test_w14_bound_to_another_uri_is_rebound_so_new_comments_keep_the_w14_prefix():
    comments_xml = _apply_to_existing_comments(
        f'xmlns:w="{_W_NS}" xmlns:w14="urn:other"', "<w:p><w:r><w:t>Earlier note</w:t></w:r></w:p>"
    )

    assert f'xmlns:w14="{w14_ns}"' in comments_xml
    assert "w14:paraId" in comments_xml
    assert "ns0:paraId" not in comments_xml


def test_a_root_already_declaring_w14_and_w15_as_ignorable_is_left_untouched():
    root_attrs = (
        f'xmlns:w="{_W_NS}" xmlns:w14="{w14_ns}" xmlns:w15="{w15_ns}" xmlns:mc="{_MC_NS}" mc:Ignorable="w14 w15"'
    )
    doc = Document(_docx_with_existing_comments("Text.", root_attrs, "<w:p><w:r><w:t>Earlier</w:t></w:r></w:p>"))
    manager = CommentsManager(doc)
    original_root = manager._get_or_create_comments_part().element

    assert manager.comments_part.element is original_root
