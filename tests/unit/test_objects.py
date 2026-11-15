import io
import struct
import zlib

import pytest
from docx import Document
from docx.shared import Mm

from docform import DocxTemplate, InlineImage, Listing, R, RichText, RichTextParagraph, RP

from conftest import saved, body_text


def _png(width=4, height=4):
    """Build a small solid PNG without needing an imaging library."""
    raw = b"".join(b"\x00" + b"\x1f\x4e\x79" * width for _ in range(height))

    def chunk(kind, data):
        body = kind + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    signature = b"\x89PNG\r\n\x1a\n"
    return signature + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


PNG = _png()


def test_richtext_formatting(make_docx):
    rt = RichText("bold", bold=True)
    rt.add(" red", color="FF0000", italic=True)
    tpl = DocxTemplate(make_docx(["{{r rt }}"]))
    tpl.render({"rt": rt})
    runs = [r for r in saved(tpl).paragraphs[0].runs if r.text]
    assert [r.text for r in runs] == ["bold", " red"]
    assert runs[0].bold
    assert runs[1].italic
    assert str(runs[1].font.color.rgb) == "FF0000"


def test_richtext_escapes_text(make_docx):
    tpl = DocxTemplate(make_docx(["{{r rt }}"]))
    tpl.render({"rt": R("<a> & b")})
    assert body_text(saved(tpl)) == ["<a> & b"]


def test_richtext_newline_and_tab(make_docx):
    tpl = DocxTemplate(make_docx(["{{r rt }}"]))
    tpl.render({"rt": RichText("a\nb\tc")})
    p = saved(tpl).paragraphs[0]
    assert "<w:br/>" in p._p.xml
    assert "<w:tab/>" in p._p.xml


def test_richtext_link(make_docx):
    tpl = DocxTemplate(make_docx(["{{r rt }}"]))
    rt = RichText("docs", url_id=tpl.build_url_id("https://example.com"))
    tpl.render({"rt": rt})
    doc = saved(tpl)
    targets = [r.target_ref for r in doc.part.rels.values() if r.is_external]
    assert "https://example.com" in targets
    assert "w:hyperlink" in doc.paragraphs[0]._p.xml


def test_richtext_paragraph(make_docx):
    rp = RichTextParagraph()
    rp.add(RichText("first", bold=True))
    rp.add("second", parastyle="Heading1")
    tpl = DocxTemplate(make_docx(["{{p rp }}"]))
    tpl.render({"rp": rp})
    doc = saved(tpl)
    assert body_text(doc) == ["first", "second"]
    assert doc.paragraphs[1].style.name == "Heading 1"


def test_rp_alias():
    assert RP is RichTextParagraph
    assert R is RichText


def test_listing_keeps_line_breaks(make_docx):
    tpl = DocxTemplate(make_docx(["{{ lst }}"]))
    tpl.render({"lst": Listing("one\ntwo & <three>")})
    p = saved(tpl).paragraphs[0]
    assert "<w:br/>" in p._p.xml
    assert "two & <three>" in p.text


def test_inline_image(make_docx):
    tpl = DocxTemplate(make_docx(["logo: {{ img }}"]))
    tpl.render({"img": InlineImage(tpl, io.BytesIO(PNG), width=Mm(20))})
    doc = saved(tpl)
    assert len(doc.inline_shapes) == 1
    assert doc.inline_shapes[0].width == Mm(20)


def test_inline_images_get_unique_ids(make_docx):
    tpl = DocxTemplate(make_docx(["{% for i in imgs %}{{ i }}{% endfor %}"]))
    imgs = [InlineImage(tpl, io.BytesIO(PNG)) for _ in range(3)]
    tpl.render({"imgs": imgs})
    doc = saved(tpl)
    ids = [el.get("id") for el in doc.element.body.iter("{*}docPr")]
    assert len(ids) == 3
    assert len(set(ids)) == 3


def test_subdoc(make_docx):
    pytest.importorskip("docxcompose")
    tpl = DocxTemplate(make_docx(["before", "{{p sd }}", "after"]))
    sd = tpl.new_subdoc()
    sd.add_paragraph("from subdoc")
    tpl.render({"sd": sd})
    assert body_text(saved(tpl)) == ["before", "from subdoc", "after"]


def test_subdoc_from_file(make_docx, tmp_path):
    pytest.importorskip("docxcompose")
    other = Document()
    other.add_paragraph("merged in")
    other_path = tmp_path / "other.docx"
    other.save(str(other_path))

    tpl = DocxTemplate(make_docx(["{{p sd }}"]))
    tpl.render({"sd": tpl.new_subdoc(str(other_path))})
    assert "merged in" in body_text(saved(tpl))
