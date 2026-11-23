import io
import struct
import zlib

import pytest
from docx import Document
from docx.oxml.ns import qn
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


def test_richtext_paragraph_rtl(make_docx):
    rp = RichTextParagraph()
    rp.add("שלום", rtl=True)
    rp.add("left to right")
    tpl = DocxTemplate(make_docx(["{{p rp }}"]))
    tpl.render({"rp": rp})
    doc = saved(tpl)
    rtl_p, ltr_p = doc.paragraphs
    assert rtl_p._p.pPr.find(qn("w:bidi")) is not None
    assert rtl_p.runs[0]._r.rPr.find(qn("w:rtl")) is not None
    assert ltr_p._p.pPr is None or ltr_p._p.pPr.find(qn("w:bidi")) is None
    assert ltr_p.runs[0]._r.rPr is None


def test_richtext_rtl_on_every_add(make_docx):
    rt = RichText("سلام", rtl=True, bold=True)
    rt.add(" دنیا", rtl=True, italic=True)
    tpl = DocxTemplate(make_docx(["{{r rt }}"]))
    tpl.render({"rt": rt})
    runs = [r for r in saved(tpl).paragraphs[0].runs if r.text]
    assert all(r._r.rPr.find(qn("w:rtl")) is not None for r in runs)
    assert runs[0]._r.rPr.find(qn("w:bCs")) is not None
    assert runs[1]._r.rPr.find(qn("w:iCs")) is not None


@pytest.mark.parametrize("autoescape", [False, True])
def test_filters_can_return_richtext_and_inline_image(make_docx, autoescape):
    from jinja2.sandbox import SandboxedEnvironment

    tpl = DocxTemplate(make_docx(["{{r title|rich(color='FF0000', bold=True) }}", "{{ logo|image(10) }}"]))
    env = SandboxedEnvironment(autoescape=autoescape)
    env.filters["rich"] = lambda value, **props: RichText(value, **props)
    env.filters["image"] = lambda data, mm: InlineImage(tpl, io.BytesIO(data), width=Mm(mm))
    tpl.render({"title": "a < b", "logo": PNG}, env, autoescape=autoescape)

    doc = saved(tpl)
    run = [r for r in doc.paragraphs[0].runs if r.text][0]
    assert run.text == "a < b"
    assert run.bold
    assert str(run.font.color.rgb) == "FF0000"
    assert len(doc.inline_shapes) == 1


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


def test_subdoc_from_document_object(make_docx):
    pytest.importorskip("docxcompose")
    other = Document()
    other.add_paragraph("built in memory")
    other.add_picture(io.BytesIO(PNG), width=Mm(10))

    tpl = DocxTemplate(make_docx(["before", "{{p sd }}", "after"]))
    tpl.render({"sd": tpl.new_subdoc(other)})
    doc = saved(tpl)
    assert body_text(doc)[0] == "before"
    assert "built in memory" in body_text(doc)
    assert body_text(doc)[-1] == "after"
    assert doc.inline_shapes[0]._inline.graphic.graphicData.pic.blipFill.blip.embed in doc.part.rels
    # the Document passed in still has its section settings
    assert other.element.body.sectPr is not None


def test_subdoc_from_rendered_template(make_docx):
    pytest.importorskip("docxcompose")
    sub_tpl = DocxTemplate(make_docx(["Hello {{ name }}"]))
    sub_tpl.render({"name": "John"})

    tpl = DocxTemplate(make_docx(["{{p sd }}"]))
    tpl.render({"sd": tpl.new_subdoc(sub_tpl)})
    assert body_text(saved(tpl)) == ["Hello John"]


def _pictures(path):
    """Map each picture's name to the bytes of the image it shows."""
    doc = Document(str(path))
    part = doc.part
    out = {}
    for pic in doc.element.body.xpath(".//pic:pic"):
        name = pic.xpath("./pic:nvPicPr/pic:cNvPr/@name")[0]
        rid = pic.xpath("./pic:blipFill/a:blip/@r:embed")[0]
        out[name] = part.related_parts[rid].blob
    return out, part


def _two_copies_of_one_image(tmp_path):
    doc = Document()
    for _ in range(2):
        doc.add_picture(io.BytesIO(PNG), width=Mm(10))
    for name, el in zip(("img1", "img2"), doc.element.body.xpath(".//pic:cNvPr")):
        el.set("name", name)
    path = tmp_path / "pics.docx"
    doc.save(str(path))
    return path


def test_replace_pic_leaves_identical_images_alone(tmp_path):
    other = _png(6, 6)
    tpl = DocxTemplate(str(_two_copies_of_one_image(tmp_path)))
    tpl.replace_pic("img1", io.BytesIO(other))
    tpl.render({})
    out = tmp_path / "out.docx"
    tpl.save(str(out))

    pics, _ = _pictures(out)
    assert pics == {"img1": other, "img2": PNG}


def test_replace_pic_drops_the_unused_image(tmp_path):
    doc = Document()
    doc.add_picture(io.BytesIO(PNG), width=Mm(10))
    doc.element.body.xpath(".//pic:cNvPr")[0].set("name", "only")
    path = tmp_path / "one.docx"
    doc.save(str(path))

    other = _png(6, 6)
    tpl = DocxTemplate(str(path))
    tpl.replace_pic("only", io.BytesIO(other))
    tpl.render({})
    out = tmp_path / "out.docx"
    tpl.save(str(out))

    pics, part = _pictures(out)
    assert pics == {"only": other}
    images = [p for p in part.related_parts.values() if p.content_type.startswith("image/")]
    assert len(images) == 1


def _picture_sizes(path):
    """Map each picture's name to its (extent, a:ext) sizes in EMU."""
    doc = Document(str(path))
    out = {}
    for inline in doc.element.body.xpath(".//wp:inline"):
        name = inline.xpath(".//pic:cNvPr/@name")[0]
        extent = inline.xpath("./wp:extent")[0]
        ext = inline.xpath(".//pic:spPr/a:xfrm/a:ext")[0]
        out[name] = (
            (int(extent.get("cx")), int(extent.get("cy"))),
            (int(ext.get("cx")), int(ext.get("cy"))),
        )
    return out


def _render_replaced(tmp_path, **size):
    doc = Document()
    doc.add_picture(io.BytesIO(PNG), width=Mm(20), height=Mm(10))
    doc.add_picture(io.BytesIO(PNG), width=Mm(20), height=Mm(10))
    for name, el in zip(("img1", "img2"), doc.element.body.xpath(".//pic:cNvPr")):
        el.set("name", name)
    path = tmp_path / "sized.docx"
    doc.save(str(path))

    tpl = DocxTemplate(str(path))
    tpl.replace_pic("img1", io.BytesIO(_png(6, 6)), **size)
    tpl.render({})
    out = tmp_path / "out.docx"
    tpl.save(str(out))
    return _picture_sizes(out)


def test_replace_pic_sets_width_and_height(tmp_path):
    sizes = _render_replaced(tmp_path, width=Mm(30), height=Mm(40))
    assert sizes["img1"] == ((Mm(30), Mm(40)), (Mm(30), Mm(40)))
    assert sizes["img2"] == ((Mm(20), Mm(10)), (Mm(20), Mm(10)))


def test_replace_pic_scales_the_missing_side(tmp_path):
    sizes = _render_replaced(tmp_path, width=Mm(40))
    assert sizes["img1"] == ((Mm(40), Mm(20)), (Mm(40), Mm(20)))
    sizes = _render_replaced(tmp_path, height=Mm(5))
    assert sizes["img1"] == ((Mm(10), Mm(5)), (Mm(10), Mm(5)))


def test_replace_pic_keeps_size_by_default(tmp_path):
    sizes = _render_replaced(tmp_path)
    assert sizes["img1"] == ((Mm(20), Mm(10)), (Mm(20), Mm(10)))


def test_replace_pic_rejects_non_positive_size():
    tpl = DocxTemplate(io.BytesIO())
    with pytest.raises(ValueError):
        tpl.replace_pic("img1", io.BytesIO(PNG), width=0)


def test_subdoc_ending_with_a_table_is_followed_by_a_paragraph(tmp_path):
    pytest.importorskip("docxcompose")
    doc = Document()
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "{{p sd }}"
    doc.add_paragraph("{{p sd }}")
    path = tmp_path / "tpl.docx"
    doc.save(str(path))

    tpl = DocxTemplate(str(path))
    sd = tpl.new_subdoc()
    sd.add_table(rows=1, cols=1).cell(0, 0).text = "inner"
    tpl.render({"sd": sd})
    body = saved(tpl).element.body

    # A table cell has to end with a paragraph, and so does the body
    # before its section properties.
    outer_cell = body.xpath("./w:tbl/w:tr/w:tc")[0]
    assert outer_cell[-1].tag.endswith("}p")
    assert body[-2].tag.endswith("}p")


def _list_num_ids(doc):
    """numId each "item" paragraph ends up with, from its own numPr or its style."""
    ids = []
    for p in doc.paragraphs:
        if not p.text.startswith("item"):
            continue
        num = p._p.xpath("./w:pPr/w:numPr/w:numId/@w:val")
        if not num:
            num = p.style.element.xpath("./w:pPr/w:numPr/w:numId/@w:val")
        ids.append(num[0] if num else None)
    return ids


def test_nested_subdoc_keeps_a_style_numbered_list_together(tmp_path):
    pytest.importorskip("docxcompose")
    first = Document()
    first.add_paragraph("{%p for i in items %}")
    first.add_paragraph("{{ i }}", style="List Number")
    first.add_paragraph("{%p endfor %}")
    first.save(str(tmp_path / "first_tpl.docx"))
    wrap = Document()
    wrap.add_paragraph("{{p inner }}")
    wrap.save(str(tmp_path / "wrap_tpl.docx"))

    tpl = DocxTemplate(str(tmp_path / "first_tpl.docx"))
    tpl.render({"items": ["item 1", "item 2", "item 3"]})
    previous = tmp_path / "level1.docx"
    tpl.save(str(previous))
    for level in (2, 3):
        tpl = DocxTemplate(str(tmp_path / "wrap_tpl.docx"))
        tpl.render({"inner": tpl.new_subdoc(str(previous))})
        previous = tmp_path / ("level%d.docx" % level)
        tpl.save(str(previous))
        ids = _list_num_ids(Document(str(previous)))
        # one list, so one numbering instance for all of its items
        assert len(ids) == 3 and len(set(ids)) == 1, (level, ids)
