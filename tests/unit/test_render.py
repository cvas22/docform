import io
import os
from pathlib import Path

import pytest
from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from jinja2 import Environment
from jinja2.exceptions import SecurityError, TemplateSyntaxError, UndefinedError
from jinja2 import StrictUndefined

import docform
from docform import DocxTemplate

from conftest import saved, body_text, table_rows


def test_variables(make_docx):
    tpl = DocxTemplate(make_docx(["Hello {{ name }}", "Total: {{ order.total }}"]))
    tpl.render({"name": "Ada", "order": {"total": "$12"}})
    assert body_text(saved(tpl)) == ["Hello Ada", "Total: $12"]


@pytest.mark.parametrize(
    ("separator", "expected"), [("", "}}"), ("\t", "}\t_}"), ("\n", "}\n_}")]
)
def test_escaped_delimiter_across_runs(tmp_path, separator, expected):
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("left }")
    paragraph.add_run(separator)
    paragraph.add_run("_").bold = True
    paragraph.add_run("} right")
    path = tmp_path / "delimiter.docx"
    doc.save(path)

    tpl = DocxTemplate(path)
    tpl.render({})
    assert body_text(saved(tpl)) == ["left " + expected + " right"]


def test_escaped_delimiter_across_proofing_marks(tmp_path):
    doc = Document()
    paragraph = doc.add_paragraph()
    paragraph.add_run("left }")
    for mark, text in (("spellStart", "_"), ("spellEnd", "} right")):
        marker = OxmlElement("w:proofErr")
        marker.set(qn("w:type"), mark)
        paragraph._p.append(marker)
        paragraph.add_run(text)
    path = tmp_path / "proofing.docx"
    doc.save(path)

    tpl = DocxTemplate(path)
    tpl.render({})
    assert body_text(saved(tpl)) == ["left }} right"]


def test_missing_variable_renders_empty(make_docx):
    tpl = DocxTemplate(make_docx(["Hi {{ missing }}!"]))
    tpl.render({})
    assert body_text(saved(tpl)) == ["Hi !"]


def test_strict_undefined_raises(make_docx):
    tpl = DocxTemplate(make_docx(["Hi {{ missing }}"]))
    with pytest.raises(UndefinedError):
        tpl.render({}, jinja_env=Environment(undefined=StrictUndefined))


def test_syntax_error_raises(make_docx):
    tpl = DocxTemplate(make_docx(["{% if %}"]))
    with pytest.raises(TemplateSyntaxError):
        tpl.render({})


def test_paragraph_if_removes_tag_paragraphs(make_docx):
    path = make_docx(["{%p if paid %}", "PAID", "{%p else %}", "DUE", "{%p endif %}", "end"])
    tpl = DocxTemplate(path)
    tpl.render({"paid": True})
    assert body_text(saved(tpl)) == ["PAID", "end"]

    tpl = DocxTemplate(path)
    tpl.render({"paid": False})
    assert body_text(saved(tpl)) == ["DUE", "end"]


def test_paragraph_for_loop(make_docx):
    tpl = DocxTemplate(make_docx(["{%p for x in items %}", "- {{ x }}", "{%p endfor %}"]))
    tpl.render({"items": ["a", "b", "c"]})
    assert body_text(saved(tpl)) == ["- a", "- b", "- c"]


def test_paragraph_tag_with_percent_and_brace(make_docx):
    tpl = DocxTemplate(make_docx([
        "{%p set d = {'label': 'even'} %}",
        "{%p for i in range(4) %}",
        "{%p if i % 2 == 0 %}",
        "{{ d.label }} {{ i }}",
        "{%p endif %}",
        "{%p endfor %}",
    ]))
    tpl.render({})
    assert body_text(saved(tpl)) == ["even 0", "even 2"]


def test_table_row_loop(make_docx):
    table = [
        ["Name", "Qty"],
        ["{%tr for i in items %}", ""],
        ["{{ i.name }}", "{{ i.qty }}"],
        ["{%tr endfor %}", ""],
        ["Total", "{{ total }}"],
    ]
    tpl = DocxTemplate(make_docx(table=table))
    tpl.render({"items": [{"name": "pen", "qty": 2}, {"name": "ink", "qty": 5}], "total": 7})
    assert table_rows(saved(tpl)) == [
        ["Name", "Qty"],
        ["pen", "2"],
        ["ink", "5"],
        ["Total", "7"],
    ]


def test_whitespace_control_before_a_table_keeps_the_table(tmp_path):
    doc = Document()
    doc.add_paragraph("{% if True -%}")
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "cell"
    doc.add_paragraph("{% endif %}after")
    path = tmp_path / "ws.docx"
    doc.save(str(path))

    tpl = DocxTemplate(str(path))
    tpl.render({})
    out = saved(tpl)
    assert table_rows(out) == [["cell"]]
    assert body_text(out)[-1] == "after"


def test_vm_with_paragraph_text_alignment(tmp_path):
    doc = Document()
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "{%tr for x in xs %}"
    table.cell(1, 0).text = "{% vm %}{{ x.g }}"
    table.cell(1, 1).text = "{{ x.v }}"
    table.cell(2, 0).text = "{%tr endfor %}"
    # Word adds this for vertically centered text, and it starts with "w:t"
    align = OxmlElement("w:textAlignment")
    align.set(qn("w:val"), "center")
    table.cell(1, 0).paragraphs[0]._p.get_or_add_pPr().append(align)
    doc.add_paragraph("after table")
    path = tmp_path / "vm.docx"
    doc.save(str(path))

    tpl = DocxTemplate(str(path))
    tpl.render({"xs": [{"g": "A", "v": n} for n in (1, 2, 3)]})
    out = saved(tpl)
    assert [row[1] for row in table_rows(out)] == ["1", "2", "3"]
    merges = out.element.body.xpath(".//w:tc/w:tcPr/w:vMerge/@w:val")
    assert merges == ["restart", "continue", "continue"]
    assert body_text(out) == ["after table"]


def test_table_row_loop_with_no_items_leaves_header(make_docx):
    table = [["Name"], ["{%tr for i in items %}"], ["{{ i }}"], ["{%tr endfor %}"]]
    tpl = DocxTemplate(make_docx(table=table))
    tpl.render({"items": []})
    assert table_rows(saved(tpl)) == [["Name"]]


def test_table_column_loop(make_docx):
    table = [["first", "{%tc for c in cols %}", "{{ c }}", "{%tc endfor %}"]]
    tpl = DocxTemplate(make_docx(table=table))
    tpl.render({"cols": ["a", "b", "c"]})
    assert table_rows(saved(tpl)) == [["first", "a", "b", "c"]]


def test_run_tag_is_removed(make_docx):
    tpl = DocxTemplate(make_docx(["x{%r if True %}y{%r endif %}z"]))
    tpl.render({})
    assert body_text(saved(tpl)) == ["xyz"]


def test_cellbg_sets_shading(make_docx):
    tpl = DocxTemplate(make_docx(table=[["{% cellbg color %}hot"]]))
    tpl.render({"color": "FF0000"})
    cell = saved(tpl).tables[0].rows[0].cells[0]
    assert cell.text == "hot"
    assert 'w:fill="FF0000"' in cell._tc.xml


def test_colspan_sets_gridspan(make_docx):
    tpl = DocxTemplate(make_docx(table=[["{% colspan n %}wide"]]))
    tpl.render({"n": 3})
    xml = saved(tpl).tables[0]._tbl.xml
    assert 'w:gridSpan w:val="3"' in xml


def test_header_footer_and_properties(make_docx):
    path = make_docx(["body {{ x }}"], header="head {{ x }}", footer="foot {{ x }}", title="Report {{ x }}")
    tpl = DocxTemplate(path)
    tpl.render({"x": "42"})
    doc = saved(tpl)
    assert body_text(doc) == ["body 42"]
    assert doc.sections[0].header.paragraphs[0].text == "head 42"
    assert doc.sections[0].footer.paragraphs[0].text == "foot 42"
    assert doc.core_properties.title == "Report 42"


UNSAFE_EXPR = "{{ cycler.__init__.__globals__.os.name }}"


@pytest.mark.parametrize("where", ["body", "header", "footer", "title"])
def test_sandboxed_by_default(make_docx, where):
    tpl = DocxTemplate(make_docx(**{
        "paragraphs": [UNSAFE_EXPR] if where == "body" else ["ok"],
        "header": UNSAFE_EXPR if where == "header" else None,
        "footer": UNSAFE_EXPR if where == "footer" else None,
        "title": UNSAFE_EXPR if where == "title" else None,
    }))
    with pytest.raises(SecurityError):
        tpl.render({})


def test_sandboxed_with_autoescape(make_docx):
    tpl = DocxTemplate(make_docx([UNSAFE_EXPR]))
    with pytest.raises(SecurityError):
        tpl.render({}, autoescape=True)


def test_plain_environment_opts_out_of_sandbox(make_docx):
    tpl = DocxTemplate(make_docx([UNSAFE_EXPR]))
    tpl.render({}, jinja_env=Environment())
    assert body_text(saved(tpl)) == [os.name]


def test_custom_filter(make_docx):
    env = Environment()
    env.filters["shout"] = lambda s: s.upper() + "!"
    tpl = DocxTemplate(make_docx(["{{ word|shout }}"]))
    tpl.render({"word": "hey"}, jinja_env=env)
    assert body_text(saved(tpl)) == ["HEY!"]


def test_xml_special_chars_need_autoescape(make_docx):
    tpl = DocxTemplate(make_docx(["{{ s }}"]))
    tpl.render({"s": "a < b & c"}, autoescape=True)
    assert body_text(saved(tpl)) == ["a < b & c"]


def test_invalid_xml_warns_instead_of_dropping_silently(make_docx):
    tpl = DocxTemplate(make_docx(["{%p for x in items %}", "{{ x }}", "{%p endfor %}", "after"]))
    with pytest.warns(UserWarning, match="autoescape=True"):
        tpl.render({"items": ["1", "<NA>", "3"]})


def test_valid_render_does_not_warn(make_docx, recwarn):
    tpl = DocxTemplate(make_docx(["{{ s }}"]))
    tpl.render({"s": "a < b"}, autoescape=True)
    assert not [w for w in recwarn if issubclass(w.category, UserWarning)]


def test_render_twice_with_init(make_docx):
    tpl = DocxTemplate(make_docx(["Hi {{ n }}"]))
    tpl.render({"n": 1})
    assert body_text(saved(tpl)) == ["Hi 1"]
    tpl.init_docx()
    tpl.render({"n": 2})
    assert body_text(saved(tpl)) == ["Hi 2"]


def test_document_attributes_before_render(make_docx):
    tpl = DocxTemplate(make_docx(["Hello {{ name }}"], title="T"))
    assert [p.text for p in tpl.paragraphs] == ["Hello {{ name }}"]
    assert tpl.core_properties.title == "T"

    tpl.render({"name": "Ada"})
    assert [p.text for p in tpl.paragraphs] == ["Hello Ada"]


def test_missing_attribute_still_raises(make_docx):
    tpl = DocxTemplate(make_docx(["x"]))
    with pytest.raises(AttributeError):
        tpl.no_such_attribute


def test_save_to_path(make_docx, tmp_path):
    tpl = DocxTemplate(make_docx(["{{ a }}"]))
    tpl.render({"a": "ok"})
    out = tmp_path / "out.docx"
    tpl.save(out)

    assert body_text(Document(str(out))) == ["ok"]


def test_undeclared_variables(make_docx):
    path = make_docx(
        ["{{ name }}", "{%p for i in items %}", "{{ i.x }}", "{%p endfor %}"],
        header="{{ company }}",
    )
    tpl = DocxTemplate(path)
    assert tpl.get_undeclared_template_variables() == {"name", "items", "company"}
    assert tpl.get_undeclared_template_variables(context={"name": 1}) == {"items", "company"}


def test_undeclared_variables_nested(make_docx):
    path = make_docx(
        [
            "{{ foo.bar.baz }} {{ foo.bar }} {{ foo['Qux'] }}",
            "{{ rows[i].x }} {{ name }}",
            "{% for k, v in data.items() %}{{ v }}{% endfor %}",
            "{% for it in items %}{{ it.price }}{% endfor %}",
        ],
        header="{{ company.name|upper }}",
    )
    tpl = DocxTemplate(path)
    assert tpl.get_undeclared_template_variables(nested=True) == {
        "foo.bar.baz", "foo.Qux", "rows", "i", "name", "data", "items", "company.name",
    }
    # plain names are unchanged without nested
    assert tpl.get_undeclared_template_variables() == {
        "foo", "rows", "i", "name", "data", "items", "company",
    }


def test_undeclared_variables_nested_with_context(make_docx):
    class Company:
        name = "ACME"

    tpl = DocxTemplate(make_docx(["{{ foo.bar }} {{ foo['Qux'] }} {{ c.name }} {{ c.city }}"]))
    context = {"foo": {"Qux": 1}, "c": Company()}
    assert tpl.get_undeclared_template_variables(context=context, nested=True) == {
        "foo.bar", "c.city",
    }
    context["foo"]["bar"] = 2
    assert tpl.get_undeclared_template_variables(context=context, nested=True) == {"c.city"}


def test_public_api_and_alias():
    assert docform.Template is docform.DocxTemplate
    assert docform.docx.Template is docform.docx.DocxTemplate
    assert isinstance(docform.__version__, str)


def test_libreoffice_core_properties_are_not_duplicated(make_docx, tmp_path):
    import collections
    import zipfile

    src = make_docx(["{{ x }}"], title="Report {{ x }}")
    lo = tmp_path / "libreoffice.docx"
    with zipfile.ZipFile(str(src)) as zin, zipfile.ZipFile(str(lo), "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "_rels/.rels":
                data = data.replace(
                    b"package/2006/relationships/metadata/core-properties",
                    b"officedocument/2006/relationships/metadata/core-properties",
                )
            zout.writestr(item, data)

    tpl = DocxTemplate(str(lo))
    tpl.render({"x": "ok"})
    out = tmp_path / "out.docx"
    tpl.save(str(out))

    with zipfile.ZipFile(str(out)) as z:
        names = collections.Counter(z.namelist())
    assert names["docProps/core.xml"] == 1
    assert Document(str(out)).core_properties.title == "Report ok"


FOOTNOTES_TPL = Path(__file__).resolve().parents[1] / "templates" / "footnotes_tpl.docx"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _footnote_doc(tmp_path, body_paragraphs, footnotes):
    """Rebuild footnotes_tpl.docx with the given body and footnote texts.

    body_paragraphs: list of plain text or ("text", footnote_id) tuples.
    footnotes: {id: text}
    """
    import zipfile
    from xml.sax.saxutils import escape

    def para(item):
        if isinstance(item, tuple):
            text, fid = item
            return (
                '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r>'
                '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                '<w:footnoteReference w:id="%s"/></w:r></w:p>' % (escape(text), fid)
            )
        return '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>' % escape(item)

    out = tmp_path / "fn.docx"
    with zipfile.ZipFile(str(FOOTNOTES_TPL)) as zin, zipfile.ZipFile(str(out), "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename).decode("utf-8")
            if item.filename == "word/document.xml":
                start = data.index("<w:body>") + len("<w:body>")
                sect = data.index("<w:sectPr", start)
                data = data[:start] + "".join(map(para, body_paragraphs)) + data[sect:]
            elif item.filename == "word/footnotes.xml":
                head, tail = data.split('<w:footnote w:id="1"', 1)
                tail = tail.split("</w:footnote>", 1)[1]
                notes = "".join(
                    '<w:footnote w:id="%s"><w:p><w:r><w:t xml:space="preserve">%s'
                    "</w:t></w:r></w:p></w:footnote>" % (fid, escape(text))
                    for fid, text in footnotes.items()
                )
                data = head + notes + tail
            zout.writestr(item, data.encode("utf-8"))
    return out


def _footnotes(doc):
    import zipfile

    with zipfile.ZipFile(io.BytesIO(_saved_bytes(doc))) as z:
        root = etree.fromstring(z.read("word/footnotes.xml"))
    notes = {}
    for note in root.findall("{%s}footnote" % W_NS):
        if note.get("{%s}type" % W_NS, "normal") == "normal":
            notes[note.get("{%s}id" % W_NS)] = "".join(note.itertext())
    return notes


def _saved_bytes(tpl):
    stream = io.BytesIO()
    tpl.save(stream)
    return stream.getvalue()


def test_footnote_of_a_removed_paragraph_is_dropped(tmp_path):
    path = _footnote_doc(
        tmp_path,
        [("kept", 1), "{%p if False %}", ("gone", 2), "{%p endif %}", ("also kept", 3)],
        {1: "one", 2: "two", 3: "three"},
    )
    tpl = DocxTemplate(str(path))
    tpl.render({})
    assert _footnotes(tpl) == {"1": "one", "3": "three"}


def test_footnotes_are_rendered_once_with_several_sections(tmp_path):
    path = _footnote_doc(tmp_path, [("text", 1)], {1: "{_{ x }_} and {{ y }}"})
    doc = Document(str(path))
    doc.add_section()
    doc.add_paragraph("second section")
    doc.save(str(path))

    tpl = DocxTemplate(str(path))
    tpl.render({"x": "evaluated", "y": "Y"})
    assert _footnotes(tpl) == {"1": "{{ x }} and Y"}


def _body_doc(tmp_path, body_xml):
    """Save a document whose body is the given paragraph XML."""
    import zipfile

    base = tmp_path / "base.docx"
    Document().save(str(base))
    out = tmp_path / "bm.docx"
    with zipfile.ZipFile(str(base)) as zin, zipfile.ZipFile(str(out), "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml = data.decode("utf-8")
                start = xml.index("<w:body>") + len("<w:body>")
                xml = xml[:start] + body_xml + xml[xml.index("<w:sectPr", start):]
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    return out


def _run(text):
    return '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % text


BM_START = '<w:bookmarkStart w:id="7" w:name="_Ref1"/>'
BM_END = '<w:bookmarkEnd w:id="7"/>'


def _bookmarks(doc):
    body = doc.element.body
    return (
        body.xpath(".//w:bookmarkStart/@w:name"),
        body.xpath(".//w:bookmarkEnd/@w:id"),
    )


@pytest.mark.parametrize(
    "body",
    [
        # start between the runs of a tag Word split
        "<w:p>" + _run("a {% if x ") + BM_START + _run("%} b {% endif %}") + BM_END + "</w:p>",
        # end inside the runs of the closing tag
        "<w:p>" + BM_START + _run("a {% if x %} b {% en") + BM_END + _run("dif %}") + "</w:p>",
        # on a paragraph tag line
        "<w:p>" + BM_START + _run("{%p if x %}") + BM_END + "</w:p><w:p>" + _run("b")
        + "</w:p><w:p>" + _run("{%p endif %}") + "</w:p>",
        # in the span {%- merges away
        "<w:p>" + BM_START + _run("a") + BM_END + "</w:p><w:p>" + _run("{%- if x %} b{% endif %}") + "</w:p>",
    ],
    ids=["inside-open-tag", "inside-close-tag", "paragraph-tag", "dash-merge"],
)
def test_bookmarks_survive_tag_clean_up(tmp_path, body):
    tpl = DocxTemplate(str(_body_doc(tmp_path, body)))
    tpl.render({"x": True})
    assert _bookmarks(saved(tpl)) == (["_Ref1"], ["7"])


def test_bookmark_in_a_loop_is_kept_once(tmp_path):
    body = (
        "<w:p>" + _run("{%p for i in items %}") + "</w:p>"
        "<w:p>" + BM_START + _run("{{ i }}") + BM_END + "</w:p>"
        "<w:p>" + _run("{%p endfor %}") + "</w:p>"
    )
    tpl = DocxTemplate(str(_body_doc(tmp_path, body)))
    tpl.render({"items": ["a", "b", "c"]})
    out = saved(tpl)
    assert body_text(out) == ["a", "b", "c"]
    assert _bookmarks(out) == (["_Ref1"], ["7"])


def test_render_and_save_async(make_docx):
    import asyncio
    from concurrent.futures import ThreadPoolExecutor

    async def build(name, executor):
        tpl = DocxTemplate(make_docx(["Hello {{ name }}"]))
        await tpl.render_async({"name": name}, executor=executor)
        out = io.BytesIO()
        await tpl.save_async(out, executor=executor)
        out.seek(0)
        return body_text(Document(out))

    async def main():
        with ThreadPoolExecutor(max_workers=2) as pool:
            return await asyncio.gather(*(build(n, pool) for n in ("Ann", "Bob", "Cy")))

    assert asyncio.run(main()) == [["Hello Ann"], ["Hello Bob"], ["Hello Cy"]]


def test_render_async_raises_render_errors(make_docx):
    import asyncio

    tpl = DocxTemplate(make_docx(["{{ broken "]))
    with pytest.raises(TemplateSyntaxError):
        asyncio.run(tpl.render_async({}))
