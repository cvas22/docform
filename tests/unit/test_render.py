import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from jinja2 import Environment
from jinja2.exceptions import TemplateSyntaxError, UndefinedError
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


def test_render_twice_with_init(make_docx):
    tpl = DocxTemplate(make_docx(["Hi {{ n }}"]))
    tpl.render({"n": 1})
    assert body_text(saved(tpl)) == ["Hi 1"]
    tpl.init_docx()
    tpl.render({"n": 2})
    assert body_text(saved(tpl)) == ["Hi 2"]


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


def test_public_api_and_alias():
    assert docform.Template is docform.DocxTemplate
    assert docform.docx.Template is docform.docx.DocxTemplate
    assert isinstance(docform.__version__, str)
