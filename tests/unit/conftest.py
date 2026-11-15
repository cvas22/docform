import io

import pytest
from docx import Document


@pytest.fixture
def make_docx(tmp_path):
    """Build a small .docx template and return its path.

    ``paragraphs`` is a list of strings, one paragraph each. ``table`` is a
    list of rows (lists of cell strings). ``header`` and ``footer`` set the
    text of the default section header and footer.
    """
    counter = {"n": 0}

    def _make(paragraphs=(), table=None, header=None, footer=None, title=None):
        doc = Document()
        for text in paragraphs:
            doc.add_paragraph(text)
        if table:
            t = doc.add_table(rows=len(table), cols=len(table[0]))
            for r, row in enumerate(table):
                for c, text in enumerate(row):
                    t.rows[r].cells[c].text = text
        section = doc.sections[0]
        if header is not None:
            section.header.paragraphs[0].text = header
        if footer is not None:
            section.footer.paragraphs[0].text = footer
        if title is not None:
            doc.core_properties.title = title
        counter["n"] += 1
        path = tmp_path / "tpl{}.docx".format(counter["n"])
        doc.save(str(path))
        return path

    return _make


def saved(tpl):
    """Save a rendered template to memory and open it with python-docx."""
    buf = io.BytesIO()
    tpl.save(buf)
    buf.seek(0)
    return Document(buf)


def body_text(doc):
    return [p.text for p in doc.paragraphs]


def table_rows(doc, index=0):
    return [[c.text for c in row.cells] for row in doc.tables[index].rows]
