import runpy
from pathlib import Path

from docx import Document


def test_multilevel_list_example_keeps_one_list_with_explicit_levels(tmp_path):
    example = Path(__file__).resolve().parents[2] / "examples" / "multilevel_list"
    code = runpy.run_path(str(example / "render.py"))
    template = example / "template.docx"
    output = tmp_path / "multilevel_list.docx"

    code["render"](template, output, code["ITEMS"])

    doc = Document(output)
    assert [paragraph.text for paragraph in doc.paragraphs] == [
        text for _, text in code["ITEMS"]
    ]
    num_ids = {paragraph._p.pPr.numPr.numId.val for paragraph in doc.paragraphs}
    assert num_ids == {Document(template).paragraphs[0]._p.pPr.numPr.numId.val}
    assert [paragraph._p.pPr.numPr.ilvl.val for paragraph in doc.paragraphs] == [
        level for level, _ in code["ITEMS"]
    ]
