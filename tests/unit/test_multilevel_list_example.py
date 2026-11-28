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
    template_num_id = Document(template).paragraphs[0]._p.pPr.numPr.numId.val
    assert num_ids == {template_num_id}
    numbering = doc.part.numbering_part.element
    abstract_id = numbering.num_having_numId(template_num_id).abstractNumId.val
    level_path = f'./w:abstractNum[@w:abstractNumId="{abstract_id}"]/w:lvl'
    assert len(numbering.xpath(level_path)) == 3
    assert not numbering.xpath(level_path + "/w:lvlRestart")
    assert [paragraph._p.pPr.numPr.ilvl.val for paragraph in doc.paragraphs] == [
        level for level, _ in code["ITEMS"]
    ]
