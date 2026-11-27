"""Render a numbered list with Word-managed restarts."""

from pathlib import Path

from docform import DocxTemplate

ITEMS = (
    (0, "Prepare the report"),
    (1, "Collect the data"),
    (1, "Review the figures"),
    (2, "Check the axis labels"),
    (2, "Check the units"),
    (0, "Publish the report"),
    (1, "Prepare the release"),
    (2, "Check the metadata"),
)


def render(template_path, output_path, items):
    tpl = DocxTemplate(template_path)
    # Read the list instance from the numbered placeholder in the template.
    placeholder = next(p for p in tpl.get_docx().paragraphs if p.text == "{{p items }}")
    num_id = placeholder._p.pPr.numPr.numId.val

    subdoc = tpl.new_subdoc()
    for level, text in items:
        paragraph = subdoc.add_paragraph(text)
        numbering = paragraph._p.get_or_add_pPr().get_or_add_numPr()
        numbering.get_or_add_ilvl().val = level
        numbering.get_or_add_numId().val = num_id

    tpl.render({"items": subdoc})
    tpl.save(output_path)


if __name__ == "__main__":
    here = Path(__file__).resolve().parent
    output = here / "multilevel_list.docx"
    render(here / "template.docx", output, ITEMS)
    print("Wrote", output)
