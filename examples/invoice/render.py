"""Render the invoice example: python render.py"""
import json
from pathlib import Path

from docform import DocxTemplate

here = Path(__file__).parent
context = json.loads((here / "data.json").read_text(encoding="utf-8"))

doc = DocxTemplate(here / "invoice_template.docx")
doc.render(context)
doc.save(here / "invoice_acme.docx")
print("Wrote", here / "invoice_acme.docx")
