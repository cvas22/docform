---
name: Bug report
about: Something renders wrong, crashes, or produces a broken document
labels: bug
---

## What happened

<!-- What went wrong? Include the full error message and traceback if there is one. -->

## What you expected

## How to reproduce

<!-- A small runnable script plus the .docx template. Attach the template to the issue,
with anything private removed. Issues we can't reproduce are hard to fix. -->

```python
from docform import DocxTemplate

doc = DocxTemplate("template.docx")
doc.render({...})
doc.save("out.docx")
```

## Versions

- Docform:
- Python:
- python-docx:
- OS, and the app you opened the result in (Word, LibreOffice, ...):
