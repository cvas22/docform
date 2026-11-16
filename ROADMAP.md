# Roadmap

Docform's goal is a dependable way to fill document templates with data from Python. This is the rough plan, not a promise of dates. Ideas and votes are welcome in [issues](https://github.com/cvas22/docform/issues).

## Now

- Bring over the reviewed fixes that are still open on python-docx-template, with credit to their authors
- Work through the open bug reports, starting with the ones that produce corrupted documents
- Grow the unit tests so every feature has assertions, not just a sample script

## Next

- Type hints across the public API
- Safer defaults for untrusted templates, such as an opt-in sandboxed rendering mode
- Clearer errors that point to where in the template rendering failed
- Documentation published as a website

## Later

- More formats next to Word, each in its own module: Excel (`docform.xlsx`), PowerPoint (`docform.pptx`) and OpenDocument text
- Optional PDF output through an external converter
