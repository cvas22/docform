# Docform

[![Python](https://img.shields.io/badge/python-3.7%2B-blue)](pyproject.toml)
[![License: LGPL-2.1](https://img.shields.io/badge/license-LGPL--2.1-green)](LICENSE.txt)

**Fill document templates with data.**

Design the document in Word with whatever layout, styles, headers and images you need, and put [Jinja2](https://jinja.palletsprojects.com/) tags where the data goes. Docform renders it into a finished `.docx`, so you can generate contracts, invoices, reports, certificates and letters in bulk from Python.

## Install

```bash
pip install docform

# with subdocument support
pip install "docform[subdoc]"
```

## Quickstart

In Word, write a template like this and save it as `invoice_template.docx`:

```text
Invoice for {{ customer }}

{%tr for item in items %}
| {{ item.name }} | {{ item.qty }} |
{%tr endfor %}

Total: {{ total }}
```

Then render it:

```python
from docform import DocxTemplate

doc = DocxTemplate("invoice_template.docx")
doc.render({
    "customer": "Acme",
    "items": [{"name": "Widget", "qty": 3}, {"name": "Gadget", "qty": 1}],
    "total": "$120.00",
})
doc.save("invoice_acme.docx")
```

Everything you set up in Word, like fonts, logos and page layout, stays intact. Only the tags are replaced.

## How it works

```mermaid
flowchart LR
    T["Word template<br/>(.docx with Jinja2 tags)"] --> R{{"Docform<br/>render()"}}
    D["Your data<br/>(dict, JSON, database rows)"] --> R
    R --> O["Finished document<br/>(.docx)"]
```

## Example

This invoice comes from [`examples/invoice`](examples/invoice). The template on the left is a normal Word file. Running `python render.py` turns it into the document on the right.

<table>
  <tr>
    <th>Template (<code>invoice_template.docx</code>)</th>
    <th>Rendered (<code>invoice_acme.docx</code>)</th>
  </tr>
  <tr>
    <td><img src="docs/images/template.png" alt="Invoice template with Jinja2 tags" width="420"></td>
    <td><img src="docs/images/rendered.png" alt="Rendered invoice with customer data and four line items" width="420"></td>
  </tr>
</table>

A few things to notice in the template:

- `{%tr for item in items %}` and `{%tr endfor %}` sit in their own table rows. Docform repeats the row between them once per item and removes the tag rows.
- `{%p if paid %}` works the same way for paragraphs, so the payment note switches between "PAID" and a reminder.
- Formatting on a tag, like the bold blue `{{ total }}`, carries over to the value.

## Features

- Full Jinja2 syntax: variables, conditions, loops, filters and custom filters
- Paragraph, table row, table cell and run level tags (`{%p %}`, `{%tr %}`, `{%tc %}`, `{%r %}`) that remove the surrounding Word element so loops don't leave empty rows behind
- Rich text with colors, fonts, bold, italic, links and right-to-left text
- Inline images, plus replacing pictures and embedded objects already in the template
- Subdocuments: build parts with python-docx or merge in existing `.docx` files
- Headers, footers, footnotes and document properties
- Cell background colors and horizontal and vertical cell merges
- Listing undeclared template variables before rendering
- A command line interface for JSON-driven rendering

## Command line

```bash
python -m docform template.docx data.json output.docx
```

Use `-o` to overwrite an existing output file and `-q` for quiet mode.

## Rendering untrusted templates

Templates run Jinja2 code. If your users can upload their own templates, render them with a sandboxed environment:

```python
from jinja2.sandbox import SandboxedEnvironment

doc.render(context, jinja_env=SandboxedEnvironment(autoescape=True))
```

## Formats

Word (`.docx`) is the only format today. Each format lives in its own module, so new ones can be added without breaking existing code:

```python
from docform.docx import Template
```

## Documentation

The full template syntax and API reference are in [`docs/`](docs/index.rst). [`examples/`](examples/) has complete, runnable examples, and [`tests/`](tests/) has a small script for every feature.

## Adopters

Using Docform in production or in an open source project? Add yourself to [ADOPTERS.md](ADOPTERS.md).

## Citing

If you use Docform in research, please cite it. GitHub's "Cite this repository" button reads [CITATION.cff](CITATION.cff).

## Credits

Docform started as a fork of [python-docx-template](https://github.com/elapouya/python-docx-template) by Eric Lapouyade and builds on [python-docx](https://github.com/python-openxml/python-docx) and [Jinja2](https://github.com/pallets/jinja).

## License

[LGPL-2.1](LICENSE.txt), the same as python-docx-template.
