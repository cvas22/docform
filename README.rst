=======
docform
=======

Fill document templates with data.

You design the document in Word, with all the layout, styles, headers and
images you want, and put Jinja2 tags where the data goes. docform renders it
into a finished .docx, so you can generate contracts, invoices, reports or
letters in bulk from Python.

.. code-block:: python

    from docform import DocxTemplate

    doc = DocxTemplate("invoice_template.docx")
    doc.render({"customer": "Acme", "items": [{"name": "Widget", "qty": 3}]})
    doc.save("invoice_acme.docx")

Besides plain variables, templates can loop over table rows, columns and
paragraphs, insert images and rich text, merge other .docx files as
subdocuments, and fill headers, footers and document properties.

Word (.docx) is the only format today. Format-specific code lives under its
own module (``docform.docx``) so other formats can be added next to it.

Install
-------

.. code-block:: bash

    pip install docform

    # with subdocument support
    pip install "docform[subdoc]"

There's also a command line interface:

.. code-block:: bash

    python -m docform template.docx data.json output.docx

Documentation
-------------

See the ``docs`` folder for the full template syntax and API.

Credits
-------

docform started as a fork of
`python-docx-template <https://github.com/elapouya/python-docx-template>`_
by Eric Lapouyade, and builds on
`python-docx <https://github.com/python-openxml/python-docx>`_ and
`Jinja2 <https://github.com/pallets/jinja>`_. It's distributed under the same
LGPL-2.1 license.
