# -*- coding: utf-8 -*-
"""
docform: fill document templates with data.

Based on python-docx-template by Eric Lapouyade.
"""

__version__ = "0.20.2"

# flake8: noqa
from .docx import InlineImage, Listing, RichText, R, RichTextParagraph, RP
from .docx import DocxTemplate, Template

try:
    from .docx import Subdoc
except ImportError:
    pass
