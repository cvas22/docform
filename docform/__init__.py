# -*- coding: utf-8 -*-
"""
Docform: fill document templates with data.
"""

__version__ = "1.0.2"

# flake8: noqa
from .docx import InlineImage, Listing, RichText, R, RichTextParagraph, RP
from .docx import DocxTemplate, Template

try:
    from .docx import Subdoc
except ImportError:
    pass
