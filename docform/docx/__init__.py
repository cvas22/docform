# -*- coding: utf-8 -*-
"""Word (.docx) templates."""

# flake8: noqa
from .inline_image import InlineImage
from .listing import Listing
from .richtext import RichText, R, RichTextParagraph, RP
from .template import DocxTemplate

Template = DocxTemplate

try:
    from .subdoc import Subdoc
except ImportError:
    pass
