# -*- coding: utf-8 -*-
"""
Created : 2015-03-12
"""
from __future__ import annotations

from os import PathLike
from typing import TYPE_CHECKING, Any, Optional, IO, Union, Dict, Set
import asyncio
import functools
import io
from lxml import etree
from docx import Document
from docx.opc.oxml import parse_xml
from docx.opc.part import XmlPart
import docx.oxml.ns
from docx.opc.constants import RELATIONSHIP_TYPE as REL_TYPE
from jinja2 import Environment, meta, nodes
from jinja2.sandbox import SandboxedEnvironment
from jinja2.exceptions import TemplateError

try:
    from html import escape  # noqa: F401
except ImportError:
    # cgi.escape is deprecated in python 3.7
    from cgi import escape  # noqa: F401
import re
import warnings
import binascii
import os
import zipfile

if TYPE_CHECKING:
    from .subdoc import Subdoc


def _lookup_path(node):
    """Return (root name, [keys]) for foo.bar['baz'], or None."""
    keys = []
    while True:
        if isinstance(node, nodes.Getattr):
            keys.append(node.attr)
            node = node.node
        elif isinstance(node, nodes.Getitem):
            if not isinstance(node.arg, nodes.Const):
                # foo[i]: only the part before the dynamic key is known
                keys = []
            else:
                keys.append(node.arg.value)
            node = node.node
        elif isinstance(node, nodes.Name):
            return node.name, keys[::-1]
        else:
            return None


def _variable_paths(ast, names):
    """Collect the lookup paths, as tuples, rooted at the given names.

    A path that is only the start of a longer one is dropped, and method
    calls like foo.items() stop at foo.
    """
    found = set()

    def visit(node, called=False):
        if isinstance(node, (nodes.Getattr, nodes.Getitem, nodes.Name)):
            path = _lookup_path(node)
            if path is not None and path[0] in names:
                keys = path[1]
                if called and keys:
                    keys = keys[:-1]
                found.add((path[0],) + tuple(keys))
                # keep looking for lookups inside dynamic keys
                for child in node.find_all(nodes.Getitem):
                    if not isinstance(child.arg, nodes.Const):
                        visit(child.arg)
                return
        for child in node.iter_child_nodes():
            visit(child, called=isinstance(node, nodes.Call) and child is node.node)

    visit(ast)
    found |= {(name,) for name in names}
    return {
        p for p in found
        if not any(len(q) > len(p) and q[: len(p)] == p for q in found)
    }


def _context_has_path(context, path):
    value = context
    for key in path:
        if isinstance(value, dict) and key in value:
            value = value[key]
        elif isinstance(key, str) and hasattr(value, key):
            value = getattr(value, key)
        else:
            try:
                value = value[key]
            except Exception:
                return False
    return True


class DocxTemplate(object):
    """Class for managing docx files as they were jinja2 templates"""

    HEADER_URI = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    )
    FOOTER_URI = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    )

    def __init__(self, template_file: Union[IO[bytes], str, PathLike]) -> None:
        self.template_file = template_file
        self.reset_replacements()
        self.docx = None
        self.is_rendered = False
        self.is_saved = False
        self.allow_missing_pics = False

    # LibreOffice links the core properties with this relationship type
    # instead of the standard one, so python-docx doesn't find them.
    LIBREOFFICE_CORE_PROPERTIES = (
        "http://schemas.openxmlformats.org/officedocument/2006/relationships"
        "/metadata/core-properties"
    )

    def init_docx(self, reload: bool = True):
        if not self.docx or (self.is_rendered and reload):
            self.docx = Document(self.template_file)
            self.fix_core_properties_rel()
            self.is_rendered = False

    def fix_core_properties_rel(self):
        """Relabel a LibreOffice core properties relationship as the standard one.

        Otherwise python-docx adds a second docProps/core.xml when the document
        properties are rendered, and LibreOffice refuses the file.
        """
        rels = self.docx.part.package.rels
        types = {rel.reltype for rel in rels.values()}
        if REL_TYPE.CORE_PROPERTIES in types:
            return
        for rel in rels.values():
            if rel.reltype == self.LIBREOFFICE_CORE_PROPERTIES:
                rel._reltype = REL_TYPE.CORE_PROPERTIES
                return

    def render_init(self):
        self.init_docx()
        self.pic_map = {}
        self.current_rendering_part = None
        self.docx_ids_index = 1000
        self.is_saved = False

    def __getattr__(self, name):
        # Forward unknown attributes to the python-docx Document, loading it
        # first so they also work before render() has been called.
        if name == "docx" or name.startswith("__"):
            raise AttributeError(name)
        self.init_docx(reload=False)
        return getattr(self.docx, name)

    def xml_to_string(self, xml, encoding="unicode"):
        # Be careful : pretty_print MUST be set to False, otherwise patch_xml()
        # won't work properly
        return etree.tostring(xml, encoding="unicode", pretty_print=False)

    def get_docx(self):
        self.init_docx()
        return self.docx

    def get_xml(self):
        body = self.docx._element.body
        self.rescue_bookmarks(body)
        return self.xml_to_string(body)

    def rescue_bookmarks(self, root):
        """Move bookmarks out of XML that patch_xml() is about to delete.

        Word cross-references (REF fields) point at bookmarks, and a missing one
        shows "Error! Reference source not found". patch_xml() drops the XML
        between the pieces of a tag Word split over several runs, the whole
        paragraph, row or cell of a {%p %}, {%tr %} or {%tc %} tag, and the XML
        that {%- and -%} merge away. Bookmarks there move just outside, which
        keeps them in the same paragraph for paragraph-number references.
        """
        w = docx.oxml.ns.nsmap["w"]
        t_tag, p_tag = "{%s}t" % w, "{%s}p" % w
        marks = ("{%s}bookmarkStart" % w, "{%s}bookmarkEnd" % w)

        def text_of(p):
            return "".join(t.text or "" for t in p.iter(t_tag))

        def is_end_tag(tag_text):
            # {% endfor %}, {%p endif %}, {%- endwith -%}: anything moved next
            # to a closing tag goes after it, so it lands outside the block
            # instead of being repeated by a loop.
            m = re.match(r"[{][%{#]-?\s*(?:(?:p|tr|tc|r)\s)?\s*(\w*)", tag_text)
            return bool(m and m.group(1).startswith("end"))

        def put(marks, anchor, after):
            for mark in reversed(marks) if after else marks:
                (anchor.addnext if after else anchor.addprevious)(mark)

        # Paragraph, row and cell tags replace their whole container.
        for p in list(root.iter(p_tag)):
            text = text_of(p)
            m = re.search(r"({%|{{)(p|tr|tc) ", text)
            if not m:
                continue
            target = p
            if m.group(2) in ("tr", "tc"):
                for anc in p.iterancestors("{%s}%s" % (w, m.group(2))):
                    target = anc
                    break
            moved = list(p.iter(*marks))
            if moved:
                put(moved, target, is_end_tag(text[m.start():]))

        # Tags split over runs, and the spans {%- and -%} merge away.
        inside, open_run, tag_text, tag_marks = False, None, "", []
        dash_marks, dash_anchor = [], None
        prev_char, prev_run = "", None
        last_text_run, since_text = None, []
        dash_close_run, since_dash = None, []
        for el in list(root.iter(t_tag, *marks)):
            if el.tag in marks:
                if inside:
                    tag_marks.append(el)
                    continue
                since_text.append(el)
                if dash_close_run is not None:
                    since_dash.append(el)
                continue
            run = el.getparent()
            if dash_close_run is not None:
                put(since_dash, dash_close_run, False)
                dash_close_run, since_dash = None, []
            text = el.text or ""
            for i, char in enumerate(text):
                pair = prev_char + char
                if inside:
                    tag_text += char
                if not inside and pair in ("{%", "{{", "{#"):
                    inside, open_run, tag_text, tag_marks = True, prev_run if i == 0 else run, pair, []
                    if pair == "{%" and text[i + 1: i + 2] == "-" and last_text_run is not None:
                        dash_marks, dash_anchor, since_text = since_text, last_text_run, []
                elif inside and pair in ("%}", "}}", "#}"):
                    inside = False
                    end_tag = is_end_tag(tag_text)
                    if tag_marks:
                        put(tag_marks, run if end_tag else open_run, end_tag)
                    if dash_marks:
                        put(dash_marks, run if end_tag else dash_anchor, end_tag)
                        dash_marks, dash_anchor = [], None
                    if pair == "%}" and tag_text[-3:-2] == "-":
                        dash_close_run = run
                prev_char, prev_run = char, run
            last_text_run, since_text = run, []

    def write_xml(self, filename):
        # XML defaults to UTF-8 : do not rely on the locale default encoding,
        # it is not UTF-8 on Windows and raises UnicodeEncodeError as soon as
        # the document contains a character it cannot represent.
        with open(filename, "w", encoding="utf-8") as fh:
            fh.write(self.get_xml())

    @staticmethod
    def _keep_section_break(m):
        """Replace a {%p %} paragraph by its tag, keeping a section break.

        A paragraph that ends a section holds the section's settings (page
        size, margins, orientation) in w:pPr/w:sectPr. Dropping it with the
        tag would merge the section into the next one. Keep an empty
        paragraph with those properties outside the block: before an
        opening tag and after a closing one, so it's never skipped or
        repeated by the tag.
        """
        tag = m.group(1) + " " + m.group(2)
        ppr = re.search(
            r"<w:pPr>(?:(?!</w:pPr>).)*<w:sectPr[ >/].*?</w:pPr>", m.group(0), flags=re.DOTALL
        )
        if ppr is None:
            return tag
        section_break = "<w:p>" + ppr.group(0) + "</w:p>"
        if re.match(r"-?\s*end", m.group(2)):
            return tag + section_break
        return section_break + tag

    def patch_xml(self, src_xml):
        """Make a lots of cleaning to have a raw xml understandable by jinja2 :
        strip all unnecessary xml tags, manage table cell background color and colspan,
        unescape html entities, etc..."""

        # Word may insert proofing marks, bookmarks or run formatting between
        # text nodes. Paragraph boundaries, tabs and line breaks end a delimiter.
        boundary = r"</?w:(?:p|tab|br|cr)(?=[\s/>])"
        split_run_gap = rf"(?:</w:t>(?:(?!{boundary})<[^>]+>)*<w:t(?: [^>]*)?>)?"
        for escaped_delimiter in ("{_{", "}_}", "{_%", "%_}"):
            pattern = split_run_gap.join(map(re.escape, escaped_delimiter))
            src_xml = re.sub(pattern, escaped_delimiter, src_xml)

        # replace {<something>{ by {{   ( works with {{ }} {% and %} {# and #})
        src_xml = re.sub(
            r"(?<={)(<[^>]*>)+(?=[\{%\#])|(?<=[%\}\#])(<[^>]*>)+(?=\})",
            "",
            src_xml,
            flags=re.DOTALL,
        )

        # replace {{<some tags>jinja2 stuff<some other tags>}} by {{jinja2 stuff}}
        # same thing with {% ... %} and {# #}
        # "jinja2 stuff" could a variable, a 'if' etc... anything jinja2 will understand
        def striptags(m):
            return re.sub(
                "</w:t>.*?(<w:t>|<w:t [^>]*>)", "", m.group(0), flags=re.DOTALL
            )

        src_xml = re.sub(
            r"{%(?:(?!%}).)*|{#(?:(?!#}).)*|{{(?:(?!}}).)*",
            striptags,
            src_xml,
            flags=re.DOTALL,
        )

        # manage table cell colspan
        def colspan(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(
                r"<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>",
                "",
                cell_xml,
                flags=re.DOTALL,
            )
            cell_xml = re.sub(r"<w:gridSpan[^/]*/>", "", cell_xml, count=1)
            # a self-closing <w:tcPr/> would leave the inserted gridSpan outside of it
            cell_xml = re.sub(
                r"<w:tcPr(\s[^>]*)?/>",
                lambda mm: "<w:tcPr%s></w:tcPr>" % (mm.group(1) or ""),
                cell_xml,
                count=1,
            )
            return re.sub(
                r"(<w:tcPr[^>]*>)",
                r'\1<w:gridSpan w:val="{{%s}}"/>' % m.group(2),
                cell_xml,
            )

        src_xml = re.sub(
            r"(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*colspan\s+([^%]*)\s*%}(.*?</w:tc>)",
            colspan,
            src_xml,
            flags=re.DOTALL,
        )

        # manage table cell background color
        def cellbg(m):
            cell_xml = m.group(1) + m.group(3)
            cell_xml = re.sub(
                r"<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>",
                "",
                cell_xml,
                flags=re.DOTALL,
            )
            cell_xml = re.sub(r"<w:shd[^/]*/>", "", cell_xml, count=1)
            # a self-closing <w:tcPr/> would leave the inserted shd outside of it
            cell_xml = re.sub(
                r"<w:tcPr(\s[^>]*)?/>",
                lambda mm: "<w:tcPr%s></w:tcPr>" % (mm.group(1) or ""),
                cell_xml,
                count=1,
            )
            return re.sub(
                r"(<w:tcPr[^>]*>)",
                r'\1<w:shd w:val="clear" w:color="auto" w:fill="{{%s}}"/>' % m.group(2),
                cell_xml,
            )

        src_xml = re.sub(
            r"(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?</w:tc>)",
            cellbg,
            src_xml,
            flags=re.DOTALL,
        )

        # ensure space preservation
        src_xml = re.sub(
            r"<w:t>((?:(?!<w:t>).)*)({{.*?}}|{%.*?%})",
            r'<w:t xml:space="preserve">\1\2',
            src_xml,
            flags=re.DOTALL,
        )
        src_xml = re.sub(
            r"({{r\s.*?}}|{%r\s.*?%})",
            r'</w:t></w:r><w:r><w:t xml:space="preserve">\1</w:t></w:r><w:r><w:t xml:space="preserve">',
            src_xml,
            flags=re.DOTALL,
        )

        # Merging never crosses a table, row or cell edge, that would drop the
        # table's own tags. The tag then stays put and only whitespace goes.
        table_edge = r"</?w:(?:tbl|tr|tc)[\s>]"
        # {%- will merge with previous paragraph text
        src_xml = re.sub(
            r"</w:t>(?:(?!</w:t>|" + table_edge + r").)*?{%-",
            "{%",
            src_xml,
            flags=re.DOTALL,
        )
        # -%} will merge with next paragraph text
        src_xml = re.sub(
            r"-%}(?:(?!<w:t[ >]|{%|{{|" + table_edge + r").)*?<w:t(?:\s[^>]*)?>",
            "%}",
            src_xml,
            flags=re.DOTALL,
        )

        for y in ["tr", "tc", "p", "r"]:
            # replace into xml code the row/paragraph/run containing
            # {%y xxx %} or {{y xxx}} template tag
            # by {% xxx %} or {{ xx }} without any surrounding <w:y> tags :
            # This is mandatory to have jinja2 generating correct xml code
            pat = (
                r"<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({%%|{{)%(y)s ((?:(?!%%}|}}).)*(?:%%}|}})).*?</w:%(y)s>"
                % {"y": y}
            )
            if y == "p":
                src_xml = re.sub(pat, self._keep_section_break, src_xml, flags=re.DOTALL)
            else:
                src_xml = re.sub(pat, r"\1 \2", src_xml, flags=re.DOTALL)

        for y in ["tr", "tc", "p"]:
            # same thing, but for {#y xxx #} (but not where y == 'r', since that
            # makes less sense to use comments in that context
            pat = (
                r"<w:%(y)s[ >](?:(?!<w:%(y)s[ >]).)*({#)%(y)s ([^}#]*(?:#})).*?</w:%(y)s>"
                % {"y": y}
            )
            src_xml = re.sub(pat, r"\1 \2", src_xml, flags=re.DOTALL)

        # add vMerge
        # use {% vm %} to make this table cell and its copies
        # be vertically merged within a {% for %}
        def v_merge_tc(m):
            def v_merge(m1):
                return (
                    '<w:vMerge w:val="{% if loop.first %}restart{% else %}continue{% endif %}"/>'
                    + m1.group(1)  # Everything between ``</w:tcPr>`` and ``<w:t>``.
                    + "{% if loop.first %}"
                    + m1.group(2)  # Everything before ``{% vm %}``.
                    + m1.group(3)  # Everything after ``{% vm %}``.
                    + "{% endif %}"
                    + m1.group(4)  # ``</w:t>``.
                )

            return re.sub(
                r"(</w:tcPr[ >].*?<w:t(?:\s[^>]*)?>)(.*?)(?:{%\s*vm\s*%})(.*?)(</w:t>)",
                v_merge,
                m.group(),
                # Everything between ``</w:tc>`` and ``</w:tc>`` with ``{% vm %}`` inside.
                flags=re.DOTALL,
            )

        src_xml = re.sub(
            r"<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*vm\s*%}.*?</w:tc[ >]",
            v_merge_tc,
            src_xml,
            flags=re.DOTALL,
        )

        # Use ``{% hm %}`` to make table cell become horizontally merged within
        # a ``{% for %}``.
        def h_merge_tc(m):
            xml_to_patch = (
                m.group()
            )  # Everything between ``</w:tc>`` and ``</w:tc>`` with ``{% hm %}`` inside.

            def with_gridspan(m1):
                return (
                    m1.group(1)  # ``w:gridSpan w:val="``.
                    + "{{ "
                    + m1.group(2)
                    + " * loop.length }}"  # Content of ``w:val``, multiplied by loop length.
                    + m1.group(3)  # Closing quotation mark.
                )

            def without_gridspan(m2):
                return (
                    '<w:gridSpan w:val="{{ loop.length }}"/>'
                    + m2.group(1)  # Everything between ``</w:tcPr>`` and ``<w:t>``.
                    + m2.group(2)  # Everything before ``{% hm %}``.
                    + m2.group(3)  # Everything after ``{% hm %}``.
                    + m2.group(4)  # ``</w:t>``.
                )

            if re.search(r"w:gridSpan", xml_to_patch):
                # Simple case, there's already ``gridSpan``, multiply its value.

                xml = re.sub(
                    r'(w:gridSpan w:val=")(\d+)(")',
                    with_gridspan,
                    xml_to_patch,
                    flags=re.DOTALL,
                )
                xml = re.sub(
                    r"{%\s*hm\s*%}",
                    "",
                    xml,  # Patched xml.
                    flags=re.DOTALL,
                )
            else:
                # There're no ``gridSpan``, add one.
                xml = re.sub(
                    r"(</w:tcPr[ >].*?<w:t(?:\s[^>]*)?>)(.*?)(?:{%\s*hm\s*%})(.*?)(</w:t>)",
                    without_gridspan,
                    xml_to_patch,
                    flags=re.DOTALL,
                )

            # Discard every other cell generated in loop.
            return "{% if loop.first %}" + xml + "{% endif %}"

        src_xml = re.sub(
            r"<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*hm\s*%}.*?</w:tc[ >]",
            h_merge_tc,
            src_xml,
            flags=re.DOTALL,
        )

        def clean_tags(m):
            return (
                m.group(0)
                .replace(r"&#8216;", "'")
                .replace("&lt;", "<")
                .replace("&gt;", ">")
                .replace("“", '"')
                .replace("”", '"')
                .replace("‘", "'")
                .replace("’", "'")
            )

        src_xml = re.sub(r"(?<=\{[\{%])(.*?)(?=[\}%]})", clean_tags, src_xml)

        return src_xml

    def render_xml_part(self, src_xml, part, context, jinja_env=None):
        src_xml = re.sub(r"<w:p([ >])", r"\n<w:p\1", src_xml)
        try:
            self.current_rendering_part = part
            if jinja_env is None:
                jinja_env = SandboxedEnvironment()
            template = jinja_env.from_string(src_xml)
            dst_xml = template.render(context)
        except TemplateError as exc:
            if hasattr(exc, "lineno") and exc.lineno is not None:
                line_number = max(exc.lineno - 4, 0)
                exc.docx_context = map(
                    lambda x: re.sub(r"<[^>]+>", "", x),
                    src_xml.splitlines()[line_number: (line_number + 7)],  # fmt: skip
                )

            raise exc
        dst_xml = re.sub(r"\n<w:p([ >])", r"<w:p\1", dst_xml)
        dst_xml = (
            dst_xml.replace("{_{", "{{")
            .replace("}_}", "}}")
            .replace("{_%", "{%")
            .replace("%_}", "%}")
        )
        dst_xml = self.resolve_listing(dst_xml)
        return dst_xml

    def render_properties(
        self, context: Dict[str, Any], jinja_env: Optional[Environment] = None
    ) -> None:
        # List of string attributes of docx.opc.coreprops.CoreProperties which are strings.
        # It seems that some attributes cannot be written as strings. Those are commented out.
        properties = [
            "author",
            # 'category',
            "comments",
            # 'content_status',
            "identifier",
            # 'keywords',
            "language",
            # 'last_modified_by',
            "subject",
            "title",
            # 'version',
        ]
        if jinja_env is None:
            jinja_env = SandboxedEnvironment()

        for prop in properties:
            initial = getattr(self.docx.core_properties, prop)
            template = jinja_env.from_string(initial)
            rendered = template.render(context)
            setattr(self.docx.core_properties, prop, rendered)

    def render_footnotes(
        self, context: Dict[str, Any], jinja_env: Optional[Environment] = None
    ) -> None:
        if jinja_env is None:
            jinja_env = SandboxedEnvironment()

        # Every section shares the same package, so render each part once.
        # Rendering twice would also evaluate escaped {_{ }_} delimiters.
        for part in self.docx.part.package.iter_parts():
            if part.content_type == self.FOOTNOTES_CONTENT_TYPE:
                xml = self.patch_xml(
                    part.blob.decode("utf-8")
                    if isinstance(part.blob, bytes)
                    else part.blob
                )
                xml = self.render_xml_part(xml, part, context, jinja_env)
                part._blob = xml.encode("utf-8")

    FOOTNOTES_CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    )
    ENDNOTES_CONTENT_TYPE = (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
    )

    def drop_unreferenced_notes(self):
        """Remove footnotes and endnotes whose reference was rendered away.

        A {%p if %} can drop the paragraph holding a footnote reference while
        the note itself stays in footnotes.xml. Word ignores it, but LibreOffice
        then pairs the remaining references with the wrong notes.
        """
        w = docx.oxml.ns.nsmap["w"]
        root = self.docx.element
        kinds = (
            (self.FOOTNOTES_CONTENT_TYPE, "footnote"),
            (self.ENDNOTES_CONTENT_TYPE, "endnote"),
        )
        for content_type, kind in kinds:
            used = set(
                etree._Element.xpath(
                    root, "//w:%sReference/@w:id" % kind, namespaces={"w": w}
                )
            )
            for part in self.docx.part.package.iter_parts():
                if part.content_type != content_type:
                    continue
                notes = etree.fromstring(part.blob)
                dropped = False
                for note in notes.findall("{%s}%s" % (w, kind)):
                    # separator and continuation notes have a type, keep them
                    if note.get("{%s}type" % w, "normal") != "normal":
                        continue
                    if note.get("{%s}id" % w) not in used:
                        notes.remove(note)
                        dropped = True
                if dropped:
                    part._blob = etree.tostring(
                        notes, xml_declaration=True, encoding="UTF-8", standalone=True
                    )

    def resolve_listing(self, xml):

        def resolve_text(run_properties, paragraph_properties, m):
            xml = m.group(0).replace(
                "\t",
                "</w:t></w:r>"
                "<w:r>%s<w:tab/></w:r>"
                '<w:r>%s<w:t xml:space="preserve">' % (run_properties, run_properties),
            )
            xml = xml.replace(
                "\a",
                "</w:t></w:r></w:p>"
                '<w:p>%s<w:r>%s<w:t xml:space="preserve">'
                % (paragraph_properties, run_properties),
            )
            xml = xml.replace("\n", '</w:t><w:br/><w:t xml:space="preserve">')
            xml = xml.replace(
                "\f",
                "</w:t></w:r></w:p>"
                '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                '<w:p>%s<w:r>%s<w:t xml:space="preserve">'
                % (paragraph_properties, run_properties),
            )
            return xml

        def resolve_run(paragraph_properties, m):
            run_properties = re.search(r"<w:rPr>.*?</w:rPr>", m.group(0))
            run_properties = run_properties.group(0) if run_properties else ""
            return re.sub(
                r"<w:t(?: [^>]*)?>.*?</w:t>",
                lambda x: resolve_text(run_properties, paragraph_properties, x),
                m.group(0),
                flags=re.DOTALL,
            )

        def resolve_paragraph(m):
            paragraph_properties = re.search(r"<w:pPr>.*?</w:pPr>", m.group(0))
            paragraph_properties = (
                paragraph_properties.group(0) if paragraph_properties else ""
            )
            return re.sub(
                r"<w:r(?: [^>]*)?>.*?</w:r>",
                lambda x: resolve_run(paragraph_properties, x),
                m.group(0),
                flags=re.DOTALL,
            )

        xml = re.sub(
            r"<w:p(?: [^>]*)?>.*?</w:p>", resolve_paragraph, xml, flags=re.DOTALL
        )

        return xml

    def build_xml(self, context, jinja_env=None):
        xml = self.get_xml()
        xml = self.patch_xml(xml)
        xml = self.render_xml_part(xml, self.docx._part, context, jinja_env)
        return xml

    def map_tree(self, tree):
        root = self.docx._element
        body = root.body
        root.replace(body, tree)
        # python-docx caches its body wrapper, drop it so reads after
        # render() see the new body instead of the template one.
        if hasattr(self.docx, "_Document__body"):
            self.docx._Document__body = None

    def get_headers_footers(self, uri):
        for relKey, val in self.docx._part.rels.items():
            if (val.reltype == uri) and (val.target_part.blob):
                yield relKey, val.target_part

    @staticmethod
    def drop_duplicate_bookmarks(root):
        """Keep one bookmark of each name.

        A bookmark inside a {% for %} block is repeated by the loop, and Word
        expects bookmark names to be unique. The first copy that has a matching
        end is kept, the others go together with their ends.
        """
        w = docx.oxml.ns.nsmap["w"]
        start, end = "{%s}bookmarkStart" % w, "{%s}bookmarkEnd" % w
        marks = list(root.iter(start, end))
        claimed = set()

        def end_of(mark):
            # first unclaimed end with the same id after this start
            mark_id = mark.get("{%s}id" % w)
            for other in marks[marks.index(mark) + 1:]:
                if other.tag == end and other.get("{%s}id" % w) == mark_id:
                    if id(other) not in claimed:
                        return other
            return None

        by_name = {}
        for mark in marks:
            if mark.tag == start:
                by_name.setdefault(mark.get("{%s}name" % w), []).append(mark)
        for group in by_name.values():
            if len(group) < 2:
                continue
            keeper = next((m for m in group if end_of(m) is not None), group[0])
            keeper_end = end_of(keeper)
            if keeper_end is not None:
                claimed.add(id(keeper_end))
            for mark in group:
                if mark is keeper:
                    continue
                mark_end = end_of(mark)
                if mark_end is not None:
                    claimed.add(id(mark_end))
                    mark_end.getparent().remove(mark_end)
                mark.getparent().remove(mark)

    def get_part_xml(self, part):
        root = parse_xml(part.blob)
        self.rescue_bookmarks(root)
        return self.xml_to_string(root)

    def get_headers_footers_encoding(self, xml):
        m = re.match(r'<\?xml[^\?]+\bencoding="([^"]+)"', xml, re.I)
        if m:
            return m.group(1)
        return "utf-8"

    def build_headers_footers_xml(self, context, uri, jinja_env=None):
        for relKey, part in self.get_headers_footers(uri):
            xml = self.get_part_xml(part)
            encoding = self.get_headers_footers_encoding(xml)
            xml = self.patch_xml(xml)
            xml = self.render_xml_part(xml, part, context, jinja_env)
            yield relKey, xml.encode(encoding)

    def map_headers_footers_xml(self, relKey, xml):
        part = self.docx._part.rels[relKey].target_part
        new_part = XmlPart.load(part.partname, part.content_type, xml, part.package)
        for rId, rel in part.rels.items():
            new_part.load_rel(rel.reltype, rel._target, rel.rId, rel.is_external)
        self.docx._part.rels[relKey]._target = new_part

    @staticmethod
    async def _run_in_executor(executor, func, *args, **kwargs):
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(executor, functools.partial(func, *args, **kwargs))

    async def render_async(
        self,
        context: Dict[str, Any],
        jinja_env: Optional[Environment] = None,
        autoescape: bool = False,
        executor=None,
    ) -> None:
        """Run render() in a worker thread so the event loop isn't blocked.

        executor is a concurrent.futures executor; None uses the loop's
        default one. Pass your own to cap how many renders run at once.
        Don't render or save the same template from two tasks at once.
        """
        await self._run_in_executor(executor, self.render, context, jinja_env, autoescape)

    async def save_async(
        self, filename: Union[IO[bytes], str, PathLike], *args, executor=None, **kwargs
    ) -> None:
        """Run save() in a worker thread, see render_async()."""
        await self._run_in_executor(executor, self.save, filename, *args, **kwargs)

    def render(
        self,
        context: Dict[str, Any],
        jinja_env: Optional[Environment] = None,
        autoescape: bool = False,
    ) -> None:
        # init template working attributes
        self.render_init()

        # Templates run in a sandbox unless the caller passes their own
        # environment, e.g. jinja2.Environment() for trusted templates.
        if jinja_env is None:
            jinja_env = SandboxedEnvironment(autoescape=autoescape)
        elif autoescape:
            jinja_env.autoescape = autoescape

        # Body
        xml_src = self.build_xml(context, jinja_env)

        # fix tables if needed
        tree = self.fix_tables(xml_src)
        self.fix_cell_properties(tree)

        # fix docPr ID's
        self.fix_docpr_ids(tree)

        # Replace body xml tree
        self.drop_duplicate_bookmarks(tree)
        self.map_tree(tree)

        # Headers
        headers = self.build_headers_footers_xml(context, self.HEADER_URI, jinja_env)
        for relKey, xml in headers:
            self.map_headers_footers_xml(relKey, xml)

        # Footers
        footers = self.build_headers_footers_xml(context, self.FOOTER_URI, jinja_env)
        for relKey, xml in footers:
            self.map_headers_footers_xml(relKey, xml)

        self.render_properties(context, jinja_env)

        self.render_footnotes(context, jinja_env)
        self.drop_unreferenced_notes()

        # set rendered flag
        self.is_rendered = True

    # using of TC tag in for cycle can cause that count of columns does not
    # correspond to real count of columns in row. This function is able to fix it.
    def fix_tables(self, xml):
        try:
            tree = etree.fromstring(xml)
        except etree.XMLSyntaxError as exc:
            # Keep going like before, but say so instead of silently
            # dropping whatever the lenient parser can't make sense of.
            warnings.warn(
                "The rendered document is not valid XML ({}), so some content "
                "may be missing from the output. This usually means a value "
                "contains a character like < or &. Render with "
                "autoescape=True to escape them.".format(exc),
                stacklevel=3,
            )
            tree = etree.fromstring(xml, parser=etree.XMLParser(recover=True))
        # get namespace
        ns = "{" + tree.nsmap["w"] + "}"
        # walk trough xml and find table
        for t in tree.iter(ns + "tbl"):
            tblGrid = t.find(ns + "tblGrid")
            columns = tblGrid.findall(ns + "gridCol")
            to_add = 0
            # walk trough all rows and try to find if there is higher cell count
            for r in t.iter(ns + "tr"):
                cells = r.findall(ns + "tc")
                if (len(columns) + to_add) < len(cells):
                    to_add = len(cells) - len(columns)
            # is necessary to add columns?
            if to_add > 0:
                # at first, calculate width of table according to columns
                # (we want to preserve it)
                width = 0.0
                new_average = None
                for c in columns:
                    if c.get(ns + "w") is not None:
                        width += float(c.get(ns + "w"))
                # try to keep proportion of table
                if width > 0:
                    old_average = width / len(columns)
                    new_average = width / (len(columns) + to_add)
                    # scale the old columns
                    for c in columns:
                        c.set(
                            ns + "w",
                            str(
                                int(float(c.get(ns + "w")) * new_average / old_average)
                            ),
                        )
                    # add new columns
                    for i in range(to_add):
                        etree.SubElement(
                            tblGrid, ns + "gridCol", {ns + "w": str(int(new_average))}
                        )

            # Refetch columns after columns addition.
            columns = tblGrid.findall(ns + "gridCol")
            columns_len = len(columns)

            cells_len_max = 0

            def get_cell_len(total, cell):
                tc_pr = cell.find(ns + "tcPr")
                grid_span = None if tc_pr is None else tc_pr.find(ns + "gridSpan")

                if grid_span is not None:
                    return total + int(grid_span.get(ns + "val"))

                return total + 1

            # Calculate max of table cells to compare with `gridCol`.
            for r in t.iter(ns + "tr"):
                cells = r.findall(ns + "tc")
                cells_len = functools.reduce(get_cell_len, cells, 0)
                cells_len_max = max(cells_len_max, cells_len)

            to_remove = columns_len - cells_len_max

            # If after the loop, there're less columns, than
            # originally was, remove extra `gridCol` declarations.
            if to_remove > 0:
                # Have to keep track of the removed width to scale the
                # table back to its original width.
                removed_width = 0.0

                for c in columns[-to_remove:]:
                    removed_width += float(c.get(ns + "w"))

                    tblGrid.remove(c)

                columns_left = tblGrid.findall(ns + "gridCol")

                # Distribute `removed_width` across all columns that has
                # left after extras removal.
                extra_space = 0
                if len(columns_left) > 0:
                    extra_space = removed_width / len(columns_left)
                    extra_space = int(extra_space)

                for c in columns_left:
                    c.set(ns + "w", str(int(float(c.get(ns + "w")) + extra_space)))

        return tree

    # Children of w:tcPr in the order the schema (CT_TcPr) requires.
    TC_PR_ORDER = {
        name: index
        for index, name in enumerate(
            (
                "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd",
                "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark",
                "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
            )
        )
    }

    def fix_cell_properties(self, tree):
        """Put w:tcPr children back in schema order.

        The vm, hm, colspan and cellbg tags insert vMerge, gridSpan and shd
        at a fixed spot in the cell properties, which is out of order when
        the cell also has borders, alignment and so on. Word refuses to open
        such a file. An element the schema list doesn't know stays right
        after the element it followed.
        """
        ns = "{" + docx.oxml.ns.nsmap["w"] + "}"
        for tc_pr in tree.iter(ns + "tcPr"):
            children = list(tc_pr)
            if len(children) < 2:
                continue
            keys = []
            last = -1
            for child in children:
                tag = child.tag if isinstance(child.tag, str) else ""
                if tag.startswith(ns):
                    last = self.TC_PR_ORDER.get(tag[len(ns):], last)
                keys.append(last)
            if keys == sorted(keys):
                continue
            ordered = [c for _, c in sorted(zip(keys, children), key=lambda kc: kc[0])]
            for child in ordered:
                tc_pr.append(child)

    def fix_docpr_ids(self, tree):
        # some Ids may have some collisions : so renumbering all of them :
        for elt in tree.xpath("//wp:docPr", namespaces=docx.oxml.ns.nsmap):
            self.docx_ids_index += 1
            elt.attrib["id"] = str(self.docx_ids_index)

    def new_subdoc(self, docpath=None) -> Subdoc:
        try:
            from .subdoc import Subdoc
        except ModuleNotFoundError as exc:
            if exc.name == "docxcompose":
                raise ModuleNotFoundError(
                    "new_subdoc() requires the optional docxcompose dependency. "
                    'Install it with: pip install "docform[subdoc]"',
                    name="docxcompose",
                ) from exc
            raise

        self.init_docx()
        return Subdoc(self, docpath)

    @staticmethod
    def get_file_crc(file_obj):
        if hasattr(file_obj, "read"):
            buf = file_obj.read()
        else:
            with open(file_obj, "rb") as fh:
                buf = fh.read()

        crc = binascii.crc32(buf) & 0xFFFFFFFF
        return crc

    def replace_media(self, src_file, dst_file):
        """Replace one media by another one into a docx

        This has been done mainly because it is not possible to add images in
        docx header/footer.
        With this function, put a dummy picture in your header/footer,
        then specify it with its replacement in this function using the file path
        or file-like objects.

        Syntax: tpl.replace_media('dummy_media_to_replace.png','media_to_paste.jpg')
            -- or --
                tpl.replace_media(io.BytesIO(image_stream), io.BytesIO(new_image_stream))

        Note: for images, the aspect ratio will be the same as the replaced image

        Note2: it is important to have the source media file as it is required
                to calculate its CRC to find them in the docx
        """

        crc = self.get_file_crc(src_file)
        if hasattr(dst_file, "read"):
            self.crc_to_new_media[crc] = dst_file.read()
        else:
            with open(dst_file, "rb") as fh:
                self.crc_to_new_media[crc] = fh.read()

    def replace_pic(self, embedded_file, dst_file, width=None, height=None):
        """Replace embedded picture with original-name given by embedded_file.
           (give only the file basename, not the full path)
           The new picture is given by dst_file (either a filename or a file-like
           object)

        width and height set the displayed size of the new picture, as EMU
        or python-docx lengths like Mm(30). If only one is given, the other
        is scaled to keep the replaced picture's aspect ratio.

        Notes:
            1) embedded_file and dst_file must have the same extension/format
               in case dst_file is a file-like object, no check is done on
               format compatibility
            2) without width or height, the size and aspect ratio will be the
               same as the replaced image
            3) There is no need to keep the original file (this is not the case
               for replace_embedded and replace_media)
        """
        for name, value in (("width", width), ("height", height)):
            if value is not None and int(value) <= 0:
                raise ValueError("%s must be a positive length" % name)
        if width is not None or height is not None:
            self.pics_to_resize[embedded_file] = (width, height)
        else:
            self.pics_to_resize.pop(embedded_file, None)

        if hasattr(dst_file, "read"):
            # NOTE: file extension not checked
            self.pics_to_replace[embedded_file] = dst_file.read()
        else:
            with open(dst_file, "rb") as fh:
                self.pics_to_replace[embedded_file] = fh.read()

    def replace_embedded(self, src_file, dst_file):
        """Replace one embedded object by another one into a docx

        This has been done mainly because it is not possible to add images
        in docx header/footer.
        With this function, put a dummy picture in your header/footer,
        then specify it with its replacement in this function

        Syntax: tpl.replace_embedded('dummy_doc.docx','doc_to_paste.docx')

        Note2 : it is important to have the source file as it is required to
                calculate its CRC to find them in the docx
        """
        with open(dst_file, "rb") as fh:
            crc = self.get_file_crc(src_file)
            self.crc_to_new_embedded[crc] = fh.read()

    def replace_zipname(self, zipname, dst_file):
        """Replace one file in the docx file

        First note that a MSWord .docx file is in fact a zip file.

        This method can be used to replace document embedded in the docx template.

        Some embedded document may have been modified by MSWord while saving
        the template : thus replace_embedded() cannot be used as CRC is not the
        same as the original file.

        This method works for embedded MSWord file like Excel or PowerPoint file,
        but won't work for others like PDF, Python or even Text files :
        For these ones, MSWord generate an oleObjectNNN.bin file which is no
        use to be replaced as it is encoded.

        Syntax:

        tpl.replace_zipname(
            'word/embeddings/Feuille_Microsoft_Office_Excel1.xlsx',
            'my_excel_file.xlsx')

        The zipname is the one you can find when you open docx with WinZip,
        7zip (Windows) or unzip -l (Linux). The zipname starts with
        "word/embeddings/". Note that the file is renamed by MSWord,
        so you have to guess a little bit...
        """
        with open(dst_file, "rb") as fh:
            self.zipname_to_replace[zipname] = fh.read()

    def reset_replacements(self):
        """Reset replacement dictionaries

        This will reset data for image/embedded/zipname replacement

        This is useful when calling several times render() with different
        image/embedded/zipname replacements without re-instantiating
        DocxTemplate object.
        In this case, the right sequence for each rendering will be :
            - reset_replacements(...)
            - replace_zipname(...), replace_media(...) and/or replace_embedded(...),
            - render(...)

        If you instantiate DocxTemplate object before each render(),
        this method is useless.
        """
        self.crc_to_new_media = {}
        self.crc_to_new_embedded = {}
        self.zipname_to_replace = {}
        self.pics_to_replace = {}
        self.pics_to_resize = {}

    def post_processing(self, docx_file):
        if self.crc_to_new_media or self.crc_to_new_embedded or self.zipname_to_replace:

            if hasattr(docx_file, "read"):
                tmp_file = io.BytesIO()
                DocxTemplate(docx_file).save(tmp_file)
                tmp_file.seek(0)
                docx_file.seek(0)
                docx_file.truncate()
                docx_file.seek(0)

            else:
                tmp_file = "%s_docform_before_replace_medias" % docx_file
                os.rename(docx_file, tmp_file)

            with zipfile.ZipFile(tmp_file) as zin:
                with zipfile.ZipFile(docx_file, "w") as zout:
                    for item in zin.infolist():
                        buf = zin.read(item.filename)
                        if item.filename in self.zipname_to_replace:
                            zout.writestr(item, self.zipname_to_replace[item.filename])
                        elif (
                            item.filename.startswith("word/media/")
                            and item.CRC in self.crc_to_new_media
                        ):
                            zout.writestr(item, self.crc_to_new_media[item.CRC])
                        elif (
                            item.filename.startswith("word/embeddings/")
                            and item.CRC in self.crc_to_new_embedded
                        ):
                            zout.writestr(item, self.crc_to_new_embedded[item.CRC])
                        else:
                            zout.writestr(item, buf)

            if not hasattr(tmp_file, "read"):
                os.remove(tmp_file)
            if hasattr(docx_file, "read"):
                docx_file.seek(0)

    def pre_processing(self):

        if self.pics_to_replace:
            self._replace_pics()

    def _replace_pics(self):
        """Replaces pictures xml tags in the docx template with pictures provided by the user"""

        replaced_pics = {key: False for key in self.pics_to_replace}

        # Main document
        part = self.docx.part
        self._replace_docx_part_pics(part, replaced_pics)

        # Header/Footer
        for relid, rel in part.rels.items():
            if rel.reltype in (REL_TYPE.HEADER, REL_TYPE.FOOTER):
                self._replace_docx_part_pics(rel.target_part, replaced_pics)

        if not self.allow_missing_pics:
            # make sure all template images defined by user were replaced
            for img_id, replaced in replaced_pics.items():
                if not replaced:
                    raise ValueError(
                        "Picture %s not found in the docx template" % img_id
                    )

    def get_pic_map(self):
        return self.pic_map

    def _replace_docx_part_pics(self, doc_part, replaced_pics):

        # Work on the live XML so a picture can be pointed at a new image part.
        et = getattr(doc_part, "_element", None)
        if et is None:
            et = etree.fromstring(doc_part.blob)
        # python-docx elements override xpath() without the namespaces argument
        def _xpath(element, path):
            return etree._Element.xpath(element, path, namespaces=docx.oxml.ns.nsmap)

        part_map = {}

        gds = _xpath(et, "//a:graphic/a:graphicData")
        for gd in gds:
            rel = None
            # Either IMAGE, CHART, SMART_ART, ...
            try:
                if gd.attrib["uri"] == docx.oxml.ns.nsmap["pic"]:
                    # Either PICTURE or LINKED_PICTURE image
                    blip = _xpath(gd, "pic:pic/pic:blipFill/a:blip")[0]
                    dest = _xpath(blip, "@r:embed")
                    if len(dest) > 0:
                        rel = dest[0]
                    else:
                        continue
                else:
                    continue

                non_visual_properties = "pic:pic/pic:nvPicPr/pic:cNvPr/"
                filename = _xpath(gd, "%s@name" % non_visual_properties)[0]
                titles = _xpath(gd, "%s@title" % non_visual_properties)
                if titles:
                    title = titles[0]
                else:
                    title = ""
                descriptions = _xpath(gd, "%s@descr" % non_visual_properties)
                if descriptions:
                    description = descriptions[0]
                else:
                    description = ""

                part_map[filename] = (
                    doc_part.rels[rel].target_ref,
                    doc_part.rels[rel].target_part,
                )

                # replace data
                for img_id, img_data in self.pics_to_replace.items():
                    if img_id == filename or img_id == title or img_id == description:
                        new_rel = self._swap_pic_image(doc_part, blip, rel, img_data)
                        if img_id in self.pics_to_resize:
                            self._resize_pic(gd, *self.pics_to_resize[img_id])
                        part_map[filename] = (
                            doc_part.rels[new_rel].target_ref,
                            doc_part.rels[new_rel].target_part,
                        )
                        replaced_pics[img_id] = True
                        break

            # FIXME: figure out what exceptions are thrown here
            # and catch more specific exceptions
            except Exception:
                continue

        self.pic_map.update(part_map)

    @staticmethod
    def _resize_pic(graphic_data, width, height):
        """Set the displayed size of the picture in graphic_data.

        Word reads the size from wp:extent on the inline or anchored drawing,
        and the picture's own a:ext has to match it.
        """
        ns = docx.oxml.ns.nsmap
        drawing = graphic_data.getparent().getparent()
        extent = drawing.find("{%s}extent" % ns["wp"])
        if extent is None:
            return
        cx, cy = int(extent.get("cx")), int(extent.get("cy"))
        if width is not None and height is not None:
            cx, cy = int(width), int(height)
        elif width is not None:
            cx, cy = int(width), int(round(cy * int(width) / cx)) if cx else cy
        else:
            cx, cy = (int(round(cx * int(height) / cy)) if cy else cx), int(height)
        extent.set("cx", str(cx))
        extent.set("cy", str(cy))
        for ext in etree._Element.xpath(
            graphic_data, "pic:pic/pic:spPr/a:xfrm/a:ext", namespaces=ns
        ):
            ext.set("cx", str(cx))
            ext.set("cy", str(cy))

    @staticmethod
    def _swap_pic_image(doc_part, blip, rel, img_data):
        """Point one picture at img_data and return its relationship id.

        Word stores identical images once, so several pictures can share an
        image part. Giving this picture its own part keeps the others intact.
        """
        try:
            new_rel, _ = doc_part.get_or_add_image(io.BytesIO(img_data))
        except Exception:
            # Formats python-docx can't read (EMF, WMF, ...): overwrite the
            # shared image data like before.
            doc_part.rels[rel].target_part._blob = img_data
            return rel
        if new_rel != rel:
            # Drop the old relationship only if this picture was its last user.
            # python-docx's drop_rel() doesn't count r:embed, so count here.
            users = etree._Element.xpath(
                blip.getroottree().getroot(),
                "//@r:embed | //@r:link | //@r:id",
                namespaces=docx.oxml.ns.nsmap,
            )
            if sum(1 for u in users if u == rel) < 2:
                del doc_part.rels[rel]
            blip.set("{%s}embed" % docx.oxml.ns.nsmap["r"], new_rel)
        return new_rel

    def build_url_id(self, url):
        self.init_docx()
        return self.docx._part.relate_to(url, REL_TYPE.HYPERLINK, is_external=True)

    def save(self, filename: Union[IO[bytes], str, PathLike], *args, **kwargs) -> None:
        # case where save() is called without doing rendering
        # ( user wants only to replace image/embedded/zipname )
        if not self.is_saved and not self.is_rendered:
            self.docx = Document(self.template_file)
        self.pre_processing()
        self.docx.save(filename, *args, **kwargs)
        self.post_processing(filename)
        self.is_saved = True

    def get_undeclared_template_variables(
        self,
        jinja_env: Optional[Environment] = None,
        context: Optional[Dict[str, Any]] = None,
        nested: bool = False,
    ) -> Set[str]:
        """Return the variables the template uses but doesn't define.

        With context, only the ones missing from it are returned. With
        nested=True, attribute and key lookups are followed too, so
        ``{{ foo.bar }}`` gives ``foo.bar`` instead of ``foo``, and with
        context a path is returned when any part of it is missing.
        """
        # Create a temporary document to analyze the template without affecting the current state
        temp_doc = Document(self.template_file)

        # Get XML from the temporary document
        xml = self.xml_to_string(temp_doc._element.body)
        xml = self.patch_xml(xml)

        # Add headers and footers
        for uri in [self.HEADER_URI, self.FOOTER_URI]:
            for relKey, val in temp_doc._part.rels.items():
                if (val.reltype == uri) and (val.target_part.blob):
                    _xml = self.xml_to_string(parse_xml(val.target_part.blob))
                    xml += self.patch_xml(_xml)

        if jinja_env:
            env = jinja_env
        else:
            env = Environment()

        parse_content = env.parse(xml)
        all_variables = meta.find_undeclared_variables(parse_content)

        if nested:
            paths = _variable_paths(parse_content, all_variables)
            if context is not None:
                paths = {p for p in paths if not _context_has_path(context, p)}
            return {".".join(str(k) for k in p) for p in paths}

        # If context is provided, return only variables that are not in the context
        if context is not None:
            provided_variables = set(context.keys())
            return all_variables - provided_variables

        # If no context provided, return all variables (original behavior)
        return all_variables
