# -*- coding: utf-8 -*-
"""
Created : 2021-07-30
"""

from copy import deepcopy

from docx import Document
from docx.oxml import CT_SectPr
from docx.oxml.ns import nsmap, qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docxcompose.properties import CustomProperties
from docxcompose.utils import xpath
from docxcompose.composer import Composer
from docxcompose.utils import NS
from lxml import etree


W14_ANCHOR_ID = (
    "{http://schemas.microsoft.com/office/word/2010/wordml}anchorId"
)


# docxcompose changes python-docx's process-wide ``cp`` prefix to the custom
# properties namespace. Restore its standard meaning so python-docx updates
# existing core properties instead of adding conflicting elements.
nsmap["cp"] = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"


class SubdocComposer(Composer):
    def attach_parts(self, doc, remove_property_fields=True):
        """Attach docx parts instead of appending the whole document
        thus subdoc insertion can be delegated to jinja2"""
        self.reset_reference_mapping()

        # Remove custom property fields but keep the values
        if remove_property_fields:
            cprops = CustomProperties(doc)
            for name in cprops.keys():
                cprops.dissolve_fields(name)

        self._create_style_id_mapping(doc)
        self.renumber_ole_objects(doc)

        for element in doc.element.body:
            if isinstance(element, CT_SectPr):
                continue
            self.add_referenced_parts(doc.part, self.doc.part, element)
            self.add_styles(doc, element)
            self.add_numberings(doc, element)
            self.restart_first_numbering(doc, element)
            self.add_images(doc, element)
            self.add_diagrams(doc, element)
            self.add_shapes(doc, element)
            self.add_footnotes(doc, element)
            self.remove_header_and_footer_references(doc, element)

        self.add_styles_from_other_parts(doc)
        self.renumber_bookmarks()
        self.renumber_docpr_ids()
        self.renumber_nvpicpr_ids()
        self.fix_section_types(doc)

    def reset_reference_mapping(self):
        super().reset_reference_mapping()
        # style id -> numPr given to the first paragraph of a restarted list
        self._style_numbering_restarts = {}

    def restart_first_numbering(self, doc, element):
        """Restart a style-numbered list for all of its paragraphs.

        docxcompose restarts a list by giving its first paragraph a new numId.
        When the numbering comes from the paragraph style, only that first
        paragraph gets it and the rest stay on the old numbering, so the list
        splits in two. Give the following paragraphs of that style the same
        numbering.
        """
        restarted = self._style_numbering_restarts

        first = self._first_styled_paragraph_props(element)
        style_id = first[0] if first else None
        had_num_pr = first is not None and first[1].find(qn("w:numPr")) is not None
        is_new = style_id is not None and style_id not in self._numbering_restarted

        super().restart_first_numbering(doc, element)

        if is_new and not had_num_pr and style_id in self._numbering_restarted:
            num_pr = first[1].find(qn("w:numPr"))
            if num_pr is not None:
                # docxcompose appends it last, after rPr, which breaks the
                # schema order Word expects. Put it in its proper place.
                first[1].remove(num_pr)
                self._set_num_pr(first[1], num_pr)
                restarted[style_id] = num_pr

        for p_pr in xpath(element, ".//w:pPr[w:pStyle]"):
            num_pr = restarted.get(p_pr.find(qn("w:pStyle")).get(qn("w:val")))
            if num_pr is not None and p_pr.find(qn("w:numPr")) is None:
                self._set_num_pr(p_pr, num_pr)

    @staticmethod
    def _set_num_pr(p_pr, num_pr):
        new = p_pr.get_or_add_numPr()
        new.get_or_add_ilvl().val = int(num_pr.find(qn("w:ilvl")).get(qn("w:val")))
        new.get_or_add_numId().val = int(num_pr.find(qn("w:numId")).get(qn("w:val")))

    @staticmethod
    def _first_styled_paragraph_props(element):
        props = xpath(element, ".//w:pPr[w:pStyle]")
        if not props:
            return None
        return props[0].find(qn("w:pStyle")).get(qn("w:val")), props[0]

    def renumber_ole_objects(self, doc):
        state = getattr(self.doc.part, "_docform_ole_id_state", None)
        if state is None:
            body = self.doc.element.body
            state = {
                "shape_ids": set(xpath(body, ".//v:shape/@id")),
                "shape_index": 1025,
                "object_ids": set(xpath(body, ".//o:OLEObject/@ObjectID")),
                "object_index": 1,
                "anchor_ids": {
                    element.get(W14_ANCHOR_ID).upper()
                    for element in body.iter()
                    if element.get(W14_ANCHOR_ID) is not None
                },
                "anchor_index": 1,
            }
            self.doc.part._docform_ole_id_state = state

        state["shape_ids"].update(
            xpath(doc.element.body, ".//v:shape/@id")
        )
        for ole_object in xpath(doc.element.body, ".//o:OLEObject"):
            old_shape_id = ole_object.get("ShapeID")
            matching_shapes = [
                shape
                for shape in xpath(ole_object.getparent(), "./v:shape")
                if (
                    old_shape_id is not None
                    and shape.get("id") == old_shape_id
                )
            ]
            if len(matching_shapes) == 1:
                shape_id = self._next_ole_id(
                    state, "shape", lambda index: "_x0000_i%d" % index
                )
                matching_shapes[0].set("id", shape_id)
                ole_object.set("ShapeID", shape_id)
            if ole_object.get("ObjectID") is not None:
                ole_object.set(
                    "ObjectID",
                    self._next_ole_id(
                        state, "object", lambda index: "_%d" % index
                    ),
                )

        for element in doc.element.body.iter():
            if element.get(W14_ANCHOR_ID) is not None:
                element.set(
                    W14_ANCHOR_ID,
                    self._next_ole_id(
                        state, "anchor", lambda index: "%08X" % index
                    ),
                )

    @staticmethod
    def _next_ole_id(state, kind, formatter):
        index_key = "%s_index" % kind
        ids_key = "%s_ids" % kind
        while formatter(state[index_key]) in state[ids_key]:
            state[index_key] += 1
        value = formatter(state[index_key])
        state[index_key] += 1
        state[ids_key].add(value)
        return value

    def add_diagrams(self, doc, element):
        # While waiting docxcompose 1.3.3
        dgm_rels = xpath(element, ".//dgm:relIds[@r:dm]")
        for dgm_rel in dgm_rels:
            for item, rt_type in (
                ("dm", RT.DIAGRAM_DATA),
                ("lo", RT.DIAGRAM_LAYOUT),
                ("qs", RT.DIAGRAM_QUICK_STYLE),
                ("cs", RT.DIAGRAM_COLORS),
            ):
                dm_rid = dgm_rel.get("{%s}%s" % (NS["r"], item))
                dm_part = doc.part.rels[dm_rid].target_part
                new_rid = self.doc.part.relate_to(dm_part, rt_type)
                dgm_rel.set("{%s}%s" % (NS["r"], item), new_rid)


class Subdoc(object):
    """Class for subdocument to insert into master document"""

    def __init__(self, tpl, docpath=None):
        self.tpl = tpl
        self.docx = tpl.get_docx()
        self.subdocx = Document(docpath)
        if docpath:
            compose = SubdocComposer(self.docx)
            compose.attach_parts(self.subdocx)
        else:
            self.subdocx._part = self.docx._part

    def __getattr__(self, name):
        return getattr(self.subdocx, name)

    def _get_xml(self):
        body = self.subdocx.element.body
        if body.sectPr is not None:
            body.remove(body.sectPr)
        if len(body) == 0:
            return ""

        destination_body = self.docx.element.body
        wrapper = etree.Element(destination_body.tag, nsmap=destination_body.nsmap)
        for element in body:
            wrapper.append(deepcopy(element))
        # Word needs a paragraph after a table at the end of a cell or of the
        # document, and {{p subdoc }} can put the subdoc's last table there.
        if wrapper[-1].tag == "{%s}tbl" % nsmap["w"]:
            etree.SubElement(wrapper, "{%s}p" % nsmap["w"])

        xml = etree.tostring(wrapper, encoding="unicode", pretty_print=False)
        return xml.partition(">")[2].rpartition("</")[0]

    def __unicode__(self):
        return self._get_xml()

    def __str__(self):
        return self._get_xml()

    def __html__(self):
        return self._get_xml()
