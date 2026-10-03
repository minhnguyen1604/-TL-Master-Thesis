# -*- coding: utf-8 -*-
"""
build_full_thesis.py
Base module for thesis automation with robust Heading location and typography preservation.
"""
import sys, os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding='utf-8')

GREY = "D9D9D9"

def shade(cell, f):
    tcPr = cell._element.get_or_add_tcPr()
    s = OxmlElement('w:shd')
    s.set(qn('w:val'), 'clear')
    s.set(qn('w:color'), 'auto')
    s.set(qn('w:fill'), f)
    tcPr.append(s)

class ThesisWriter:
    def __init__(self, doc):
        self.doc = doc
        self.cursor = None

    def at(self, heading_text):
        target = heading_text.strip()
        paragraphs = list(self.doc.paragraphs)
        candidates = []
        for idx, p in enumerate(paragraphs):
            txt = p.text.strip()
            # Ignore table of contents lines which have tab characters or page numbers at the end
            if '\t' in txt and any(txt.endswith(str(d)) for d in range(10)):
                continue
            if txt == target:
                candidates.append((p, 3 if p.style.name.startswith('Heading') else 1, idx))
            elif txt.startswith(target) or target.startswith(txt[:min(len(txt), 20)]):
                candidates.append((p, 2 if p.style.name.startswith('Heading') else 0, idx))

        if candidates:
            best_p = max(candidates, key=lambda x: (x[1], x[2]))[0]
            self.cursor = best_p
            return self

        raise ValueError("Khong tim thay de muc: " + heading_text)

    def _new_par(self):
        el = OxmlElement('w:p')
        self.cursor._element.addnext(el)
        p = Paragraph(el, self.cursor._parent)
        self.cursor = p
        return p

    def para(self, text, italic=False, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             size=12, after=8, indent=None, first_line=0.25):
        if self.cursor is not None and self.cursor.text == "" and len(self.cursor.runs) == 0:
            p = self.cursor
        else:
            p = self._new_par()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_after = Pt(after)
        pf.line_spacing = 1.5
        if indent is not None:
            pf.left_indent = Inches(indent)
        if first_line is not None:
            pf.first_line_indent = Inches(first_line)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(size)
        r.italic = italic
        r.bold = bold
        return p

    def bullet(self, text, size=12):
        p = self._new_par()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.space_after = Pt(5)
        pf.line_spacing = 1.4
        pf.left_indent = Inches(0.4)
        pf.first_line_indent = Inches(-0.2)
        r = p.add_run("– " + text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(size)
        return p

    def caption(self, text):
        return self.para(text, italic=False, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                         size=11, after=4, first_line=0)

    def source(self, text):
        return self.para(text, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER,
                         size=10, after=10, first_line=0)

    def table(self, rows, widths, font=9.5):
        tb = self.doc.add_table(rows=len(rows), cols=len(rows[0]))
        tb.style = 'Table Grid'
        tb.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(rows):
            for k, v in enumerate(row):
                cell = tb.rows[i].cells[k]
                cell.width = Inches(widths[k])
                cell.text = ""
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.0
                if i == 0 or k >= len(row) - 1:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(str(v))
                r.font.name = 'Times New Roman'
                r.font.size = Pt(font)
                if i == 0:
                    r.bold = True
                    shade(cell, GREY)
        self.cursor._element.addnext(tb._element)
        el = OxmlElement('w:p')
        tb._element.addnext(el)
        self.cursor = Paragraph(el, self.cursor._parent)
        return tb

print("ThesisWriter loaded successfully.")
