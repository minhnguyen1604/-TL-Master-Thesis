# -*- coding: utf-8 -*-
"""
sync_docx_to_markdown.py
Dong bo toan bo noi dung tu DTL_Master_Thesis_Draft.docx sang DTL_Master_Thesis_Draft.md
Dam bao file Markdown phan anh 100% noi dung hoc thuat cua file Word.
"""
import sys, os
import docx

sys.stdout.reconfigure(encoding='utf-8')

DOCX_FILE = "workingfile/DTL_Master_Thesis_Draft.docx"
MD_FILE = "workingfile/DTL_Master_Thesis_Draft.md"

def docx_to_markdown(docx_path, md_path):
    doc = docx.Document(docx_path)
    lines = []

    # Lap qua cac phan tu trong body theo thu tu XML thuc te
    for child in doc.element.body:
        tag = child.tag
        if tag.endswith('}p'):
            p = docx.text.paragraph.Paragraph(child, doc)
            txt = p.text.strip()
            style = p.style.name

            if not txt:
                lines.append("")
                continue

            if style == 'Heading 1':
                lines.append(f"\n# {txt}\n")
            elif style == 'Heading 2':
                lines.append(f"\n## {txt}\n")
            elif style == 'Heading 3':
                lines.append(f"\n### {txt}\n")
            elif style == 'Heading 4':
                lines.append(f"\n#### {txt}\n")
            elif p.runs and p.runs[0].bold and ('Table ' in txt or 'Figure ' in txt):
                lines.append(f"\n**{txt}**\n")
            elif p.runs and p.runs[0].italic and ('Source:' in txt or txt.startswith('Note:')):
                lines.append(f"*{txt}*\n")
            elif txt.startswith('– '):
                lines.append(f"- {txt[2:]}")
            else:
                lines.append(f"{txt}\n")

        elif tag.endswith('}tbl'):
            tb = docx.table.Table(child, doc)
            # Chuyen bang sang markdown table
            rows_data = []
            for r in tb.rows:
                row_cells = [cell.text.strip().replace('\n', ' ') for cell in r.cells]
                rows_data.append(row_cells)

            if rows_data:
                # Header
                header_line = "| " + " | ".join(rows_data[0]) + " |"
                sep_line = "| " + " | ".join(["---"] * len(rows_data[0])) + " |"
                lines.append(header_line)
                lines.append(sep_line)
                for row_cells in rows_data[1:]:
                    lines.append("| " + " | ".join(row_cells) + " |")
                lines.append("")

    with open(md_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    print(f"Da dong bo thanh cong toan bo sang {md_path} ({len(lines)} dong).")

if __name__ == '__main__':
    docx_to_markdown(DOCX_FILE, MD_FILE)
