# -*- coding: utf-8 -*-
"""
populate_frontmatter.py
Viet Abstract, List of Tables va List of Figures vao file Word DTL_Master_Thesis_Draft.docx
Giu dung 100% typography: Times New Roman, Body 12pt, 1.5 line spacing, 8pt after.
"""
import sys, os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding='utf-8')

FILE = "workingfile/DTL_Master_Thesis_Draft.docx"

def main():
    doc = docx.Document(FILE)

    # 1. ABSTRACT
    p_abstract = None
    p_keywords = None
    for i, p in enumerate(doc.paragraphs):
        if p.text.strip() == "ABSTRACT":
            p_abstract = p
            p_keywords = doc.paragraphs[i+1]
            break

    if p_abstract is None:
        raise ValueError("Khong tim thay tieu de ABSTRACT")

    # Xoa cac paragraph cu giua ABSTRACT va TABLE OF CONTENTS
    # Hien tai p_keywords la paragraph ngay sau ABSTRACT
    # Chen 3 doan noi dung Abstract vao giua p_abstract va p_keywords
    abstract_paras = [
        "This Master's thesis investigates the determinants of corporate customers' decisions to select "
        "bank guarantee services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank), "
        "operating within the evolving regulatory framework defined by the Law on Credit Institutions 2024 and "
        "Circular No. 61/2024/TT-NHNN of the State Bank of Vietnam. Integrating financial intermediation theory, "
        "credit risk contingent claim pricing, relationship banking, service quality (SERVQUAL), and the "
        "Technology Acceptance Model (TAM), the study conceptualises corporate selection decision as an "
        "attitudinal patronage priority construct influenced by seven core dimensions: Price Competitiveness, "
        "Processing Speed, Digital eFAST Convenience, Bank Reputation, Relationship Banking and Limits, "
        "Staff Professionalism, and Collateral and Margin Policy.",

        "A structured quantitative survey was administered across 800 corporate clients currently utilising "
        "guarantee services at VietinBank's nationwide network of 155 branches. Psychometric evaluation demonstrates "
        "rigorous internal consistency reliability across all eight measurement scales (Cronbach's alpha ranging "
        "from 0.864 to 0.888, with all corrected item-total correlations exceeding 0.69). Exploratory factor "
        "analysis confirms construct unidimensionality and distinct factorial validity, extracting seven orthogonal "
        "independent factors that account for 73.13% of the total variance, alongside a single dominant dependent "
        "factor explaining 74.10% of variance. Ordinary least squares (OLS) multiple regression indicates that "
        "the conceptual model explains 59.6% of the variance in corporate guarantee selection decisions "
        "(R² = 0.596, Adjusted R² = 0.593, F(7, 792) = 167.16, p < 0.001). All seven research hypotheses (H1 to H7) "
        "receive strong empirical support: Price Competitiveness emerges as the strongest driver (β = 0.266, p < 0.001), "
        "followed by Relationship Banking and Limits (β = 0.211, p < 0.001), Bank Reputation (β = 0.191, p < 0.001), "
        "Processing Speed (β = 0.171, p < 0.001), Digital eFAST Convenience (β = 0.167, p < 0.001), Collateral Policy "
        "(β = 0.137, p < 0.001), and Staff Professionalism (β = 0.101, p < 0.001). Sub-group difference tests "
        "(ANOVA and independent-samples t-test) confirm significant perceptual variations across ownership types and "
        "multi-banking status (H8 supported).",

        "Based on these empirical findings, the thesis formulates a coherent set of managerial recommendations "
        "for VietinBank—including tiered volume pricing, service level agreements (SLAs) for rapid issuance, "
        "flexible cash-flow-based collateral options for SMEs, and end-to-end straight-through processing on the "
        "eFAST platform—as well as policy recommendations for the State Bank of Vietnam to foster a transparent "
        "and secure digital guarantee market."
    ]

    # Chen truoc p_keywords
    for text in abstract_paras:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.space_after = Pt(8)
        pf.line_spacing = 1.5
        pf.first_line_indent = Inches(0.25)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        # Move element right before p_keywords
        p_keywords._element.addprevious(p._element)

    # Format p_keywords
    p_keywords.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf_k = p_keywords.paragraph_format
    pf_k.space_after = Pt(12)
    pf_k.line_spacing = 1.5
    pf_k.first_line_indent = Inches(0.25)
    for r in p_keywords.runs:
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.italic = True
    print("Da cap nhat Abstract hoan chinh!")

    # 2. LIST OF TABLES
    p_lot = None
    for p in doc.paragraphs:
        if p.text.strip() == "LIST OF TABLES":
            p_lot = p
            break
    
    tables_list = [
        ("Table 2.1", "Principal corporate bank guarantee products", "12"),
        ("Table 2.2", "Research gaps and the response of this study", "23"),
        ("Table 3.1", "Operationalisation of constructs and measurement indicators", "33"),
        ("Table 3.2", "Criteria applied in exploratory factor analysis", "41"),
        ("Table 3.3", "Grouping variables and corresponding difference tests", "45"),
        ("Table 4.1", "Distribution of sample by enterprise ownership type", "48"),
        ("Table 4.2", "Distribution of sample by annual revenue scale", "49"),
        ("Table 4.3", "Distribution of sample by operating experience (tenure)", "50"),
        ("Table 4.4", "Distribution of primary guarantee product used", "51"),
        ("Table 4.5", "Multi-banking status and respondent corporate positions", "52"),
        ("Table 4.6", "Scale reliability analysis results (Cronbach's Alpha and Item-Total Statistics)", "54"),
        ("Table 4.7", "KMO and Bartlett's Test of Sphericity for independent variables", "56"),
        ("Table 4.8", "Rotated Component Matrix and Variance Explained (28 Independent Items)", "57"),
        ("Table 4.9", "Component Matrix for Dependent Variable (DEC)", "59"),
        ("Table 4.10", "Pearson correlation matrix and multicollinearity diagnostics (VIF & Tolerance)", "60"),
        ("Table 4.11", "OLS Multiple Regression Model Summary and ANOVA", "62"),
        ("Table 4.12", "Regression coefficients and research hypothesis testing decisions", "64"),
        ("Table 4.13", "Robustness check regression model with corporate control variables", "66"),
        ("Table 4.14", "One-Way ANOVA of construct evaluations across enterprise ownership types", "68"),
        ("Table 4.15", "One-Way ANOVA across enterprise revenue scales and operating experience", "70"),
        ("Table 4.16", "One-Way ANOVA across primary guarantee product categories", "71"),
        ("Table 4.17", "Independent samples t-test between single-bank and multi-bank users", "72"),
    ]

    cur = p_lot
    for t_id, t_title, t_page in tables_list:
        el = OxmlElement('w:p')
        cur._element.addnext(el)
        p = Paragraph(el, cur._parent)
        cur = p
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_after = Pt(4)
        pf.line_spacing = 1.3
        
        # Format text with tabs
        r1 = p.add_run(f"{t_id}: {t_title}")
        r1.font.name = 'Times New Roman'; r1.font.size = Pt(11)
        r2 = p.add_run(f"\t{t_page}")
        r2.font.name = 'Times New Roman'; r2.font.size = Pt(11)
    print("Da cap nhat List of Tables:", len(tables_list), "bang")

    # 3. LIST OF FIGURES
    p_lof = None
    for p in doc.paragraphs:
        if p.text.strip() == "LIST OF FIGURES":
            p_lof = p
            break

    figures_list = [
        ("Figure 2.1", "Conceptual research framework and hypothesised relationships", "27"),
        ("Figure 4.1", "Empirical regression results and standardized path coefficients", "65"),
    ]

    cur = p_lof
    for f_id, f_title, f_page in figures_list:
        el = OxmlElement('w:p')
        cur._element.addnext(el)
        p = Paragraph(el, cur._parent)
        cur = p
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_after = Pt(4)
        pf.line_spacing = 1.3
        
        r1 = p.add_run(f"{f_id}: {f_title}")
        r1.font.name = 'Times New Roman'; r1.font.size = Pt(11)
        r2 = p.add_run(f"\t{f_page}")
        r2.font.name = 'Times New Roman'; r2.font.size = Pt(11)
    print("Da cap nhat List of Figures:", len(figures_list), "hinh")

    doc.save(FILE)
    print("Da luu file docx thanh cong.")

if __name__ == '__main__':
    main()
