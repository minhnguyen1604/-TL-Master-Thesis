import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    
    borders = {'top': top, 'bottom': bottom, 'left': left, 'right': right}
    for border_name, border_style in borders.items():
        if border_style:
            b_el = OxmlElement(f'w:{border_name}')
            b_el.set(qn('w:val'), border_style.get('val', 'single'))
            b_el.set(qn('w:sz'), str(border_style.get('sz', 4)))
            b_el.set(qn('w:space'), '0')
            b_el.set(qn('w:color'), border_style.get('color', '000000'))
            tcBorders.append(b_el)
        else:
            b_el = OxmlElement(f'w:{border_name}')
            b_el.set(qn('w:val'), 'none')
            tcBorders.append(b_el)
    tcPr.append(tcBorders)

def add_bottom_border_to_paragraph(paragraph, color_hex="000000", size="6"):
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)

def generate_black_pristine_thesis_template():
    doc = docx.Document()

    # 1. Page Margins (1.0 inch = 2.54 cm standard)
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # 2. Running Header Configuration (Different First Page)
    section.different_first_page_header_footer = True
    header = section.header
    p_head = header.paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_head = p_head.add_run("Master Thesis – MDE31 – Dang Tu Linh")
    r_head.font.name = 'Times New Roman'
    r_head.font.size = Pt(10)
    r_head.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r_head.italic = True
    add_bottom_border_to_paragraph(p_head, color_hex="000000", size="4")

    # Configure Styles
    styles = doc.styles

    # Normal Style
    style_normal = styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    # Heading 1 Style
    style_h1 = styles['Heading 1']
    style_h1.font.name = 'Times New Roman'
    style_h1.font.size = Pt(14)
    style_h1.font.bold = True
    style_h1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    style_h1.paragraph_format.space_before = Pt(14)
    style_h1.paragraph_format.space_after = Pt(6)

    # Heading 2 Style
    style_h2 = styles['Heading 2']
    style_h2.font.name = 'Times New Roman'
    style_h2.font.size = Pt(12.5)
    style_h2.font.bold = True
    style_h2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    style_h2.paragraph_format.space_before = Pt(10)
    style_h2.paragraph_format.space_after = Pt(4)

    # Heading 3 Style
    style_h3 = styles['Heading 3']
    style_h3.font.name = 'Times New Roman'
    style_h3.font.size = Pt(12)
    style_h3.font.bold = True
    style_h3.font.italic = True
    style_h3.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    style_h3.paragraph_format.space_before = Pt(6)
    style_h3.paragraph_format.space_after = Pt(2)

    def add_p(text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(text)
        r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    def add_toc_line(title, page_num, level=1):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        
        # Add tab stop at 6.5 inches (right aligned with dots)
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        
        if level == 1:
            p.paragraph_format.left_indent = Inches(0.0)
            p.paragraph_format.space_before = Pt(4)
            r1 = p.add_run(title)
            r1.bold = True
            r1.font.size = Pt(11)
            r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            r2 = p.add_run(f"\t{page_num}")
            r2.bold = True
            r2.font.size = Pt(11)
            r2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif level == 2:
            p.paragraph_format.left_indent = Inches(0.25)
            r1 = p.add_run(title)
            r1.font.size = Pt(10.5)
            r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            r2 = p.add_run(f"\t{page_num}")
            r2.font.size = Pt(10.5)
            r2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        elif level == 3:
            p.paragraph_format.left_indent = Inches(0.5)
            r1 = p.add_run(title)
            r1.font.size = Pt(10)
            r1.italic = True
            r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            r2 = p.add_run(f"\t{page_num}")
            r2.font.size = Pt(10)
            r2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        return p

    # ==================== COVER PAGE ====================
    p_cov1 = doc.add_paragraph()
    p_cov1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov1.paragraph_format.space_after = Pt(4)
    r = p_cov1.add_run("NATIONAL ECONOMICS UNIVERSITY\nVIETNAM-NETHERLANDS MASTER’S PROGRAM IN DEVELOPMENT ECONOMICS (MDE)")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_cov2 = doc.add_paragraph()
    p_cov2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov2.paragraph_format.space_before = Pt(60)
    p_cov2.paragraph_format.space_after = Pt(24)
    r = p_cov2.add_run("MASTER THESIS")
    r.bold = True
    r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_cov3 = doc.add_paragraph()
    p_cov3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov3.paragraph_format.space_after = Pt(60)
    r = p_cov3.add_run("FACTORS AFFECTING CORPORATE CUSTOMERS’ DECISION TO CHOOSE BANK GUARANTEE SERVICES AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)")
    r.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_cov4 = doc.add_paragraph()
    p_cov4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov4.paragraph_format.space_after = Pt(4)
    r = p_cov4.add_run("Student: DANG TU LINH")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_cov5 = doc.add_paragraph()
    p_cov5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov5.paragraph_format.space_after = Pt(4)
    r = p_cov5.add_run("Student ID / Class: MDE Class 31")
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_cov6 = doc.add_paragraph()
    p_cov6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cov6.paragraph_format.space_after = Pt(60)
    r = p_cov6.add_run("Academic Supervisor: Dr. HOANG THI THUY NGA")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_cov7 = doc.add_paragraph()
    p_cov7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_cov7.add_run("Hanoi, 2026")
    r.italic = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    doc.add_page_break()

    # ==================== TITLE PAGE ====================
    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1.paragraph_format.space_after = Pt(4)
    r = p_t1.add_run("NATIONAL ECONOMICS UNIVERSITY\nERASMUS UNIVERSITY ROTTERDAM - ISS")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2.paragraph_format.space_before = Pt(40)
    p_t2.paragraph_format.space_after = Pt(20)
    r = p_t2.add_run("MASTER’S THESIS IN DEVELOPMENT ECONOMICS")
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_t3 = doc.add_paragraph()
    p_t3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t3.paragraph_format.space_after = Pt(40)
    r = p_t3.add_run("FACTORS AFFECTING CORPORATE CUSTOMERS’ DECISION TO CHOOSE BANK GUARANTEE SERVICES AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)")
    r.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_t4 = doc.add_paragraph()
    p_t4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t4.paragraph_format.space_after = Pt(6)
    r = p_t4.add_run("Author: DANG TU LINH")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_t5 = doc.add_paragraph()
    p_t5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t5.paragraph_format.space_after = Pt(40)
    r = p_t5.add_run("Academic Supervisor: Dr. HOANG THI THUY NGA")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    p_t6 = doc.add_paragraph()
    p_t6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_t6.add_run("A thesis submitted in partial fulfillment of the requirements for the degree of\nMaster of Arts in Development Economics\nHanoi, June 2026")
    r.italic = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    doc.add_page_break()

    # ==================== STATEMENT OF AUTHORSHIP ====================
    doc.add_heading("STATEMENT OF AUTHORSHIP", level=1)
    add_p("I hereby declare that this Master’s thesis entitled \"Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)\" is my own independent research work conducted under the academic supervision of Dr. Hoang Thi Thuy Nga. The survey dataset (n = 800) and econometric findings presented in this thesis are original, transparent, and have not been submitted for any other degree or qualification at any academic institution.", space_after=24)
    
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.paragraph_format.space_after = Pt(4)
    r_s1 = p_sig.add_run("Hanoi, June 2026\nStudent Author\n\n\n\n")
    r_s1.italic = True
    r_s1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r_s2 = p_sig.add_run("Dang Tu Linh")
    r_s2.bold = True
    r_s2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    doc.add_page_break()

    # ==================== ACKNOWLEDGEMENTS ====================
    doc.add_heading("ACKNOWLEDGEMENTS", level=1)
    doc.add_page_break()

    # ==================== ABSTRACT ====================
    doc.add_heading("ABSTRACT", level=1)
    p_abs = doc.add_paragraph()
    p_abs.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_abs.paragraph_format.space_before = Pt(12)
    p_abs.paragraph_format.space_after = Pt(12)
    r_abs1 = p_abs.add_run("Keywords: ")
    r_abs1.bold = True
    r_abs1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    r_abs2 = p_abs.add_run("Bank Guarantee Services, Corporate Selection Decision, Price Competitiveness, eFAST Digital Adoption, OLS Multiple Regression, VietinBank.")
    r_abs2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    doc.add_page_break()

    # ==================== TABLE OF CONTENTS ====================
    doc.add_heading("TABLE OF CONTENTS", level=1)

    # Add TOC entries
    toc_items = [
        ("STATEMENT OF AUTHORSHIP", "iii", 1),
        ("ACKNOWLEDGEMENTS", "iv", 1),
        ("ABSTRACT", "v", 1),
        ("LIST OF ABBREVIATIONS", "vii", 1),
        ("LIST OF TABLES", "viii", 1),
        ("LIST OF FIGURES", "ix", 1),
        ("CHAPTER 1: INTRODUCTION", "1", 1),
        ("1.1. Research Rationales & Background", "1", 2),
        ("1.2. Research Problem & Industry Context at VietinBank", "3", 2),
        ("1.3. Research Objectives", "4", 2),
        ("1.3.1. General Objective", "4", 3),
        ("1.3.2. Specific Objectives", "4", 3),
        ("1.4. Research Questions", "5", 2),
        ("1.5. Scope and Boundaries of the Study", "5", 2),
        ("1.6. Significance & Contributions of the Study", "6", 2),
        ("1.7. Structure of the Thesis", "6", 2),
        ("CHAPTER 2: LITERATURE REVIEW AND THEORETICAL FRAMEWORK", "8", 1),
        ("2.1. Overview of Bank Guarantee Services in Commercial Banking", "8", 2),
        ("2.1.1. Nature and Economic Functions of Bank Guarantees", "8", 3),
        ("2.1.2. Main Types of Corporate Bank Guarantees", "9", 3),
        ("2.1.3. Legal and Regulatory Framework", "11", 3),
        ("2.2. Theoretical Foundations", "13", 2),
        ("2.2.1. Financial Intermediation & Delegated Monitoring Theory", "13", 3),
        ("2.2.2. Credit Risk Pricing & Contingent Claim Theory", "14", 3),
        ("2.2.3. Service Quality Theory & SERVQUAL Model", "16", 3),
        ("2.2.4. Relationship Banking Theory", "18", 3),
        ("2.2.5. Technology Acceptance Model (TAM) & Digital Banking", "20", 3),
        ("2.3. Empirical Literature on Corporate Bank Selection", "22", 2),
        ("2.3.1. International Empirical Studies", "22", 3),
        ("2.3.2. Empirical Studies in the Vietnamese Banking Context", "24", 3),
        ("2.4. Research Gaps", "26", 2),
        ("2.5. Conceptual Framework and Research Hypotheses", "27", 2),
        ("2.5.1. Conceptual Research Framework", "27", 3),
        ("2.5.2. Hypothesis Development", "28", 3),
        ("CHAPTER 3: RESEARCH METHODOLOGY AND EMPIRICAL DESIGN", "31", 1),
        ("3.1. Overall Research Design & Analytical Process", "31", 2),
        ("3.2. Questionnaire Design & Measurement Scales", "32", 2),
        ("3.2.1. Operationalization of Variables", "32", 3),
        ("3.2.2. Mapping Scales with the Official 34-Item Survey Questionnaire", "35", 3),
        ("3.3. Population, Sampling Strategy and Data Collection", "36", 2),
        ("3.3.1. Target Population & Sampling Method", "36", 3),
        ("3.3.2. Sample Size Determination", "37", 3),
        ("3.3.3. Survey Administration across 155 VietinBank Branches", "38", 3),
        ("3.4. Econometric & Quantitative Analytical Methods", "39", 2),
        ("3.4.1. Descriptive Statistics", "39", 3),
        ("3.4.2. Scale Reliability Testing (Cronbach’s Alpha)", "40", 3),
        ("3.4.3. Exploratory Factor Analysis (EFA)", "41", 3),
        ("3.4.4. Factor Scores Extraction Method", "42", 3),
        ("3.4.5. Pearson Correlation Analysis & Multicollinearity Diagnostics (VIF)", "43", 3),
        ("3.4.6. Multiple Linear Regression Model Specification (OLS)", "44", 3),
        ("3.4.7. Sub-Group Difference Testing Methods (ANOVA & t-test)", "45", 3),
        ("CHAPTER 4: EMPIRICAL RESULTS, DISCUSSION AND MANAGERIAL RECOMMENDATIONS", "47", 1),
        ("4.1. Descriptive Statistics of the Sample (n = 800)", "47", 2),
        ("4.1.1. Ownership Type Distribution", "47", 3),
        ("4.1.2. Firm Revenue Scale Distribution", "49", 3),
        ("4.1.3. Operating Experience Distribution", "50", 3),
        ("4.1.4. Usage Distribution of Bank Guarantee Products", "51", 3),
        ("4.2. Scale Reliability Analysis Results (Cronbach’s Alpha)", "53", 2),
        ("4.3. Exploratory Factor Analysis Results (EFA)", "55", 2),
        ("4.3.1. EFA for Independent Variables", "55", 3),
        ("4.3.2. EFA for Dependent Variable (DEC)", "58", 3),
        ("4.4. Correlation Analysis & Multicollinearity Diagnostics (VIF)", "59", 2),
        ("4.5. Multiple Linear Regression Results (OLS)", "61", 2),
        ("4.5.1. Model Summary & Goodness of Fit", "61", 3),
        ("4.5.2. Estimated Coefficients and Hypothesis Testing", "63", 3),
        ("4.6. Sub-Group Difference Analysis Results (ANOVA & t-test)", "66", 2),
        ("4.6.1. Selection Differences across Ownership Types", "66", 3),
        ("4.6.2. Selection Differences across Firm Scales and Operating Experience", "68", 3),
        ("4.6.3. Selection Differences across Guarantee Product Types", "70", 3),
        ("4.7. Discussion of Empirical Findings", "72", 2),
        ("4.8. Managerial Implications & Policy Recommendations for VietinBank", "75", 2),
        ("4.8.1. Enhancing Core Service Capabilities (Response to Objective 1)", "75", 3),
        ("4.8.2. Strategies for Price Competitiveness & Processing Speed (Response to Objective 2)", "77", 3),
        ("4.8.3. Tailored Guarantee Packages for Corporate Segments (Response to Objective 3)", "79", 3),
        ("4.8.4. Breakthrough Digital Transformation Strategy via VietinBank eFAST (Response to Objective 4)", "81", 3),
        ("4.9. Policy Recommendations for the State Bank of Vietnam", "83", 2),
        ("4.10. Research Limitations and Suggestions for Future Research", "85", 2),
        ("REFERENCES", "87", 1),
        ("APPENDICES", "92", 1),
        ("Appendix 1: Official 34-Item Survey Questionnaire", "92", 2),
        ("Appendix 2: Sample Demographic Characteristics Output", "96", 2),
        ("Appendix 3: Cronbach’s Alpha Reliability Analysis Output", "98", 2),
        ("Appendix 4: EFA Total Variance Explained & Rotated Component Matrix Output", "100", 2),
        ("Appendix 5: OLS Multiple Regression, VIF & Sub-group ANOVA Output", "103", 2)
    ]

    for t_title, t_page, t_lvl in toc_items:
        add_toc_line(t_title, t_page, level=t_lvl)

    doc.add_page_break()

    # ==================== LIST OF ABBREVIATIONS ====================
    doc.add_heading("LIST OF ABBREVIATIONS", level=1)
    
    t_abbr = doc.add_table(rows=1, cols=2)
    t_abbr.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w = [Inches(2.0), Inches(4.5)]
    for i, w in enumerate(col_w): t_abbr.rows[0].cells[i].width = w
    hdr = t_abbr.rows[0].cells
    
    r1 = hdr[0].paragraphs[0].add_run("Abbreviation")
    r1.bold = True
    r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    
    r2 = hdr[1].paragraphs[0].add_run("Full Term / Meaning")
    r2.bold = True
    r2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    
    set_cell_borders(hdr[0], top={'sz': 12, 'val': 'single', 'color': '000000'}, bottom={'sz': 8, 'val': 'single', 'color': '000000'})
    set_cell_borders(hdr[1], top={'sz': 12, 'val': 'single', 'color': '000000'}, bottom={'sz': 8, 'val': 'single', 'color': '000000'})
    
    sample_abbr = [
        ("ANOVA", "Analysis of Variance"),
        ("APG", "Advance Payment Guarantee"),
        ("Big4", "Four Major State-Owned Commercial Banks in Vietnam (VietinBank, VCB, BIDV, Agribank)"),
        ("CCF", "Credit Conversion Factor (Basel Accord)"),
        ("EFA", "Exploratory Factor Analysis"),
        ("eFAST", "VietinBank Corporate Digital Banking Platform"),
        ("FDI", "Foreign Direct Investment Enterprises"),
        ("KMO", "Kaiser-Meyer-Olkin Measure of Sampling Adequacy"),
        ("OLS", "Ordinary Least Squares"),
        ("PG", "Performance Guarantee"),
        ("RM", "Relationship Manager"),
        ("SERVQUAL", "Service Quality Measurement Framework"),
        ("SME", "Small and Medium-sized Enterprises"),
        ("SOE", "State-Owned Enterprises"),
        ("STP", "Straight-Through Processing"),
        ("TAM", "Technology Acceptance Model"),
        ("TG", "Tender Guarantee / Bid Bond"),
        ("URDG 758", "Uniform Rules for Demand Guarantees, ICC Publication No. 758"),
        ("VIF", "Variance Inflation Factor")
    ]
    for idx, (a, e) in enumerate(sample_abbr):
        rc = t_abbr.add_row().cells
        for i, w in enumerate(col_w): rc[i].width = w
        ra = rc[0].paragraphs[0].add_run(a)
        ra.bold = True
        ra.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        re = rc[1].paragraphs[0].add_run(e)
        re.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        
        # Bottom border on the last row
        if idx == len(sample_abbr) - 1:
            set_cell_borders(rc[0], bottom={'sz': 12, 'val': 'single', 'color': '000000'})
            set_cell_borders(rc[1], bottom={'sz': 12, 'val': 'single', 'color': '000000'})

    doc.add_page_break()

    # ==================== LIST OF TABLES ====================
    doc.add_heading("LIST OF TABLES", level=1)
    doc.add_page_break()

    # ==================== LIST OF FIGURES ====================
    doc.add_heading("LIST OF FIGURES", level=1)
    doc.add_page_break()

    # ==================== CHAPTER 1 ====================
    doc.add_heading("CHAPTER 1: INTRODUCTION", level=1)

    doc.add_heading("1.1. Research Rationales & Background", level=2)
    doc.add_heading("1.2. Research Problem & Industry Context at VietinBank", level=2)
    
    doc.add_heading("1.3. Research Objectives", level=2)
    doc.add_heading("1.3.1. General Objective", level=3)
    add_p("The general objective of this study is to identify and assess the factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, and to propose managerial recommendations for improving VietinBank’s attractiveness and competitiveness in the corporate bank guarantee market.")

    doc.add_heading("1.3.2. Specific Objectives", level=3)
    add_p("1. Identify the key factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, based on relevant theories, previous empirical studies, and the characteristics of bank guarantee services.")
    add_p("2. Assess the direction and relative importance of these factors in explaining corporate customers’ selection decisions.")
    add_p("3. Examine whether corporate customers’ selection decisions differ across major firm characteristics, such as ownership type, firm size, operating experience, and types of bank guarantees used.")
    add_p("4. Propose managerial recommendations for VietinBank to improve its bank guarantee products and services and strengthen its ability to attract and retain corporate customers.")

    doc.add_heading("1.4. Research Questions", level=2)
    add_p("• Research Question 1: What key factors significantly influence corporate customers’ decision to choose VietinBank for bank guarantee services?")
    add_p("• Research Question 2: What is the direction and relative importance of each factor in explaining corporate selection decisions?")
    add_p("• Research Question 3: Do corporate selection decisions significantly differ across corporate ownership types, firm revenue sizes, operating tenure, and guarantee product lines?")
    add_p("• Research Question 4: What actionable managerial recommendations should VietinBank implement to enhance its competitive standing, product attractiveness, and client retention?")

    doc.add_heading("1.5. Scope and Boundaries of the Study", level=2)
    doc.add_heading("1.6. Significance & Contributions of the Study", level=2)
    doc.add_heading("1.7. Structure of the Thesis", level=2)

    doc.add_page_break()

    # ==================== CHAPTER 2 ====================
    doc.add_heading("CHAPTER 2: LITERATURE REVIEW AND THEORETICAL FRAMEWORK", level=1)

    doc.add_heading("2.1. Overview of Bank Guarantee Services in Commercial Banking", level=2)
    doc.add_heading("2.1.1. Nature and Economic Functions of Bank Guarantees", level=3)
    doc.add_heading("2.1.2. Main Types of Corporate Bank Guarantees", level=3)
    doc.add_heading("2.1.3. Legal and Regulatory Framework", level=3)

    doc.add_heading("2.2. Theoretical Foundations", level=2)
    doc.add_heading("2.2.1. Financial Intermediation & Delegated Monitoring Theory", level=3)
    doc.add_heading("2.2.2. Credit Risk Pricing & Contingent Claim Theory", level=3)
    doc.add_heading("2.2.3. Service Quality Theory & SERVQUAL Model", level=3)
    doc.add_heading("2.2.4. Relationship Banking Theory", level=3)
    doc.add_heading("2.2.5. Technology Acceptance Model (TAM) & Digital Banking", level=3)

    doc.add_heading("2.3. Empirical Literature on Corporate Bank Selection", level=2)
    doc.add_heading("2.3.1. International Empirical Studies", level=3)
    doc.add_heading("2.3.2. Empirical Studies in the Vietnamese Banking Context", level=3)

    doc.add_heading("2.4. Research Gaps", level=2)

    doc.add_heading("2.5. Conceptual Framework and Research Hypotheses", level=2)
    doc.add_heading("2.5.1. Conceptual Research Framework", level=3)
    doc.add_heading("2.5.2. Hypothesis Development", level=3)
    add_p("• Hypothesis H1: Price Competitiveness (COST_COMP) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H2: Processing Speed (PROC_SPEED) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H3: Digital eFAST Convenience (DIGITAL_CONV) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H4: Bank Reputation (BANK_REP) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H5: Relationship Banking & Limits (RELATIONSHIP) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H6: Staff Professionalism (STAFF_QUAL) has a positive impact on corporate customers' decision to choose VietinBank.")
    add_p("• Hypothesis H7: Collateral & Margin Flexibility (COLL_POLICY) has a positive impact on corporate customers' decision to choose VietinBank.")

    doc.add_page_break()

    # ==================== CHAPTER 3 ====================
    doc.add_heading("CHAPTER 3: RESEARCH METHODOLOGY AND EMPIRICAL DESIGN", level=1)

    doc.add_heading("3.1. Overall Research Design & Analytical Process", level=2)

    doc.add_heading("3.2. Questionnaire Design & Measurement Scales", level=2)
    doc.add_heading("3.2.1. Operationalization of Variables", level=3)

    # Table 3.1 - Clean Academic Black & White Style
    t_v = doc.add_table(rows=1, cols=5)
    t_v.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_w_v = [Inches(1.2), Inches(1.8), Inches(2.5), Inches(0.9), Inches(0.6)]
    for i, w in enumerate(col_w_v): t_v.rows[0].cells[i].width = w
    hdr_v = t_v.rows[0].cells
    
    h_titles = ["Code", "Variable Name", "Measurement Content (34 Items)", "Type", "Sign"]
    for i, title in enumerate(h_titles):
        rh = hdr_v[i].paragraphs[0].add_run(title)
        rh.bold = True
        rh.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        set_cell_borders(hdr_v[i], top={'sz': 12, 'val': 'single', 'color': '000000'}, bottom={'sz': 8, 'val': 'single', 'color': '000000'})

    v_data = [
        ("DEC", "Selection Decision", "Preference & priority choice of VietinBank over competitors (4 items)", "Dependent (Y)", "N/A"),
        ("COST_COMP", "Price Competitiveness", "Fee reasonableness, competitive pricing & discount incentives (4 items)", "Independent (X1)", "+"),
        ("PROC_SPEED", "Processing Speed", "Turnaround time, prompt issuance & simplified paperwork (4 items)", "Independent (X2)", "+"),
        ("DIGITAL_CONV", "Digital eFAST Convenience", "Online submission, 24/7 e-guarantees & digital status tracking (4 items)", "Independent (X3)", "+"),
        ("BANK_REP", "Bank Reputation", "Big4 brand prestige, financial strength & 100% acceptance (5 items)", "Independent (X4)", "+"),
        ("RELATIONSHIP", "Relationship & Limits", "Credit limit flexibility, relationship tenure & VIP care (5 items)", "Independent (X5)", "+"),
        ("STAFF_QUAL", "Staff Professionalism", "RM competence, legal advisory on Bidding Law & TT61 (4 items)", "Independent (X6)", "+"),
        ("COLL_POLICY", "Collateral Flexibility", "Flexible cash margin ratio & diverse pledged collateral (4 items)", "Independent (X7)", "+")
    ]
    for idx, (c, n, m, t, s) in enumerate(v_data):
        rc = t_v.add_row().cells
        for i, w in enumerate(col_w_v): rc[i].width = w
        r_c = rc[0].paragraphs[0].add_run(c); r_c.bold = True; r_c.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r_n = rc[1].paragraphs[0].add_run(n); r_n.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r_m = rc[2].paragraphs[0].add_run(m); r_m.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r_t = rc[3].paragraphs[0].add_run(t); r_t.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r_s = rc[4].paragraphs[0].add_run(s); r_s.bold = True; r_s.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        
        # Bottom border on the last row
        if idx == len(v_data) - 1:
            for cell in rc:
                set_cell_borders(cell, bottom={'sz': 12, 'val': 'single', 'color': '000000'})

    doc.add_heading("3.2.2. Mapping Scales with the Official 34-Item Survey Questionnaire", level=3)

    doc.add_heading("3.3. Population, Sampling Strategy and Data Collection", level=2)
    doc.add_heading("3.3.1. Target Population & Sampling Method", level=3)
    doc.add_heading("3.3.2. Sample Size Determination", level=3)
    doc.add_heading("3.3.3. Survey Administration across 155 VietinBank Branches", level=3)

    doc.add_heading("3.4. Econometric & Quantitative Analytical Methods", level=2)
    doc.add_heading("3.4.1. Descriptive Statistics", level=3)
    doc.add_heading("3.4.2. Scale Reliability Testing (Cronbach’s Alpha)", level=3)
    doc.add_heading("3.4.3. Exploratory Factor Analysis (EFA)", level=3)
    doc.add_heading("3.4.4. Factor Scores Extraction Method", level=3)
    doc.add_heading("3.4.5. Pearson Correlation Analysis & Multicollinearity Diagnostics (VIF)", level=3)
    
    doc.add_heading("3.4.6. Multiple Linear Regression Model Specification (OLS)", level=3)
    p_eq = doc.add_paragraph()
    p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq.paragraph_format.space_after = Pt(8)
    r_eq = p_eq.add_run(
        "DEC = β0 + β1*COST_COMP + β2*PROC_SPEED + β3*DIGITAL_CONV + β4*BANK_REP + β5*RELATIONSHIP + β6*STAFF_QUAL + β7*COLL_POLICY + ε     (3.1)"
    )
    r_eq.bold = True
    r_eq.font.size = Pt(10.5)
    r_eq.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    doc.add_heading("3.4.7. Sub-Group Difference Testing Methods (ANOVA & t-test)", level=3)

    doc.add_page_break()

    # ==================== CHAPTER 4 ====================
    doc.add_heading("CHAPTER 4: EMPIRICAL RESULTS, DISCUSSION AND MANAGERIAL RECOMMENDATIONS", level=1)

    doc.add_heading("4.1. Descriptive Statistics of the Sample (n = 800)", level=2)
    doc.add_heading("4.1.1. Ownership Type Distribution", level=3)
    doc.add_heading("4.1.2. Firm Revenue Scale Distribution", level=3)
    doc.add_heading("4.1.3. Operating Experience Distribution", level=3)
    doc.add_heading("4.1.4. Usage Distribution of Bank Guarantee Products", level=3)

    doc.add_heading("4.2. Scale Reliability Analysis Results (Cronbach’s Alpha)", level=2)

    doc.add_heading("4.3. Exploratory Factor Analysis Results (EFA)", level=2)
    doc.add_heading("4.3.1. EFA for Independent Variables", level=3)
    doc.add_heading("4.3.2. EFA for Dependent Variable (DEC)", level=3)

    doc.add_heading("4.4. Correlation Analysis & Multicollinearity Diagnostics (VIF)", level=2)

    doc.add_heading("4.5. Multiple Linear Regression Results (OLS)", level=2)
    doc.add_heading("4.5.1. Model Summary & Goodness of Fit", level=3)
    doc.add_heading("4.5.2. Estimated Coefficients and Hypothesis Testing", level=3)

    doc.add_heading("4.6. Sub-Group Difference Analysis Results (ANOVA & t-test)", level=2)
    doc.add_heading("4.6.1. Selection Differences across Ownership Types", level=3)
    doc.add_heading("4.6.2. Selection Differences across Firm Scales and Operating Experience", level=3)
    doc.add_heading("4.6.3. Selection Differences across Guarantee Product Types", level=3)

    doc.add_heading("4.7. Discussion of Empirical Findings", level=2)

    doc.add_heading("4.8. Managerial Implications & Policy Recommendations for VietinBank", level=2)
    doc.add_heading("4.8.1. Enhancing Core Service Capabilities (Response to Objective 1)", level=3)
    doc.add_heading("4.8.2. Strategies for Price Competitiveness & Processing Speed (Response to Objective 2)", level=3)
    doc.add_heading("4.8.3. Tailored Guarantee Packages for Corporate Segments (Response to Objective 3)", level=3)
    doc.add_heading("4.8.4. Breakthrough Digital Transformation Strategy via VietinBank eFAST (Response to Objective 4)", level=3)

    doc.add_heading("4.9. Policy Recommendations for the State Bank of Vietnam", level=2)
    doc.add_heading("4.10. Research Limitations and Suggestions for Future Research", level=2)

    doc.add_page_break()

    # ==================== REFERENCES ====================
    doc.add_heading("REFERENCES", level=1)

    refs = [
        "Al-Sabbagh, M. and Al-Khathlan, K. (2018) 'Factors influencing corporate clients’ choice of commercial banks for trade finance services', Journal of Financial Services Marketing, 23(2), pp. 71–82.",
        "Baltagi, B.H. (2008) Econometric Analysis of Panel Data. 4th edn. Chichester: Wiley.",
        "Barru, D.J. (2005) 'How to Guarantee Contractor Performance on International Construction Projects: Comparing Surety Bonds with Bank Guarantees and Standby Letters of Credit', The George Washington International Law Review, 37(1), pp. 51–94.",
        "Berger, A.N. and Udell, G.F. (1995) 'Relationship Lending and Lines of Credit in Small Firm Finance', Journal of Business, 68(3), pp. 351–381.",
        "Bertrams, R.I.V.F. (2013) Bank Guarantees in International Trade. 4th edn. The Hague: Kluwer Law International.",
        "Boot, A.W.A. (2000) 'Relationship Banking: What Do We Know?', Journal of Financial Intermediation, 9(1), pp. 7–25.",
        "Carletti, E., Leonello, A. and Marquez, R. (2023) 'Loan guarantees, bank underwriting policies and financial fragility', Journal of Financial Economics, 149(2), pp. 260–295.",
        "Davis, F.D. (1989) 'Perceived Usefulness, Perceived Ease of Use, and User Acceptance of Information Technology', MIS Quarterly, 13(3), pp. 319–340.",
        "DeYoung, R. and Roland, K.P. (2001) 'Product Mix, Revenue Mix, and Risk at Commercial Banks', Journal of Financial Intermediation, 10(2), pp. 115–144.",
        "Diamond, D.W. (1984) 'Financial Intermediation and Delegated Monitoring', Review of Economic Studies, 51(3), pp. 393–414.",
        "Hassan, A.A. et al. (2018) 'The problems and abuse of performance bond in the construction industry', IOP Conference Series: Earth and Environmental Science, 143, p. 012045.",
        "Ho Dinh Phi et al. (2023) 'Effect of Service Quality on Customer Loyalty: the Mediation of Customer Satisfaction, and Corporate Reputation in Banking Industry', Eurasian Journal of Business and Management, 11(3), pp. 145–160.",
        "International Chamber of Commerce (2010) Uniform Rules for Demand Guarantees (URDG 758). ICC Publication No. 758. Paris: ICC.",
        "Kaur, M. et al. (2021) 'The determinants of bank selection criteria of SMEs: a fuzzy analytic hierarchy approach', Journal of Science and Technology Policy Management, 12(4), pp. 580–605.",
        "Le Van Dung (2021) 'The nature of payment guarantee relationships at credit institutions', Industry and Trade Magazine, 8(April), pp. 45–52.",
        "Merton, R.C. (1974) 'On the Pricing of Corporate Debt: The Risk Structure of Interest Rates', Journal of Finance, 29(2), pp. 449–470.",
        "Narteh, B. (2013) 'SME bank selection and patronage behaviour in the Ghanaian banking industry', Management Research Review, 36(11), pp. 1061–1080.",
        "Nguyen, H. et al. (2024) 'The impact of service innovation on customer satisfaction and customer loyalty: a case in Vietnamese retail banks', Future Business Journal, 10(1), p. 14.",
        "Nguyen Thi Nhung and Nguyen Duy Phu (2015) 'Payment guarantees at Vietnamese commercial banks', Development and Integration Magazine, 25(35), pp. 62–67.",
        "Oke, A.E. (2018) 'Bonding capability of Nigerian contracting firms', Engineering, Construction and Architectural Management, 25(8), pp. 1012–1024.",
        "Parasuraman, A., Zeithaml, V.A. and Berry, L.L. (1988) 'SERVQUAL: A Multiple-Item Scale for Measuring Consumer Perceptions of Service Quality', Journal of Retailing, 64(1), pp. 12–40.",
        "Phan Thi Hang Nga et al. (2024) 'Service quality, customer satisfaction and loyalty: a case study in Vietnamese SMEs', Cogent Business & Management, 11(1), p. 2304512.",
        "Ramakrishnan, R.T.S. and Thakor, A.V. (1984) 'Information Reliability and a Theory of Financial Intermediation', Review of Economic Studies, 51(3), pp. 415–432.",
        "State Bank of Vietnam (2024) Circular No. 61/2024/TT-NHNN dated December 31, 2024, providing regulations on bank guarantees (effective April 1, 2025). Hanoi: SBV.",
        "Stiglitz, J.E. and Weiss, A. (1981) 'Credit Rationing in Markets with Imperfect Information', American Economic Review, 71(3), pp. 393–410.",
        "Turnbull, P.W. and Gibbs, M.L. (1989) 'The Selection of Banks and Banking Services among Corporate Customers in South Africa', International Journal of Bank Marketing, 7(5), pp. 36–42.",
        "Venkatesh, V. et al. (2003) 'User Acceptance of Information Technology: Toward a Unified View', MIS Quarterly, 27(3), pp. 425–478.",
        "Zeithaml, V.A. (1988) 'Consumer Perceptions of Price, Quality, and Value: A Means-End Model and Synthesis of Evidence', Journal of Marketing, 52(3), pp. 2–22.",
        "Zeithaml, V.A., Berry, L.L. and Parasuraman, A. (1996) 'The Behavioral Consequences of Service Quality', Journal of Marketing, 60(2), pp. 31–46.",
        "Zelie, E.M. (2023) 'Factors determining bank selection by micro- and small-sized enterprises: evidence from Ethiopia', International Journal of Bank Marketing, 41(5), pp. 1120–1142."
    ]

    for ref in refs:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_ref.paragraph_format.left_indent = Inches(0)
        p_ref.paragraph_format.first_line_indent = Inches(0)
        p_ref.paragraph_format.space_after = Pt(6)
        r_ref = p_ref.add_run(ref)
        r_ref.font.size = Pt(10.5)
        r_ref.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    doc.add_page_break()

    # ==================== APPENDICES ====================
    doc.add_heading("APPENDICES", level=1)
    doc.add_heading("Appendix 1: Official 34-Item Survey Questionnaire", level=2)
    doc.add_heading("Appendix 2: Sample Demographic Characteristics Output", level=2)
    doc.add_heading("Appendix 3: Cronbach’s Alpha Reliability Analysis Output", level=2)
    doc.add_heading("Appendix 4: EFA Total Variance Explained & Rotated Component Matrix Output", level=2)
    doc.add_heading("Appendix 5: OLS Multiple Regression, VIF & Sub-group ANOVA Output", level=2)

    # Output file path
    output_filename = "c:/Users/nguyen.tuan.minh/Desktop/DTL-Master-Project/DTL_Master_Thesis_Full_Structure_Template.docx"
    doc.save(output_filename)
    print(f"Successfully generated 100% Pure Black Pristine Thesis Template with Complete Table of Contents at {output_filename}!")

if __name__ == "__main__":
    generate_black_pristine_thesis_template()
