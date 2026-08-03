import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

def add_bottom_border_to_paragraph(paragraph, color_hex="888888", size="6"):
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), size)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color_hex)
    pBdr.append(bottom)
    pPr.append(pBdr)

def generate_final_docx():
    doc = docx.Document()
    
    # 1. Page Margins (1.0 inch = 2.54 cm standard)
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # 2. Running Header Configuration (Different First Page)
    section.different_first_page_header_footer = True
    
    # Running Header for pages 2 onwards (Right aligned, Italic, with GRAY bottom border)
    header = section.header
    p_head = header.paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_head = p_head.add_run("Thesis Design – MDE31 – Dang Tu Linh")
    r_head.font.name = 'Times New Roman'
    r_head.font.size = Pt(10)
    r_head.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    r_head.italic = True
    add_bottom_border_to_paragraph(p_head, color_hex="888888", size="6")

    # Configure Styles
    styles = doc.styles

    # Normal / Body Text Style
    style_normal = styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(6)

    # Heading 1 Style
    style_h1 = styles['Heading 1']
    style_h1.font.name = 'Times New Roman'
    style_h1.font.size = Pt(14)
    style_h1.font.bold = True
    style_h1.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    style_h1.paragraph_format.space_before = Pt(14)
    style_h1.paragraph_format.space_after = Pt(6)

    # Heading 2 Style
    style_h2 = styles['Heading 2']
    style_h2.font.name = 'Times New Roman'
    style_h2.font.size = Pt(12.5)
    style_h2.font.bold = True
    style_h2.font.color.rgb = RGBColor(0x00, 0x33, 0x66)
    style_h2.paragraph_format.space_before = Pt(10)
    style_h2.paragraph_format.space_after = Pt(4)

    # ==================== COVER / TITLE PAGE ====================
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.paragraph_format.space_after = Pt(2)
    r_hdr = p_header.add_run("VIETNAM-NETHERLANDS MASTER’S PROGRAM IN DEVELOPMENT ECONOMICS (MDE)")
    r_hdr.bold = True
    r_hdr.font.size = Pt(13)
    r_hdr.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_td = doc.add_paragraph()
    p_td.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_td.paragraph_format.space_after = Pt(18)
    r_td = p_td.add_run("THESIS DESIGN")
    r_td.bold = True
    r_td.font.size = Pt(15)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(36)
    r_title = p_title.add_run("FACTORS AFFECTING CORPORATE CUSTOMERS' DECISION TO CHOOSE BANK GUARANTEE PRODUCTS AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)")
    r_title.bold = True
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    p_meta1 = doc.add_paragraph()
    p_meta1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta1.paragraph_format.space_after = Pt(4)
    r_sup = p_meta1.add_run("Supervisor(s): Dr. Hoang Thi Thuy Nga")
    r_sup.font.size = Pt(12)
    r_sup.bold = True

    p_meta2 = doc.add_paragraph()
    p_meta2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta2.paragraph_format.space_after = Pt(18)
    r_stu = p_meta2.add_run("Student: Dang Tu Linh, MDE Class 31")
    r_stu.font.size = Pt(12)
    r_stu.bold = True

    p_meta3 = doc.add_paragraph()
    p_meta3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_loc = p_meta3.add_run("Hanoi, 06/2026")
    r_loc.font.size = Pt(11)
    r_loc.italic = True

    doc.add_page_break()

    # Helper function for justified body paragraphs
    def add_body(text, space_after=6):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(space_after)
        p.add_run(text)
        return p

    def add_bullet(text, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.space_after = Pt(space_after)
        p.add_run("• " + text)
        return p

    def add_numbered(num_str, text, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.space_after = Pt(space_after)
        r_num = p.add_run(num_str + " ")
        r_num.bold = True
        p.add_run(text)
        return p

    # ==================== SECTION I: INTRODUCTION ====================
    doc.add_heading("I. Introduction", level=1)

    # 1.1 Research Rationales
    doc.add_heading("1.1. Research Rationales", level=2)

    add_body(
        "Bank guarantee is an indispensable security instrument closely tied to almost all high-value commercial transactions of corporate enterprises. "
        "In the construction and engineering sector, Article 18 of Decree No. 37/2015/ND-CP (amended and supplemented by Decree No. 35/2023/ND-CP) mandates "
        "that contractors must submit an advance payment guarantee for contracts with advance payments exceeding 1 billion VND; Bidding Law No. 22/2023/QH15 "
        "strictly regulates tender security and performance security obligations. In international trade, bank guarantees issued under the Uniform Rules for Demand "
        "Guarantees (URDG 758) substitute cash margin deposits, liberating corporate working capital. Demand for bank guarantee products is therefore mandatory "
        "rather than a discretionary corporate choice."
    )

    add_body(
        "For commercial banks, guarantee issuance is a signature credit operation—commonly referred to as an off-balance sheet commitment. "
        "At the time of issuing a letter of guarantee, the bank does not disburse cash upfront but provides an irrevocable commitment to pay on behalf of the principal "
        "if the principal defaults. This contingent liability is monitored off-balance sheet and is converted into an on-balance sheet forced loan only when the bank "
        "actually executes the payout."
    )

    add_body(
        "This economic mechanism creates three strategic benefits for commercial banks: (i) banks generate steady issuing and maintenance fee income throughout "
        "the guarantee term without disbursing capital, building high-margin non-interest income that is resilient to interest rate volatility; (ii) guarantee limits "
        "anchor corporate clients into long-term credit relationships, opening cross-selling avenues for deposit, payment, trade finance, and FX products; and "
        "(iii) default risk materialises with low probability while off-balance sheet commitments consume regulatory capital at lower Credit Conversion Factors (CCF) "
        "compared to funded loans of equivalent value."
    )

    add_body(
        "This low-capital, high-fee, customer-anchoring profile makes the guarantee segment intensely competitive among commercial banks. Banks simultaneously "
        "lower fee schedules, compress appraisal and issuance lead times, relax collateral and margin requirements, and digitalise procedures to retain and acquire "
        "corporate clients. Given that corporate switching costs between issuing banks are relatively low, identifying which criteria corporate clients evaluate "
        "and which criteria carry decisive weight provides direct empirical value for VietinBank's product strategy."
    )

    add_body(
        "Vietnam's regulatory landscape for guarantee operations has recently undergone fundamental transformation. Circular No. 61/2024/TT-NHNN (effective April 1, 2025, "
        "replacing Circular No. 11/2022/TT-NHNN) establishes an updated legal framework for electronic guarantees (e-guarantees). Concurrently, commercial bank digital "
        "transformation and cross-border guarantee demand are reshaping corporate selection criteria—factors virtually absent from earlier empirical models. "
        "Existing literature shows that bank guarantee products remain under-researched, particularly in Vietnam, where studies focus on retail banking or legal aspects. "
        "VietinBank—a Big4 state-owned commercial bank with its VietinBank eFAST digital platform and global correspondent network—provides an ideal empirical context. "
        "Hence, the author selects the topic: \"Factors Affecting Corporate Customers' Decision to Choose Bank Guarantee Products at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)\"."
    )

    # 1.2 General Objectives
    doc.add_heading("1.2. General Objectives", level=2)

    add_body(
        "The general objective of this research is to identify and measure the impact magnitude of factors influencing corporate customers' decision to choose bank "
        "guarantee products at VietinBank, and on that basis, to propose actionable managerial implications helping VietinBank enhance competitiveness and expand "
        "market share in this strategic product segment."
    )

    # 1.3 Specific Objectives
    doc.add_heading("1.3. Specific Objectives", level=2)

    add_body("The following sub-questions will be investigated to answer the general objective:", space_after=4)
    
    sub_q = [
        ("1.", "What factors affect corporate customers' decision to choose bank guarantee products at VietinBank?"),
        ("2.", "What is the direction and magnitude of each factor's impact on the selection decision?"),
        ("3.", "Do impact magnitudes differ across corporate subgroups categorized by ownership structure, revenue scale, operating age, and guarantee types used?"),
        ("4.", "What managerial implications should VietinBank prioritize to enhance its selection probability as the preferred issuing bank?")
    ]
    for num, q_text in sub_q:
        add_numbered(num, q_text, space_after=4)

    # 1.4 Thesis Structure
    doc.add_heading("1.4. Thesis Structure", level=2)

    add_body("This thesis comprises four main chapters:", space_after=4)
    add_bullet("Chapter 1: Introduction – Presents research rationales, general and specific objectives, sub-questions, and thesis structure.")
    add_bullet("Chapter 2: Literature Review and Theoretical Framework – Synthesizes legal and economic foundations of bank guarantees; presents theoretical frameworks (SERVQUAL, TRA/TPB, Oliver's Loyalty Model); reviews empirical studies; and identifies four research gaps underlying the proposed 10-independent-variable model.")
    add_bullet("Chapter 3: Research Methodology and Empirical Design – Presents the multiple linear regression model and 10 hypotheses; constructs the 5-point Likert scale with 37 items; describes data sources from VietinBank's existing corporate survey dataset; and details SPSS statistical procedures.")
    add_bullet("Chapter 4: Empirical Results, Discussion and Policy Recommendations – Reports descriptive statistics, reliability tests, EFA, and regression results; discusses findings; and details managerial recommendations for VietinBank.", space_after=8)

    # ==================== SECTION II: LITERATURE REVIEW ====================
    doc.add_heading("II. Literature Review", level=1)

    # 2.1 Theoretical Foundations
    doc.add_heading("2.1. Theoretical Foundations and Legal Context", level=2)

    add_body(
        "Under Vietnamese law, Circular No. 61/2024/TT-NHNN (effective April 1, 2025, replacing Circular No. 11/2022/TT-NHNN and Circular No. 49/2024/TT-NHNN) "
        "defines bank guarantee as a form of credit extension wherein the guarantor (credit institution) commits to the beneficiary to perform financial obligations "
        "on behalf of the principal should the principal fail to perform. Article 335 of the 2015 Civil Code establishes general guarantee principles. "
        "Economically, bank guarantees are non-funded credit instruments that reallocate default risk, enable trust in commercial dealings, and conserve corporate cash flow. "
        "International practice under ICC URDG 758 highlights the principle of independence and documentary character."
    )

    add_body(
        "Guarantee operations involve a tri-partite relationship comprising the Guarantor (VietinBank), Principal (Corporate Client), and Beneficiary (Project Owner/Buyer). "
        "Key product lines include Tender Guarantee (TG), Performance Guarantee (PG), Advance Payment Guarantee, Payment Guarantee (BG), Maintenance Guarantee, and Counter Guarantee."
    )

    add_body("Corporate bank selection represents organizational buying behavior explained by three integrated theoretical models:", space_after=4)
    add_numbered("1.", "Theory of Reasoned Action (TRA, Fishbein & Ajzen, 1975) & Theory of Planned Behavior (TPB, Ajzen, 1991): Establishes that beliefs regarding service attributes form attitudes and perceived behavioral control, driving selection intention.")
    add_numbered("2.", "SERVQUAL Model (Parasuraman, Zeithaml & Berry, 1985, 1988): Maps service quality dimensions—Reliability, Responsiveness, Assurance, Empathy, and Tangibles—to bank guarantee attributes (SPE, STA, REP, CUS, DIG).")
    add_numbered("3.", "Customer Perceived Value (Zeithaml, 1988) & Oliver's 4-Stage Loyalty Model (1999) / Zeithaml et al. (1996): Ground the price dimension (COST) and operationalize the dependent variable (DEC) across Cognitive, Affective, Conative, and Action dimensions (DEC1–DEC4).", space_after=8)

    # 2.2 Empirical Review & Gaps
    doc.add_heading("2.2. Empirical Review and Research Gaps", level=2)

    add_body(
        "Empirical literature spans corporate bank selection (Turnbull & Gibbs, 1989; Narteh, 2013; Kaur et al., 2021; Zelie, 2023), Vietnamese banking service quality "
        "(Phan Thi Hang Nga et al., 2024; Ho Dinh Phi et al., 2023; Nguyen et al., 2024), and trade/construction guarantees (Oke, 2018; Hassan et al., 2018; Carletti et al., 2023). "
        "Four critical research gaps emerge: (i) separation between general bank selection literature and guarantee-specific empirical quantitative models; "
        "(ii) lack of guarantee-focused corporate selection studies in Vietnam; (iii) omission of modern digital (e-guarantees), customization, and legal risk advisory dimensions; "
        "and (iv) absence of large-scale system-wide empirical studies at VietinBank."
    )

    # ==================== SECTION III: EMPIRICAL ANALYSIS ====================
    doc.add_heading("III. Empirical Analysis", level=1)

    # 3.1 Methodology
    doc.add_heading("3.1. Methodology and Proposed Model", level=2)

    add_body(
        "The study employs a multiple linear regression model with one dependent variable—Corporate Selection Decision (DEC)—and 10 independent variables (X1 to X10). "
        "All variables are latent constructs measured via 5-point Likert scale items. The primary regression equation is specified as:"
    )

    p_eq = doc.add_paragraph()
    p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq.paragraph_format.space_after = Pt(10)
    r_eq = p_eq.add_run(
        "DEC = β0 + β1*COST + β2*COL + β3*SPE + β4*REP + β5*STA + β6*REL + β7*DIG + β8*CUS + β9*RSK + β10*NET + ε                             (1)"
    )
    r_eq.bold = True
    r_eq.font.size = Pt(10.5)

    # Table 1: Variable Definitions
    p_t1_title = doc.add_paragraph()
    p_t1_title.paragraph_format.space_after = Pt(4)
    r_t1 = p_t1_title.add_run("Table 1: Variable Definitions and Measurement Items")
    r_t1.bold = True
    r_t1.font.size = Pt(11)

    t1 = doc.add_table(rows=1, cols=5)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    col_w1 = [Inches(0.8), Inches(1.8), Inches(2.8), Inches(0.9), Inches(0.6)]

    hdr1 = t1.rows[0].cells
    hdr_titles1 = ["Code", "Variable Name", "Measurement Content", "Type", "Sign"]
    for idx, t in enumerate(hdr_titles1):
        hdr1[idx].width = col_w1[idx]
        p = hdr1[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(t)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(hdr1[idx], "003366")

    vars_info = [
        ("DEC", "Selection Decision", "4 items (DEC1-DEC4): Preference, Future growth, Recommendation, Overall relative choice", "Dependent", "N/A"),
        ("COST", "Fee & Cost Structure", "3 items (COST1-COST3): Fee reasonableness, competitiveness, discount policy", "Independent", "–"),
        ("COL", "Collateral & Margin", "3 items (COL1-COL3): Flexible margin ratio, collateral diversity, appraisal speed", "Independent", "+"),
        ("SPE", "Processing Speed", "3 items (SPE1-SPE3): Fast appraisal, simplified procedures, lead time compliance", "Independent", "+"),
        ("REP", "Bank Reputation", "3 items (REP1-REP3): Big4 brand power, beneficiary acceptance, financial standing", "Independent", "+"),
        ("STA", "Staff Competence", "3 items (STA1-STA3): Technical expertise, industry understanding, service attitude", "Independent", "+"),
        ("REL", "Relationship Banking", "3 items (REL1-REL3): Credit history, payroll/account ties, existing credit line", "Independent", "+"),
        ("DIG", "Digitalisation (e-FAST)", "3 items (DIG1-DIG3): 24/7 online submission, CA digital signature, QR verification", "Independent (New)", "+"),
        ("CUS", "Customised Solutions", "3 items (CUS1-CUS3): Tailored wording, multi-party structure, progress adjustment", "Independent (New)", "+"),
        ("RSK", "Legal Risk Advisory", "3 items (RSK1-RSK3): Base contract legal review, unfair calling protection, URDG 758", "Independent (New)", "+"),
        ("NET", "Global Correspondent Net", "3 items (NET1-NET3): Global bank network, counter-guarantee acceptance, FX handling", "Independent (New)", "+")
    ]

    for code, name, desc, vtype, sign in vars_info:
        r_cells = t1.add_row().cells
        for i, w in enumerate(col_w1): r_cells[i].width = w
        
        p0 = r_cells[0].paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.add_run(code).bold = True
        
        r_cells[1].paragraphs[0].add_run(name)
        r_cells[2].paragraphs[0].add_run(desc)
        
        p3 = r_cells[3].paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.add_run(vtype)

        p4 = r_cells[4].paragraphs[0]
        p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p4.add_run(sign).bold = True

    # 3.2 Data and Pre-Tests
    doc.add_heading("3.2. Data Collection and Pre-Tests", level=2)

    add_body(
        "Rather than administering a new small-scale survey, this study inherits an existing system-wide corporate survey dataset conducted by VietinBank "
        "across its corporate client base nationwide. The sampling frame covers corporate enterprises actively utilizing guarantee services across 155 VietinBank branches. "
        "The target dataset comprises n = 800 to n = 10,000 corporate respondents, far exceeding minimum sample size thresholds (n >= 185 for EFA; "
        "n >= 138 for regression). Pre-tests include data cleaning (removing missing values and straight-lining responses), descriptive statistics, "
        "Cronbach's Alpha reliability testing (threshold >= 0.6), Exploratory Factor Analysis (KMO >= 0.5, Bartlett sig < 0.05, Variance Explained > 50%), "
        "Pearson correlation, and VIF multicollinearity diagnostics (VIF < 10, ideally < 2). All estimations are executed in SPSS 26."
    )

    # 3.3 Findings and Interpretation
    doc.add_heading("3.3. Anticipated Findings and Interpretation", level=2)

    add_body("Based on theoretical frameworks and banking context, the following results are anticipated:", space_after=4)
    add_bullet("COST (X1) is expected to exert a statistically significant negative impact (β1 < 0), confirming price sensitivity.")
    add_bullet("All positive-dimension attributes (X2 to X10) are expected to yield positive coefficients (β2 to β10 > 0). Digitalisation (DIG) and Processing Speed (SPE) are projected to emerge among the strongest determinants, reflecting corporate urgency in tender deadlines. Customisation (CUS) and Legal Risk Advisory (RSK) will validate the strategic value of tailored banking solutions.", space_after=8)

    # ==================== SECTION IV: CONCLUSIONS ====================
    doc.add_heading("IV. Conclusions", level=1)

    add_body(
        "This thesis design establishes a comprehensive empirical framework investigating corporate bank guarantee selection decisions at VietinBank. "
        "By integrating classical SERVQUAL and TPB frameworks with four novel operational dimensions (e-guarantees, customization, legal risk advisory, and global correspondent network), "
        "the study fills critical empirical gaps in Vietnam's banking literature. Managerial recommendations will guide VietinBank in risk-based fee optimization, "
        "eFAST Straight-Through Processing (STP) auto-approval limits, flexible margin policies, and international correspondent network expansion."
    )

    # ==================== SECTION V: REFERENCES ====================
    # Clean LEFT alignment for all references (just like Pham Minh Hieu sample paper!)
    doc.add_heading("V. References", level=1)

    refs = [
        "Ajzen, I. (1991) 'The Theory of Planned Behavior', Organizational Behavior and Human Decision Processes, 50(2), pp. 179–211.",
        "Al-Sabbagh, M. and Al-Khathlan, K. (2018) 'Factors influencing corporate clients’ choice of commercial banks for trade finance services', Journal of Financial Services Marketing, 23(2), pp. 71–82.",
        "Baltagi, B.H. (2008) Econometric Analysis of Panel Data. 4th edn. Chichester: Wiley.",
        "Barru, D.J. (2005) 'How to Guarantee Contractor Performance on International Construction Projects: Comparing Surety Bonds with Bank Guarantees and Standby Letters of Credit', The George Washington International Law Review, 37(1), pp. 51–94.",
        "Bertrams, R.I.V.F. (2013) Bank Guarantees in International Trade. 4th edn. The Hague: Kluwer Law International.",
        "Carletti, E., Leonello, A. and Marquez, R. (2023) 'Loan guarantees, bank underwriting policies and financial fragility', Journal of Financial Economics, 149(2), pp. 260–295.",
        "Fishbein, M. and Ajzen, I. (1975) Belief, Attitude, Intention and Behavior: An Introduction to Theory and Research. Reading, MA: Addison-Wesley.",
        "Hassan, A.A. et al. (2018) 'The problems and abuse of performance bond in the construction industry', IOP Conference Series: Earth and Environmental Science, 143, p. 012045.",
        "Ho Dinh Phi et al. (2023) 'Effect of Service Quality on Customer Loyalty: the Mediation of Customer Satisfaction, and Corporate Reputation in Banking Industry', Eurasian Journal of Business and Management, 11(3), pp. 145–160.",
        "International Chamber of Commerce (2010) Uniform Rules for Demand Guarantees (URDG 758). ICC Publication No. 758. Paris: ICC.",
        "Kaur, M. (2015) 'Bank Selection Process and Market Segmentation: Evidence from Indian Exporting SMEs', Vision: The Journal of Business Perspective, 19(3), pp. 214–226.",
        "Kaur, M. et al. (2021) 'The determinants of bank selection criteria of SMEs: a fuzzy analytic hierarchy approach', Journal of Science and Technology Policy Management, 12(4), pp. 580–605.",
        "Le Van Dung (2021) 'The nature of payment guarantee relationships at credit institutions', Industry and Trade Magazine, 8(April), pp. 45–52.",
        "Narteh, B. (2013) 'SME bank selection and patronage behaviour in the Ghanaian banking industry', Management Research Review, 36(11), pp. 1061–1080.",
        "Ngo, M.V. and Nguyen, H.H. (2016) 'The Relationship between Service Quality, Customer Satisfaction and Customer Loyalty: An Investigation in Vietnamese Retail Banking Sector', Journal of Competitiveness, 8(2), pp. 103–116.",
        "Nguyen, H. et al. (2024) 'The impact of service innovation on customer satisfaction and customer loyalty: a case in Vietnamese retail banks', Future Business Journal, 10(1), p. 14.",
        "Nguyen Thi Nhung and Nguyen Duy Phu (2015) 'Payment guarantees at Vietnamese commercial banks', Development and Integration Magazine, 25(35), pp. 62–67.",
        "Oke, A.E. (2018) 'Bonding capability of Nigerian contracting firms', Engineering, Construction and Architectural Management, 25(8), pp. 1012–1024.",
        "Oliver, R.L. (1999) 'Whence Consumer Loyalty?', Journal of Marketing, 63(Special Issue), pp. 33–44.",
        "Parasuraman, A., Zeithaml, V.A. and Berry, L.L. (1985) 'A Conceptual Model of Service Quality and Its Implications for Future Research', Journal of Marketing, 49(4), pp. 41–50.",
        "Parasuraman, A., Zeithaml, V.A. and Berry, L.L. (1988) 'SERVQUAL: A Multiple-Item Scale for Measuring Consumer Perceptions of Service Quality', Journal of Retailing, 64(1), pp. 12–40.",
        "Phan Thi Hang Nga et al. (2024) 'Service quality, customer satisfaction and loyalty: a case study in Vietnamese SMEs', Cogent Business & Management, 11(1), p. 2304512.",
        "State Bank of Vietnam (2024) Circular No. 61/2024/TT-NHNN dated December 31, 2024, providing regulations on bank guarantees (effective April 1, 2025). Hanoi: SBV.",
        "Turnbull, P.W. and Gibbs, M.L. (1989) 'The Selection of Banks and Banking Services among Corporate Customers in South Africa', International Journal of Bank Marketing, 7(5), pp. 36–42.",
        "Zeithaml, V.A. (1988) 'Consumer Perceptions of Price, Quality, and Value: A Means-End Model and Synthesis of Evidence', Journal of Marketing, 52(3), pp. 2–22.",
        "Zeithaml, V.A. (1996) 'The Behavioral Consequences of Service Quality', Journal of Marketing, 60(2), pp. 31–46.",
        "Zelie, E.M. (2023) 'Factors determining bank selection by micro- and small-sized enterprises: evidence from Ethiopia', International Journal of Bank Marketing, 41(5), pp. 1120–1142."
    ]

    for ref in refs:
        p_ref = doc.add_paragraph()
        p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_ref.paragraph_format.left_indent = Inches(0)
        p_ref.paragraph_format.first_line_indent = Inches(0)
        p_ref.paragraph_format.space_after = Pt(6)
        p_ref.add_run(ref).font.size = Pt(10.5)

    # Save output file
    output_filename = "c:/Users/nguyen.tuan.minh/Desktop/DTL-Master-Project/DTL_Thesis_Design_NEU_MDE_Final.docx"
    doc.save(output_filename)
    print(f"Successfully updated Section V References with clean LEFT alignment (straight line along left margin)!")

if __name__ == "__main__":
    generate_final_docx()
