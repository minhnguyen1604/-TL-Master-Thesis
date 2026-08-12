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

def generate_option1_docx():
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
    r_title = p_title.add_run("FACTORS AFFECTING GUARANTEE FEE INCOME AND FINANCIAL EFFICIENCY OF CORPORATE CUSTOMERS AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)")
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
        "Bank guarantee is a fundamental credit instrument closely tied to almost all high-value commercial transactions of corporate enterprises. "
        "In the construction and engineering sector, Article 18 of Decree No. 37/2015/ND-CP (amended and supplemented by Decree No. 35/2023/ND-CP) mandates "
        "that contractors must submit an advance payment guarantee for contracts with advance payments exceeding 1 billion VND; Bidding Law No. 22/2023/QH15 "
        "strictly regulates tender security and performance security obligations. In international trade, bank guarantees issued under the Uniform Rules for Demand "
        "Guarantees (URDG 758) substitute cash margin deposits, liberating corporate working capital. Demand for bank guarantee products is therefore mandatory "
        "rather than a discretionary corporate choice."
    )

    add_body(
        "For commercial banks, guarantee issuance represents a signature credit operation—commonly referred to as an off-balance sheet commitment. "
        "At the time of issuing a letter of guarantee, the bank does not disburse cash upfront but provides an irrevocable commitment to pay on behalf of the principal "
        "if the principal defaults. This contingent liability is monitored off-balance sheet and is converted into an on-balance sheet forced loan only when the bank "
        "actually executes the payout."
    )

    add_body(
        "This economic mechanism creates three strategic financial benefits for commercial banks: (i) banks generate steady issuing and maintenance fee income throughout "
        "the guarantee term without disbursing capital, building high-margin non-interest fee income that is resilient to interest rate volatility; (ii) guarantee limits "
        "anchor corporate clients into long-term credit relationships, opening cross-selling avenues for deposit, payment, trade finance, and FX products; and "
        "(iii) default risk materialises with low probability while off-balance sheet commitments consume regulatory capital at lower Credit Conversion Factors (CCF) "
        "compared to funded loans of equivalent value."
    )

    add_body(
        "In commercial banking practice, Guarantee Fee Income (THU_PHI_BL) and Fee Yield Ratio serve as the primary financial metrics measuring off-balance sheet profitability "
        "and corporate customer relationship value. However, fee income generation is constrained by intense interbank competition, client price sensitivity, credit risk ratings, "
        "collateral margin requirements, and digital transformation. Identifying which financial and operational drivers significantly influence guarantee fee revenue and yield "
        "provides direct empirical value for VietinBank's asset-liability management and fee pricing strategies."
    )

    add_body(
        "Vietnam's regulatory landscape for guarantee operations has recently undergone fundamental transformation. Circular No. 61/2024/TT-NHNN (effective April 1, 2025, "
        "replacing Circular No. 11/2022/TT-NHNN) establishes an updated legal framework for electronic guarantees (e-guarantees). Concurrently, commercial bank digital "
        "transformation through VietinBank eFAST and global correspondent networks is reshaping fee extraction capabilities. "
        "Existing literature focuses predominantly on retail banking or overall bank selection, leaving transaction-level guarantee fee income under-researched. "
        "VietinBank—a Big4 state-owned commercial bank with its extensive corporate banking database—provides an ideal empirical context. "
        "Hence, the author selects Option 1: \"Factors Affecting Guarantee Fee Income and Financial Efficiency of Corporate Customers at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)\"."
    )

    # 1.2 General Objectives
    doc.add_heading("1.2. General Objectives", level=2)

    add_body(
        "The general objective of this research is to identify and measure the impact magnitude of transaction-level and firm-level financial determinants influencing "
        "guarantee fee income and fee yield efficiency at VietinBank, and on that basis, to propose actionable managerial implications helping VietinBank optimize fee pricing, "
        "enhance off-balance sheet profitability, and expand market share in this strategic product segment."
    )

    # 1.3 Specific Objectives
    doc.add_heading("1.3. Specific Objectives", level=2)

    add_body("The following sub-questions will be investigated to answer the general objective:", space_after=4)
    
    sub_q = [
        ("1.", "What transaction-level and firm-level financial factors significantly affect corporate guarantee fee income at VietinBank?"),
        ("2.", "What is the direction and magnitude of each financial factor's impact on total fee revenue and fee yield ratio?"),
        ("3.", "How do credit rating tiers, collateral margin ratios, and digital eFAST adoption moderate the fee income relationship?"),
        ("4.", "What risk-adjusted fee pricing policies and managerial recommendations should VietinBank implement to maximize off-balance sheet fee returns?")
    ]
    for num, q_text in sub_q:
        add_numbered(num, q_text, space_after=4)

    # 1.4 Thesis Structure
    doc.add_heading("1.4. Thesis Structure", level=2)

    add_body("This thesis comprises four main chapters:", space_after=4)
    add_bullet("Chapter 1: Introduction – Presents research rationales, general and specific objectives, sub-questions, and thesis structure.")
    add_bullet("Chapter 2: Literature Review and Theoretical Framework – Synthesizes legal, economic, and banking foundations of guarantee pricing; presents theoretical frameworks (Financial Intermediation, Credit Pricing, Relationship Banking); reviews empirical studies; and identifies key research gaps.")
    add_bullet("Chapter 3: Research Methodology and Empirical Design – Specifies the multiple linear regression model with log-transformed fee income; defines financial and firm variables; describes VietinBank's transaction dataset (n = 800); and details Python econometric procedures.")
    add_bullet("Chapter 4: Empirical Results, Discussion and Policy Recommendations – Reports descriptive statistics, correlation matrix, Pooled OLS and Robust regression results; discusses financial findings; and details managerial fee pricing recommendations for VietinBank.", space_after=8)

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
        "Key product lines include Tender Guarantee (TG), Performance Guarantee (PG), Advance Payment Guarantee (APG), Payment Guarantee (BG), Maintenance Guarantee, and Counter Guarantee."
    )

    add_body("Commercial bank guarantee fee pricing and off-balance sheet financial efficiency are grounded in four theoretical models:", space_after=4)
    add_numbered("1.", "Financial Intermediation Theory (Diamond, 1984; Ramakrishnan & Thakor, 1984): Explains how commercial banks act as delegated monitors, producing creditworthiness signals that justify charging guarantee issuing and maintenance fees.")
    add_numbered("2.", "Credit Risk Pricing & Contingent Claim Theory (Stiglitz & Weiss, 1981; Merton, 1974): Models guarantee fee schedules as a function of default risk probabilities, collateral margin protection, and underlying contract exposure.")
    add_numbered("3.", "Relationship Banking Theory (Boot, 2000; Berger & Udell, 1995): Posits that long-term credit history and multi-product ties (payroll, payment accounts, trade finance) reduce information asymmetry, enabling risk-adjusted fee discounting.")
    add_numbered("4.", "Customer Lifetime Value (CLV) & Non-Interest Revenue Theory (DeYoung & Roland, 2001): Operationalizes off-balance sheet fee income (THU_PHI_BL) and fee yield efficiency as primary components of non-interest bank profitability.", space_after=8)

    # 2.2 Empirical Review & Gaps
    doc.add_heading("2.2. Empirical Review and Research Gaps", level=2)

    add_body(
        "Empirical literature spans corporate bank selection (Turnbull & Gibbs, 1989; Narteh, 2013; Kaur et al., 2021), Vietnamese banking service quality "
        "(Phan Thi Hang Nga et al., 2024; Ho Dinh Phi et al., 2023; Nguyen et al., 2024), and trade finance fee pricing (Al-Sabbagh & Al-Khathlan, 2018; Carletti et al., 2023; Le Van Dung, 2021). "
        "Four critical research gaps emerge: (i) separation between survey-based perception studies and quantitative transaction-level fee revenue models; "
        "(ii) lack of guarantee-focused financial yield studies in Vietnam; (iii) omission of digital eFAST channel adoption and legal risk advisory impacts on fee realization; "
        "and (iv) absence of large-scale system-wide empirical transaction studies at VietinBank."
    )

    # ==================== SECTION III: EMPIRICAL ANALYSIS ====================
    doc.add_heading("III. Empirical Analysis", level=1)

    # 3.1 Methodology
    doc.add_heading("3.1. Methodology and Proposed Model", level=2)

    add_body(
        "The study employs a multiple linear regression model with primary dependent variable—Log of Guarantee Fee Income (ln_FEE)—and 8 independent financial and firm variables. "
        "To ensure econometric robustness, a secondary model evaluating Fee Yield Ratio (FEE_YIELD = FEE_INCOME / GUARANTEE_VOLUME * 100%) is specified. "
        "The primary regression equation is specified as:"
    )

    p_eq = doc.add_paragraph()
    p_eq.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_eq.paragraph_format.space_after = Pt(10)
    r_eq = p_eq.add_run(
        "ln(FEE_INCOME) = β0 + β1*ln(LIMIT) + β2*MARGIN_RATIO + β3*TENOR + β4*FIRM_SIZE + β5*FIRM_AGE + β6*CREDIT_RATING + β7*RELATIONSHIP + β8*DIGITAL + ε     (1)"
    )
    r_eq.bold = True
    r_eq.font.size = Pt(10.5)

    # Table 1: Variable Definitions
    p_t1_title = doc.add_paragraph()
    p_t1_title.paragraph_format.space_after = Pt(4)
    r_t1 = p_t1_title.add_run("Table 1: Variable Definitions and Measurement Specifications (Option 1)")
    r_t1.bold = True
    r_t1.font.size = Pt(11)

    t1 = doc.add_table(rows=1, cols=5)
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    t1.autofit = False
    col_w1 = [Inches(1.1), Inches(1.7), Inches(2.6), Inches(0.9), Inches(0.6)]

    hdr1 = t1.rows[0].cells
    hdr_titles1 = ["Code", "Variable Name", "Measurement Content / Source", "Type", "Sign"]
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
        ("ln_FEE", "Guarantee Fee Income", "Log of total annual guarantee fee revenue in VND (Core Banking MIS)", "Dependent (Y1)", "N/A"),
        ("FEE_YIELD", "Fee Yield Ratio", "Ratio of fee income to total guarantee volume (%) = (FEE/VOLUME)*100%", "Dependent (Y2)", "N/A"),
        ("ln_LIMIT", "Guarantee Credit Limit", "Log of approved guarantee limit amount in VND", "Independent (X1)", "+"),
        ("MARGIN_RATIO", "Collateral Margin Ratio", "Ratio of cash margin / collateral value to guarantee limit (%)", "Independent (X2)", "–"),
        ("TENOR", "Average Guarantee Tenor", "Average duration of guarantee commitments in months", "Independent (X3)", "+"),
        ("FIRM_SIZE", "Corporate Revenue Scale", "Log of corporate client annual revenue (VND)", "Independent (X4)", "+"),
        ("FIRM_AGE", "Corporate Operating Age", "Number of operating years since formal business registration", "Independent (X5)", "+"),
        ("CREDIT_RATING", "Internal Credit Rating", "VietinBank internal rating tier (Numerical score: 1=AAA to 10=C)", "Independent (X6)", "–"),
        ("RELATIONSHIP", "Banking Relationship Tenure", "Years of active credit and deposit relationship with VietinBank", "Independent (X7)", "–"),
        ("DIGITAL", "eFAST Digital Adoption", "Dummy variable = 1 if eFAST online guarantee user; 0 otherwise", "Independent (X8)", "+")
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

    # Table 1 Source Line (Gray Italic text below table)
    p_src = doc.add_paragraph()
    p_src.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_src.paragraph_format.space_before = Pt(4)
    p_src.paragraph_format.space_after = Pt(12)
    r_src = p_src.add_run("Source: Author's design based on VietinBank transactional database and literature.")
    r_src.font.name = 'Times New Roman'
    r_src.font.size = Pt(9.5)
    r_src.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r_src.italic = True

    # 3.2 Data and Pre-Tests
    doc.add_heading("3.2. Data Collection and Econometric Pre-Tests", level=2)

    add_body(
        "This study inherits an existing system-wide corporate transactional dataset extracted from VietinBank's Core Banking and Management Information System (MIS) "
        "across 155 VietinBank branches nationwide. The sampling frame covers active corporate guarantee accounts, yielding a clean dataset of n = 800 corporate observations. "
        "This sample size far exceeds minimum econometric thresholds (n >= 170 for multi-variable EFA; n >= 130 for multiple linear regression). "
        "Econometric procedures include anonymization, outlier winsorization at 1st and 99th percentiles, log transformation of skewed monetary variables, "
        "descriptive statistics, Pearson correlation matrix, Pooled OLS estimation, Breusch-Pagan heteroskedasticity test, VIF multicollinearity test (VIF < 5), "
        "and Huber-White robust standard errors. All data processing and regression estimations are executed in Python."
    )

    # 3.3 Findings and Interpretation
    doc.add_heading("3.3. Anticipated Findings and Interpretation", level=2)

    add_body("Based on theoretical credit pricing frameworks and banking context, the following results are anticipated:", space_after=4)
    add_bullet("Guarantee Limit (ln_LIMIT) and Tenor (TENOR) are projected to exert positive, highly significant impacts (β1 > 0, β3 > 0), confirming volume and duration as key revenue drivers.")
    add_bullet("Higher Credit Rating (better rating score) and longer Relationship Tenure (RELATIONSHIP) are expected to yield negative coefficients (β6 < 0, β7 < 0), reflecting preferential fee discount policies granted to high-quality, long-term corporate clients. Conversely, eFAST Digital Adoption (DIGITAL) will demonstrate a positive impact (β8 > 0), validating digital channel efficiency in driving overall fee realization.", space_after=8)

    # ==================== SECTION IV: CONCLUSIONS ====================
    doc.add_heading("IV. Conclusions", level=1)

    add_body(
        "This thesis design establishes a rigorous quantitative framework evaluating guarantee fee income and financial efficiency at VietinBank. "
        "By integrating transaction-level financial metrics with credit risk and relationship banking determinants, the study fills major empirical gaps in off-balance sheet banking literature. "
        "Managerial recommendations will guide VietinBank executive leadership in designing risk-adjusted fee pricing schedules, optimizing eFAST digital fee incentives, "
        "implementing relationship-based limit bundling, and balancing collateral margin requirements to maximize off-balance sheet non-interest returns."
    )

    # ==================== SECTION V: REFERENCES ====================
    doc.add_heading("V. References", level=1)

    refs = [
        "Al-Sabbagh, M. and Al-Khathlan, K. (2018) 'Factors influencing corporate clients’ choice of commercial banks for trade finance services', Journal of Financial Services Marketing, 23(2), pp. 71–82.",
        "Baltagi, B.H. (2008) Econometric Analysis of Panel Data. 4th edn. Chichester: Wiley.",
        "Barru, D.J. (2005) 'How to Guarantee Contractor Performance on International Construction Projects: Comparing Surety Bonds with Bank Guarantees and Standby Letters of Credit', The George Washington International Law Review, 37(1), pp. 51–94.",
        "Berger, A.N. and Udell, G.F. (1995) 'Relationship Lending and Lines of Credit in Small Firm Finance', Journal of Business, 68(3), pp. 351–381.",
        "Bertrams, R.I.V.F. (2013) Bank Guarantees in International Trade. 4th edn. The Hague: Kluwer Law International.",
        "Boot, A.W.A. (2000) 'Relationship Banking: What Do We Know?', Journal of Financial Intermediation, 9(1), pp. 7–25.",
        "Carletti, E., Leonello, A. and Marquez, R. (2023) 'Loan guarantees, bank underwriting policies and financial fragility', Journal of Financial Economics, 149(2), pp. 260–295.",
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
        "Phan Thi Hang Nga et al. (2024) 'Service quality, customer satisfaction and loyalty: a case study in Vietnamese SMEs', Cogent Business & Management, 11(1), p. 2304512.",
        "Ramakrishnan, R.T.S. and Thakor, A.V. (1984) 'Information Reliability and a Theory of Financial Intermediation', Review of Economic Studies, 51(3), pp. 415–432.",
        "State Bank of Vietnam (2024) Circular No. 61/2024/TT-NHNN dated December 31, 2024, providing regulations on bank guarantees (effective April 1, 2025). Hanoi: SBV.",
        "Stiglitz, J.E. and Weiss, A. (1981) 'Credit Rationing in Markets with Imperfect Information', American Economic Review, 71(3), pp. 393–410.",
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
    output_filename = "c:/Users/nguyen.tuan.minh/Desktop/DTL-Master-Project/DTL_Thesis_Design_Option1_Fee_Income_Final.docx"
    doc.save(output_filename)
    print(f"Successfully generated Option 1 Thesis Design Word document at {output_filename}!")

if __name__ == "__main__":
    generate_option1_docx()
