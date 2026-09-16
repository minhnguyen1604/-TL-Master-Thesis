# -*- coding: utf-8 -*-
"""Sinh file Word De cuong Luan van (Thesis Design) - chuan NEU MDE - mo hinh 7 bien."""
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

NAVY = RGBColor(0x00, 0x33, 0x66)
GREY = "D9D9D9"
J = WD_ALIGN_PARAGRAPH.JUSTIFY
C = WD_ALIGN_PARAGRAPH.CENTER


def shade(cell, hexfill):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hexfill)
    tcPr.append(shd)


def add_header(doc):
    hp = doc.sections[0].header.paragraphs[0]
    hp.text = "Thesis Design – MDE31 – Dang Tu Linh"
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.runs[0].font.size = Pt(9.5)
    hp.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    hp.runs[0].italic = True
    pPr = hp._element.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr'); bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single'); bot.set(qn('w:sz'), '6')
    bot.set(qn('w:space'), '1'); bot.set(qn('w:color'), '888888')
    pbdr.append(bot); pPr.append(pbdr)


def P(doc, text, size=11.5, bold=False, italic=False, align=J, color=None, after=6, indent=None):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.3
    if indent: p.paragraph_format.left_indent = Inches(indent)
    r = p.add_run(text)
    r.bold = bold; r.italic = italic; r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return p


def H(doc, text, lvl=1):
    sizes = {1: 13.5, 2: 12}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if lvl == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run(text); r.bold = True
    r.font.size = Pt(sizes[lvl]); r.font.color.rgb = NAVY
    return p


def table(doc, rows, widths, header=True):
    tb = doc.add_table(rows=len(rows), cols=len(rows[0]))
    tb.style = 'Table Grid'; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for k, v in enumerate(row):
            cell = tb.rows[i].cells[k]; cell.width = Inches(widths[k])
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(2)
            if k >= len(row) - 1 or (i == 0 and header):
                p.alignment = C
            r = p.add_run(v)
            r.font.size = Pt(10)
            if i == 0 and header:
                r.bold = True; shade(cell, GREY)
    return tb


REFS = [
 "Al-Sabbagh, M. and Al-Khathlan, K. (2018) 'Factors influencing corporate clients' choice of commercial banks for trade finance services', Journal of Financial Services Marketing, 23(2), pp. 71–82.",
 "Berger, A.N. and Udell, G.F. (1995) 'Relationship Lending and Lines of Credit in Small Firm Finance', Journal of Business, 68(3), pp. 351–381.",
 "Bertrams, R.I.V.F. (2013) Bank Guarantees in International Trade. 4th edn. The Hague: Kluwer Law International.",
 "Boot, A.W.A. (2000) 'Relationship Banking: What Do We Know?', Journal of Financial Intermediation, 9(1), pp. 7–25.",
 "Carletti, E., Leonello, A. and Marquez, R. (2023) 'Loan guarantees, bank underwriting policies and financial fragility', Journal of Financial Economics, 149(2), pp. 260–295.",
 "Davis, F.D. (1989) 'Perceived Usefulness, Perceived Ease of Use, and User Acceptance of Information Technology', MIS Quarterly, 13(3), pp. 319–340.",
 "Diamond, D.W. (1984) 'Financial Intermediation and Delegated Monitoring', Review of Economic Studies, 51(3), pp. 393–414.",
 "Hair, J.F., Black, W.C., Babin, B.J. and Anderson, R.E. (2019) Multivariate Data Analysis. 8th edn. Andover: Cengage Learning.",
 "Ho Dinh Phi et al. (2023) 'Effect of Service Quality on Customer Loyalty: the Mediation of Customer Satisfaction and Corporate Reputation in Banking Industry', Eurasian Journal of Business and Management, 11(3), pp. 145–160.",
 "International Chamber of Commerce (2010) Uniform Rules for Demand Guarantees (URDG 758). ICC Publication No. 758. Paris: ICC.",
 "Kaur, M. et al. (2021) 'The determinants of bank selection criteria of SMEs: a fuzzy analytic hierarchy approach', Journal of Science and Technology Policy Management, 12(4), pp. 580–605.",
 "Le Van Dung (2021) 'The nature of payment guarantee relationships at credit institutions', Industry and Trade Magazine, 8(April), pp. 45–52.",
 "Merton, R.C. (1974) 'On the Pricing of Corporate Debt: The Risk Structure of Interest Rates', Journal of Finance, 29(2), pp. 449–470.",
 "Narteh, B. (2013) 'SME bank selection and patronage behaviour in the Ghanaian banking industry', Management Research Review, 36(11), pp. 1061–1080.",
 "National Assembly of Vietnam (2015) Civil Code No. 91/2015/QH13. Hanoi.",
 "National Assembly of Vietnam (2023) Law on Bidding No. 22/2023/QH15, as amended by Law No. 57/2024/QH15. Hanoi.",
 "Nguyen, H. et al. (2024) 'The impact of service innovation on customer satisfaction and customer loyalty: a case in Vietnamese retail banks', Future Business Journal, 10(1), p. 14.",
 "Oliver, R.L. (1999) 'Whence Consumer Loyalty?', Journal of Marketing, 63(Special Issue), pp. 33–44.",
 "Parasuraman, A., Zeithaml, V.A. and Berry, L.L. (1988) 'SERVQUAL: A Multiple-Item Scale for Measuring Consumer Perceptions of Service Quality', Journal of Retailing, 64(1), pp. 12–40.",
 "Phan Thi Hang Nga et al. (2024) 'Service quality, customer satisfaction and loyalty: a case study in Vietnamese SMEs', Cogent Business & Management, 11(1), p. 2304512.",
 "Ramakrishnan, R.T.S. and Thakor, A.V. (1984) 'Information Reliability and a Theory of Financial Intermediation', Review of Economic Studies, 51(3), pp. 415–432.",
 "State Bank of Vietnam (2024) Circular No. 61/2024/TT-NHNN dated 31 December 2024 on bank guarantees (effective 1 April 2025). Hanoi: SBV.",
 "Stiglitz, J.E. and Weiss, A. (1981) 'Credit Rationing in Markets with Imperfect Information', American Economic Review, 71(3), pp. 393–410.",
 "Turnbull, P.W. and Gibbs, M.L. (1989) 'The Selection of Banks and Banking Services among Corporate Customers in South Africa', International Journal of Bank Marketing, 7(5), pp. 36–42.",
 "Venkatesh, V. et al. (2003) 'User Acceptance of Information Technology: Toward a Unified View', MIS Quarterly, 27(3), pp. 425–478.",
 "Vietnamese Government (2015) Decree No. 37/2015/ND-CP on construction contracts, as amended by Decree No. 35/2023/ND-CP. Hanoi.",
 "Zeithaml, V.A., Berry, L.L. and Parasuraman, A. (1996) 'The Behavioral Consequences of Service Quality', Journal of Marketing, 60(2), pp. 31–46.",
 "Zelie, E.M. (2023) 'Factors determining bank selection by micro- and small-sized enterprises: evidence from Ethiopia', International Journal of Bank Marketing, 41(5), pp. 1120–1142.",
]
print("Refs:", len(REFS))


def build():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Inches(1.0)
        s.left_margin = s.right_margin = Inches(1.0)
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'; st.font.size = Pt(11.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    add_header(doc)

    # ---- Title block
    P(doc, "VIETNAM-NETHERLANDS MASTER'S PROGRAM\nIN DEVELOPMENT ECONOMICS (MDE)", 12, True, align=C, color=NAVY, after=2)
    P(doc, "THESIS DESIGN", 15, True, align=C, color=NAVY, after=10)
    P(doc, "FACTORS AFFECTING CORPORATE CUSTOMERS' DECISION TO CHOOSE BANK GUARANTEE SERVICES "
           "AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)",
      13, True, align=C, color=NAVY, after=12)
    P(doc, "Supervisor: Dr. Hoang Thi Thuy Nga", 11.5, align=C, after=1)
    P(doc, "Student: Dang Tu Linh, MDE Class 31", 11.5, align=C, after=1)
    P(doc, "Hanoi, September 2026", 11.5, align=C, after=14)

    # ================= I. INTRODUCTION
    H(doc, "I. INTRODUCTION")
    H(doc, "1.1. Research Rationales", 2)
    P(doc, "Bank guarantees are an indispensable security instrument in high-value commercial transactions. Under "
           "Article 18 of Decree No. 37/2015/ND-CP (amended by Decree No. 35/2023/ND-CP), contractors must provide "
           "advance payment guarantees for construction contracts with advances exceeding VND 1 billion, while the "
           "Law on Bidding No. 22/2023/QH15 (as amended by Law No. 57/2024/QH15) mandates tender security and "
           "performance security. In cross-border trade, guarantees issued under ICC Uniform Rules for Demand "
           "Guarantees (URDG 758) substitute for cash margin deposits and release corporate working capital "
           "(Bertrams, 2013). Corporate demand for guarantee services is therefore largely compulsory rather than "
           "discretionary; what remains discretionary is the choice of issuing bank.")
    P(doc, "For commercial banks, guarantee issuance is an off-balance sheet credit commitment: the bank disburses "
           "no funds at issuance but assumes a contingent liability. This generates recurring fee income without "
           "funding cost, anchors corporate clients into long-term credit relationships, and consumes less "
           "regulatory capital than funded lending of equivalent value. These economics make the guarantee segment "
           "intensely contested, and because corporate switching costs between issuing banks are low, the criteria "
           "that drive bank choice carry direct strategic value.")
    P(doc, "Two contextual shifts make the question timely. First, Circular No. 61/2024/TT-NHNN (effective 1 April "
           "2025, replacing Circular No. 11/2022/TT-NHNN) established a full legal framework for electronic "
           "guarantees, turning digital issuance from a convenience into a competitive instrument. Second, "
           "empirical research on corporate bank selection has concentrated on lending and general service quality "
           "(Turnbull and Gibbs, 1989; Narteh, 2013; Kaur et al., 2021; Zelie, 2023), while guarantee-specific work "
           "remains predominantly legal and qualitative. VietinBank, a leading Vietnamese commercial bank operating "
           "155 domestic branches with the VietinBank eFAST digital platform, offers an appropriate empirical "
           "setting. The author therefore selects the topic stated above.")

    H(doc, "1.2. General Objective", 2)
    P(doc, "The general objective of this study is to identify and assess the factors associated with corporate "
           "customers' decision to choose VietinBank for bank guarantee services, and to propose managerial "
           "recommendations for improving VietinBank's attractiveness and competitiveness in the corporate bank "
           "guarantee market.")

    H(doc, "1.3. Specific Objectives", 2)
    for t in [
        "Objective 1: Identify the key factors associated with corporate customers' decision to choose VietinBank "
        "for bank guarantee services, based on relevant theories, previous empirical studies, and the "
        "characteristics of bank guarantee services.",
        "Objective 2: Assess the direction and relative importance of these factors in explaining corporate "
        "customers' selection decisions.",
        "Objective 3: Examine whether corporate customers' selection decisions differ across major firm "
        "characteristics, such as ownership type, firm size, operating experience, and types of bank guarantees used.",
        "Objective 4: Propose managerial recommendations for VietinBank to improve its bank guarantee products and "
        "services and strengthen its ability to attract and retain corporate customers."]:
        P(doc, "• " + t, 11.5, indent=0.25, after=4)

    H(doc, "1.4. Research Questions", 2)
    P(doc, "Each specific objective is operationalised through one corresponding research question:", after=4)
    for t in [
        "RQ1: Which factors significantly influence corporate customers' decision to choose VietinBank for bank "
        "guarantee services?",
        "RQ2: What is the direction and relative importance of each factor in explaining that decision?",
        "RQ3: Do selection decisions differ significantly across corporate ownership types, firm revenue sizes, "
        "operating tenure, and primary guarantee product lines?",
        "RQ4: Which managerial actions should VietinBank prioritise to strengthen corporate customer attraction "
        "and retention in the guarantee segment?"]:
        P(doc, "• " + t, 11.5, indent=0.25, after=4)

    H(doc, "1.5. Scope of the Study", 2)
    P(doc, "The study surveys corporate customers currently using bank guarantee services at VietinBank branches "
           "nationwide. Accordingly, the dependent construct measures the degree of selection priority and "
           "patronage intention towards VietinBank relative to competing banks, rather than a first-time binary "
           "choice between banks. The managerial scope is correspondingly defined as customer retention and "
           "expansion of guarantee wallet share within the existing corporate client base. This boundary is stated "
           "explicitly so that findings are not over-generalised to the market-wide bank selection process.")

    H(doc, "1.6. Thesis Structure", 2)
    for t in ["Chapter 1: Introduction — rationales, objectives, research questions, scope and structure.",
              "Chapter 2: Literature Review and Theoretical Framework — legal and economic foundations, five "
              "supporting theories, empirical review, research gaps, conceptual framework and hypotheses.",
              "Chapter 3: Research Methodology and Empirical Design — the seven-step quantitative procedure, the "
              "OLS model, questionnaire operationalisation, sampling plan, and the sub-group analysis design.",
              "Chapter 4: Empirical Results, Discussion and Managerial Recommendations — descriptive statistics, "
              "reliability and factor analysis, regression and sub-group results, discussion, and recommendations."]:
        P(doc, "• " + t, 11.5, indent=0.25, after=4)

    # ================= II. LITERATURE REVIEW
    H(doc, "II. LITERATURE REVIEW AND THEORETICAL FRAMEWORK")
    H(doc, "2.1. Legal Context and Theoretical Foundations", 2)
    P(doc, "Circular No. 61/2024/TT-NHNN defines a bank guarantee as a form of credit extension under which the "
           "guarantor commits to the beneficiary to discharge financial obligations on behalf of the principal upon "
           "default. Article 335 of the Civil Code 2015 sets the general civil law basis, while URDG 758 governs "
           "the independence and documentary character of international demand guarantees. Five theories support "
           "the proposed model:")
    for t in [
        "Financial Intermediation and Delegated Monitoring (Diamond, 1984; Ramakrishnan and Thakor, 1984): banks "
        "act as specialised information producers, so a guarantee from a reputable bank signals creditworthiness "
        "to third-party beneficiaries. This grounds Bank Reputation (BANK_REP).",
        "Credit Risk Pricing and Contingent Claims (Merton, 1974; Stiglitz and Weiss, 1981): guarantee fees and "
        "margin requirements price the underlying contingent credit exposure. This grounds Price Competitiveness "
        "(COST_COMP) and Collateral Policy (COLL_POLICY).",
        "SERVQUAL Service Quality Model (Parasuraman, Zeithaml and Berry, 1988): the responsiveness and assurance "
        "dimensions ground Processing Speed (PROC_SPEED) and Staff Professionalism (STAFF_QUAL). In corporate "
        "banking, institutional reputation is conceptually distinct from operational reliability and is therefore "
        "modelled as a separate construct rather than as a SERVQUAL dimension.",
        "Relationship Banking Theory (Boot, 2000; Berger and Udell, 1995): multi-product ties and relationship "
        "length reduce information asymmetry and enable flexible limits and pricing. This grounds RELATIONSHIP.",
        "Technology Acceptance Model (Davis, 1989; Venkatesh et al., 2003): perceived usefulness and ease of use "
        "drive adoption of digital channels, grounding Digital eFAST Convenience (DIGITAL_CONV) in the new legal "
        "environment for e-guarantees."]:
        P(doc, "• " + t, 11.5, indent=0.25, after=5)

    H(doc, "2.2. Empirical Review and Research Gaps", 2)
    P(doc, "Corporate bank selection has been studied by Turnbull and Gibbs (1989), Narteh (2013), Kaur et al. "
           "(2021) and Zelie (2023); Vietnamese service quality and loyalty by Phan Thi Hang Nga et al. (2024), Ho "
           "Dinh Phi et al. (2023) and Nguyen et al. (2024); trade finance selection by Al-Sabbagh and Al-Khathlan "
           "(2018); and guarantee economics by Carletti, Leonello and Marquez (2023) and Le Van Dung (2021). Four "
           "gaps follow: (i) off-balance sheet guarantee services are rarely isolated as the selection object; "
           "(ii) no published model combines price competitiveness with digital e-guarantee adoption under the new "
           "regulatory regime; (iii) sub-group comparison across corporate ownership types and guarantee product "
           "lines is absent in the Vietnamese literature; and (iv) no large-scale system-wide corporate study "
           "exists for a Vietnamese commercial bank in this segment.")

    H(doc, "2.3. Conceptual Framework and Research Hypotheses", 2)
    P(doc, "Following the supervisor's guidance on model parsimony, the framework retains seven independent "
           "constructs. Consistent with the measurement direction of the questionnaire — where a higher Likert "
           "score denotes a more favourable evaluation, including for fee competitiveness — all seven hypotheses "
           "are directional and positive:", after=5)
    for t in ["H1: Price Competitiveness (COST_COMP) has a positive effect on corporate selection priority.",
              "H2: Processing Speed (PROC_SPEED) has a positive effect on corporate selection priority.",
              "H3: Digital eFAST Convenience (DIGITAL_CONV) has a positive effect on corporate selection priority.",
              "H4: Bank Reputation (BANK_REP) has a positive effect on corporate selection priority.",
              "H5: Relationship Banking and Limits (RELATIONSHIP) has a positive effect on corporate selection priority.",
              "H6: Staff Professionalism (STAFF_QUAL) has a positive effect on corporate selection priority.",
              "H7: Collateral and Margin Flexibility (COLL_POLICY) has a positive effect on corporate selection priority."]:
        P(doc, "• " + t, 11.5, indent=0.25, after=4)
    P(doc, "H8 is stated as a difference hypothesis serving Objective 3: corporate selection priority differs "
           "significantly across ownership type, revenue size, operating tenure, and primary guarantee product line.",
      11.5, indent=0.25, after=6)

    # ================= III. METHODOLOGY
    H(doc, "III. RESEARCH METHODOLOGY AND EMPIRICAL DESIGN")
    H(doc, "3.1. Research Design and Analytical Procedure", 2)
    P(doc, "The study applies a quantitative cross-sectional survey design. Following the seven-step procedure "
           "approved by the supervisor: (1) descriptive statistics; (2) reliability testing by Cronbach's Alpha "
           "(alpha >= 0.70, corrected item-total correlation >= 0.30); (3) Exploratory Factor Analysis with Varimax "
           "rotation (KMO >= 0.50, Bartlett's test p < 0.05, eigenvalue >= 1.0, cumulative variance >= 50%, factor "
           "loading >= 0.50); (4) extraction of representative factor scores; (5) Pearson correlation analysis; "
           "(6) multicollinearity diagnostics with a conservative threshold of VIF < 3.0; and (7) OLS multiple "
           "regression together with sub-group difference testing. Thresholds follow Hair et al. (2019). All "
           "analysis is executed in Python (pandas, factor_analyzer, statsmodels, pingouin).")

    H(doc, "3.2. Econometric Model Specification", 2)
    P(doc, "The estimated equation is:", after=4)
    P(doc, "DEC = β0 + β1·COST_COMP + β2·PROC_SPEED + β3·DIGITAL_CONV + β4·BANK_REP "
           "+ β5·RELATIONSHIP + β6·STAFF_QUAL + β7·COLL_POLICY + ε     (1)",
      11.5, True, align=C, after=8)
    P(doc, "Table 1: Variable definitions, measurement items and expected signs", 11, True, align=C, after=4)
    table(doc, [
        ["Code", "Construct", "Measurement content", "Items", "Type", "Sign"],
        ["DEC", "Selection Priority & Patronage Intention", "Degree of priority given to VietinBank relative to competing banks when guarantee needs arise", "DEC1–DEC4", "Dependent", "—"],
        ["COST_COMP", "Price Competitiveness", "Fee reasonableness, competitiveness versus other banks, discount policy", "COMP1–COMP3", "X1", "+"],
        ["PROC_SPEED", "Processing Speed", "Appraisal turnaround, procedural simplicity, issuance lead time", "SPEED1–SPEED3", "X2", "+"],
        ["DIGITAL_CONV", "Digital eFAST Convenience", "24/7 online submission, e-guarantee with digital signature, online verification", "DIGI1–DIGI3", "X3", "+"],
        ["BANK_REP", "Bank Reputation", "Market standing, beneficiary acceptance of the guarantee, financial strength", "REPU1–REPU3", "X4", "+"],
        ["RELATIONSHIP", "Relationship & Limits", "Credit history, multi-product ties, flexibility of guarantee limits", "RELA1–RELA3", "X5", "+"],
        ["STAFF_QUAL", "Staff Professionalism", "Technical competence, legal advisory (Bidding Law, Circular 61, URDG 758), responsiveness", "STAFF1–STAFF3", "X6", "+"],
        ["COLL_POLICY", "Collateral & Margin Flexibility", "Flexible margin ratio, diversity of accepted collateral, valuation procedure", "COLL1–COLL3", "X7", "+"],
    ], [0.95, 1.25, 2.55, 0.85, 0.6, 0.4])
    P(doc, "Source: author's design. Total 25 Likert items = 7 constructs × 3 items + 4 dependent items.",
      9.5, italic=True, align=C, after=8)

    H(doc, "3.3. Operationalisation of the Dependent Construct", 2)
    P(doc, "Addressing the supervisor's comment that the dependent variable must capture the priority given to "
           "VietinBank relative to competing banks, all four observed items are framed comparatively and are "
           "introduced in the questionnaire by the instruction \"compared with other banks the enterprise has used "
           "or considered\":", after=5)
    table(doc, [
        ["Item", "Statement (translated)", "Dimension captured"],
        ["DEC1", "When guarantee needs arise, the enterprise prioritises VietinBank ahead of other banks.", "First-choice preference"],
        ["DEC2", "The enterprise allocates the majority of its guarantee value and volume to VietinBank rather than to other banks.", "Wallet share allocation"],
        ["DEC3", "For upcoming tenders and contracts, the enterprise intends to continue with VietinBank rather than switch.", "Continuation intention"],
        ["DEC4", "The enterprise would recommend VietinBank to partners and joint-venture contractors as the best issuing bank.", "Advocacy"],
    ], [0.7, 4.3, 1.6])
    P(doc, "The four indicators are adapted from the behavioural-consequences battery of Zeithaml, Berry and "
           "Parasuraman (1996). Oliver's (1999) loyalty framework is invoked at the conceptual level only, to "
           "justify a multi-indicator attitudinal construct; no one-to-one mapping to its four stages is claimed, "
           "since advocacy is not a stage in that framework and action loyalty is not observable in a "
           "cross-sectional design. Because the construct is a continuous multi-item attitudinal scale rather than "
           "a binary act of choice, OLS estimation is appropriate; a discrete-choice specification would be "
           "required only if the outcome were an observed switch between banks.", after=8)

    H(doc, "3.4. Population, Sampling Strategy and Data Collection", 2)
    P(doc, "The target population comprises corporate enterprises currently using guarantee services at VietinBank "
           "branches nationwide. Data will be collected through the structured questionnaire presented above, "
           "distributed via relationship managers across 155 domestic branches, with a target of n = 800 valid "
           "responses. Respondents are the officers responsible for guarantee transactions (CFO or board member, "
           "chief accountant, tender manager, or the assigned specialist). The target sample comfortably exceeds "
           "conventional minima — five observations per item (5 × 25 = 125) for EFA and 50 + 8k = 106 for "
           "regression with seven predictors. Data screening will remove incomplete forms and straight-lined "
           "responses before analysis. Because the questionnaire is distributed through the bank's own channel, "
           "social desirability bias is acknowledged as a limitation; descriptive statistics will be inspected for "
           "compressed standard deviations, and interpretation will rest on between-respondent variance rather "
           "than on absolute means.")

    H(doc, "3.5. Sub-group Analysis Design", 2)
    P(doc, "Objective 3 is served by four grouping variables collected in Part I of the questionnaire, each "
           "designed to be mutually exclusive so that one-way ANOVA is admissible:", after=5)
    table(doc, [
        ["Grouping criterion", "Categories", "Test"],
        ["Ownership type", "Private/LLC; Joint-stock (non-state); State-owned; FDI; Other", "One-way ANOVA + Tukey HSD"],
        ["Revenue size", "< 20bn; 20–100bn; 100–500bn; >= 500bn VND", "One-way ANOVA + Tukey HSD"],
        ["Operating tenure", "< 3; 3–5; 5–10; >= 10 years", "One-way ANOVA + Tukey HSD"],
        ["Primary guarantee type", "Tender (TG); Performance (PG); Advance payment (APG); Payment (BG); Other", "One-way ANOVA + Tukey HSD"],
        ["Multi-banking status", "VietinBank only vs. two or more banks", "Independent-samples t-test"],
    ], [1.5, 3.3, 1.8])
    P(doc, "Ownership type separates state-owned from FDI enterprises as distinct categories, and the guarantee "
           "product question is single-response, so that group membership is unambiguous. As a robustness check, "
           "ownership and size dummies together with multi-banking status are added to equation (1) to verify that "
           "the estimated coefficients β1 to β7 remain stable once composition effects are controlled for.",
      after=8)

    H(doc, "3.6. Anticipated Findings", 2)
    P(doc, "The seven constructs are expected to show positive and statistically significant coefficients, with "
           "Price Competitiveness and Bank Reputation anticipated among the strongest drivers and Digital eFAST "
           "Convenience expected to be significant given the new regulatory framework for e-guarantees. Sub-group "
           "tests are expected to reveal differences by ownership type and by primary guarantee product line. "
           "These are prior expectations derived from theory and are subject to empirical confirmation; no "
           "coefficient magnitudes are asserted in advance.")

    # ================= IV. CONCLUSIONS
    H(doc, "IV. EXPECTED CONTRIBUTIONS AND MANAGERIAL RECOMMENDATIONS")
    P(doc, "Managerial recommendations are organised so that each responds to one research question:", after=5)
    for t in [
        "Response to RQ1 — Factor identification. Consolidate performance management around the seven empirically "
        "identified drivers, and align branch-level service standards and product design with the constructs that "
        "prove significant.",
        "Response to RQ2 — Prioritisation by effect size. Allocate investment according to standardised beta "
        "rankings rather than statistical significance alone, concentrating on fee schedule design and reduction "
        "of issuance turnaround time where these emerge as dominant.",
        "Response to RQ3 — Differentiated segment strategies. Design distinct service packages for the corporate "
        "segments that show significant differences in selection priority, and tailor guarantee product offerings "
        "to the primary product line each segment uses.",
        "Response to RQ4 — Retention and digital channel development. Expand straight-through processing limits on "
        "VietinBank eFAST, strengthen legal advisory capability for Bidding Law and Circular 61 compliance, and "
        "apply flexible margin policies to reinforce wallet share among existing corporate clients."]:
        P(doc, "• " + t, 11.5, indent=0.25, after=5)
    P(doc, "In addition to bank-level actions, the study will formulate recommendations addressed to the State "
           "Bank of Vietnam, covering guidance for the consistent implementation of Circular 61/2024/TT-NHNN on "
           "electronic guarantees, interoperability standards for e-guarantee verification across credit "
           "institutions, and disclosure requirements on guarantee fee schedules that would strengthen price "
           "transparency in the corporate guarantee market.", after=6)
    P(doc, "Academically, the study contributes the first quantitative model in Vietnam that isolates bank "
           "guarantee services as the selection object and integrates digital e-guarantee adoption under the "
           "Circular 61/2024 regime. Practically, it supplies VietinBank with a ranked, evidence-based priority "
           "list for the corporate guarantee segment.", after=8)

    # ================= V. REFERENCES
    H(doc, "V. REFERENCES")
    for r in REFS:
        p = doc.add_paragraph()
        p.alignment = J
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.first_line_indent = Inches(-0.3)
        run = p.add_run(r); run.font.size = Pt(10.5)

    doc.save("DTL_Thesis_Design_NEU_MDE_Final.docx")
    print("Da tao DTL_Thesis_Design_NEU_MDE_Final.docx | so tai lieu tham khao:", len(REFS))


if __name__ == "__main__":
    build()
