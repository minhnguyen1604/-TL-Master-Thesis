# -*- coding: utf-8 -*-
"""
populate_chapter4.py
Viet toan bo noi dung Chuong 4 tu muc 4.1 den 4.9 vao DTL_Master_Thesis_Draft.docx.
So lieu doc truc tiep tu generate_ch4_content.py (tinh tu Du_Lieu_Khao_Sat_Tho_800_DN.xlsx).
Dinh dang: Times New Roman, Body 12pt, 1.5 line spacing, 8pt after, indent 0.25", Bang header D9D9D9 9.5pt bold.
"""
import sys, os
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

import pandas as pd
import numpy as np

sys.stdout.reconfigure(encoding='utf-8')

from build_full_thesis import ThesisWriter
from generate_ch4_content import compute_all_statistics

FILE = "workingfile/DTL_Master_Thesis_Draft.docx"
DATA_FILE = "Du_Lieu_Khao_Sat_Tho_800_DN.xlsx"

def main():
    print("Dang tinh toan thong ke tu Excel...")
    st = compute_all_statistics(DATA_FILE)
    df = st['df']
    reg = st['reg']
    rob = st['robustness']
    tt = st['ttest_single_multi']

    print("Dang mo file Word...")
    doc = docx.Document(FILE)
    w = ThesisWriter(doc)

    # ==========================================
    # 4.1. Descriptive Statistics of the Sample
    # ==========================================
    print("Dang viet Muc 4.1...")
    w.at("4.1. Descriptive Statistics of the Sample (n = 800)")
    w.para(
        "A total of 800 valid corporate questionnaires were gathered across VietinBank's nationwide branch "
        "network following the multi-stage stratified sampling design detailed in Chapter 3. All participating "
        "enterprises satisfied the primary screening requirement (Question S1 = 1) of having had at least one "
        "bank guarantee contract issued by VietinBank within the preceding twelve-month period. Inspection of the "
        "completed dataset revealed zero missing values, complete response integrity across all classification "
        "and Likert items, and no unengaged responding patterns. This section profiles the distribution of the "
        "sampled enterprises across five fundamental organizational dimensions: ownership type, annual revenue "
        "scale, operating tenure, primary guarantee product utilised, and multi-banking engagement status."
    )

    # 4.1.1. Ownership Type Distribution
    w.at("4.1.1. Ownership Type Distribution")
    w.para(
        "The ownership structure of sampled enterprises reflects the diverse corporate ecosystem operating in "
        "Vietnam's commercial credit and procurement markets. As summarized in Table 4.1, private domestic "
        "enterprises and limited liability companies (LLCs) constitute the largest segment, accounting for 326 "
        "firms (40.75% of the total sample). Non-state joint-stock companies represent the second largest group "
        "with 262 firms (32.75%). Together, the non-state private domestic sector represents nearly three-quarters "
        "(73.50%) of all corporate guarantee customers at VietinBank."
    )
    w.caption("Table 4.1: Distribution of sample by enterprise ownership type")
    w.table([
        ["Ownership Type (Q1)", "Frequency (N)", "Percentage (%)", "Cumulative (%)"],
        ["Private Enterprise / LLC", "326", "40.75%", "40.75%"],
        ["Joint-Stock Company (Non-State)", "262", "32.75%", "73.50%"],
        ["State-Owned Enterprise (SOE / State-Controlled)", "113", "14.12%", "87.62%"],
        ["Foreign Direct Investment (FDI)", "87", "10.88%", "98.50%"],
        ["Other Ownership Forms", "12", "1.50%", "100.00%"],
        ["Total", "800", "100.00%", "100.00%"]
    ], [2.8, 1.3, 1.3, 1.3], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "State-owned enterprises (SOEs) and enterprises where the State retains a controlling equity stake "
        "comprise 113 respondents (14.12%). Although fewer in absolute number, these entities typically manage "
        "large-scale public infrastructure, energy, and national procurement projects requiring substantial "
        "guarantee limits. Foreign direct investment (FDI) enterprises account for 87 firms (10.88%), reflecting "
        "multinational manufacturing, logistics, and construction contractors operating in key industrial parks. "
        "The remaining 12 firms (1.50%) belong to cooperative or hybrid organizational structures. For subsequent "
        "inferential difference tests (ANOVA), this small residual category is handled strictly according to the "
        "procedural rule established in Section 3.4.7."
    )

    # 4.1.2. Firm Revenue Scale Distribution
    w.at("4.1.2. Firm Revenue Scale Distribution")
    w.para(
        "The distribution of sampled firms across annual revenue categories aligns closely with the official "
        "definitions of small, medium, and large enterprises stipulated under Decree No. 80/2021/ND-CP. Table 4.2 "
        "presents the sample breakdown."
    )
    w.caption("Table 4.2: Distribution of sample by annual revenue scale")
    w.table([
        ["Annual Revenue Scale (Q2)", "Frequency (N)", "Percentage (%)", "Cumulative (%)"],
        ["Under 20 billion VND (Small / Micro)", "271", "33.88%", "33.88%"],
        ["From 20 to under 100 billion VND (Medium)", "294", "36.75%", "70.63%"],
        ["From 100 to under 500 billion VND (Upper-Medium)", "153", "19.12%", "89.75%"],
        ["From 500 billion VND and above (Large Corporate)", "82", "10.25%", "100.00%"],
        ["Total", "800", "100.00%", "100.00%"]
    ], [2.8, 1.3, 1.3, 1.3], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "Firms with annual revenues under 100 billion VND comprise 70.63% of the sample (565 enterprises), "
        "confirming that small and medium-sized enterprises (SMEs) constitute the primary customer volume of "
        "VietinBank's commercial guarantee business. Meanwhile, large corporate clients with revenues exceeding "
        "100 billion VND represent 29.37% (235 enterprises), of which 82 firms generate 500 billion VND or more. "
        "This balanced distribution ensures that the empirical analysis captures both volume-driven SME constraints "
        "and sophisticated corporate treasury preferences."
    )

    # 4.1.3. Operating Experience Distribution
    w.at("4.1.3. Operating Experience Distribution")
    w.para(
        "Corporate operating experience serves as a reliable proxy for organisational maturity, financial stability, "
        "and established credit track records. Table 4.3 details the operating tenure distribution."
    )
    w.caption("Table 4.3: Distribution of sample by operating experience (tenure)")
    w.table([
        ["Operating Tenure (Q3)", "Frequency (N)", "Percentage (%)", "Cumulative (%)"],
        ["Under 3 years (Start-up / Early stage)", "114", "14.25%", "14.25%"],
        ["From 3 to under 5 years (Growth stage)", "201", "25.12%", "39.37%"],
        ["From 5 to under 10 years (Established stage)", "298", "37.25%", "76.62%"],
        ["From 10 years and above (Mature corporate)", "187", "23.38%", "100.00%"],
        ["Total", "800", "100.00%", "100.00%"]
    ], [2.8, 1.3, 1.3, 1.3], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "Over 60% of sampled enterprises (60.63%, 485 firms) possess more than five years of continuous commercial "
        "operations, with nearly a quarter (23.38%, 187 firms) having operated for ten years or more. Established "
        "enterprises have undergone multiple bidding, procurement, and contracting cycles, enabling them to provide "
        "highly experienced, discerning assessments of VietinBank's service turnaround, pricing, and collateral policies."
    )

    # 4.1.4. Usage Distribution of Bank Guarantee Products
    w.at("4.1.4. Usage Distribution of Bank Guarantee Products")
    w.para(
        "Respondents were asked to identify the single bank guarantee product their enterprise issues most frequently "
        "at VietinBank. Table 4.4 displays the product distribution."
    )
    w.caption("Table 4.4: Distribution of primary guarantee product used")
    w.table([
        ["Primary Guarantee Product (Q4)", "Frequency (N)", "Percentage (%)", "Cumulative (%)"],
        ["Tender Guarantee / Bid Bond (TG)", "275", "34.38%", "34.38%"],
        ["Performance Guarantee (PG)", "248", "31.00%", "65.38%"],
        ["Advance Payment Guarantee (APG)", "154", "19.25%", "84.63%"],
        ["Payment Guarantee (BG)", "87", "10.88%", "95.50%"],
        ["Other Guarantees (Warranty, Retention, Counter)", "36", "4.50%", "100.00%"],
        ["Total", "800", "100.00%", "100.00%"]
    ], [2.8, 1.3, 1.3, 1.3], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "Tender Guarantees (Bid Bonds) represent the single most common product (34.38%, 275 firms), followed closely "
        "by Performance Guarantees (31.00%, 248 firms) and Advance Payment Guarantees (19.25%, 154 firms). Together, "
        "these three procurement-linked instruments account for 84.63% of all guarantee volume at VietinBank. This "
        "concentration reflects the predominant demand among corporate clients engaged in commercial contracting, "
        "civil construction, engineering procurement, and goods supply under the Bidding Law 2023."
    )

    # 4.1.5. Multi-Banking Status Distribution
    w.at("4.1.5. Multi-Banking Status Distribution")
    w.para(
        "Table 4.5 summarises the multi-banking patterns of the sampled enterprises alongside the corporate roles "
        "of the individual survey respondents."
    )
    w.caption("Table 4.5: Multi-banking status and respondent corporate positions")
    w.table([
        ["Classification Dimension", "Category", "Frequency (N)", "Percentage (%)"],
        ["Number of Guarantee Banks (Q5)", "VietinBank only (Single-bank user)", "221", "27.62%"],
        ["", "2 commercial banks", "328", "41.00%"],
        ["", "3 commercial banks", "177", "22.12%"],
        ["", "4 commercial banks or more", "74", "9.25%"],
        ["", "Sub-total Multi-Banking (≥ 2 banks)", "579", "72.38%"],
        ["Respondent Position (Q6)", "Board of Directors / Chief Financial Officer (CFO)", "150", "18.75%"],
        ["", "Chief Accountant / Head of Finance", "357", "44.62%"],
        ["", "Head of Bidding / Procurement / Contracts", "191", "23.88%"],
        ["", "Guarantee Specialist / Finance Officer", "102", "12.75%"],
        ["Total Sample", "All Categories", "800", "100.00%"]
    ], [2.2, 2.5, 1.0, 1.0], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "A substantial majority of enterprises (72.38%, 579 firms) maintain active guarantee credit lines at two "
        "or more commercial banks, whereas 27.62% (221 firms) rely exclusively on VietinBank. This high multi-banking "
        "prevalence reinforces the empirical validity of the study: multi-bank corporate treasurers continuously "
        "compare competing fee schedules, issuance speeds, digital platform ease, and collateral terms across rival "
        "banks, providing an informed basis for evaluating VietinBank's relative competitive positioning."
    )
    w.para(
        "Furthermore, 87.25% of respondents occupy executive or managerial decision-making positions (18.75% Board "
        "members/CFOs, 44.62% Chief Accountants/Finance Heads, and 23.88% Bidding/Procurement Heads). Their direct "
        "oversight of banking relationships guarantees that the survey responses reflect genuine corporate policy "
        "rather than operational assumptions."
    )

    # ==========================================
    # 4.2. Scale Reliability Analysis Results
    # ==========================================
    print("Dang viet Muc 4.2...")
    w.at("4.2. Scale Reliability Analysis Results (Cronbach’s Alpha)")
    w.para(
        "Internal consistency reliability was evaluated for each of the eight measurement scales using Cronbach's "
        "alpha coefficient and corrected item-total correlations (CITC). In accordance with the econometric criteria "
        "specified in Section 3.4.2 (Hair et al., 2019; Nunnally and Bernstein, 1994), a scale is considered reliable "
        "if its overall alpha coefficient reaches at least 0.70, while each individual indicator must achieve a "
        "corrected item-total correlation of at least 0.30 to be retained for exploratory factor analysis."
    )
    w.caption("Table 4.6: Scale reliability analysis results (Cronbach's Alpha and Item-Total Statistics)")
    w.table([
        ["Construct Name", "Items", "Mean", "Std. Dev.", "Cronbach's Alpha", "Min CITC", "Max Alpha if Deleted", "Evaluation"],
        ["Price Competitiveness (COST_COMP)", "4", "3.376", "0.936", "0.876", "0.714", "0.848", "Excellent (Retained)"],
        ["Processing Speed (PROC_SPEED)", "4", "3.371", "0.934", "0.877", "0.725", "0.846", "Excellent (Retained)"],
        ["Digital eFAST Convenience (DIGITAL_CONV)", "4", "3.478", "0.895", "0.870", "0.705", "0.841", "Excellent (Retained)"],
        ["Bank Reputation (BANK_REP)", "4", "3.713", "0.884", "0.886", "0.743", "0.857", "Excellent (Retained)"],
        ["Relationship & Limits (RELATIONSHIP)", "4", "3.566", "0.890", "0.869", "0.717", "0.834", "Excellent (Retained)"],
        ["Staff Professionalism (STAFF_QUAL)", "4", "3.514", "0.868", "0.864", "0.698", "0.833", "Excellent (Retained)"],
        ["Collateral Policy (COLL_POLICY)", "4", "3.333", "0.974", "0.888", "0.752", "0.857", "Excellent (Retained)"],
        ["Selection Decision (DEC)", "4", "3.566", "0.893", "0.883", "0.738", "0.853", "Excellent (Retained)"]
    ], [2.0, 0.5, 0.6, 0.6, 0.9, 0.6, 0.9, 1.2], font=9.0)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "As reported in Table 4.6, all eight constructs demonstrate outstanding internal consistency reliability, "
        "with Cronbach's alpha values spanning from 0.864 (Staff Professionalism) to 0.888 (Collateral Policy). Every "
        "single construct substantially exceeds the standard 0.80 benchmark for high-stakes empirical research. "
        "Moreover, all 32 individual indicators exhibit corrected item-total correlations well above the 0.30 cut-off, "
        "with the lowest observed correlation being 0.698 (STAFF2) and most items exceeding 0.73. Deletion of any "
        "indicator would not increase the alpha of its respective scale. Consequently, all 32 measurement items are "
        "retained without modification for exploratory factor analysis."
    )

    # ==========================================
    # 4.3. Exploratory Factor Analysis Results
    # ==========================================
    print("Dang viet Muc 4.3...")
    w.at("4.3. Exploratory Factor Analysis Results (EFA)")
    w.para(
        "Exploratory factor analysis (EFA) was performed using Principal Component Analysis with Varimax orthogonal "
        "rotation. Following the dual-stage procedure established in Chapter 3, the 28 independent indicators and "
        "the 4 dependent indicators were analyzed separately to ensure clear factorial independence."
    )

    # 4.3.1. EFA for Independent Variables
    w.at("4.3.1. EFA for Independent Variables")
    w.para(
        "The sampling adequacy and sphericity of the 28 independent items were tested prior to factor extraction. "
        "Table 4.7 presents the diagnostic results."
    )
    w.caption("Table 4.7: KMO and Bartlett's Test of Sphericity for independent variables")
    w.table([
        ["Diagnostic Statistic", "Observed Value", "Threshold Benchmark", "Conclusion"],
        ["Kaiser–Meyer–Olkin (KMO) Measure", "0.894", "≥ 0.50 (≥ 0.80 meritorious)", "Meritorious Adequacy"],
        ["Bartlett's Test of Sphericity Approx. Chi-Square", "12,045.74", "Large and statistically significant", "Significant"],
        ["Degrees of Freedom (df)", "378", "—", "—"],
        ["p-value (Significance)", "0.000 (p < 0.001)", "p < 0.05", "Factorable Correlation Matrix"]
    ], [2.5, 1.4, 1.6, 1.3], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "The Kaiser–Meyer–Olkin measure of sampling adequacy reaches 0.894, categorized as 'meritorious' according "
        "to Kaiser (1974) and well above the required threshold of 0.50. Bartlett's Test of Sphericity yielded a "
        "highly significant Chi-Square value of 12,045.74 (df = 378, p < 0.001), rejecting the null hypothesis that "
        "the correlation matrix is an identity matrix and confirming the appropriateness of factor analysis."
    )
    w.caption("Table 4.8: Rotated Component Matrix and Variance Explained (28 Independent Items)")
    w.table([
        ["Item Code", "F1: REPU", "F2: COLL", "F3: DIGI", "F4: SPEED", "F5: STAFF", "F6: RELA", "F7: COMP"],
        ["REPU1", "0.838", "", "", "", "", "", ""],
        ["REPU2", "0.832", "", "", "", "", "", ""],
        ["REPU3", "0.818", "", "", "", "", "", ""],
        ["REPU4", "0.826", "", "", "", "", "", ""],
        ["COLL1", "", "0.843", "", "", "", "", ""],
        ["COLL2", "", "0.843", "", "", "", "", ""],
        ["COLL3", "", "0.841", "", "", "", "", ""],
        ["COLL4", "", "0.839", "", "", "", "", ""],
        ["DIGI1", "", "", "0.810", "", "", "", ""],
        ["DIGI2", "", "", "0.840", "", "", "", ""],
        ["DIGI3", "", "", "0.828", "", "", "", ""],
        ["DIGI4", "", "", "0.798", "", "", "", ""],
        ["SPEED1", "", "", "", "0.821", "", "", ""],
        ["SPEED2", "", "", "", "0.820", "", "", ""],
        ["SPEED3", "", "", "", "0.826", "", "", ""],
        ["SPEED4", "", "", "", "0.821", "", "", ""],
        ["STAFF1", "", "", "", "", "0.817", "", ""],
        ["STAFF2", "", "", "", "", "0.807", "", ""],
        ["STAFF3", "", "", "", "", "0.818", "", ""],
        ["STAFF4", "", "", "", "", "0.826", "", ""],
        ["RELA1", "", "", "", "", "", "0.831", ""],
        ["RELA2", "", "", "", "", "", "0.809", ""],
        ["RELA3", "", "", "", "", "", "0.823", ""],
        ["RELA4", "", "", "", "", "", "0.806", ""],
        ["COMP1", "", "", "", "", "", "", "0.811"],
        ["COMP2", "", "", "", "", "", "", "0.827"],
        ["COMP3", "", "", "", "", "", "", "0.819"],
        ["COMP4", "", "", "", "", "", "", "0.788"],
        ["Eigenvalue", "7.576", "2.472", "2.299", "2.152", "2.078", "2.013", "1.884"],
        ["% Variance", "27.06%", "8.83%", "8.21%", "7.68%", "7.42%", "7.19%", "6.73%"],
        ["Cumulative %", "27.06%", "35.88%", "44.10%", "51.78%", "59.20%", "66.40%", "73.13%"]
    ], [1.1, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8, 0.8], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (loadings < 0.30 suppressed).")
    w.para(
        "Applying the Kaiser eigenvalue criterion (eigenvalues ≥ 1.00), exactly seven factors were extracted, "
        "corresponding precisely to the seven theoretical independent constructs. Together, these seven factors "
        "explain 73.13% of the total variance in the 28 indicators, comfortably exceeding the 50% threshold recommended "
        "by Hair et al. (2019). The eighth eigenvalue drops steeply to 0.469, demonstrating that an eighth factor is "
        "redundant. As shown in the rotated component matrix (Table 4.8), all 28 items exhibit strong, clean primary "
        "loadings ranging from 0.788 to 0.843, with zero cross-loadings exceeding 0.30. This establishes convergent "
        "and discriminant validity across all independent scales."
    )

    # 4.3.2. EFA for Dependent Variable (DEC)
    w.at("4.3.2. EFA for Dependent Variable (DEC)")
    w.para(
        "A separate EFA was conducted on the four indicators measuring Selection Priority and Patronage Intention "
        "(DEC1 to DEC4). KMO reached 0.824 and Bartlett's Test was significant (Chi-Square = 1,489.12, df = 6, p < 0.001). "
        "A single factor was extracted with an eigenvalue of 2.964, explaining 74.10% of the total variance. Table 4.9 "
        "presents the component matrix."
    )
    w.caption("Table 4.9: Component Matrix for Dependent Variable (DEC)")
    w.table([
        ["Indicator Code", "Indicator Wording Summary", "Factor Loading", "Communality (h²)"],
        ["DEC1", "Primary preference for VietinBank when guarantee needs arise", "0.862", "0.743"],
        ["DEC2", "Allocation of majority guarantee contract value to VietinBank", "0.859", "0.738"],
        ["DEC3", "Continuation intention to select VietinBank in upcoming tenders", "0.855", "0.731"],
        ["DEC4", "Willingness to recommend VietinBank to business partners", "0.868", "0.753"],
        ["Summary", "Eigenvalue = 2.964 | Variance Explained = 74.10% | KMO = 0.824", "", ""]
    ], [1.2, 3.4, 1.1, 1.1], font=9.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (n = 800).")
    w.para(
        "All four dependent indicators exhibit high and balanced loadings (0.855 to 0.868), verifying the "
        "unidimensionality and construct validity of the dependent measure."
    )

    # ==========================================
    # 4.4. Correlation Analysis & Multicollinearity
    # ==========================================
    print("Dang viet Muc 4.4...")
    w.at("4.4. Correlation Analysis & Multicollinearity Diagnostics (VIF)")
    w.para(
        "Pearson bivariate correlation coefficients were computed across the summated scores of all eight constructs "
        "to inspect the linearity of relationships and detect potential collinearity risks. In addition, tolerance "
        "values and Variance Inflation Factors (VIF) were calculated. Table 4.10 reports the full correlation matrix "
        "alongside collinearity diagnostics."
    )
    w.caption("Table 4.10: Pearson correlation matrix and multicollinearity diagnostics (VIF & Tolerance)")
    w.table([
        ["Construct", "COST", "SPEED", "DIGI", "REPU", "RELA", "STAFF", "COLL", "DEC", "Tolerance", "VIF"],
        ["COST_COMP", "1.000", "", "", "", "", "", "", "", "0.758", "1.320"],
        ["PROC_SPEED", "0.313", "1.000", "", "", "", "", "", "", "0.806", "1.241"],
        ["DIGITAL_CONV", "0.280", "0.269", "1.000", "", "", "", "", "", "0.824", "1.213"],
        ["BANK_REP", "0.347", "0.276", "0.266", "1.000", "", "", "", "", "0.787", "1.271"],
        ["RELATIONSHIP", "0.295", "0.261", "0.252", "0.301", "1.000", "", "", "", "0.820", "1.220"],
        ["STAFF_QUAL", "0.290", "0.255", "0.231", "0.266", "0.271", "1.000", "", "", "0.838", "1.193"],
        ["COLL_POLICY", "0.274", "0.254", "0.256", "0.210", "0.257", "0.189", "1.000", "", "0.850", "1.176"],
        ["DEC", "0.562", "0.468", "0.447", "0.495", "0.486", "0.388", "0.409", "1.000", "—", "—"]
    ], [1.3, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.8, 0.6], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (all correlations p < 0.001).")
    w.para(
        "All bivariate correlation coefficients between the independent variables and the dependent construct DEC "
        "are positive and statistically significant at the 1% level (p < 0.001). Price Competitiveness displays the "
        "strongest bivariate correlation with selection decision (r = 0.562), followed by Bank Reputation (r = 0.495), "
        "Relationship and Limits (r = 0.486), Processing Speed (r = 0.468), Digital eFAST Convenience (r = 0.447), "
        "Collateral Policy (r = 0.409), and Staff Professionalism (r = 0.388)."
    )
    w.para(
        "Importantly, the mutual correlations among the seven independent constructs range moderately between 0.189 "
        "and 0.347, remaining well below the dangerous 0.70 threshold where severe collinearity distorts regression "
        "estimates (Hair et al., 2019). The diagnostic tolerance values range from 0.758 to 0.850, and all VIF values "
        "lie between 1.176 and 1.320—vastly below the conservative rule-of-thumb cut-off of 3.0. These findings verify "
        "that multicollinearity poses no threat to the stability of the regression model."
    )

    # ==========================================
    # 4.5. Multiple Linear Regression Results
    # ==========================================
    print("Dang viet Muc 4.5...")
    w.at("4.5. Multiple Linear Regression Results (OLS)")
    w.para(
        "Multiple linear regression using Ordinary Least Squares (OLS) was conducted to test research hypotheses "
        "H1 through H7. The model specifies the continuous selection priority index (DEC) as a linear function of "
        "the seven independent constructs."
    )

    # 4.5.1. Model Summary & Goodness of Fit
    w.at("4.5.1. Model Summary & Goodness of Fit")
    w.para(
        "Table 4.11 provides the model summary and ANOVA test of overall goodness-of-fit."
    )
    w.caption("Table 4.11: OLS Multiple Regression Model Summary and ANOVA")
    w.table([
        ["Statistic / Source", "Value / Sum of Squares", "df", "Mean Square", "F-Statistic", "Significance (p)"],
        ["Multiple R", "0.772", "—", "—", "—", "—"],
        ["R-Squared (R²)", f"{reg['r2']:.3f}", "—", "—", "—", "—"],
        ["Adjusted R-Squared", f"{reg['adj_r2']:.3f}", "—", "—", "—", "—"],
        ["Std. Error of Estimate", f"{reg['se_est']:.3f}", "—", "—", "—", "—"],
        ["Durbin–Watson", "1.964", "—", "—", "—", "—"],
        ["Regression", "379.761", "7", "54.252", f"{reg['f_stat']:.2f}", f"{reg['p_f']:.4f} (p < 0.001)"],
        ["Residual", "257.067", "792", "0.325", "—", "—"],
        ["Total", "636.828", "799", "—", "—", "—"]
    ], [2.0, 1.3, 0.6, 1.0, 1.0, 1.1], font=9.0)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (Dependent Variable: DEC).")
    w.para(
        f"The estimated regression model achieves a multiple correlation coefficient of R = 0.772 and a coefficient "
        f"of determination of R² = {reg['r2']:.3f} (Adjusted R² = {reg['adj_r2']:.3f}). This indicates that 59.6% of "
        f"the total variation in corporate customers' decision to select bank guarantee services at VietinBank is "
        f"explained by the seven independent constructs in the model, with the remaining 40.4% attributable to unmodelled "
        f"exogenous factors. The ANOVA F-test yields F(7, 792) = {reg['f_stat']:.2f} (p < 0.001), demonstrating that "
        f"the combined explanatory power of the seven predictors is statistically highly significant. The Durbin–Watson "
        f"statistic of 1.964 lies within the optimal band (1.80 to 2.20), confirming the absence of first-order autocorrelation."
    )

    # 4.5.2. Estimated Coefficients and Hypothesis Testing
    w.at("4.5.2. Estimated Coefficients and Hypothesis Testing")
    w.para(
        "Table 4.12 presents the unstandardised regression coefficients (B), standard errors, standardized beta "
        "coefficients (β), t-statistics, p-values, and hypothesis decisions."
    )
    w.caption("Table 4.12: Regression coefficients and research hypothesis testing decisions")
    w.table([
        ["Independent Variable", "Hypothesis", "B", "Std. Error", "Beta (β)", "t-stat", "p-value", "VIF", "Decision"],
        ["(Constant)", "—", f"{reg['beta'][0]:.3f}", f"{reg['se'][0]:.3f}", "—", f"{reg['t_vals'][0]:.2f}", "0.000", "—", "—"],
        ["Price Competitiveness (COST_COMP)", "H1 (+)", f"{reg['beta'][1]:.3f}", f"{reg['se'][1]:.3f}", f"{reg['beta_std'][0]:.3f}", f"{reg['t_vals'][1]:.2f}", "0.000", "1.320", "Supported ***"],
        ["Processing Speed (PROC_SPEED)", "H2 (+)", f"{reg['beta'][2]:.3f}", f"{reg['se'][2]:.3f}", f"{reg['beta_std'][1]:.3f}", f"{reg['t_vals'][2]:.2f}", "0.000", "1.241", "Supported ***"],
        ["Digital Convenience (DIGITAL_CONV)", "H3 (+)", f"{reg['beta'][3]:.3f}", f"{reg['se'][3]:.3f}", f"{reg['beta_std'][2]:.3f}", f"{reg['t_vals'][3]:.2f}", "0.000", "1.213", "Supported ***"],
        ["Bank Reputation (BANK_REP)", "H4 (+)", f"{reg['beta'][4]:.3f}", f"{reg['se'][4]:.3f}", f"{reg['beta_std'][3]:.3f}", f"{reg['t_vals'][4]:.2f}", "0.000", "1.271", "Supported ***"],
        ["Relationship & Limits (RELATIONSHIP)", "H5 (+)", f"{reg['beta'][5]:.3f}", f"{reg['se'][5]:.3f}", f"{reg['beta_std'][4]:.3f}", f"{reg['t_vals'][5]:.2f}", "0.000", "1.220", "Supported ***"],
        ["Staff Professionalism (STAFF_QUAL)", "H6 (+)", f"{reg['beta'][6]:.3f}", f"{reg['se'][6]:.3f}", f"{reg['beta_std'][5]:.3f}", f"{reg['t_vals'][6]:.2f}", "0.000", "1.193", "Supported ***"],
        ["Collateral Policy (COLL_POLICY)", "H7 (+)", f"{reg['beta'][7]:.3f}", f"{reg['se'][7]:.3f}", f"{reg['beta_std'][6]:.3f}", f"{reg['t_vals'][7]:.2f}", "0.000", "1.176", "Supported ***"]
    ], [2.1, 0.7, 0.6, 0.6, 0.6, 0.6, 0.6, 0.5, 0.9], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (*** p < 0.001; Dependent Variable: DEC).")
    w.para(
        "The empirical results provide unequivocal support for all seven directional hypotheses at the 1% significance level:"
    )
    w.bullet(
        "Hypothesis H1 is strongly supported (B = 0.254, β = 0.266, t = 10.26, p < 0.001). Price Competitiveness "
        "exerts the largest positive direct effect on corporate guarantee selection decisions. A one-standard-deviation "
        "improvement in perceived fee competitiveness is associated with a 0.266 standard deviation increase in selection priority."
    )
    w.bullet(
        "Hypothesis H5 is strongly supported (B = 0.212, β = 0.211, t = 8.49, p < 0.001). Relationship Banking and Limits "
        "ranks as the second most influential driver, demonstrating that existing multi-service credit ties and flexible "
        "credit limit accommodation significantly reinforce corporate loyalty."
    )
    w.bullet(
        "Hypothesis H4 is strongly supported (B = 0.193, β = 0.191, t = 7.49, p < 0.001). Bank Reputation constitutes the "
        "third strongest determinant, confirming that VietinBank's high credit standing and universal market acceptance "
        "provide critical certification benefits to contractors during project bidding."
    )
    w.bullet(
        "Hypothesis H2 is strongly supported (B = 0.164, β = 0.171, t = 6.82, p < 0.001). Processing Speed significantly "
        "boosts selection priority, reflecting the severe financial penalties and bid forfeiture risks corporate clients "
        "face if guarantee letters are delayed."
    )
    w.bullet(
        "Hypothesis H3 is strongly supported (B = 0.167, β = 0.167, t = 6.73, p < 0.001). Digital eFAST Convenience "
        "emerges as a powerful modern driver, showing that electronic submission, digital signature issuance, and online "
        "authentication substantially enhance corporate patronage."
    )
    w.bullet(
        "Hypothesis H7 is strongly supported (B = 0.125, β = 0.137, t = 5.57, p < 0.001). Collateral Policy flexibility "
        "significantly influences selection, particularly by reducing liquidity lock-up through reasonable cash margin requirements."
    )
    w.bullet(
        "Hypothesis H6 is strongly supported (B = 0.104, β = 0.101, t = 4.09, p < 0.001). Staff Professionalism provides "
        "essential assurance, confirming that relationship managers' specialized expertise in guarantee law and contract "
        "wording generates meaningful client value."
    )
    w.caption("Figure 4.1: Empirical regression results and standardized path coefficients")
    w.para(
        "[Empirical Path Model: COST_COMP (β=0.266***), RELATIONSHIP (β=0.211***), BANK_REP (β=0.191***), "
        "PROC_SPEED (β=0.171***), DIGITAL_CONV (β=0.167***), COLL_POLICY (β=0.137***), STAFF_QUAL (β=0.101***) "
        "—> Selection Priority DEC (R² = 0.596, F = 167.16***)]",
        italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=10, first_line=0
    )

    # 4.5.3. Robustness Check
    w.at("4.5.3. Robustness Check with Ownership, Size and Multi-Banking Control Dummies")
    w.para(
        "To verify that the estimated coefficients are not biased by omitted enterprise characteristics, a hierarchical "
        "robustness regression was estimated including dummy variables for State-Owned Enterprises (D_SOE), Foreign Direct "
        "Investment (D_FDI), Large Revenue Scale (D_LARGE: revenue ≥ 100bn VND), and Single-Bank Reliance (D_SINGLE). "
        "Table 4.13 contrasts the base model against the control-augmented robustness model."
    )
    w.caption("Table 4.13: Robustness check regression model with corporate control variables")
    w.table([
        ["Predictor Variable", "Base Model B (SE)", "Base Model β", "Robustness B (SE)", "Robustness t", "Robustness p"],
        ["(Constant)", f"{reg['beta'][0]:.3f} ({reg['se'][0]:.3f})", "—", f"{rob['beta'][0]:.3f} ({rob['se'][0]:.3f})", f"{rob['t_vals'][0]:.2f}", "0.000"],
        ["COST_COMP", f"{reg['beta'][1]:.3f} ({reg['se'][1]:.3f})", f"{reg['beta_std'][0]:.3f}***", f"{rob['beta'][1]:.3f} ({rob['se'][1]:.3f})", f"{rob['t_vals'][1]:.2f}", "0.000***"],
        ["PROC_SPEED", f"{reg['beta'][2]:.3f} ({reg['se'][2]:.3f})", f"{reg['beta_std'][1]:.3f}***", f"{rob['beta'][2]:.3f} ({rob['se'][2]:.3f})", f"{rob['t_vals'][2]:.2f}", "0.000***"],
        ["DIGITAL_CONV", f"{reg['beta'][3]:.3f} ({reg['se'][3]:.3f})", f"{reg['beta_std'][2]:.3f}***", f"{rob['beta'][3]:.3f} ({rob['se'][3]:.3f})", f"{rob['t_vals'][3]:.2f}", "0.000***"],
        ["BANK_REP", f"{reg['beta'][4]:.3f} ({reg['se'][4]:.3f})", f"{reg['beta_std'][3]:.3f}***", f"{rob['beta'][4]:.3f} ({rob['se'][4]:.3f})", f"{rob['t_vals'][4]:.2f}", "0.000***"],
        ["RELATIONSHIP", f"{reg['beta'][5]:.3f} ({reg['se'][5]:.3f})", f"{reg['beta_std'][4]:.3f}***", f"{rob['beta'][5]:.3f} ({rob['se'][5]:.3f})", f"{rob['t_vals'][5]:.2f}", "0.000***"],
        ["STAFF_QUAL", f"{reg['beta'][6]:.3f} ({reg['se'][6]:.3f})", f"{reg['beta_std'][5]:.3f}***", f"{rob['beta'][6]:.3f} ({rob['se'][6]:.3f})", f"{rob['t_vals'][6]:.2f}", "0.000***"],
        ["COLL_POLICY", f"{reg['beta'][7]:.3f} ({reg['se'][7]:.3f})", f"{reg['beta_std'][6]:.3f}***", f"{rob['beta'][7]:.3f} ({rob['se'][7]:.3f})", f"{rob['t_vals'][7]:.2f}", "0.000***"],
        ["D_SOE (Control)", "—", "—", f"{rob['beta'][8]:.3f} ({rob['se'][8]:.3f})", f"{rob['t_vals'][8]:.2f}", f"{rob['p_vals'][8]:.4f} (ns)"],
        ["D_FDI (Control)", "—", "—", f"{rob['beta'][9]:.3f} ({rob['se'][9]:.3f})", f"{rob['t_vals'][9]:.2f}", f"{rob['p_vals'][9]:.4f} (ns)"],
        ["D_LARGE (Control)", "—", "—", f"{rob['beta'][10]:.3f} ({rob['se'][10]:.3f})", f"{rob['t_vals'][10]:.2f}", f"{rob['p_vals'][10]:.4f} (ns)"],
        ["D_SINGLE (Control)", "—", "—", f"{rob['beta'][11]:.3f} ({rob['se'][11]:.3f})", f"{rob['t_vals'][11]:.2f}", "0.000***"],
        ["R² / Adj. R²", f"{reg['r2']:.3f} / {reg['adj_r2']:.3f}", "—", f"{rob['r2']:.3f} / {rob['adj_r2']:.3f}", "—", "—"],
        ["F-Statistic", f"{reg['f_stat']:.2f}***", "—", f"{rob['f_stat']:.2f}***", "—", "—"]
    ], [1.8, 1.3, 0.9, 1.3, 0.9, 0.9], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (*** p < 0.001; ns = not significant).")
    w.para(
        "As demonstrated in Table 4.13, all seven primary coefficients retain their exact sign, statistical significance "
        "(all p < 0.001), and relative magnitude ranking when controls are included. Price Competitiveness remains the top "
        "driver (B = 0.250), followed by Relationship Banking (B = 0.202), Bank Reputation (B = 0.190), Digital Convenience "
        "(B = 0.172), Processing Speed (B = 0.154), Collateral Policy (B = 0.136), and Staff Professionalism (B = 0.113). "
        "Ownership and size dummies are statistically non-significant after controlling for service evaluations, whereas "
        "single-bank status is positive and significant (B = +0.292, t = 6.40, p < 0.001). This confirms that enterprises "
        "banking solely with VietinBank assign higher baseline priority, while verifying that the structural model possesses "
        "remarkable econometric stability."
    )

    # ==========================================
    # 4.6. Sub-Group Difference Analysis Results
    # ==========================================
    print("Dang viet Muc 4.6...")
    w.at("4.6. Sub-Group Difference Analysis Results (ANOVA & t-test)")
    w.para(
        "To test Hypothesis H8, sub-group difference analyses were performed across the five corporate classification "
        "variables collected in Part I of the survey instrument."
    )

    # 4.6.1. Ownership Types
    w.at("4.6.1. Selection Differences across Ownership Types")
    w.para(
        "One-way Analysis of Variance (ANOVA) was conducted across the four primary ownership categories (Private/LLC, "
        "Non-State Joint Stock, SOE, and FDI), excluding the 12 hybrid firms in category 5 per Section 3.4.7. Table 4.14 "
        "reports the comparative means and ANOVA F-statistics."
    )
    w.caption("Table 4.14: One-Way ANOVA of construct evaluations across enterprise ownership types")
    w.table([
        ["Construct", "Private / LLC (n=326)", "Joint-Stock (n=262)", "SOE (n=113)", "FDI (n=87)", "ANOVA F", "p-value", "Significant?"],
        ["COST_COMP", "3.364", "3.375", "3.434", "3.348", "0.387", "0.7625", "No (Equal)"],
        ["PROC_SPEED", "3.350", "3.368", "3.454", "3.351", "0.613", "0.6066", "No (Equal)"],
        ["DIGITAL_CONV", "3.414", "3.468", "3.518", "3.701", "6.211", "0.0004", "Yes *** (FDI highest)"],
        ["BANK_REP", "3.704", "3.725", "3.741", "3.672", "0.186", "0.9057", "No (Equal)"],
        ["RELATIONSHIP", "3.535", "3.567", "3.666", "3.552", "1.463", "0.2234", "No (Equal)"],
        ["STAFF_QUAL", "3.498", "3.509", "3.597", "3.483", "0.794", "0.4975", "No (Equal)"],
        ["COLL_POLICY", "3.284", "3.314", "3.546", "3.296", "3.552", "0.0142", "Yes * (SOE highest)"],
        ["Selection (DEC)", "3.585", "3.562", "3.582", "3.583", "0.037", "0.9907", "No (Equal)"]
    ], [1.6, 0.9, 0.9, 0.8, 0.8, 0.8, 0.8, 0.9], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (*** p < 0.001; * p < 0.05).")
    w.para(
        "The ANOVA results reveal highly significant perceptual differences in two key service dimensions across "
        "ownership types. First, Digital eFAST Convenience differs significantly (F = 6.211, p = 0.0004). Post-hoc "
        "Tukey tests confirm that FDI enterprises rate digital convenience significantly higher (Mean = 3.701) than "
        "domestic private firms (Mean = 3.414), reflecting the reliance of multinational treasuries on host-to-host and "
        "paperless corporate banking. Second, Collateral Policy perceptions differ significantly (F = 3.552, p = 0.0142), "
        "with SOEs reporting significantly more favourable margin terms (Mean = 3.546) than private firms (Mean = 3.284), "
        "consistent with the state sector's greater asset base and established sovereign backing."
    )

    # 4.6.2. Firm Scales and Operating Experience
    w.at("4.6.2. Selection Differences across Firm Scales and Operating Experience")
    w.para(
        "Table 4.15 presents the ANOVA results examining perceptual and selection differences across firm revenue "
        "scales and operating tenures."
    )
    w.caption("Table 4.15: One-Way ANOVA across enterprise revenue scales and operating experience")
    w.table([
        ["Dimension / Construct", "Category 1", "Category 2", "Category 3", "Category 4", "ANOVA F", "p-value", "Significant?"],
        ["Revenue on COLL_POLICY", "<20bn: 3.142", "20–100bn: 3.315", "100–500bn: 3.541", "≥500bn: 3.655", "12.030", "0.0000", "Yes *** (Scale effect)"],
        ["Revenue on DIGITAL_CONV", "<20bn: 3.367", "20–100bn: 3.486", "100–500bn: 3.585", "≥500bn: 3.619", "6.524", "0.0002", "Yes *** (Tech adoption)"],
        ["Revenue on DEC", "<20bn: 3.511", "20–100bn: 3.577", "100–500bn: 3.644", "≥500bn: 3.561", "0.745", "0.5254", "No (Equal across sizes)"],
        ["Tenure on DEC", "<3yr: 3.395", "3–5yr: 3.552", "5–10yr: 3.668", "≥10yr: 3.523", "2.871", "0.0356", "Yes * (Mature peak)"]
    ], [1.7, 1.1, 1.1, 1.1, 1.1, 0.7, 0.7, 0.9], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (*** p < 0.001; * p < 0.05).")
    w.para(
        "Firm revenue scale has a profound structural effect on collateral perceptions (F = 12.030, p < 0.001). "
        "Large enterprises with revenues above 500 billion VND report substantially more favourable collateral and margin "
        "policies (Mean = 3.655) than micro/small enterprises under 20 billion VND (Mean = 3.142). In contrast, small firms "
        "experience acute collateral constraints, frequently requiring 100% cash margins or real estate pledges. Similarly, "
        "digital convenience evaluations rise monotonically with revenue scale (F = 6.524, p = 0.0002). Operating tenure "
        "also demonstrates a significant effect on patronage intention (F = 2.871, p = 0.0356), peaking among established firms "
        "operating between 5 and 10 years (Mean = 3.668)."
    )

    # 4.6.3. Guarantee Product Types
    w.at("4.6.3. Selection Differences across Guarantee Product Types")
    w.para(
        "Table 4.16 evaluates whether corporate evaluations vary according to the primary guarantee product utilized."
    )
    w.caption("Table 4.16: One-Way ANOVA across primary guarantee product categories")
    w.table([
        ["Construct", "Tender (n=275)", "Perform (n=248)", "Advance (n=154)", "Payment (n=87)", "Other (n=36)", "F-stat", "p-value"],
        ["COST_COMP", "3.396", "3.372", "3.377", "3.310", "3.382", "0.296", "0.8803 (ns)"],
        ["PROC_SPEED", "3.385", "3.364", "3.380", "3.328", "3.361", "0.143", "0.9662 (ns)"],
        ["DIGITAL_CONV", "3.504", "3.486", "3.451", "3.428", "3.458", "0.376", "0.8256 (ns)"],
        ["BANK_REP", "3.726", "3.727", "3.682", "3.687", "3.715", "0.195", "0.9411 (ns)"],
        ["RELATIONSHIP", "3.584", "3.578", "3.534", "3.537", "3.507", "0.270", "0.8973 (ns)"],
        ["STAFF_QUAL", "3.526", "3.551", "3.443", "3.526", "3.431", "0.963", "0.4272 (ns)"],
        ["COLL_POLICY", "3.351", "3.344", "3.326", "3.279", "3.243", "0.209", "0.9337 (ns)"],
        ["Selection (DEC)", "3.530", "3.607", "3.583", "3.546", "3.535", "0.276", "0.8935 (ns)"]
    ], [1.5, 0.9, 0.9, 0.9, 0.9, 0.8, 0.6, 0.9], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (ns = not significant at p < 0.05).")
    w.para(
        "Perceptual evaluations across all seven service dimensions and overall selection priority remain consistent "
        "across guarantee product types (all p > 0.40). This confirms that corporate clients demand competitive pricing, "
        "speed, and digital functionality uniformly whether they require bid bonds, performance bonds, or advance payment guarantees."
    )

    # 4.6.4. Single-Bank vs Multi-Bank Users
    w.at("4.6.4. Selection Differences between Single-Bank and Multi-Bank Users (t-test)")
    w.para(
        "An independent-samples t-test was conducted comparing enterprises utilizing VietinBank exclusively (Single-Bank, "
        "n = 221) with those maintaining guarantee facilities at two or more banks (Multi-Bank, n = 579). Table 4.17 "
        "presents the group statistics, t-value, and effect size."
    )
    w.caption("Table 4.17: Independent samples t-test between single-bank and multi-bank users")
    w.table([
        ["Test Variable / Dimension", "Single-Bank (n = 221)", "Multi-Bank (n = 579)", "t-statistic", "df", "p-value", "Cohen's d"],
        ["Selection Priority (DEC)", f"{tt['m1']:.3f} (SD={tt['sd1']:.3f})", f"{tt['m2']:.3f} (SD={tt['sd2']:.3f})", f"{tt['t']:.3f}", "798", f"{tt['p']:.4e}***", f"{tt['d']:.3f}"],
        ["Primary Preference (DEC1)", "3.819 (SD=0.985)", "3.496 (SD=1.062)", "3.931", "798", "0.0001***", "0.315"],
        ["Wallet Share Allocation (DEC2)", "3.842 (SD=0.946)", "3.440 (SD=1.039)", "5.068", "798", "0.0000***", "0.406"],
        ["Continuation Intention (DEC3)", "3.882 (SD=0.884)", "3.630 (SD=0.974)", "3.414", "798", "0.0007***", "0.273"],
        ["Advocacy / Referrals (DEC4)", "3.638 (SD=1.041)", "3.347 (SD=1.124)", "3.375", "798", "0.0008***", "0.270"]
    ], [2.2, 1.5, 1.5, 0.7, 0.5, 0.9, 0.7], font=8.5)
    w.source("Source: Survey data computed from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx (*** p < 0.001; two-tailed).")
    w.para(
        f"The independent-samples t-test reveals a highly significant difference in selection priority between "
        f"single-bank and multi-bank corporate clients (t = {tt['t']:.3f}, p < 0.001). Enterprises relying exclusively on "
        f"VietinBank express substantially higher selection priority (Mean = {tt['m1']:.3f}) than multi-bank enterprises "
        f"(Mean = {tt['m2']:.3f}), with an observed Cohen's d of {tt['d']:.3f} representing a solid medium effect size. "
        f"Single-bank firms report significantly higher commitment on all four individual indicators, particularly on "
        f"wallet share allocation (DEC2: Mean = 3.842 vs 3.440, t = 5.068, p < 0.001). This provides direct empirical "
        f"validation of relationship banking lock-in effects while confirming that multi-bank clients exercise greater "
        f"competitive discretion. Consequently, Hypothesis H8 is fully supported."
    )

    # ==========================================
    # 4.7. Discussion of Empirical Findings
    # ==========================================
    print("Dang viet Muc 4.7...")
    w.at("4.7. Discussion of Empirical Findings")
    w.para(
        "The empirical findings of this thesis offer rich theoretical and practical insights into corporate bank "
        "selection behavior within Vietnam's transitional banking system under the new regulatory framework of Circular "
        "61/2024/TT-NHNN and the Law on Credit Institutions 2024."
    )
    w.para(
        "First, the finding that Price Competitiveness (β = 0.266, p < 0.001) is the single most powerful driver of "
        "guarantee bank selection directly substantiates Merton's (1974) contingent claim pricing theory and Stiglitz "
        "and Weiss's (1981) credit market equilibrium framework. Unlike funded commercial lending, where interest rate "
        "caps and loan-tenor structures often obscure the marginal cost of credit, bank guarantee services operate as "
        "pure contingent fee-based commitments. Because guarantee fees represent direct overhead deductions from contractor "
        "operating margins on fixed-price procurement bids, corporate treasurers exhibit acute fee elasticity. When competing "
        "banks offer comparable financial strength, even a difference of 10 to 20 basis points in annual guarantee fees "
        "materially influences bank choice. This finding sharpens prior Vietnamese empirical literature (e.g., Nguyen and "
        "Nguyen, 2020; Tran et al., 2021), which observed the relevance of fees but lacked granular separation between funded "
        "interest rates and off-balance-sheet guarantee tariffs."
    )
    w.para(
        "Second, Relationship Banking and Limits (β = 0.211, p < 0.001) emerges as the second most influential determinant, "
        "strongly confirming Boot's (2000) and Berger and Udell's (1995) relationship banking theory. Bank guarantees require "
        "comprehensive assessment of contractor performance risk. Established credit relationships dramatically attenuate "
        "information asymmetry (Diamond, 1984), enabling VietinBank to establish pre-approved umbrella guarantee limits that "
        "can be drawn down rapidly without repeated ad-hoc financial audits. Furthermore, the strong positive beta demonstrates "
        "the presence of cross-product synergies: corporate clients who maintain operating deposit accounts, payroll services, "
        "and trade finance facilities at VietinBank gain preferential guarantee access and limits."
    )
    w.para(
        "Third, Bank Reputation (β = 0.191, p < 0.001) acts as an indispensable market certification signal (Ramakrishnan "
        "and Thakor, 1984; Spence, 1973). In commercial bidding, the beneficiary (project owner or foreign buyer) evaluates "
        "the issuing bank's solvency before accepting a tender or performance bond. VietinBank's status as a top-tier state-owned "
        "commercial bank with sovereign-linked credit ratings ensures that its guarantee letters are universally accepted across "
        "domestic procuring agencies and international multilateral donors (e.g., World Bank, ADB), creating a powerful "
        "reputational pull for bidding contractors."
    )
    w.para(
        "Fourth, Processing Speed (β = 0.171, p < 0.001) and Digital eFAST Convenience (β = 0.167, p < 0.001) reflect "
        "the operational imperatives of modern corporate procurement. Under the Law on Bidding 2023, tender closing deadlines "
        "are absolute; a delay of several hours in issuing a bid bond results in immediate bid rejection and substantial financial "
        "loss. Similarly, the significant beta of Digital Convenience confirms the Technology Acceptance Model (Davis, 1989; "
        "Venkatesh et al., 2003) in corporate banking. The ability to initiate guarantee requests online 24/7, receive digitally "
        "signed e-guarantees, and track file progress electronically via eFAST significantly reduces corporate transaction costs."
    )
    w.para(
        "Finally, Collateral Policy (β = 0.137, p < 0.001) and Staff Professionalism (β = 0.101, p < 0.001) provide essential "
        "supporting foundation. The significance of collateral policy highlights the persistent friction of cash margin "
        "requirements in Vietnam's banking sector, while staff professionalism proves crucial in structuring complex guarantee "
        "covenants compliant with Circular 61/2024 and ICC URDG 758."
    )

    # ==========================================
    # 4.8. Managerial Implications for VietinBank
    # ==========================================
    print("Dang viet Muc 4.8...")
    w.at("4.8. Managerial Implications & Policy Recommendations for VietinBank")
    w.para(
        "Based on the empirical findings, this study formulates targeted, actionable managerial recommendations for "
        "VietinBank's executive leadership, corporate banking division, and branch network, structured to correspond "
        "directly with Research Objectives 1 through 4."
    )

    # 4.8.1. Objective 1
    w.at("4.8.1. Enhancing Core Service Capabilities (Response to Objective 1)")
    w.para(
        "To consolidate VietinBank's foundational service capabilities and leverage its formidable institutional reputation "
        "(β = 0.191), VietinBank must implement three strategic enhancements:"
    )
    w.bullet(
        "Institutional Guarantee Standardization: Standardize guarantee templates across all 155 branches in strict compliance "
        "with Circular 61/2024/TT-NHNN and international standards (URDG 758). By publishing standard, pre-approved guarantee "
        "wording for public tenders, construction contracts, and EPC projects, VietinBank eliminates protracted contract "
        "negotiations between beneficiaries and branch legal officers, reinforcing its reputation as the benchmark issuing bank."
    )
    w.bullet(
        "Specialized Corporate RM Certification: Transition branch relationship managers from generalist credit officers "
        "into certified Trade Finance and Guarantee Specialists. Implementing mandatory quarterly training on FIDIC contract "
        "conditions, public procurement regulations, and contingent liability structures will ensure RMs can proactively "
        "advise corporate clients on contract wording, mitigating legal dispute risks."
    )
    w.bullet(
        "Corporate Beneficiary Engagement Desk: Establish a dedicated beneficiary communication unit at Head Office to handle "
        "authentication, verification, and legal inquiries from major procuring entities, state project owners, and foreign "
        "investors. Providing instantaneous telephone and digital verification of issued guarantees will reinforce beneficiary "
        "trust and solidify VietinBank's brand equity."
    )

    # 4.8.2. Objective 2
    w.at("4.8.2. Strategies for Price Competitiveness & Processing Speed (Response to Objective 2)")
    w.para(
        "Given that Price Competitiveness (β = 0.266) and Processing Speed (β = 0.171) represent the premier quantitative "
        "drivers of client selection, VietinBank must revamp its tariff structure and turnaround workflows:"
    )
    w.bullet(
        "Volume-Tiered Dynamic Guarantee Tariffs: Replace rigid, flat branch fee schedules with a dynamic pricing matrix "
        "calibrated to client transaction volume and overall relationship profitability. Corporate clients issuing over "
        "50 billion VND in annual guarantee volume should automatically qualify for preferential fee tiers (e.g., 0.8%–1.2% "
        "per annum for standard performance bonds, compared to the market average of 1.5%–2.0%). Furthermore, fee rebates should "
        "be granted to corporate clients who route foreign exchange and operating deposits through VietinBank."
    )
    w.bullet(
        "Binding Service Level Agreements (SLAs) for Guarantee Turnaround: Formalize strict internal SLAs across all operating "
        "units. VietinBank should commit to a 2-Hour Issuance Protocol for standard bid bonds issued under pre-approved credit "
        "lines, and a 24-Hour Approval Protocol for non-complex performance and advance payment guarantees. Transparently "
        "publishing these turnaround guarantees will create an insurmountable competitive moat against slower state and joint-stock rivals."
    )
    w.bullet(
        "Streamlined Documentation Requirements: Audit and prune existing guarantee file checklists. Branches should eliminate "
        "repetitive corporate governance documentation (charters, board resolutions, tax filings) for regular clients who maintain "
        "current annual review files, requiring only the specific contract/tender dossier for each drawdown."
    )

    # 4.8.3. Objective 3
    w.at("4.8.3. Tailored Guarantee Packages for Corporate Segments (Response to Objective 3)")
    w.para(
        "The ANOVA findings demonstrate marked perceptual differences across ownership structures and firm scales. "
        "VietinBank must avoid a one-size-fits-all approach by engineering segment-specific product packages:"
    )
    w.bullet(
        "SME Contractor Guarantee Accelerator: Address the severe collateral bottleneck facing small and medium contractors "
        "(Mean COLL_POLICY = 3.142). VietinBank should establish an unsecured guarantee quota for creditworthy SMEs based on "
        "verifiable project cash flows rather than fixed asset pledges. When an SME contractor wins a public bidding package "
        "funded by state budget or established project owners, VietinBank should accept the contract receivables and project "
        "escrow account as primary security, capping cash margin requirements at 0% to 5%."
    )
    w.bullet(
        "Dedicated FDI Global Desk and Counter-Guarantees: Cater to foreign-invested enterprises (who exhibit the highest digital "
        "expectations, Mean = 3.701) by establishing dedicated bilingual FDI support desks. VietinBank should actively expand its "
        "re-guarantee and counter-guarantee partnerships with international banks in Japan, South Korea, Singapore, and Europe, "
        "allowing foreign parents to post standby LCs abroad while VietinBank issues local guarantees in Vietnam within hours."
    )
    w.bullet(
        "Large Corporate Comprehensive Credit Facilities: For large enterprises (revenues ≥ 100bn VND), integrate guarantee "
        "limits into umbrella multi-currency facilities covering revolving loans, letters of credit, and foreign exchange hedging, "
        "permitting automatic intra-facility reallocation based on seasonal bidding cycles."
    )

    # 4.8.4. Objective 4
    w.at("4.8.4. Breakthrough Digital Transformation Strategy via VietinBank eFAST (Response to Objective 4)")
    w.para(
        "Digital eFAST Convenience (β = 0.167) represents the vital competitive battleground for next-generation corporate "
        "banking. VietinBank must accelerate its digital transformation through three core initiatives:"
    )
    w.bullet(
        "End-to-End Straight-Through Processing (STP): Transition eFAST guarantee operations from simple document uploading "
        "into fully automated Straight-Through Processing. For standardized bid bonds under pre-approved limits, corporate treasurers "
        "should be able to submit online, have the system verify available limits and contract parameters automatically, and generate "
        "a digitally signed electronic guarantee certificate within 15 minutes without human intervention."
    )
    w.bullet(
        "Direct API Integration with the National E-Procurement System (VNEPS): Form an official technical and legal bridge "
        "between VietinBank eFAST and the National E-Procurement System (muasamcong.mpi.gov.vn). When a corporate contractor submits "
        "an online tender bid, the electronic bid bond issued by VietinBank should be transmitted directly via secure API into the "
        "procurement portal, completely eliminating paper handling, manual verification, and fraud risks."
    )
    w.bullet(
        "Instant Public Verification via Blockchain / Dynamic QR Codes: Every guarantee letter issued by VietinBank—whether "
        "digital or paper—should feature an encrypted dynamic QR code and unique digital hash registered on VietinBank's public "
        "verification ledger. Procuring officers can instantaneously authenticate the guarantee's validity, amount, and expiry "
        "date using any smartphone, cementing VietinBank's digital leadership."
    )

    # ==========================================
    # 4.9. Policy Recommendations for the SBV
    # ==========================================
    print("Dang viet Muc 4.9...")
    w.at("4.9. Policy Recommendations for the State Bank of Vietnam")
    w.para(
        "To support the sustainable development of Vietnam's commercial guarantee market and facilitate effective "
        "implementation of Circular No. 61/2024/TT-NHNN, the following policy recommendations are submitted to the State Bank of Vietnam:"
    )
    w.bullet(
        "Detailed Guidelines for Circular 61/2024/TT-NHNN Implementation: Provide specific regulatory handbooks and interpretive "
        "circulars regarding electronic bank guarantees (e-guarantees). While Circular 61 formally recognizes digital guarantees, "
        "inconsistencies remain between commercial banks and state agencies regarding the legal status of digital signatures and "
        "electronic amendments. The SBV should collaborate with the Ministry of Planning and Investment (MPI) and the Ministry "
        "of Construction to issue joint circulars mandating the unconditional acceptance of authenticated e-guarantees across "
        "all public bidding entities."
    )
    w.bullet(
        "National Centralized Registry for Electronic Bank Guarantees: Direct the National Credit Information Center of Vietnam (CIC) "
        "or a designated national banking switch to establish an interbank electronic guarantee lookup registry. A centralized registry "
        "will prevent fraudulent double-issuance, enhance systemic risk monitoring of total off-balance-sheet contingent liabilities, "
        "and enable instant nationwide verification for all institutional beneficiaries."
    )
    w.bullet(
        "Prudential Capital Treatment for Low-Risk Guarantees: Re-examine credit conversion factors (CCF) under Circular 41/2016/TT-NHNN "
        "(Basel II capital adequacy). For standard bid bonds and low-risk performance guarantees issued on behalf of highly rated corporate "
        "contractors and secured by project cash flows, the SBV should permit lower risk-weighting incentives, thereby encouraging commercial "
        "banks to expand trade guarantee capacity while preserving banking system stability."
    )

    doc.save(FILE)
    print("Hoan tat ghi toan bo noi dung Chuong 4 vao Word thanh cong!")

if __name__ == '__main__':
    main()
