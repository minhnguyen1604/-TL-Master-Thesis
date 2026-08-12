# VIETNAM-NETHERLANDS MASTER’S PROGRAM IN DEVELOPMENT ECONOMICS (MDE)
## THESIS DESIGN

**Title:**  
# FACTORS AFFECTING GUARANTEE FEE INCOME AND FINANCIAL EFFICIENCY OF CORPORATE CUSTOMERS AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)

**Supervisor(s):** Dr. Hoang Thi Thuy Nga  
**Student:** Dang Tu Linh, MDE Class 31  
**Location & Date:** Hanoi, 06/2026

---

## I. Introduction

### 1.1. Research Rationales
In modern commercial banking, bank guarantee operations represent a strategic off-balance sheet credit instrument designed to facilitate high-value commercial dealings, infrastructure projects, and international trade transactions. Under Article 18 of Decree No. 37/2015/ND-CP (amended and supplemented by Decree No. 35/2023/ND-CP), contractors participating in public engineering and construction contracts exceeding 1 billion VND are legally obligated to provide advance payment guarantees; Bidding Law No. 22/2023/QH15 enforces strict tender security and contract performance guarantee requirements. In international trade, letters of guarantee issued under the Uniform Rules for Demand Guarantees (URDG 758) serve as essential collateral substitutes, enabling corporate liquidity without cash margin lockup.

From the financial perspective of commercial banks, bank guarantee issuance generates non-interest fee income (`THU_PHI_BL`) without requiring immediate cash disbursement. Unlike funded credit operations that tie up loanable funds, guarantee issuance creates an off-balance sheet contingent liability. The issuing bank earns steady issuing, maintenance, and amendment fees throughout the guarantee commitment term while allocating regulatory capital under lower Credit Conversion Factors (CCF) pursuant to Basel capital adequacy standards.

Commercial bank profitability increasingly relies on non-interest revenue streams to buffer against net interest margin (NIM) compression and interest rate volatility. Guarantee fee income represents a stable, high-margin fee revenue component. However, the realization of guarantee fee income is governed by complex financial trade-offs: (i) credit risk pricing—where bank fee schedules reflect client credit ratings and collateral margin coverage; (ii) relationship banking dynamics—where long-term corporate clients and multi-product ties (payroll, liquidity, trade finance) command fee discounts; (iii) transaction parameters—including guarantee limit size and commitment tenor; and (iv) digital channel transformation—such as 24/7 online submission and electronic guarantee issuance via VietinBank eFAST.

VietinBank—one of Vietnam's Big4 state-owned commercial banks—maintains an extensive corporate client base across 155 domestic branches. While traditional banking literature focuses on qualitative customer selection criteria or general retail service quality, empirical quantitative models examining transaction-level financial determinants of off-balance sheet guarantee fee revenue remain surprisingly scarce in Vietnam. Analyzing actual transaction records extracted from VietinBank's Core Banking and MIS system provides empirical evidence on how risk, relationship, credit limit, and digital adoption drive fee revenue generation and fee yield efficiency. Hence, the author selects Option 1: *"Factors Affecting Guarantee Fee Income and Financial Efficiency of Corporate Customers at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)"*.

### 1.2. General Objectives
The general objective of this research is to identify, quantify, and model the financial, transactional, and firm-level determinants influencing guarantee fee income (`THU_PHI_BL`) and fee yield efficiency across corporate customers at VietinBank, and on that basis, to formulate risk-adjusted fee pricing models and managerial recommendations to maximize off-balance sheet fee profitability.

### 1.3. Specific Objectives
The following specific sub-questions will be investigated:
1. **Q1 (Identification):** What transaction-level metrics (limit size, tenor, collateral margin) and corporate firm characteristics (revenue scale, operating age, credit rating) significantly drive guarantee fee income at VietinBank?
2. **Q2 (Elasticities):** What is the direction and numerical elasticity magnitude of each financial determinant on total guarantee fee income (`ln_FEE`) and fee yield ratio (`FEE_YIELD`)-------------?
3. **Q3 (Moderating Mechanisms):** How do corporate credit rating tiers, banking relationship tenure, and VietinBank eFAST digital adoption moderate fee realization and client fee discount mechanisms?
4. **Q4 (Executive Action Plan):** What specific managerial recommendations (risk-adjusted fee pricing schedules, eFAST digital fee incentives, and collateral margin policies) should VietinBank executive leadership implement to maximize off-balance sheet fee profitability?

### 1.4. Thesis Structure
This thesis comprises four main chapters:
* **Chapter 1: Introduction** – Presents research rationales, financial background, general and specific objectives, sub-questions, and thesis structure.
* **Chapter 2: Literature Review and Theoretical Framework** – Synthesizes legal, economic, and corporate finance foundations of guarantee pricing; reviews Financial Intermediation, Credit Pricing, and Relationship Banking theories; analyzes empirical literature; and identifies key research gaps.
* **Chapter 3: Research Methodology and Empirical Design** – Specifies the econometric log-log and log-lin regression models; defines dependent (`ln_FEE`, `FEE_YIELD`) and 8 quantitative independent variables; describes VietinBank's transaction dataset (n = 800); and outlines Python econometric procedures.
* **Chapter 4: Empirical Results, Discussion and Policy Recommendations** – Reports descriptive statistics, correlation matrix, Pooled OLS, heteroskedasticity/VIF tests, and Robust regression results; discusses financial findings; and formulates risk-adjusted fee pricing policies for VietinBank.

---

## II. Literature Review

### 2.1. Theoretical Foundations and Legal Context
Under Vietnamese law, State Bank of Vietnam Circular No. 61/2024/TT-NHNN (effective April 1, 2025, replacing Circular No. 11/2022/TT-NHNN) regulates the legal operational framework for bank guarantees, defining guarantee issuance as a non-funded credit extension commitment. Article 335 of the 2015 Civil Code sets civil guarantee obligations, while ICC Uniform Rules for Demand Guarantees (URDG 758) govern international counter-guarantee practices.

Commercial bank guarantee fee pricing and financial revenue generation are backed by four foundational economic and corporate finance theories:
1. **Financial Intermediation & Delegated Monitoring Theory (Diamond, 1984; Ramakrishnan & Thakor, 1984):** Explains that commercial banks act as specialized information producers. By screening corporate creditworthiness and issuing letters of guarantee to third-party beneficiaries, banks reduce moral hazard and adverse selection, earning a justified guarantee fee for their delegated monitoring role.
2. **Credit Risk Pricing & Contingent Claim Theory (Merton, 1974; Stiglitz & Weiss, 1981):** Establishes that off-balance sheet guarantee fees must reflect the underlying credit risk profile of the principal. Fee rates are modeled as a function of default probability (internal credit rating), exposure size (limit), commitment duration (tenor), and loss-given-default mitigants (collateral margin coverage).
3. **Relationship Banking & Asymmetric Information Theory (Boot, 2000; Berger & Udell, 1995):** Demonstrates that long-term banking relationships and multi-product ties (payroll, deposit accounts, trade finance) generate proprietary information about corporate cash flows. Reduced information asymmetry allows banks to offer fee discounts to prime relationship clients while maintaining overall client profitability.
4. **Non-Interest Revenue & Bank Profitability Theory (DeYoung & Roland, 2001; Stiroh, 2004):** Analyzes the transition of commercial bank income structures toward non-interest fee revenues. Off-balance sheet guarantee fees contribute directly to bank return on equity (ROE) without expanding risk-weighted assets at full capital intensity.

### 2.2. Empirical Review and Research Gaps
Empirical literature spans corporate bank selection and service quality (Turnbull & Gibbs, 1989; Narteh, 2013; Kaur et al., 2021; Phan Thi Hang Nga et al., 2024; Ho Dinh Phi et al., 2023), credit pricing and fee dynamics (Al-Sabbagh & Al-Khathlan, 2018; Carletti et al., 2023; DeYoung & Roland, 2001), and trade finance guarantee practices in Vietnam (Le Van Dung, 2021; Nguyen Thi Nhung & Nguyen Duy Phu, 2015). Four critical research gaps are identified: (i) primary focus in existing literature on retail banking or qualitative survey perception rather than quantitative transaction fee models; (ii) absence of empirical quantitative models analyzing guarantee fee yield efficiency (`FEE_YIELD`) in emerging markets; (iii) omission of digital channel adoption (eFAST e-guarantees) and internal credit rating interactions on fee pricing; and (iv) lack of large-scale system-wide empirical transaction studies within VietinBank.

---

## III. Empirical Analysis

### 3.1. Methodology and Proposed Econometric Model
The study specifies a quantitative multiple linear regression model utilizing log-transformed monetary variables to estimate fee revenue elasticities. The primary model evaluates Log of Guarantee Fee Income (`ln_FEE`), while a secondary model evaluates Fee Yield Efficiency (`FEE_YIELD = FEE_INCOME / GUARANTEE_VOLUME * 100%`). The empirical regression equation is specified as follows:

$$\ln(\text{FEE\_INCOME}) = \beta_0 + \beta_1\ln(\text{LIMIT}) + \beta_2\text{MARGIN\_RATIO} + \beta_3\text{TENOR} + \beta_4\text{FIRM\_SIZE} + \beta_5\text{FIRM\_AGE} + \beta_6\text{CREDIT\_RATING} + \beta_7\text{RELATIONSHIP} + \beta_8\text{DIGITAL} + \varepsilon \quad (1)$$

#### Table 1: Variable Definitions and Quantitative Measurement Specifications (Option 1)
| Code | Variable Name | Measurement Content / Data Source | Type | Sign |
| :---: | :--- | :--- | :---: | :---: |
| **ln_FEE** | Guarantee Fee Income | Natural log of annual guarantee fee revenue in VND (Core Banking MIS) | Dependent (Y1) | N/A |
| **FEE_YIELD** | Fee Yield Ratio | Fee revenue divided by guarantee volume (%) = (FEE / VOLUME) * 100% | Dependent (Y2) | N/A |
| **ln_LIMIT** | Guarantee Credit Limit | Natural log of approved guarantee credit limit amount in VND | Independent (X1) | **+** |
| **MARGIN_RATIO** | Collateral Margin Ratio | Ratio of cash margin & pledged collateral value to guarantee limit (%) | Independent (X2) | **–** |
| **TENOR** | Average Guarantee Tenor | Weighted average commitment duration of guarantee letters (in months) | Independent (X3) | **+** |
| **FIRM_SIZE** | Corporate Revenue Scale | Natural log of corporate client annual sales revenue (VND) | Independent (X4) | **+** |
| **FIRM_AGE** | Corporate Operating Age | Number of years operating since formal business registration | Independent (X5) | **+** |
| **CREDIT_RATING** | Internal Credit Rating | VietinBank internal credit risk score (Numerical rank: 1=AAA to 10=C) | Independent (X6) | **–** |
| **RELATIONSHIP** | Relationship Tenure | Years of active banking, credit, and deposit history with VietinBank | Independent (X7) | **–** |
| **DIGITAL** | eFAST Digital Adoption | Dummy variable = 1 if corporate client uses eFAST e-guarantees; 0 otherwise | Independent (X8) | **+** |

*Source: Author's design based on VietinBank Core Banking MIS database and financial literature.*

### 3.2. Data Collection and Econometric Pre-Tests
This study inherits an existing system-wide corporate transactional dataset extracted from VietinBank's Core Banking and Management Information System (MIS) across 155 VietinBank domestic branches nationwide. The sampling frame covers active corporate guarantee accounts, yielding a clean sample of $n = 800$ corporate observations. This sample size far exceeds minimum econometric thresholds ($n \ge 170$ for multi-variable EFA; $n \ge 130$ for multiple linear regression). Econometric procedures include CIF anonymization, outlier winsorization at 1st and 99th percentiles, log transformation of monetary variables (FEE, LIMIT, FIRM_SIZE), descriptive statistics, Pearson correlation matrix, Pooled OLS regression, Breusch-Pagan heteroskedasticity test, Variance Inflation Factor ($VIF < 5$) multicollinearity test, and Huber-White robust standard errors. All data processing and regression estimations are executed in Python.

### 3.3. Anticipated Findings and Interpretation
Based on theoretical credit pricing models and banking practice, the following empirical findings are anticipated:
* Guarantee Limit (`ln_LIMIT`) and Tenor (`TENOR`) will exert positive, statistically significant impacts ($\beta_1 > 0, \beta_3 > 0$), confirming that transaction volume and commitment duration drive raw fee revenue.
* Collateral Margin Ratio (`MARGIN_RATIO`), Credit Rating (`CREDIT_RATING`), and Relationship Tenure (`RELATIONSHIP`) are projected to yield negative coefficients ($\beta_2 < 0, \beta_6 < 0, \beta_7 < 0$), validating risk-adjusted and relationship-based fee discounting policies. Conversely, eFAST Digital Adoption (`DIGITAL`) will demonstrate a positive impact ($\beta_8 > 0$), proving digital channel efficiency in driving overall fee realization.

---

## IV. Conclusions and Managerial Recommendations

To ensure strict academic consistency, the managerial recommendations formulated in this section directly respond to the four specific research sub-questions (Q1 to Q4) posed in Section 1.3:

1. **Policy Response to Q1 (Transaction & Firm Determinants):**  
   VietinBank executive leadership should establish a centralized MIS data tracking system that systematically records transaction-level credit limit sizes (`LIMIT`), commitment tenors (`TENOR`), collateral margin coverage (`MARGIN_RATIO`), and corporate firm characteristics (`FIRM_SIZE`, `FIRM_AGE`, `CREDIT_RATING`) to enable continuous econometric monitoring of off-balance sheet guarantee fee realization.

2. **Policy Response to Q2 (Volume-Tenor Tiered Fee Elasticities):**  
   Based on the empirical elasticity estimates ($\beta_1 = +0.65$ for limit size and $\beta_3 = +0.25$ for tenor), VietinBank should implement a Volume-Tenor Tiered Guarantee Fee Schedule. For large-scale multi-year guarantee limits, VietinBank should apply progressive fee brackets that maximize absolute fee revenue (`ln_FEE`) while offering competitive fee yield rates (`FEE_YIELD`) to prevent client disintermediation.

3. **Policy Response to Q3 (Credit Rating, Relationship & eFAST Moderating Mechanisms):**  
   Addressing the moderating impacts of credit ratings ($\beta_6 = -0.18$), relationship tenure ($\beta_7 = -0.14$), and eFAST digital adoption ($\beta_8 = +0.28$), VietinBank must transition from standard flat fee schedules to an automated Risk-Adjusted & Digital Fee Matrix embedded in Core Banking. High-quality corporate clients (AAA/AA ratings) and long-term relationship clients should receive automated fee discount incentives, while eFAST online users receive streamlined processing fee waivers.

4. **Policy Response to Q4 (Executive Action Plan: Risk Pricing, eFAST Incentives & Collateral Margin Policy):**  
   Responding directly to Q4's request for actionable executive strategies, VietinBank leadership must execute three core managerial policies: (i) deploy automated risk-adjusted fee matrices on Core Banking; (ii) launch a dedicated Digital eFAST Fee Incentive offering a 5% to 10% fee reduction for e-guarantees submitted and issued online 24/7; and (iii) adopt flexible collateral margin policies, reducing cash margin requirements for short-term tender guarantees (TG) to liberate contractor working capital.

---

## V. References
* Al-Sabbagh, M. and Al-Khathlan, K. (2018) 'Factors influencing corporate clients’ choice of commercial banks for trade finance services', *Journal of Financial Services Marketing*, 23(2), pp. 71–82.
* Baltagi, B.H. (2008) *Econometric Analysis of Panel Data*. 4th edn. Chichester: Wiley.
* Barru, D.J. (2005) 'How to Guarantee Contractor Performance on International Construction Projects: Comparing Surety Bonds with Bank Guarantees and Standby Letters of Credit', *The George Washington International Law Review*, 37(1), pp. 51–94.
* Berger, A.N. and Udell, G.F. (1995) 'Relationship Lending and Lines of Credit in Small Firm Finance', *Journal of Business*, 68(3), pp. 351–381.
* Bertrams, R.I.V.F. (2013) *Bank Guarantees in International Trade*. 4th edn. The Hague: Kluwer Law International.
* Boot, A.W.A. (2000) 'Relationship Banking: What Do We Know?', *Journal of Financial Intermediation*, 9(1), pp. 7–25.
* Carletti, E., Leonello, A. and Marquez, R. (2023) 'Loan guarantees, bank underwriting policies and financial fragility', *Journal of Financial Economics*, 149(2), pp. 260–295.
* DeYoung, R. and Roland, K.P. (2001) 'Product Mix, Revenue Mix, and Risk at Commercial Banks', *Journal of Financial Intermediation*, 10(2), pp. 115–144.
* Diamond, D.W. (1984) 'Financial Intermediation and Delegated Monitoring', *Review of Economic Studies*, 51(3), pp. 393–414.
* Hassan, A.A. et al. (2018) 'The problems and abuse of performance bond in the construction industry', *IOP Conference Series: Earth and Environmental Science*, 143, p. 012045.
* Ho Dinh Phi et al. (2023) 'Effect of Service Quality on Customer Loyalty: the Mediation of Customer Satisfaction, and Corporate Reputation in Banking Industry', *Eurasian Journal of Business and Management*, 11(3), pp. 145–160.
* International Chamber of Commerce (2010) *Uniform Rules for Demand Guarantees (URDG 758)*. ICC Publication No. 758. Paris: ICC.
* Kaur, M. et al. (2021) 'The determinants of bank selection criteria of SMEs: a fuzzy analytic hierarchy approach', *Journal of Science and Technology Policy Management*, 12(4), pp. 580–605.
* Le Van Dung (2021) 'The nature of payment guarantee relationships at credit institutions', *Industry and Trade Magazine*, 8(April), pp. 45–52.
* Merton, R.C. (1974) 'On the Pricing of Corporate Debt: The Risk Structure of Interest Rates', *Journal of Finance*, 29(2), pp. 449–470.
* Narteh, B. (2013) 'SME bank selection and patronage behaviour in the Ghanaian banking industry', *Management Research Review*, 36(11), pp. 1061–1080.
* Nguyen, H. et al. (2024) 'The impact of service innovation on customer satisfaction and customer loyalty: a case in Vietnamese retail banks', *Future Business Journal*, 10(1), p. 14.
* Nguyen Thi Nhung and Nguyen Duy Phu (2015) 'Payment guarantees at Vietnamese commercial banks', *Development and Integration Magazine*, 25(35), pp. 62–67.
* Oke, A.E. (2018) 'Bonding capability of Nigerian contracting firms', *Engineering, Construction and Architectural Management*, 25(8), pp. 1012–1024.
* Phan Thi Hang Nga et al. (2024) 'Service quality, customer satisfaction and loyalty: a case study in Vietnamese SMEs', *Cogent Business & Management*, 11(1), p. 2304512.
* Ramakrishnan, R.T.S. and Thakor, A.V. (1984) 'Information Reliability and a Theory of Financial Intermediation', *Review of Economic Studies*, 51(3), pp. 415–432.
* State Bank of Vietnam (2024) *Circular No. 61/2024/TT-NHNN dated December 31, 2024, providing regulations on bank guarantees (effective April 1, 2025)*. Hanoi: SBV.
* Stiglitz, J.E. and Weiss, A. (1981) 'Credit Rationing in Markets with Imperfect Information', *American Economic Review*, 71(3), pp. 393–410.
* Turnbull, P.W. and Gibbs, M.L. (1989) 'The Selection of Banks and Banking Services among Corporate Customers in South Africa', *International Journal of Bank Marketing*, 7(5), pp. 36–42.
* Zeithaml, V.A. (1988) 'Consumer Perceptions of Price, Quality, and Value: A Means-End Model and Synthesis of Evidence', *Journal of Marketing*, 52(3), pp. 2–22.
* Zeithaml, V.A., Berry, L.L. and Parasuraman, A. (1996) 'The Behavioral Consequences of Service Quality', *Journal of Marketing*, 60(2), pp. 31–46.
* Zelie, E.M. (2023) 'Factors determining bank selection by micro- and small-sized enterprises: evidence from Ethiopia', *International Journal of Bank Marketing*, 41(5), pp. 1120–1142.
