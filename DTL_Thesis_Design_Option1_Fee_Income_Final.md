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
Bank guarantee is a fundamental credit instrument closely tied to almost all high-value commercial transactions of corporate enterprises. In the construction and engineering sector, Article 18 of Decree No. 37/2015/ND-CP (amended and supplemented by Decree No. 35/2023/ND-CP) mandates that contractors must submit an advance payment guarantee for contracts with advance payments exceeding 1 billion VND; Bidding Law No. 22/2023/QH15 strictly regulates tender security and performance security obligations. In international trade, bank guarantees issued under the Uniform Rules for Demand Guarantees (URDG 758) substitute cash margin deposits, liberating corporate working capital. Demand for bank guarantee products is therefore mandatory rather than a discretionary corporate choice.

For commercial banks, guarantee issuance represents a signature credit operation—commonly referred to as an off-balance sheet commitment. At the time of issuing a letter of guarantee, the bank does not disburse cash upfront but provides an irrevocable commitment to pay on behalf of the principal if the principal defaults. This contingent liability is monitored off-balance sheet and is converted into an on-balance sheet forced loan only when the bank actually executes the payout.

This economic mechanism creates three strategic financial benefits for commercial banks: (i) banks generate steady issuing and maintenance fee income throughout the guarantee term without disbursing capital, building high-margin non-interest fee income that is resilient to interest rate volatility; (ii) guarantee limits anchor corporate clients into long-term credit relationships, opening cross-selling avenues for deposit, payment, trade finance, and FX products; and (iii) default risk materialises with low probability while off-balance sheet commitments consume regulatory capital at lower Credit Conversion Factors (CCF) compared to funded loans of equivalent value.

In commercial banking practice, Guarantee Fee Income (`THU_PHI_BL`) and Fee Yield Ratio serve as the primary financial metrics measuring off-balance sheet profitability and corporate customer relationship value. However, fee income generation is constrained by intense interbank competition, client price sensitivity, credit risk ratings, collateral margin requirements, and digital transformation. Identifying which financial and operational drivers significantly influence guarantee fee revenue and yield provides direct empirical value for VietinBank's asset-liability management and fee pricing strategies.

Vietnam's regulatory landscape for guarantee operations has recently undergone fundamental transformation. Circular No. 61/2024/TT-NHNN (effective April 1, 2025, replacing Circular No. 11/2022/TT-NHNN) establishes an updated legal framework for electronic guarantees (e-guarantees). Concurrently, commercial bank digital transformation through VietinBank eFAST and global correspondent networks is reshaping fee extraction capabilities. Existing literature focuses predominantly on retail banking or overall bank selection, leaving transaction-level guarantee fee income under-researched. VietinBank—a Big4 state-owned commercial bank with its extensive corporate banking database—provides an ideal empirical context. Hence, the author selects Option 1: *"Factors Affecting Guarantee Fee Income and Financial Efficiency of Corporate Customers at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)"*.

### 1.2. General Objectives
The general objective of this research is to identify and measure the impact magnitude of transaction-level and firm-level financial determinants influencing guarantee fee income and fee yield efficiency at VietinBank, and on that basis, to propose actionable managerial implications helping VietinBank optimize fee pricing, enhance off-balance sheet profitability, and expand market share in this strategic product segment.

### 1.3. Specific Objectives
The following sub-questions will be investigated to answer the general objective:
1. What transaction-level and firm-level financial factors significantly affect corporate guarantee fee income at VietinBank?
2. What is the direction and magnitude of each financial factor's impact on total fee revenue and fee yield ratio?
3. How do credit rating tiers, collateral margin ratios, and digital eFAST adoption moderate the fee income relationship?
4. What risk-adjusted fee pricing policies and managerial recommendations should VietinBank implement to maximize off-balance sheet fee returns?

### 1.4. Thesis Structure
This thesis comprises four main chapters:
* **Chapter 1: Introduction** – Presents research rationales, general and specific objectives, sub-questions, and thesis structure.
* **Chapter 2: Literature Review and Theoretical Framework** – Synthesizes legal, economic, and banking foundations of guarantee pricing; presents theoretical frameworks (Financial Intermediation, Credit Pricing, Relationship Banking); reviews empirical studies; and identifies key research gaps.
* **Chapter 3: Research Methodology and Empirical Design** – Specifies the multiple linear regression model with log-transformed fee income; defines financial and firm variables; describes VietinBank's transaction dataset (n = 800); and details Python econometric procedures.
* **Chapter 4: Empirical Results, Discussion and Policy Recommendations** – Reports descriptive statistics, correlation matrix, Pooled OLS and Robust regression results; discusses financial findings; and details managerial fee pricing recommendations for VietinBank.

---

## II. Literature Review

### 2.1. Theoretical Foundations and Legal Context
Under Vietnamese law, Circular No. 61/2024/TT-NHNN (effective April 1, 2025, replacing Circular No. 11/2022/TT-NHNN and Circular No. 49/2024/TT-NHNN) defines bank guarantee as a form of credit extension wherein the guarantor (credit institution) commits to the beneficiary to perform financial obligations on behalf of the principal should the principal fail to perform. Article 335 of the 2015 Civil Code establishes general guarantee principles. Economically, bank guarantees are non-funded credit instruments that reallocate default risk, enable trust in commercial dealings, and conserve corporate cash flow. International practice under ICC URDG 758 highlights the principle of independence and documentary character.

Guarantee operations involve a tri-partite relationship comprising the Guarantor (VietinBank), Principal (Corporate Client), and Beneficiary (Project Owner/Buyer). Key product lines include Tender Guarantee (TG), Performance Guarantee (PG), Advance Payment Guarantee (APG), Payment Guarantee (BG), Maintenance Guarantee, and Counter Guarantee.

Commercial bank guarantee fee pricing and off-balance sheet financial efficiency are grounded in four theoretical models:
1. **Financial Intermediation Theory (Diamond, 1984; Ramakrishnan & Thakor, 1984):** Explains how commercial banks act as delegated monitors, producing creditworthiness signals that justify charging guarantee issuing and maintenance fees.
2. **Credit Risk Pricing & Contingent Claim Theory (Stiglitz & Weiss, 1981; Merton, 1974):** Models guarantee fee schedules as a function of default risk probabilities, collateral margin protection, and underlying contract exposure.
3. **Relationship Banking Theory (Boot, 2000; Berger & Udell, 1995):** Posits that long-term credit history and multi-product ties (payroll, payment accounts, trade finance) reduce information asymmetry, enabling risk-adjusted fee discounting.
4. **Customer Lifetime Value (CLV) & Non-Interest Revenue Theory (DeYoung & Roland, 2001):** Operationalizes off-balance sheet fee income (`THU_PHI_BL`) and fee yield efficiency as primary components of non-interest bank profitability.

### 2.2. Empirical Review and Research Gaps
Empirical literature spans corporate bank selection (Turnbull & Gibbs, 1989; Narteh, 2013; Kaur et al., 2021), Vietnamese banking service quality (Phan Thi Hang Nga et al., 2024; Ho Dinh Phi et al., 2023; Nguyen et al., 2024), and trade finance fee pricing (Al-Sabbagh & Al-Khathlan, 2018; Carletti et al., 2023; Le Van Dung, 2021). Four critical research gaps emerge: (i) separation between survey-based perception studies and quantitative transaction-level fee revenue models; (ii) lack of guarantee-focused financial yield studies in Vietnam; (iii) omission of digital eFAST channel adoption and legal risk advisory impacts on fee realization; and (iv) absence of large-scale system-wide empirical transaction studies at VietinBank.

---

## III. Empirical Analysis

### 3.1. Methodology and Proposed Model
The study employs a multiple linear regression model with primary dependent variable—Log of Guarantee Fee Income (`ln_FEE`)—and 8 independent financial and firm variables. To ensure econometric robustness, a secondary model evaluating Fee Yield Ratio (`FEE_YIELD = FEE_INCOME / GUARANTEE_VOLUME * 100%`) is specified. The primary regression equation is specified as:

$$\ln(\text{FEE\_INCOME}) = \beta_0 + \beta_1\ln(\text{LIMIT}) + \beta_2\text{MARGIN\_RATIO} + \beta_3\text{TENOR} + \beta_4\text{FIRM\_SIZE} + \beta_5\text{FIRM\_AGE} + \beta_6\text{CREDIT\_RATING} + \beta_7\text{RELATIONSHIP} + \beta_8\text{DIGITAL} + \varepsilon \quad (1)$$

#### Table 1: Variable Definitions and Measurement Specifications (Option 1)
| Code | Variable Name | Measurement Content / Source | Type | Sign |
| :---: | :--- | :--- | :---: | :---: |
| **ln_FEE** | Guarantee Fee Income | Log of total annual guarantee fee revenue in VND (Core Banking MIS) | Dependent (Y1) | N/A |
| **FEE_YIELD** | Fee Yield Ratio | Ratio of fee income to total guarantee volume (%) = (FEE/VOLUME)*100% | Dependent (Y2) | N/A |
| **ln_LIMIT** | Guarantee Credit Limit | Log of approved guarantee limit amount in VND | Independent (X1) | **+** |
| **MARGIN_RATIO** | Collateral Margin Ratio | Ratio of cash margin / collateral value to guarantee limit (%) | Independent (X2) | **–** |
| **TENOR** | Average Guarantee Tenor | Average duration of guarantee commitments in months | Independent (X3) | **+** |
| **FIRM_SIZE** | Corporate Revenue Scale | Log of corporate client annual revenue (VND) | Independent (X4) | **+** |
| **FIRM_AGE** | Corporate Operating Age | Number of operating years since formal business registration | Independent (X5) | **+** |
| **CREDIT_RATING** | Internal Credit Rating | VietinBank internal rating tier (Numerical score: 1=AAA to 10=C) | Independent (X6) | **–** |
| **RELATIONSHIP** | Banking Relationship Tenure | Years of active credit and deposit relationship with VietinBank | Independent (X7) | **–** |
| **DIGITAL** | eFAST Digital Adoption | Dummy variable = 1 if eFAST online guarantee user; 0 otherwise | Independent (X8) | **+** |

*Source: Author's design based on VietinBank transactional database and literature.*

### 3.2. Data Collection and Econometric Pre-Tests
This study inherits an existing system-wide corporate transactional dataset extracted from VietinBank's Core Banking and Management Information System (MIS) across 155 VietinBank branches nationwide. The sampling frame covers active corporate guarantee accounts, yielding a clean dataset of $n = 800$ corporate observations. This sample size far exceeds minimum econometric thresholds ($n \ge 170$ for multi-variable EFA; $n \ge 130$ for multiple linear regression). Econometric procedures include anonymization, outlier winsorization at 1st and 99th percentiles, log transformation of skewed monetary variables, descriptive statistics, Pearson correlation matrix, Pooled OLS estimation, Breusch-Pagan heteroskedasticity test, VIF multicollinearity test ($VIF < 5$), and Huber-White robust standard errors. All data processing and regression estimations are executed in Python.

### 3.3. Anticipated Findings and Interpretation
Based on theoretical credit pricing frameworks and banking context, the following results are anticipated:
* Guarantee Limit (`ln_LIMIT`) and Tenor (`TENOR`) are projected to exert positive, highly significant impacts ($\beta_1 > 0, \beta_3 > 0$), confirming volume and duration as key revenue drivers.
* Higher Credit Rating (better rating score) and longer Relationship Tenure (`RELATIONSHIP`) are expected to yield negative coefficients ($\beta_6 < 0, \beta_7 < 0$), reflecting preferential fee discount policies granted to high-quality, long-term corporate clients. Conversely, eFAST Digital Adoption (`DIGITAL`) will demonstrate a positive impact ($\beta_8 > 0$), validating digital channel efficiency in driving overall fee realization.

---

## IV. Conclusions
This thesis design establishes a rigorous quantitative framework evaluating guarantee fee income and financial efficiency at VietinBank. By integrating transaction-level financial metrics with credit risk and relationship banking determinants, the study fills major empirical gaps in off-balance sheet banking literature. Managerial recommendations will guide VietinBank executive leadership in designing risk-adjusted fee pricing schedules, optimizing eFAST digital fee incentives, implementing relationship-based limit bundling, and balancing collateral margin requirements to maximize off-balance sheet non-interest returns.

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
* Zeithaml, V.A. (1996) 'The Behavioral Consequences of Service Quality', *Journal of Marketing*, 60(2), pp. 31–46.
* Zelie, E.M. (2023) 'Factors determining bank selection by micro- and small-sized enterprises: evidence from Ethiopia', *International Journal of Bank Marketing*, 41(5), pp. 1120–1142.
