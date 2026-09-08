# NATIONAL ECONOMICS UNIVERSITY
## VIETNAM-NETHERLANDS MASTER’S PROGRAM IN DEVELOPMENT ECONOMICS (MDE)

---

# MASTER THESIS

# FACTORS AFFECTING CORPORATE CUSTOMERS’ DECISION TO CHOOSE BANK GUARANTEE SERVICES AT VIETNAM JOINT STOCK COMMERCIAL BANK FOR INDUSTRY AND TRADE (VIETINBANK)

**Author / Student:** Dang Tu Linh  
**Student ID / Class:** MDE Class 31  
**Academic Supervisor:** Dr. Hoang Thi Thuy Nga  
**Location & Year:** Hanoi, June 2026  

---

## STATEMENT OF AUTHORSHIP (LỜI CAM ĐOAN)
> 👉 *[Nội dung hướng dẫn]: Trình bày lời cam đoan của học viên về tính trung thực, độc lập và đạo đức nghiên cứu khoa học. Xác nhận toàn bộ dữ liệu khảo sát 800 doanh nghiệp và kết quả ước lượng hồi quy OLS là số liệu thu thập thực tế, không sao chép bất hợp pháp từ bất kỳ công trình nào.*

I hereby declare that this Master’s thesis entitled "Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)" is my own independent research work conducted under the academic supervision of Dr. Hoang Thi Thuy Nga. The survey dataset ($n = 800$) and econometric findings presented in this thesis are original, transparent, and have not been submitted for any other degree or qualification at any academic institution.

Hanoi, June 2026  
*Student Author*  
**Dang Tu Linh**

---

## ACKNOWLEDGEMENTS (LỜI CẢM ƠN)
> 👉 *[Nội dung hướng dẫn]: Dành lời cảm ơn trân trọng tới: (1) Ban Giám hiệu NEU, Viện MDE và Đại học Erasmus Rotterdam / ISS; (2) TS. Hoàng Thị Thúy Nga - GVHD đã tận tình định hướng và phản biện; (3) Ban Lãnh đạo, cán bộ Khối KHDN Ngân hàng VietinBank trên 155 chi nhánh đã hỗ trợ phát phiếu khảo sát; (4) Đại diện 800 doanh nghiệp tham gia trả lời phiếu.*

*[Nội dung lời cảm ơn sẽ được hoàn thiện tại đây...]*

---

## ABSTRACT
> 👉 *[Nội dung hướng dẫn]: Tóm tắt luận văn bằng Tiếng Anh (khoảng 300 - 500 từ) gồm 5 phần: (1) Background & Objective; (2) Methodology (OLS, EFA, Cronbach's Alpha, ANOVA, n = 800); (3) Key Findings (7 nhân tố tác động dương +, COST_COMP và BANK_REP mạnh nhất, eFAST có ý nghĩa, khác biệt nhóm SOE/SME/FDI); (4) Policy Implications (4 nhóm giải pháp cho VietinBank); (5) Keywords.*

*[English Abstract content will be inserted here...]*

**Keywords:** Bank Guarantee Services, Corporate Selection Decision, Price Competitiveness, eFAST Digital Adoption, OLS Regression, VietinBank.

---

## TÓM TẮT LUẬN VĂN (BẢN TIẾNG VIỆT)
> 👉 *[Nội dung hướng dẫn]: Bản tóm tắt tiếng Việt tương ứng với Abstract, trình bày rõ bối cảnh, mô hình 7 biến độc lập, cỡ mẫu 800 doanh nghiệp và các kiến nghị giải pháp then chốt cho VietinBank.*

*[Nội dung Tóm tắt tiếng Việt sẽ được hoàn thiện tại đây...]*

**Từ khóa:** Bảo lãnh ngân hàng, Khách hàng doanh nghiệp, Quyết định lựa chọn, Tính cạnh tranh biểu phí, Chuyển đổi số eFAST, Hồi quy OLS, VietinBank.

---

## TABLE OF CONTENTS (MỤC LỤC)
> 👉 *[Nội dung hướng dẫn]: Trang Mục lục tự động trong MS Word (References -> Table of Contents).*

*[Table of Contents will be generated automatically in MS Word...]*

---

## LIST OF ABBREVIATIONS (DANH MỤC TỪ VIẾT TẮT)
| Abbreviation | Full English Term | Vietnamese Equivalent |
| :--- | :--- | :--- |
| **APG** | Advance Payment Guarantee | Bảo lãnh hoàn trả tiền tạm ứng |
| **EFA** | Exploratory Factor Analysis | Phân tích nhân tố khám phá |
| **eFAST** | VietinBank Corporate Digital Banking | Nền tảng ngân hàng số doanh nghiệp VietinBank |
| **FDI** | Foreign Direct Investment | Doanh nghiệp có vốn đầu tư trực tiếp nước ngoài |
| **OLS** | Ordinary Least Squares | Phương pháp bình phương bé nhất thông thường |
| **PG** | Performance Guarantee | Bảo lãnh thực hiện hợp đồng |
| **TG** | Tender Guarantee / Bid Bond | Bảo lãnh dự thầu |
| **URDG 758** | Uniform Rules for Demand Guarantees 758 | Quy tắc thống nhất về bảo lãnh theo yêu cầu (ICC) |
| **VIF** | Variance Inflation Factor | Hệ số phóng đại phương sai (Đo đa cộng tuyến) |

---

## LIST OF TABLES (DANH MỤC BẢNG BIỂU)
> 👉 *[Nội dung hướng dẫn]: Danh mục thống kê tự động tất cả các bảng trong luận văn kèm số trang.*

---

## LIST OF FIGURES (DANH MỤC HÌNH VẼ & SƠ ĐỒ)
> 👉 *[Nội dung hướng dẫn]: Danh mục thống kê tự động tất cả các sơ đồ, hình vẽ và biểu đồ trong luận văn.*

---

# CHAPTER 1: INTRODUCTION

### 1.1. Research Rationales & Background
> 👉 *[Nội dung hướng dẫn]: Trình bày bối cảnh kinh tế vĩ mô và tính cấp thiết: (1) Bản chất bảo lãnh ngân hàng là công cụ bảo đảm nghĩa vụ hợp đồng bắt buộc trong xây dựng, đấu thầu theo Luật Đấu thầu số 22/2023/QH15 và Nghị định số 35/2023/NĐ-CP; (2) Thông tư 61/2024/TT-NHNN tạo hành lang pháp lý mới cho bảo lãnh điện tử (e-guarantees); (3) Tầm quan trọng của thu nhập phí bảo lãnh phi tín dụng đối với ngân hàng thương mại; (4) Áp lực cạnh tranh gay gắt khiến VietinBank cần hiểu rõ các yếu tố quyết định lựa chọn của doanh nghiệp.*

### 1.2. Research Problem & Industry Context at VietinBank
> 👉 *[Nội dung hướng dẫn]: Nhận diện vấn đề nghiên cứu thực tế tại VietinBank: Khách hàng doanh nghiệp ngày càng khắt khe về thời gian phát hành, tính cạnh tranh của biểu phí và tính linh hoạt của tài sản thế chấp. Làm rõ bối cảnh mạng lưới 155 chi nhánh của VietinBank và khoảng trống trong việc nghiên cứu hành vi lựa chọn dịch vụ bảo lãnh B2B.*

### 1.3. Research Objectives
#### 1.3.1. General Objective
> 👉 *[Nội dung hướng dẫn]: Cập nhật nguyên văn tiếng Anh của cô giáo hướng dẫn:*

"The general objective of this study is to identify and assess the factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, and to propose managerial recommendations for improving VietinBank’s attractiveness and competitiveness in the corporate bank guarantee market."

#### 1.3.2. Specific Objectives
> 👉 *[Nội dung hướng dẫn]: Cập nhật nguyên văn 4 mục tiêu cụ thể cô giáo đã phê duyệt (bao gồm phân tích khác biệt nhóm ở Objective 3):*

1. **Objective 1:** Identify the key factors associated with corporate customers’ decision to choose VietinBank for bank guarantee services, based on relevant theories, previous empirical studies, and the characteristics of bank guarantee services.
2. **Objective 2:** Assess the direction and relative importance of these factors in explaining corporate customers’ selection decisions.
3. **Objective 3:** Examine whether corporate customers’ selection decisions differ across major firm characteristics, such as ownership type, firm size, operating experience, and types of bank guarantees used.
4. **Objective 4:** Propose managerial recommendations for VietinBank to improve its bank guarantee products and services and strengthen its ability to attract and retain corporate customers.

### 1.4. Research Questions
> 👉 *[Nội dung hướng dẫn]: Thiết lập 4 câu hỏi nghiên cứu tương ứng trực tiếp với 4 mục tiêu cụ thể:*

* **Question 1:** What key factors significantly influence corporate customers’ decision to choose VietinBank for bank guarantee services?
* **Question 2:** What is the direction and relative importance of each factor in explaining corporate selection decisions?
* **Question 3:** Do corporate selection decisions significantly differ across corporate ownership types, firm revenue sizes, operating tenure, and guarantee product lines?
* **Question 4:** What actionable managerial recommendations should VietinBank implement to enhance its competitive standing, product attractiveness, and client retention?

### 1.5. Scope and Boundaries of the Study
> 👉 *[Nội dung hướng dẫn]: Xác định phạm vi nghiên cứu: Khách hàng doanh nghiệp, 155 chi nhánh VietinBank toàn quốc, dữ liệu khảo sát 2025 - 2026, tập trung vào 7 nhân tố dịch vụ và quyết định lựa chọn.*

### 1.6. Significance & Contributions of the Study
> 👉 *[Nội dung hướng dẫn]: Nêu rõ 2 đóng góp lớn: Đóng góp học thuật (bổ sung khoảng trống bảo lãnh B2B và chuyển đổi số eFAST) và Đóng góp thực tiễn (cung cấp bằng chứng định lượng giúp Ban Lãnh đạo VietinBank tối ưu biểu phí, quy trình và phân khúc KH).*

### 1.7. Structure of the Thesis
> 👉 *[Nội dung hướng dẫn]: Giới thiệu tóm tắt kết cấu 4 chương chính của luận văn.*

---

# CHAPTER 2: LITERATURE REVIEW AND THEORETICAL FRAMEWORK

### 2.1. Overview of Bank Guarantee Services in Commercial Banking
#### 2.1.1. Nature and Economic Functions of Bank Guarantees
> 👉 *[Nội dung hướng dẫn]: Phân tích bản chất kinh tế: Cam kết tín dụng ngoại bảng (Off-balance sheet commitment), tạo uy tín trung gian giúp bảo toàn vốn lưu động cho nhà thầu.*

#### 2.1.2. Main Types of Corporate Bank Guarantees
> 👉 *[Nội dung hướng dẫn]: Trình bày đặc điểm của 4 sản phẩm chính: Bảo lãnh dự thầu (TG), Bảo lãnh thực hiện hợp đồng (PG), Bảo lãnh tạm ứng (APG), Bảo lãnh thanh toán (BG).*

#### 2.1.3. Legal and Regulatory Framework
> 👉 *[Nội dung hướng dẫn]: Phân tích khung pháp lý: Thông tư 61/2024/TT-NHNN (bảo lãnh điện tử), Luật Đấu thầu 22/2023, Nghị định 35/2023 và URDG 758.*

### 2.2. Theoretical Foundations
#### 2.2.1. Financial Intermediation & Delegated Monitoring Theory
> 👉 *[Nội dung hướng dẫn]: Lý thuyết Trung gian Tài chính (Diamond, 1984; Ramakrishnan & Thakor, 1984): Ngân hàng đóng vai trò sản xuất thông tin và bảo chứng tín nhiệm cho doanh nghiệp.*

#### 2.2.2. Credit Risk Pricing & Contingent Claim Theory
> 👉 *[Nội dung hướng dẫn]: Lý thuyết Định giá Rủi ro (Merton, 1974; Stiglitz & Weiss, 1981): Mức phí và tỷ lệ ký quỹ phản ánh xác suất rủi ro và giá trị tài sản đảm bảo.*

#### 2.2.3. Service Quality Theory & SERVQUAL Model
> 👉 *[Nội dung hướng dẫn]: Mô hình SERVQUAL (Parasuraman et al., 1988): Đo lường chất lượng dịch vụ. Tách riêng Uy tín thương hiệu (BANK_REP) khỏi Độ tin cậy (Reliability) theo đúng nhắc nhở của GVHD.*

#### 2.2.4. Relationship Banking Theory
> 👉 *[Nội dung hướng dẫn]: Lý thuyết Ngân hàng Quan hệ (Boot, 2000; Berger & Udell, 1995): Quan hệ lâu năm giúp giảm bất cân xứng thông tin, hưởng hạn mức lớn và ưu đãi phí.*

#### 2.2.5. Technology Acceptance Model (TAM) & Digital Banking
> 👉 *[Nội dung hướng dẫn]: Mô hình Chấp nhận Công nghệ TAM (Davis, 1989; Venkatesh et al., 2003): Cơ sở bảo chứng cho biến Chuyển đổi số eFAST (DIGITAL_CONV) dựa trên tính hữu ích và dễ sử dụng.*

### 2.3. Empirical Literature on Corporate Bank Selection
#### 2.3.1. International Empirical Studies
> 👉 *[Nội dung hướng dẫn]: Tổng quan các công trình quốc tế tiêu biểu: Turnbull & Gibbs (1989), Narteh (2013), Al-Sabbagh & Al-Khathlan (2018), Kaur et al. (2021), Zelie (2023), Carletti et al. (2023).*

#### 2.3.2. Empirical Studies in the Vietnamese Banking Context
> 👉 *[Nội dung hướng dẫn]: Tổng quan các nghiên cứu tại Việt Nam: Phan Thị Hằng Nga et al. (2024), Hồ Đình Phi et al. (2023), Nguyễn et al. (2024), Lê Văn Dũng (2021).*

### 2.4. Research Gaps
> 👉 *[Nội dung hướng dẫn]: Chỉ ra 4 khoảng trống nghiên cứu: Thiếu nghiên cứu sâu về bảo lãnh ngoại bảng; Chưa có mô hình tích hợp biến eFAST; Thiếu so sánh khác biệt nhóm doanh nghiệp; Chưa có nghiên cứu quy mô n = 800 tại VietinBank.*

### 2.5. Conceptual Framework and Research Hypotheses
#### 2.5.1. Conceptual Research Framework
> 👉 *[Nội dung hướng dẫn]: Mô tả sơ đồ khung phân tích gồm 7 biến độc lập tác động lên biến phụ thuộc DEC.*

#### 2.5.2. Hypothesis Development
> 👉 *[Nội dung hướng dẫn]: Lập luận cơ sở lý thuyết và phát biểu 7 giả thuyết nghiên cứu (tất cả đều mang dấu dương +):*

* **Hypothesis H1:** Price Competitiveness (`COST_COMP`) has a positive impact on corporate customers' decision to choose VietinBank.
* **Hypothesis H2:** Processing Speed (`PROC_SPEED`) has a positive impact on corporate customers' decision to choose VietinBank.
* **Hypothesis H3:** Digital eFAST Convenience (`DIGITAL_CONV`) has a positive impact on corporate customers' decision to choose VietinBank.
* **Hypothesis H4:** Bank Reputation (`BANK_REP`) has a positive impact on corporate customers' decision to choose VietinBank.
* **Hypothesis H5:** Relationship Banking & Limits (`RELATIONSHIP`) has a positive impact on corporate customers' decision to choose VietinBank.
* **Hypothesis H6:** Staff Professionalism (`STAFF_QUAL`) has a positive impact on corporate customers' decision to choose VietinBank.
* **Hypothesis H7:** Collateral & Margin Flexibility (`COLL_POLICY`) has a positive impact on corporate customers' decision to choose VietinBank.

---

# CHAPTER 3: RESEARCH METHODOLOGY AND EMPIRICAL DESIGN

### 3.1. Overall Research Design & Analytical Process
> 👉 *[Nội dung hướng dẫn]: Mô tả sơ đồ quy trình 7 bước tuần tự: Thống kê mô tả -> Cronbach's Alpha -> EFA -> Factor Scores -> Tương quan Pearson -> Kiểm định VIF -> Hồi quy OLS và Kiểm định ANOVA/t-test theo gợi ý của GVHD.*

### 3.2. Questionnaire Design & Measurement Scales
#### 3.2.1. Operationalization of Variables
> 👉 *[Nội dung hướng dẫn]: Bảng định nghĩa thang đo Likert 1-5 của 7 biến độc lập và 1 biến phụ thuộc:*

| Code | Variable Name | Measurement Content (34 Items) | Type | Sign |
| :---: | :--- | :--- | :---: | :---: |
| **DEC** | Selection Decision | Preference & priority choice of VietinBank over competitors (4 items) | Dependent (Y) | N/A |
| **COST_COMP** | Price Competitiveness | Fee reasonableness, competitive pricing & discount incentives (4 items) | Independent (X1) | **+** |
| **PROC_SPEED** | Processing Speed | Turnaround time, prompt issuance & simplified paperwork (4 items) | Independent (X2) | **+** |
| **DIGITAL_CONV** | Digital eFAST Convenience | Online submission, 24/7 e-guarantees & digital status tracking (4 items) | Independent (X3) | **+** |
| **BANK_REP** | Bank Reputation | Big4 brand prestige, financial strength & 100% acceptance (5 items) | Independent (X4) | **+** |
| **RELATIONSHIP** | Relationship & Limits | Credit limit flexibility, relationship tenure & VIP care (5 items) | Independent (X5) | **+** |
| **STAFF_QUAL** | Staff Professionalism | RM competence, legal advisory on Bidding Law & TT61 (4 items) | Independent (X6) | **+** |
| **COLL_POLICY** | Collateral Flexibility | Flexible cash margin ratio & diverse pledged collateral (4 items) | Independent (X7) | **+** |

#### 3.2.2. Mapping Scales with the Official 34-Item Survey Questionnaire
> 👉 *[Nội dung hướng dẫn]: Chứng minh nguyên tắc "Có bột mới gột nên hồ": Ánh xạ chi tiết từng mã câu hỏi CP1-CP4, TD1-TD4, CS1-CS4, UT1-UT5, QH1-QH5, CB1-CB4, TSBĐ1-TSBĐ4 và LC1-LC4 tương ứng với 7 biến X và biến Y trong bộ dữ liệu gốc.*

### 3.3. Population, Sampling Strategy and Data Collection
#### 3.3.1. Target Population & Sampling Method
> 👉 *[Nội dung hướng dẫn]: Mô tả tổng thể doanh nghiệp và phương pháp chọn mẫu phân tầng kết hợp thuận tiện.*

#### 3.3.2. Sample Size Determination
> 👉 *[Nội dung hướng dẫn]: Biện luận quy mô mẫu n = 800: Vượt xa tiêu chuẩn tối thiểu của EFA (n >= 170) và OLS (n >= 106).*

#### 3.3.3. Survey Administration across 155 VietinBank Branches
> 👉 *[Nội dung hướng dẫn]: Mô tả quy trình thu thập dữ liệu qua mạng lưới 155 chi nhánh VietinBank.*

### 3.4. Econometric & Quantitative Analytical Methods
#### 3.4.1. Descriptive Statistics
> 👉 *[Nội dung hướng dẫn]: Phân tích tần số, tỷ lệ phần trăm, Mean và Std Dev.*

#### 3.4.2. Scale Reliability Testing (Cronbach’s Alpha)
> 👉 *[Nội dung hướng dẫn]: Tiêu chuẩn Alpha >= 0.60 (tốt >= 0.80) và Corrected Item-Total Correlation >= 0.30.*

#### 3.4.3. Exploratory Factor Analysis (EFA)
> 👉 *[Nội dung hướng dẫn]: Tiêu chuẩn KMO >= 0.50, Bartlett sig < 0.05, Varimax rotation, Factor Loading >= 0.50, Cumulative Variance >= 50%.*

#### 3.4.4. Factor Scores Extraction Method
> 👉 *[Nội dung hướng dẫn]: Phương pháp tính giá trị biến đại diện (Mean Score / Regression Factor Score).*

#### 3.4.5. Pearson Correlation Analysis & Multicollinearity Diagnostics (VIF)
> 👉 *[Nội dung hướng dẫn]: Phân tích tương quan Pearson và kiểm định VIF < 5.*

#### 3.4.6. Multiple Linear Regression Model Specification (OLS)
> 👉 *[Nội dung hướng dẫn]: Phương trình hồi quy OLS chính thức:*

$$\text{DEC} = \beta_0 + \beta_1\text{COST\_COMP} + \beta_2\text{PROC\_SPEED} + \beta_3\text{DIGITAL\_CONV} + \beta_4\text{BANK\_REP} + \beta_5\text{RELATIONSHIP} + \beta_6\text{STAFF\_QUAL} + \beta_7\text{COLL\_POLICY} + \varepsilon \quad (3.1)$$

#### 3.4.7. Sub-Group Difference Testing Methods (ANOVA & t-test)
> 👉 *[Nội dung hướng dẫn]: Trình bày phương pháp Independent Samples t-test và One-Way ANOVA (Tukey post-hoc) so sánh khác biệt nhóm theo Objective 3.*

---

# CHAPTER 4: EMPIRICAL RESULTS, DISCUSSION AND MANAGERIAL RECOMMENDATIONS

### 4.1. Descriptive Statistics of the Sample ($n = 800$)
#### 4.1.1. Ownership Type Distribution
> 👉 *[Nội dung hướng dẫn]: Bảng và biểu đồ phân bổ mẫu theo SOEs, Doanh nghiệp Tư nhân, và Doanh nghiệp FDI.*

#### 4.1.2. Firm Revenue Scale Distribution
> 👉 *[Nội dung hướng dẫn]: Phân bổ mẫu theo quy mô doanh thu: Siêu nhỏ, Nhỏ, Vừa (SMEs) và Lớn (Corporate).*

#### 4.1.3. Operating Experience Distribution
> 👉 *[Nội dung hướng dẫn]: Phân bổ mẫu theo số năm hoạt động (< 3 năm, 3-5 năm, 5-10 năm, > 10 năm).*

#### 4.1.4. Usage Distribution of Bank Guarantee Products
> 👉 *[Nội dung hướng dẫn]: Phân bổ tần suất sử dụng bảo lãnh Dự thầu (TG), Tạm ứng (APG), Thực hiện HĐ (PG), Thanh toán (BG).*

### 4.2. Scale Reliability Analysis Results (Cronbach’s Alpha)
> 👉 *[Nội dung hướng dẫn]: Bảng kết quả hệ số Cronbach's Alpha cho 7 biến độc lập và biến DEC (tất cả đều đạt > 0.80).*

### 4.3. Exploratory Factor Analysis Results (EFA)
#### 4.3.1. EFA for Independent Variables
> 👉 *[Nội dung hướng dẫn]: Bảng KMO, Bartlett's Test, Tổng phương sai trích và Ma trận nhân tố xoay (Rotated Component Matrix).*

#### 4.3.2. EFA for Dependent Variable (DEC)
> 👉 *[Nội dung hướng dẫn]: Kết quả EFA hội tụ 4 biến quan sát vào 1 nhân tố DEC duy nhất.*

### 4.4. Correlation Analysis & Multicollinearity Diagnostics (VIF)
> 👉 *[Nội dung hướng dẫn]: Bảng ma trận tương quan Pearson và bảng kiểm định VIF < 2.0 (khẳng định không có đa cộng tuyến).*

### 4.5. Multiple Linear Regression Results (OLS)
#### 4.5.1. Model Summary & Goodness of Fit
> 👉 *[Nội dung hướng dẫn]: Bảng tóm tắt mô hình: R, R-Square, Adjusted R-Square, kiểm định F (sig < 0.001).*

#### 4.5.2. Estimated Coefficients and Hypothesis Testing
> 👉 *[Nội dung hướng dẫn]: Bảng hệ số B, Beta, p-value. Khẳng định chấp nhận 7 giả thuyết H1 đến H7 mang dấu dương (+), xếp hạng tầm quan trọng tương đối (theo Objective 2).*

### 4.6. Sub-Group Difference Analysis Results (ANOVA & t-test)
#### 4.6.1. Selection Differences across Ownership Types
> 👉 *[Nội dung hướng dẫn]: Kết quả ANOVA so sánh sự khác biệt trong quyết định chọn giữa SOE, Tư nhân và FDI (theo Objective 3).*

#### 4.6.2. Selection Differences across Firm Scales and Operating Experience
> 👉 *[Nội dung hướng dẫn]: Kết quả ANOVA so sánh giữa các nhóm quy mô doanh thu và thâm niên.*

#### 4.6.3. Selection Differences across Guarantee Product Types
> 👉 *[Nội dung hướng dẫn]: Kết quả ANOVA so sánh giữa các nhóm sử dụng bảo lãnh Dự thầu, Tạm ứng, Thực hiện HĐ và Thanh toán.*

### 4.7. Discussion of Empirical Findings
> 👉 *[Nội dung hướng dẫn]: Thảo luận học thuật: Giải thích vai trò dẫn dắt của COST_COMP và BANK_REP, sự trỗi dậy của kênh số eFAST (DIGITAL_CONV).*

### 4.8. Managerial Implications & Policy Recommendations for VietinBank
#### 4.8.1. Enhancing Core Service Capabilities (Response to Objective 1)
> 👉 *[Nội dung hướng dẫn]: Tập trung nguồn lực phát triển 7 giá trị cốt lõi.*

#### 4.8.2. Strategies for Price Competitiveness & Processing Speed (Response to Objective 2)
> 👉 *[Nội dung hướng dẫn]: Biểu phí chiết khấu linh hoạt và tinh giản quy trình thẩm định để phát hành siêu tốc.*

#### 4.8.3. Tailored Guarantee Packages for Corporate Segments (Response to Objective 3)
> 👉 *[Nội dung hướng dẫn]: May đo sản phẩm: Gói hạn mức lớn cho SOEs, Gói eFAST nhanh gọn cho SMEs, Gói URDG 758 cho FDI.*

#### 4.8.4. Breakthrough Digital Transformation Strategy via VietinBank eFAST (Response to Objective 4)
> 👉 *[Nội dung hướng dẫn]: Đẩy mạnh phát hành e-guarantee 24/7 qua eFAST, tích hợp tư vấn pháp lý Luật Đấu thầu.*

### 4.9. Policy Recommendations for the State Bank of Vietnam
> 👉 *[Nội dung hướng dẫn]: Kiến nghị với NHNN về hoàn thiện khung pháp lý thực thi Thông tư 61/2024/TT-NHNN.*

### 4.10. Research Limitations and Suggestions for Future Research
> 👉 *[Nội dung hướng dẫn]: Giới hạn về mẫu và các hướng nghiên cứu mở rộng trong tương lai.*

---

# REFERENCES (TÀI LIỆU THAM KHẢO)
> 👉 *[Nội dung hướng dẫn]: Danh mục 27 tài liệu tham khảo có thật 100% theo chuẩn Harvard A-Z (Diamond, Merton, Boot, Parasuraman, Davis, Berger, Carletti, Thông tư 61/2024/TT-NHNN, URDG 758...).*

---

# APPENDICES (HỆ THỐNG PHỤ LỤC)
* **Appendix 1: Official 34-Item Survey Questionnaire (English & Vietnamese)**
* **Appendix 2: Sample Demographic Characteristics Output**
* **Appendix 3: Cronbach’s Alpha Reliability Analysis Output**
* **Appendix 4: EFA Total Variance Explained & Rotated Component Matrix Output**
* **Appendix 5: OLS Multiple Regression, VIF & Sub-group ANOVA Output**
