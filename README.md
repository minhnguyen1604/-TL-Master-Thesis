# VIETNAM-NETHERLANDS MASTER’S PROGRAM IN DEVELOPMENT ECONOMICS (MDE)
## MASTER THESIS PROJECT – DANG TU LINH (MDE CLASS 31)
**National Economics University (NEU) & Erasmus University Rotterdam / International Institute of Social Studies (ISS)**  
**Supervisor:** Dr. Hoang Thi Thuy Nga  
**Student:** Dang Tu Linh  

---

## 📌 OVERVIEW OF THE RESEARCH PROJECT

Dự án Luận văn Thạc sĩ Kinh tế Phát triển (MDE Khóa 31) tập trung nghiên cứu hoạt động **Bảo lãnh Ngân hàng dành cho Khách hàng Doanh nghiệp** tại **Ngân hàng TMCP Công thương Việt Nam (VietinBank)**.

Dự án được xây dựng và chuẩn hóa theo **2 Phương án nghiên cứu học thuật chiến lược**:

1. **PHƯƠNG ÁN 2 (BẢN CHÍNH THỨC ĐÃ CẬP NHẬT THEO GÓP Ý CỦA GVHD):**
   * **Title:** *“Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)”*
   * **Phương pháp:** Hồi quy OLS trên dữ liệu khảo sát sơ cấp $n = 800$ khách hàng doanh nghiệp tại 155 chi nhánh VietinBank.
   * **Hệ thống 7 biến độc lập cốt lõi:** `COST_COMP` (+), `PROC_SPEED` (+), `DIGITAL_CONV` (+), `BANK_REP` (+), `RELATIONSHIP` (+), `STAFF_QUAL` (+), `COLL_POLICY` (+).
   * **Biến phụ thuộc:** `DEC` – Corporate Selection Decision (Mức độ ưu tiên lựa chọn VietinBank so với các ngân hàng đối thủ).
   * **Bổ sung:** Phân tích khác biệt nhóm (Sub-group Analysis / ANOVA / t-test) theo loại hình sở hữu, quy mô, thâm niên và loại sản phẩm bảo lãnh.

2. **PHƯƠNG ÁN 1 (HƯỚNG TÀI CHÍNH & DOANH THU THU NHẬP PHÍ NGOẠI BẢNG):**
   * **Title:** *“Factors Affecting Guarantee Fee Income and Financial Efficiency of Corporate Customers at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)”*
   * **Phương pháp:** Hồi quy OLS 2 phương trình trên dữ liệu thứ cấp trích xuất từ Core Banking MIS ($n = 800$ doanh nghiệp).
   * **Biến mục tiêu:** $\ln(\text{FEE\_INCOME})$ (Tổng thu nhập phí tuyệt đối) và $\text{FEE\_YIELD}$ (Tỷ lệ thu phí / Biên lợi nhuận phí ròng - Robustness Check).
   * **8 biến độc lập tài chính:** `ln_LIMIT`, `MARGIN_RATIO`, `TENOR`, `FIRM_SIZE`, `FIRM_AGE`, `CREDIT_RATING`, `RELATIONSHIP`, `DIGITAL`.

---

## 📂 DANH MỤC TỆP TIN TRONG KHO LƯU TRỮ (REPOSITORY MAP)

### 📄 1. Hồ sơ Đề cương Luận văn (Thesis Design Proposals)
* **`DTL_Thesis_Design_NEU_MDE_Final.docx`**: **File Word Đề cương Luận văn chính thức (Phương án 2)** – Đã cập nhật 100% theo nhận xét của TS. Hoàng Thị Thúy Nga (chuẩn 7 biến, 4 mục tiêu tiếng Anh NEU MDE, căn lề $1.0\text{ inch}$, running header có gạch chân xám `#888888`).
* **`DTL_Thesis_Design_NEU_MDE_Final.md`**: Bản xem trước Markdown của Đề cương Phương án 2.
* **`DTL_Thesis_Design_Option1_Fee_Income_Final.docx`**: **File Word Đề cương Luận văn Phương án 1** – Hướng Thu nhập Phí và Hiệu quả Tài chính trên dữ liệu giao dịch Core Banking.
* **`DTL_Thesis_Design_Option1_Fee_Income_Final.md`**: Bản xem trước Markdown của Đề cương Phương án 1.

### 📋 2. Báo cáo Giải trình & Nhật ký Phát triển Đề tài
* **`Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.docx`**: **Bản Word Báo cáo Giải trình Độc lập** – Tổng hợp chi tiết các điểm cô đồng tình, các điểm yêu cầu điều chỉnh, và Mục 1.3 giải trình cơ sở giữ lại biến Chuyển đổi số eFAST (`DIGITAL_CONV`).
* **`Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.md`**: Bản xem trước Markdown của Báo cáo giải trình.
* **`Nhat_Ky_Trao_Doi_Va_Phat_Trien_De_Tai_Master.md`**: **Nhật ký Đóng gói Toàn bộ Quá trình Trao đổi** – Hệ thống hóa toàn diện quá trình phát triển cả 2 phương án, đối soát câu hỏi - giải pháp, giải thích kinh tế lượng và giải trình học thuật.
* **`Tong_Hop_Bao_Cao_De_Tai_Master_VietinBank.docx`** & **`.md`**: Báo cáo tổng hợp bối cảnh ban đầu của đề tài Master VietinBank.
* **`Discussion_Product_Variables_and_Survey_Design.md`**: Tài liệu phân tích chuyên sâu về tối ưu độ dài bảng hỏi và cấu trúc biến sản phẩm bảo lãnh.

### 📝 3. Phiếu Khảo sát Chính thức (Survey Questionnaire)
* **`Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx`**: Mẫu Phiếu khảo sát chính thức 34 câu hỏi Likert 1-5 chuẩn định dạng Word bảng biểu, đã loại bỏ biến thừa ESG.
* **`Phieu_Khao_Sat_Chinh_Thuc_VietinBank.md`**: Bản xem trước Markdown của Phiếu khảo sát.

### 🐍 4. Mã nguồn Python Tự động hóa (Python Scripts)
* **`generate_final_mde_thesis_design.py`**: Script sinh tự động file Word Đề cương chính thức Phương án 2.
* **`generate_option1_mde_thesis_design.py`**: Script sinh tự động file Word Đề cương Phương án 1.
* **`generate_advisor_feedback_report_docx.py`**: Script sinh tự động file Word Báo cáo giải trình góp ý của GVHD.
* **`create_survey_docx.py`**: Script sinh tự động file Word Phiếu khảo sát 34 câu.

---

## ⚙️ HƯỚNG DẪN TÁI TẠO VÀ XUẤT BẢN FILE WORD

Để sinh lại toàn bộ các file Word chuẩn định dạng từ mã nguồn Python:

```bash
# 1. Sinh file Đề cương Luận văn chính thức (Phương án 2 - 7 biến)
python generate_final_mde_thesis_design.py

# 2. Sinh file Đề cương Luận văn Phương án 1 (Thu nhập Phí)
python generate_option1_mde_thesis_design.py

# 3. Sinh file Báo cáo Giải trình Góp ý GVHD (có mục eFAST)
python generate_advisor_feedback_report_docx.py

# 4. Sinh file Phiếu khảo sát chính thức 34 câu
python create_survey_docx.py
```

---

*Kho lưu trữ được quản lý và bảo vệ tại GitHub: [https://github.com/minhnguyen1604/-TL-Master-Thesis](https://github.com/minhnguyen1604/-TL-Master-Thesis)*
