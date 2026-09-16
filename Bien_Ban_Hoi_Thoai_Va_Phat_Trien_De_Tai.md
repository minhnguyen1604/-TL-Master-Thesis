# BIÊN BẢN HỘI THOẠI & NHẬT KÝ PHÁT TRIỂN ĐỀ TÀI LUẬN VĂN THẠC SĨ
## CHƯƠNG TRÌNH CAO HỌC KINH TẾ PHÁT TRIỂN VIỆT NAM - HÀ LAN (MDE - KHÓA 31)
**Trường Đại học Kinh tế Quốc dân (NEU) & Viện Nghiên cứu Xã hội Quốc tế (ISS) - Đại học Erasmus Rotterdam**

* **Học viên:** Đặng Tú Linh (Lớp MDE Khóa 31)
* **Giảng viên hướng dẫn:** TS. Hoàng Thị Thúy Nga
* **Tên đề tài tiếng Anh chính thức:** *Factors Affecting Corporate Customers’ Decision to Choose Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade (VietinBank)*
* **Tên đề tài tiếng Việt:** *Các yếu tố ảnh hưởng đến quyết định lựa chọn dịch vụ bảo lãnh ngân hàng của khách hàng doanh nghiệp tại Ngân hàng TMCP Công thương Việt Nam (VietinBank)*
* **Thư mục làm việc cục bộ:** `c:\Users\Admin\OneDrive\Desktop\ĐTL-Master-Thesis`
* **Kho lưu trữ GitHub:** [https://github.com/minhnguyen1604/-TL-Master-Thesis](https://github.com/minhnguyen1604/-TL-Master-Thesis)

---

# MỤC LỤC
1. [Phiên 1: Đồng bộ & Clone Dữ liệu Dự án từ GitHub](#1-phiên-1-đồng-bộ--clone-dữ-liệu-dự-án-từ-github)
2. [Phiên 2: Xác định Hệ thống File Form Mẫu Tạo Sẵn để Triển khai Viết](#2-phiên-2-xác-định-hệ-thống-file-form-mẫu-tạo-sẵn-để-triển-khai-viết)
3. [Phiên 3: Đánh giá Chuyên sâu về Độ Chặt Chẽ Học Thuật & 5 Điểm Hở Tiềm Ẩn](#3-phiên-3-đánh-giá-chuyên-sâu-về-độ-chặt-chẽ-học-thuật--5-điểm-hở-tiềm-ẩn)
4. [Phiên 4: Phân loại Chi tiết Các Phần Viết Ngay (Không Cần Chạy Mô Hình)](#4-phiên-4-phân-loại-chi-tiết-các-phần-viết-ngay-không-cần-chạy-mô-hình)
5. [Phiên 5: Quyết định Điều chỉnh Cấu trúc Phụ lục (Loại Bỏ Phiếu Khảo Sát Khỏi Phụ Lục)](#5-phiên-5-quyết-định-điều-chỉnh-cấu-trúc-phụ-lục-loại-bỏ-phiếu-khảo-sát-khỏi-phụ-lục)
6. [Kế hoạch & Lộ trình Triển khai Viết Tiếp Theo](#6-kế-hoạch--lộ-trình-triển-khai-viết-tiếp-theo)

---

## 1. PHIÊN 1: ĐỒNG BỘ & CLONE DỮ LIỆU DỰ ÁN TỪ GITHUB

* **Yêu cầu từ Người dùng:**
  > *"clone tất cả dữ liệu của dự án này xuống folder này nhé"* (kèm ảnh chụp giao diện kho lưu trữ GitHub).

* **Phân tích & Xử lý kỹ thuật:**
  * Trích xuất thông tin định danh từ ảnh chụp:
    * Chủ sở hữu: `minhnguyen1604`
    * Tên repository: `-TL-Master-Thesis` (viết tắt của ĐTL-Master-Thesis trên hệ thống GitHub)
    * Tác giả commit: Nguyễn Tuấn Minh (BVGI Act...)
  * Thực thi lệnh `git clone https://github.com/minhnguyen1604/-TL-Master-Thesis.git .` trực tiếp vào thư mục gốc `c:\Users\Admin\OneDrive\Desktop\ĐTL-Master-Thesis`.
  * Kết quả: Tải về đầy đủ toàn bộ 13 tệp tin tài liệu, mã nguồn và dữ liệu:
    1. `README.md`
    2. `ĐTL-Master Thesis.docx`
    3. `Du_Lieu_Mau_Khao_Sat_SPSS.csv`
    4. `Du_Lieu_SPSS_Hoan_Chinh.csv`
    5. `Ma_Tran_Hoi_Quy_11_Bien.csv`
    6. `Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx` & `.md`
    7. `Tai_Lieu_Nghien_Cuu_Bao_Lanh_Ngan_Hang.md`
    8. `Tong_Hop_Bao_Cao_De_Tai_Master_VietinBank.docx` & `.md`
    9. `create_conversation_summary.py`
    10. `create_survey_docx.py`
    11. `excel_summary.txt`

---

## 2. PHIÊN 2: XÁC ĐỊNH HỆ THỐNG FILE FORM MẪU TẠO SẴN ĐỂ TRIỂN KHAI VIẾT

* **Yêu cầu từ Người dùng:**
  > *"em biết file nào là file tạo sẵn form để viết chưa? chỉ ra cho anh nhé, xong anh chỉ dẫn tiếp"*

* **Phân tích & Kết luận Hệ thống Tệp Tin:**
  Trong thư mục dự án có 3 bộ file form mẫu được thiết kế sẵn theo từng cấp độ:

  1. **File Khung Cấu trúc Luận văn Hoàn chỉnh (File Chính để Viết Luận văn):**
     * **File Word:** `DTL_Master_Thesis_Full_Structure_Template.docx`
     * **File Markdown:** `DTL_Master_Thesis_Full_Structure_Template.md`
     * *Đặc điểm:* Đã định dạng chuẩn NEU MDE (Margin 1.0 inch, Running Header có gạch chân, Spacing 1.15, Font Times New Roman), dựng sẵn 100% cấu trúc các đề mục phân cấp (Heading 1, 2, 3), từ Trang Bìa, Lời cam đoan, Lời cảm ơn, Tóm tắt, Mục lục, 4 Chương nội dung đến Tài liệu tham khảo và Phụ lục. Dưới mỗi đề mục đều có hướng dẫn định hướng nội dung cần điền.

  2. **File Đề cương Luận văn Chi tiết (File Nộp Báo cáo GVHD):**
     * **File Word:** `DTL_Thesis_Design_NEU_MDE_Final.docx`
     * **File Markdown:** `DTL_Thesis_Design_NEU_MDE_Final.md`
     * *Đặc điểm:* Đề cương chi tiết (Phương án 2 – Khảo sát $n = 800$, OLS 7 biến) đã được chỉnh sửa chuẩn hóa theo góp ý của GVHD TS. Hoàng Thị Thúy Nga.

  3. **File Mẫu Phiếu Khảo sát Doanh nghiệp:**
     * **File Word:** `Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx`
     * **File Markdown:** `Phieu_Khao_Sat_Chinh_Thuc_VietinBank.md`
     * *Đặc điểm:* Mẫu phiếu khảo sát gồm 34 câu hỏi thang đo Likert 1-5 dành cho khách hàng doanh nghiệp VietinBank.

---

## 3. PHIÊN 3: ĐÁNH GIÁ CHUYÊN SÂU VỀ ĐỘ CHẶT CHẼ HỌC THUẬT & 5 ĐIỂM HỞ TIỀM ẨN

* **Yêu cầu từ Người dùng:**
  > *"đọc các phần chính này và đánh giá cho anh mọi thứ đã chặt chẽ chưa nhé, có điểm nào hở không? nếu có thì chỉ ra, sau đó anh sẽ chỉ em các bước tiếp theo"*

### 3.1. Các Điểm Mạnh Cốt Lõi Về Độ Chặt Chẽ (Strengths)
1. **Tiếp thu 100% Góp ý của GVHD (TS. Hoàng Thị Thúy Nga):**
   * Rút gọn từ 10 biến xuống đúng **7 biến độc lập trọng tâm** (nằm trong ngưỡng khuyến nghị 6–7 biến).
   * Đổi tên biến `COST` thành **`COST_COMP` (Price Competitiveness)** và đổi dấu kỳ vọng từ Âm (–) sang **DƯƠNG (+)** phù hợp với câu hỏi Likert 5 điểm ("Biểu phí rất hợp lý và cạnh tranh").
   * Lập luận bảo vệ thành công biến **Chuyển đổi số eFAST (`DIGITAL_CONV`)** dựa trên Mô hình Chấp nhận Công nghệ TAM (Davis, 1989) và bối cảnh Thông tư 61/2024/TT-NHNN.
   * Tách riêng Uy tín thương hiệu (`BANK_REP`) khỏi biến Năng lực cán bộ (`STAFF_QUAL`).
2. **Cơ sở Lý thuyết Bảo chứng 1-1 cho 7 Biến Độc Lập:**
   * `BANK_REP`: Lý thuyết Trung gian Tài chính & Giám sát Ủy thác (*Diamond, 1984*).
   * `COST_COMP` & `COLL_POLICY`: Lý thuyết Định giá Rủi ro Tín dụng (*Merton, 1974; Stiglitz & Weiss, 1981*).
   * `PROC_SPEED` & `STAFF_QUAL`: Mô hình Chất lượng Dịch vụ SERVQUAL (*Parasuraman et al., 1988*).
   * `RELATIONSHIP`: Lý thuyết Ngân hàng Quan hệ (*Boot, 2000; Berger & Udell, 1995*).
   * `DIGITAL_CONV`: Mô hình Chấp nhận Công nghệ TAM (*Davis, 1989; Venkatesh et al., 2003*).
3. **Tính Nhất Quán Giữa Mục Tiêu và Giải Pháp:**
   * 4 Mục tiêu nghiên cứu (Objectives 1–4) ánh xạ trực tiếp 1-đổi-1 sang 4 câu hỏi nghiên cứu và 4 nhóm khuyến nghị quản trị tại Chương 4.
4. **Cập nhật Pháp lý Mới Nhất:**
   * Thông tư 61/2024/TT-NHNN (hiệu lực 01/4/2025 về nghiệp vụ bảo lãnh ngân hàng & bảo lãnh điện tử), Luật Đấu thầu 22/2023/QH15 và ICC URDG 758.

### 3.2. 5 Điểm Hở Học Thuật Tiềm Ẩn & Phương Án Phòng Thủ Khi Bảo Vệ

| Điểm hở / Nguy cơ | Bản chất vấn đề học thuật | Rủi ro trước Hội đồng | Phương án Phòng thủ & Khắc phục |
| :--- | :--- | :--- | :--- |
| **1. Đa cộng tuyến (Multicollinearity)** | Trong B2B, `RELATIONSHIP`, `BANK_REP` và `COLL_POLICY` thường gắn chặt với nhau; DN quan hệ lâu năm thì được hạn mức ký quỹ tốt và tin tưởng uy tín Big4. | Hệ số tương quan $r > 0.7$; VIF có thể tăng cao hoặc bị gộp nhân tố khi chạy EFA. | Trong Chương 3, kiểm soát $VIF < 3.0$; nếu bị gộp nhân tố trong EFA thì giải trình bằng lý thuyết B2B Relationship Banking (Berger & Udell, 1995). |
| **2. Thuật ngữ Biến DEC (Likert vs Choice)** | Quyết định lựa chọn (Choice) trong kinh tế lượng thuần túy là biến nhị phân $0/1$ (Logit/Probit), nhưng ở đây đo bằng thang Likert 1-5 và chạy OLS. | Giám khảo hỏi: *"Tại sao gọi là Choice Decision mà lại chạy OLS trên thang đo Likert?"* | Làm rõ bản chất của `DEC` là **"Selection Priority & Preference Level"** (Mức độ ưu tiên & Thiên hướng lựa chọn). Việc dùng OLS cho biến Likert trung bình là chuẩn mực trong nghiên cứu B2B Service Quality (Zeithaml, Parasuraman; Narteh, 2013). |
| **3. Thiên vị Chọn mẫu (Selection Bias)** | Bộ dữ liệu $n = 800$ được thu thập từ các khách hàng đang giao dịch tại 155 chi nhánh VietinBank. | Mẫu chỉ gồm doanh nghiệp *đã chọn* VietinBank nên điểm `DEC` bị lệch cao (Survivorship Bias). | Ghi rõ trong mục Hạn chế (Limitations): Nghiên cứu tập trung vào **"Hành vi duy trì & Gia tăng thị phần lựa chọn của KHDN hiện hữu"** (Customer Retention & Patronage Share), đo mức độ ưu tiên so với đối thủ. |
| **4. Sót Biến kiểm soát (Control Variables)** | Ở Mục tiêu 3 ta làm ANOVA so sánh khác biệt nhóm (loại hình sở hữu, quy mô, loại bảo lãnh) nhưng phương trình OLS gốc chưa đưa biến kiểm soát vào. | Có thể bị nhận xét là thiếu biến kiểm soát (Omitted Variable Bias) trong hồi quy OLS. | Bổ sung phần Kiểm định Độ vững (Robustness Check) tại Mục 4.5 đưa thêm các biến giả kiểm soát (`D_SOE`, `D_FDI`, `D_LARGE`, `D_TENDER`) để khẳng định các hệ số $\beta_1 \dots \beta_7$ vẫn vững. |
| **5. Phân hóa Nội địa vs Quốc tế (URDG 758)** | Doanh nghiệp nội địa quan tâm Luật Đấu thầu & Thông tư 61, doanh nghiệp FDI/XNK quan tâm Bảo lãnh đối ứng theo URDG 758. | Nhóm FDI có thể có hành vi lựa chọn khác biệt lớn so với doanh nghiệp trong nước. | Sử dụng kết quả phân tích nhóm ở Mục tiêu 3 để luận giải sâu sắc sự khác biệt này ở Chương 4. |

---

## 4. PHIÊN 4: PHÂN LOẠI CHI TIẾT CÁC PHẦN VIẾT NGAY (KHÔNG CẦN CHẠY MÔ HÌNH)

* **Yêu cầu từ Người dùng:**
  > *"đánh giá cho anh trong form mẫu của bài, file word ý nhé, những phần nào có thể viết luôn được mà không cần chạy mô hình, liệt kê ra giúp anh, sau đó anh sẽ duyệt nhé"*

* **Bảng Phân loại Khối lượng Công việc:**

### Nhóm A: CÓ THỂ TRIỂN KHAI VIẾT NGAY 100% (Chiếm ~60 - 65% Luận văn)
1. **Phần Đầu Luận văn (Front Matter):**
   * Trang Bìa chuẩn NEU MDE
   * Lời cam đoan (*Statement of Authorship*)
   * Lời cảm ơn (*Acknowledgements*)
   * Danh mục Từ viết tắt (*List of Abbreviations*)
2. **CHƯƠNG 1: INTRODUCTION (Viết trọn vẹn 100%):**
   * 1.1. Research Rationales & Background
   * 1.2. Research Problem & Industry Context at VietinBank
   * 1.3. Research Objectives (General Objective & 4 Specific Objectives)
   * 1.4. Research Questions (4 Câu hỏi nghiên cứu)
   * 1.5. Scope and Boundaries of the Study
   * 1.6. Significance & Contributions of the Study
   * 1.7. Structure of the Thesis
3. **CHƯƠNG 2: LITERATURE REVIEW AND THEORETICAL FRAMEWORK (Viết trọn vẹn 100%):**
   * 2.1. Tổng quan hoạt động Bảo lãnh Ngân hàng (Khái niệm, phân loại, khung pháp lý TT 61/2024/TT-NHNN, Luật Đấu thầu 2023, URDG 758)
   * 2.2. Khung Lý thuyết Nền tảng (5 Lý thuyết: Diamond 1984, Merton 1974, SERVQUAL 1988, Boot 2000, TAM 1989)
   * 2.3. Tổng quan Nghiên cứu Thực nghiệm (Quốc tế và Việt Nam)
   * 2.4. Khoảng trống Nghiên cứu (Research Gaps)
   * 2.5. Khung Khái niệm & Phát triển 7 Giả thuyết Nghiên cứu ($H_1 \dots H_7$)
4. **CHƯƠNG 3: RESEARCH METHODOLOGY AND EMPIRICAL DESIGN (Viết trọn vẹn 100%):**
   * 3.1. Thiết kế Nghiên cứu Tổng thể & Quy trình 7 bước
   * 3.2. Thiết kế Bảng hỏi & Vận hành hóa Biến (Operationalization of Variables - 34 items)
   * 3.3. Tổng thể, Chiến lược Chọn mẫu & Thu thập Dữ liệu ($n = 800$, 155 chi nhánh)
   * 3.4. Phương pháp Phân tích Định lượng (Tiêu chuẩn Cronbach's Alpha, EFA, VIF, OLS Equation, Sub-group ANOVA/t-test)

### Nhóm B: PHẦN TRIỂN KHAI SAU KHI CHẠY MÔ HÌNH (Chiếm ~35 - 40% Luận văn)
1. **CHƯƠNG 4: EMPIRICAL RESULTS, DISCUSSION AND MANAGERIAL RECOMMENDATIONS:**
   * 4.1. Thống kê mô tả mẫu ($n = 800$)
   * 4.2. Kết quả Cronbach's Alpha
   * 4.3. Kết quả Phân tích Nhân tố Khám phá EFA
   * 4.4. Ma trận tương quan Pearson & Đa cộng tuyến VIF
   * 4.5. Kết quả Hồi quy Tuyến tính OLS (Hệ số $\beta$, $R^2$, $F$-test, $p$-value)
   * 4.6. Kết quả Kiểm định Khác biệt Nhóm ANOVA & t-test
   * 4.7. Thảo luận Kết quả Thực nghiệm
   * 4.8. Hàm ý Quản trị & Giải pháp cho VietinBank (Khớp 4 mục tiêu)
   * 4.9. Kiến nghị với Ngân hàng Nhà nước
   * 4.10. Hạn chế của Đề tài & Hướng nghiên cứu tương lai
2. **Abstract (Tóm tắt) & Danh mục Bảng biểu:** Điền các con số định lượng cụ thể ($R^2$, hệ số $\beta$ lớn nhất...).

---

## 5. PHIÊN 5: QUYẾT ĐỊNH ĐIỀU CHỈNH CẤU TRÚC PHỤ LỤC (LOẠI BỎ PHIẾU KHẢO SÁT KHỎI PHỤ LỤC)

* **Yêu cầu từ Người dùng:**
  > *"phụ lục không đưa phần khảo sát vào nhé"*

* **Tiếp thu & Thực thi:**
  * Đã loại bỏ hoàn toàn mẫu Phiếu khảo sát 34 câu hỏi ra khỏi phần Phụ lục (Appendices).
  * Quy chuẩn lại Phụ lục chỉ tập trung vào **các Bảng Biểu Đầu ra Thống kê & Kinh tế lượng Thực nghiệm (Statistical Outputs)**:
    * **Appendix 1:** Sample Demographic Characteristics Output *(Bảng tần số cơ cấu mẫu n = 800)*
    * **Appendix 2:** Cronbach’s Alpha Reliability Analysis Output *(Độ tin cậy thang đo 7 biến & DEC)*
    * **Appendix 3:** EFA Total Variance Explained & Rotated Component Matrix Output *(Ma trận xoay nhân tố EFA)*
    * **Appendix 4:** OLS Multiple Regression, VIF & Sub-group ANOVA Output *(Kết quả hồi quy OLS, VIF và ANOVA)*
  * Đã cập nhật đồng bộ thay đổi này vào:
    * File Markdown: `DTL_Master_Thesis_Full_Structure_Template.md`
    * File Script sinh Word: `generate_full_thesis_structure_template_docx.py`

---

## 6. KẾ HOẠCH & LỘ TRÌNH TRIỂN KHAI VIẾT TIẾP THEO

1. **Giai đoạn 1 (Viết hoàn thiện Khối lượng Độc lập - Nhóm A):**
   * Triển khai viết chi tiết toàn bộ nội dung học thuật cho **Chương 1 (Introduction)**, **Chương 2 (Literature Review)** và **Chương 3 (Methodology)** vào file Word theo đúng định dạng chuẩn mực của Chương trình Thạc sĩ MDE NEU.
2. **Giai đoạn 2 (Chạy Thực nghiệm & Hoàn thiện Chương 4 - Nhóm B):**
   * Xử lý bộ dữ liệu khảo sát $n = 800$, chạy các lệnh phân tích định lượng (Cronbach's Alpha, EFA, OLS Regression, ANOVA).
   * Lắp các bảng biểu kết quả vào Chương 4 và Phụ lục (Appendix 1 – 4).
   * Viết phần Thảo luận và 4 Nhóm Khuyến nghị Quản trị cho VietinBank.
3. **Giai đoạn 3 (Hoàn thiện Tổng thể & Rà soát Đóng gói):**
   * Hoàn thiện trang Abstract, đồng bộ Danh mục Bảng biểu (List of Tables), Danh mục Hình (List of Figures).
   * Xuất bản bản in hoàn chỉnh phục vụ nộp chấm Luận văn Thạc sĩ.

---
*Biên bản được lưu trữ tự động tại thư mục dự án: `c:\Users\Admin\OneDrive\Desktop\ĐTL-Master-Thesis\Bien_Ban_Hoi_Thoai_Va_Phat_Trien_De_Tai.md`.*
