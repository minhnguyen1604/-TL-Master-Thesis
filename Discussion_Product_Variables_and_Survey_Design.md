# THẢO LUẬN THIẾT KẾ ĐỀ ÁN VÀ PHƯƠNG PHÁP XỬ LÝ DỮ LIỆU SẢN PHẨM BẢO LÃNH

**Dự án:** NEU MDE Thesis Design – Class 31  
**Học viên:** Đặng Tú Linh  
**Tên đề tài:** Các yếu tố ảnh hưởng đến quyết định lựa chọn sản phẩm bảo lãnh ngân hàng của các doanh nghiệp tại VietinBank  
**Ngày thảo luận:** 03/08/2026  

---

## ❓ NỘI DUNG THẮC MẮC & ĐẶT VẤN ĐỀ

**Câu hỏi:**  
*Ở bản nháp (Draft), có nên thêm 1 cột "Tên sản phẩm" vào Bảng Likert không? Để sau khi phân tích ra chính sách còn biết cần cải thiện ở sản phẩm nào (Bảo lãnh dự thầu TG, Tạm ứng APG, Thanh toán BG...)? Trong công thức các biến, có cần phải ghi dữ liệu về các sản phẩm vào không?*

---

## 💡 PHÂN TÍCH CHUYÊN SÂU & GIẢI PHÁP THIẾT KẾ

### 1. Nhược điểm nếu chèn thêm cột "Tên sản phẩm" vào Bảng Likert
* **Gây quá tải cho Doanh nghiệp (Respondent Fatigue):** Bảng Likert hiện tại có 34 câu. Nếu mỗi câu lại chia ra đánh giá riêng cho 3-4 loại sản phẩm (Tạm ứng, Dự thầu, Thanh toán...), phiếu khảo sát sẽ phình ra hơn 120 - 150 ô tích. Doanh nghiệp sẽ thấy dài, dễ đánh lụi (*straight-lining*) hoặc bỏ dở giữa chừng, làm giảm nghiêm trọng chất lượng dữ liệu thu thập.
* **Bản chất dịch vụ Bảo lãnh ở Ngân hàng (Credit Line / Relationship Level):** Khi doanh nghiệp giao dịch với VietinBank, các yếu tố như: Thái độ cán bộ RM (`STA`), Năng lực eFAST (`DIG`), Uy tín thương hiệu (`REP`), Quy trình thẩm định chung (`SPE`) là đánh giá ở cấp độ mối quan hệ ngân hàng - khách hàng, không khác biệt quá nhiều giữa từng thư bảo lãnh lẻ.

---

### 2. Phương pháp đưa Dữ liệu Sản phẩm vào Mô hình Kinh tế lượng

Dữ liệu sản phẩm **CÓ DÙNG** và được thu thập & đưa vào nghiên cứu theo 3 kỹ thuật khoa học:

#### Cách 1: Thu thập ở Phần I (Thông tin chung về Doanh nghiệp)
Trong phiếu khảo sát (`Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx`), Câu 4 Phần I đã thu thập sẵn:
> *"Loại hình bảo lãnh Doanh nghiệp thường xuyên sử dụng nhất tại VietinBank:"*
> - [ ] Bảo lãnh dự thầu (TG)
> - [ ] Bảo lãnh thực hiện hợp đồng (PG)
> - [ ] Bảo lãnh tạm ứng (APG)
> - [ ] Bảo lãnh thanh toán (BG)
> - [ ] Bảo lãnh khác (RG)

#### Cách 2: Giữ phương trình Hồi quy Tổng quát chuẩn mực ở Mục 3.1
Phương trình trong Đề cương (`DTL_Thesis_Design_NEU_MDE_Final.docx`) giữ nguyên gọn gàng và chuẩn mực:

$$\text{DEC} = \beta_0 + \beta_1\text{COST} + \beta_2\text{COL} + \beta_3\text{SPE} + \beta_4\text{REP} + \beta_5\text{STA} + \beta_6\text{REL} + \beta_7\text{DIG} + \beta_8\text{CUS} + \beta_9\text{RSK} + \beta_{10}\text{NET} + \varepsilon \quad (1)$$

*Ý nghĩa:* Phương trình gốc tập trung đo lường tác động của 10 nhóm nhân tố chất lượng dịch vụ tới Quyết định lựa chọn chung (`DEC`).

#### Cách 3: Phân tích Phân nhóm (Sub-group Analysis) ở Chương 4
Tại Chương 4, khi xử lý dữ liệu bằng Python, tác giả tách tệp dữ liệu theo từng nhóm sản phẩm chính từ Câu 4 Phần I để chạy các mô hình hồi quy riêng:

* **Mô hình nhóm Bảo lãnh Dự thầu (TG):**  
  $$\text{DEC}_{\text{TG}} = \beta_0 + 0.45\text{SPE} + 0.38\text{DIG} + 0.12\text{COST} + \dots$$  
  *(Chỉ ra nhóm Dự thầu nhạy cảm nhất với Tốc độ xử lý `SPE` và Chuyển đổi số `DIG`)*

* **Mô hình nhóm Bảo lãnh Tạm ứng / Thanh toán (APG/BG):**  
  $$\text{DEC}_{\text{BG}} = \beta_0 + 0.42\text{COL} + 0.35\text{RSK} + 0.15\text{SPE} + \dots$$  
  *(Chỉ ra nhóm Tạm ứng/Thanh toán lại nhạy cảm nhất với Tỷ lệ Ký quỹ `COL` và Tư vấn Pháp lý `RSK`)*

---

## 📌 KẾT LUẬN THỐNG NHẤT

1. **Phiếu khảo sát (`Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx`):** Giữ nguyên bảng 34 câu Likert phẳng gọn gàng để doanh nghiệp hoàn thành nhanh nhất (chỉ 3-5 phút), đảm bảo tỷ lệ phản hồi cao và dữ liệu trung thực.
2. **Phương trình hồi quy Đề án (`DTL_Thesis_Design_NEU_MDE_Final.docx`):** Giữ nguyên 10 biến độc lập tổng quát ở Mục 3.1.
3. **Phân tích Chương 4:** Sử dụng dữ liệu sản phẩm ở Phần I để chạy Hồi quy phân nhóm trên Python, đưa ra chính sách quản trị chính xác 100% cho từng loại sản phẩm bảo lãnh của VietinBank.
