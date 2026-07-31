# TỔNG HỢP TOÀN BỘ NỘI DUNG TRAO ĐỔI & ĐỊNH HƯỚNG LUẬN VĂN THẠC SĨ
**Chủ đề:** Nghiên cứu Dịch vụ Bảo lãnh Ngân hàng dành cho Khách hàng Doanh nghiệp tại VietinBank

---

## I. TÊN ĐỀ TÀI & PHẠM VI NGHIÊN CỨU TỔNG THỂ

* **Tên đề tài Tiếng Việt chính thức (Phương án Rộng & Sang):**
  > **"Các yếu tố ảnh hưởng đến phát triển hoạt động bảo lãnh ngân hàng cho khách hàng doanh nghiệp tại Ngân hàng Thương mại Cổ phần Công Thương Việt Nam"**

* **Tên đề tài Tiếng Anh tương ứng:**
  > **"Factors Affecting the Development of Corporate Bank Guarantee Services at Vietnam Joint Stock Commercial Bank for Industry and Trade"**

* **Tầm vóc đề tài:**
  Đề tài mang tính bao quát chiến lược toàn hệ thống VietinBank, kết hợp giữa quy mô doanh số phát hành, hiệu quả thu nhập phí bảo lãnh, và định hướng chuyển đổi số bảo lãnh điện tử (VietinBank eFAST).

---

## II. PHƯƠNG ÁN 1: NGHIÊN CỨU BẰNG DỮ LIỆU KHẢO SÁT (PRIMARY SURVEY DATA)

1. **Bản chất phương pháp:** 
   Phát phiếu khảo sát dạng thang đo Likert 5 điểm (1: Hoàn toàn không đồng ý -> 5: Hoàn toàn đồng ý) cho người đại diện doanh nghiệp (CFO, Kế toán trưởng, Trưởng phòng thầu).

2. **Cấu trúc Biến mục tiêu Y (Quyết định lựa chọn - DEC):**
   Gồm 4 biến quan sát đo lường 4 giai đoạn lòng trung thành theo khung lý thuyết Oliver (1999) & Zeithaml et al. (1996):
   - `DEC1`: Doanh nghiệp luôn ưu tiên lựa chọn VietinBank để phát hành thư bảo lãnh (Trung thành Nhận thức).
   - `DEC2`: Doanh nghiệp sẵn sàng tăng quy mô và tần suất sử dụng bảo lãnh VietinBank trong tương lai (Trung thành Ý định).
   - `DEC3`: Doanh nghiệp sẵn sàng giới thiệu dịch vụ bảo lãnh VietinBank cho các đối tác khác (Trung thành Lan tỏa).
   - `DEC4`: Nhìn chung, VietinBank là lựa chọn tối ưu nhất khi phát sinh nhu cầu bảo lãnh (Trung thành Thái độ so với đối thủ).

3. **Các biến độc lập (X1 -> X7):**
   Bao gồm `COST` (Chi phí), `COL` (Tài sản đảm bảo), `SPE` (Tốc độ xử lý), `REP` (Uy tín), `STA` (Cán bộ RM), `REL` (Mối quan hệ), và `DIG` (Chuyển đổi số & e-Guarantee - Biến mới).

---

## III. PHƯƠNG ÁN 2: NGHIÊN CỨU BẰNG DỮ LIỆU GIAO DỊCH THỰC TẾ (SECONDARY TRANSACTION DATA)

1. **Nguồn dữ liệu thực tế:**
   Trích xuất từ 2 File Excel dữ liệu gốc của VietinBank:
   - File 1: `doanh_số_BL_6T2026_theo_ĐCTC_tại_CN_108_ (2).xlsx` (Chi tiết Doanh số cấp bảo lãnh quy đổi VNĐ).
   - File 2: `Phí BL chi tiết từng CIF tại CN PK 98.xlsx` (Chi tiết Thu phí bảo lãnh, Loại bảo lãnh, Chi nhánh, Khu vực).
   -> Ghép dữ liệu từ 2 file về 1 Bảng duy nhất theo Khóa chung là Mã `Số CIF`.

2. **Các phương án xác định Biến mục tiêu (Y):**
   - **Phương án 1 (Biến đơn):** $Y = \ln(\text{Thu\_Phi\_BL\_VND})$
   - **Phương án 2 (Kỹ thuật PCA):** Gộp Doanh số và Thu phí bằng Chuẩn hóa Z-score & Phân tích nhân tố PCA -> `Y_PERFORMANCE`
   - **Phương án 3 (Tỷ lệ Thu phí):** $Y = (\text{Số tiền Phí thu được VNĐ} / \text{Doanh số Bảo lãnh cấp ra VNĐ}) \times 100\%$
   - **Phương án 4 (Tổng Logarit):** $Y = \ln(\text{Doanh số Bảo lãnh cấp ra VNĐ}) + \ln(\text{Số tiền Phí bảo lãnh thu được VNĐ})$

3. **Danh mục 6 Biến độc lập (X1 -> X6) chuẩn hóa trong mô hình:**
   - `TY_LE_PHI` ($X_1$): Tỷ lệ phí bảo lãnh thực tế (%)
   - `TAN_SUAT_GD` ($X_2$): Tổng số lượng thư bảo lãnh/hợp đồng thực hiện trong kỳ
   - `LOAI_BL` ($X_3$): Phân loại dòng bảo lãnh (BG, PG, TG, RG)
   - `PHAN_KHUC` ($X_4$): Phân khúc khách hàng (98 - Định chế tài chính, SMEs, KHDN Lớn)
   - `KHU_VUC` ($X_5$): Địa bàn vùng miền của Chi nhánh (KV1, KV2, KV3, KV7...)
   - `KENH_EFAST` ($X_6$): Biến giả Chuyển đổi số ($1$ = giao dịch qua eFAST, $0$ = làm tại quầy)

---

## IV. ĐÓNG GÓP HÀM Ý QUẢN TRỊ & GIẢI PHÁP CHO VIETINBANK

1. **Tối ưu hóa Biểu phí & Phân hóa giá:** Phân tích độ co giãn của phí để tìm mức phí tối ưu cho từng phân khúc; giảm phí cho khách hàng xếp hạng tín nhiệm cao và khách hàng phát hành qua eFAST.
2. **Đẩy mạnh Số hóa & Phê duyệt tự động STP:** Đề xuất nâng hạn mức duyệt tự động trên VietinBank eFAST cho các hợp đồng dưới 5 tỷ VNĐ không cần qua thẩm định thủ công.
3. **Cơ chế Ký quỹ linh hoạt:** Giảm hoặc miễn 100% ký quỹ cho bảo lãnh dự thầu (TG) với doanh nghiệp uy tín để giải phóng dòng tiền cho nhà thầu.
