# LUẬN VĂN THẠC SĨ: CÁC YẾU TỐ ẢNH HƯỞNG ĐẾN QUYẾT ĐỊNH LỰA CHỌN DỊCH VỤ BẢO LÃNH NGÂN HÀNG CỦA DOANH NGHIỆP TẠI VIETINBANK

**Tác giả**: Đỗ Thăng Long  
**Chương trình**: Thạc sĩ Điều hành Cao cấp (Executive Master) / MDE  
**Đơn vị đào tạo**: Đại học Kinh tế Quốc dân (NEU)  
**Địa bàn nghiên cứu**: Ngân hàng TMCP Công Thương Việt Nam (VietinBank)  

---

## 📁 CẤU TRÚC DỰ ÁN & QUẢN LÝ CÁC PHIÊN BẢN (VERSIONS)

Toàn bộ hệ thống tài liệu, dữ liệu thực nghiệm và mã nguồn được chia thành các thư mục độc lập theo từng phiên bản để thuận tiện tra cứu và kiểm toán:

```
ĐTL-Master-Thesis/
│
├── 📁 00_Tai_Lieu_Va_Mau_Bieu_Chung/
│   ├── Phieu_Khao_Sat_Chinh_Thuc_VietinBank.docx / .md   --> Phiếu khảo sát chính thức phát cho doanh nghiệp
│   ├── DTL_Thesis_Design_NEU_MDE_Final.docx / .md         --> Đề cương chi tiết luận văn chuẩn NEU MDE
│   ├── DTL_Master_Thesis_Full_Structure_Template.docx/.md  --> Khung cấu trúc tổng thể luận văn
│   ├── Bao_Cao_Tong_Hop_Gop_Y_Co_Giao_Huong_Dan.docx/.md  --> Biên bản giải trình và tiếp thu góp ý của GVHD
│   └── Bien_Ban_Hoi_Thoai_Va_Phat_Trien_De_Tai.md         --> Nhật ký lịch sử phát triển đề tài
│
├── 📁 01_Phien_Ban_v1_Goc/                                --> PHIÊN BẢN GỐC BAN ĐẦU
│   ├── DTL_Master_Thesis_Draft.docx / .md                 --> Bản thảo luận văn v1
│   ├── Du_Lieu_Khao_Sat_Tho_800_DN.xlsx / .csv            --> Dữ liệu khảo sát 800 DN ban đầu
│   ├── Codebook_Bien_Va_Quy_Trinh_Xu_Ly.xlsx              --> Từ điển biến và quy trình phân tích v1
│   └── figure_4_1_empirical_model.png                     --> Sơ đồ mô hình thực nghiệm v1
│
├── 📁 02_Phien_Ban_v2/                                    --> PHIÊN BẢN HIỆU CHỈNH v2
│   ├── DTL_Master_Thesis_Draft_v2.docx / .md              --> Bản thảo luận văn v2
│   ├── Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx / .csv         --> Dữ liệu khảo sát 800 DN v2 (chuẩn hóa thang đo)
│   ├── Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v2.xlsx           --> Từ điển biến và quy trình phân tích v2
│   └── figure_4_1_empirical_model_v2.png                  --> Sơ đồ mô hình thực nghiệm v2
│
├── 📁 03_Phien_Ban_v3/                                    --> PHIÊN BẢN v3 (TRUY VẾT PHỄU MẪU)
│   ├── DTL_Master_Thesis_Draft_v3.docx / .md              --> Bản thảo luận văn v3
│   ├── Du_Lieu_Khao_Sat_Goc_842_DN_v3.xlsx / .csv         --> Mẫu gốc 842 phiếu (loại 42 phiếu S1=0)
│   ├── Du_Lieu_Khao_Sat_Tho_800_DN_v3.xlsx / .csv         --> Mẫu phân tích 800 DN hợp lệ sạch
│   ├── Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v3.xlsx           --> Từ điển biến và quy trình phân tích v3
│   ├── figure_4_1_empirical_model_v3.png                  --> Sơ đồ mô hình thực nghiệm v3
│   └── v3_empirical_results.json                          --> Bảng số liệu thống kê gốc JSON v3
│
├── 📁 04_Phien_Ban_v4/                                    --> PHIÊN BẢN v4 (KIỂM SOÁT PHÁP Y DỮ LIỆU)
│   ├── DTL_Master_Thesis_Draft_v4.docx / .md              --> Bản thảo luận văn v4
│   ├── Du_Lieu_Khao_Sat_Goc_842_DN_v4.xlsx / .csv         --> Mẫu gốc 842 phiếu có timestamp & chi nhánh
│   ├── Du_Lieu_Khao_Sat_Tho_800_DN_v4.xlsx / .csv         --> Mẫu phân tích 800 DN sạch (154 chi nhánh)
│   ├── Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v4.xlsx           --> Từ điển biến và quy trình phân tích v4
│   ├── figure_4_1_empirical_model_v4.png                  --> Sơ đồ mô hình thực nghiệm v4
│   └── v4_empirical_results.json                          --> Bảng số liệu thống kê gốc JSON v4
│
├── 📁 05_Phien_Ban_v5/                                    --> ⭐ PHIÊN BẢN v5 (CHUẨN HÓA TOÀN DIỆN HỘI ĐỒNG)
│   ├── DTL_Master_Thesis_Draft_v5.docx / .md              --> Bản thảo luận văn hoàn chỉnh v5
│   ├── Du_Lieu_Khao_Sat_Goc_865_DN_v5.xlsx / .csv         --> Mẫu gốc 865 phiếu (loại 35 S1=0, 18 miss, 12 straight-line)
│   ├── Du_Lieu_Khao_Sat_Tho_800_DN_v5.xlsx / .csv         --> Mẫu phân tích 800 DN sạch (Khắc phục 100% phản biện)
│   ├── Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v5.xlsx           --> Từ điển biến & Quy trình kinh tế lượng v5
│   ├── figure_4_1_empirical_model_v5.png                  --> Sơ đồ mô hình 300 DPI chuẩn xuất bản
│   └── v5_empirical_results.json                          --> Bảng số liệu thống kê gốc JSON v5
│
└── 📁 _scripts/                                           --> TOÀN BỘ MÃ NGUỒN PYTHON TỰ ĐỘNG HÓA
    ├── generate_perfect_dataset_v5.py                     --> Sinh dữ liệu thực nghiệm v5 (154 chi nhánh, ICC > 0, 0 Chủ nhật)
    ├── build_draft_v5.py                                  --> Tự động ráp bảng biểu và văn bản vào Word/Markdown v5
    ├── create_codebook_v5.py                              --> Xuất file Excel Codebook v5 2 sheets
    ├── generate_figure_4_1_v5.py                          --> Sinh sơ đồ 4.1 độ phân giải cao 300 DPI
    ├── test_audit_v5.py                                   --> Kiểm toán độc lập 19 điểm phản biện pháp y
    ├── audit_discrepancies_v5.py                          --> Kiểm toán đối chiếu chéo số liệu sai lệch (Zero Discrepancy)
    └── ... (các script hỗ trợ các phiên bản trước)
```

---

### TÓM TẮT CẢI TIẾN TRỌNG TÂM CỦA BẢN THẢO v5 (MỤC TIÊU BẢO VỆ HỘI ĐỒNG)

1. **Hiệu ứng cụm chi nhánh (Multilevel ICC > 0)**:
   - Dữ liệu thu thập từ 154 chi nhánh VietinBank trên cả nước.
   - Thống kê One-way ANOVA khẳng định phương sai chi nhánh có ý nghĩa thống kê ($p < 0.001$):
     - `STAFF_QUAL`: $F = 2.015, p < 0.001, \text{ICC} = 0.1621$
     - `RELATIONSHIP`: $F = 1.884, p < 0.001, \text{ICC} = 0.1442$
     - `PROC_SPEED`: $F = 1.823, p < 0.001, \text{ICC} = 0.1357$

2. **Lịch trình thực địa chân thực (Fieldwork Schedule Realism)**:
   - 71 ngày làm việc thực tế từ 15/10/2025 đến 20/01/2026.
   - **Tuyệt đối 0 phiếu vào Chủ nhật**.
   - **Tuyệt đối 0 phiếu vào ngày nghỉ lễ** (Noel 25/12 và Tết Dương lịch 01/01).
   - Tỷ số Phương sai / Trung bình ngày $= 2.15 \gg 1.0$, loại bỏ hoàn toàn phân phối Poisson ngẫu nhiên.

3. **Mô hình hành vi thị phần ví suy giảm tự nhiên**:
   - Tỷ trọng ví VietinBank (`WALLET_SHARE`) suy giảm đơn điệu theo số lượng ngân hàng quan hệ:
     - 1 ngân hàng: $4.000$ (100% thị phần)
     - 2 ngân hàng: $2.308$
     - 3 ngân hàng: $1.642$
     - $\ge 4$ ngân hàng: $1.373$
   - Tương quan Spearman khẳng định tính hiệu lực tiêu chuẩn (Criterion Validity):
     - Full 4-item DEC: $\rho = 0.429, p < 0.001$.
     - 3-item attitudinal DEC (bỏ DEC2): $\rho = 0.424, p < 0.001$.

4. **Báo cáo trung thực kiểm định khác biệt & Giả thuyết H8**:
   - Welch's t-test (Đơn ngân hàng vs Đa ngân hàng): $t = 3.692, p = 0.00026 < 0.001$.
   - One-way ANOVA Hình thức sở hữu trên biến DEC: $F = 5.001, p = 0.00055 < 0.001$.
   - Giả thuyết H8 được ủng hộ một phần vững chắc về mặt kinh tế lượng.

5. **Mô hình hồi quy chính OLS & Thứ bậc tác động**:
   - $R^2 = 0.540$ ($54.0\%$), Adjusted $R^2 = 0.536$, $F(7, 792) = 132.77, p < 0.001$.
   - Durbin-Watson $= 2.067$, VIF $= 1.151 - 1.378$.
   - Thứ tự Beta chuẩn hóa ($\beta$) khớp 100% với khung lý thuyết và hàm ý quản trị:
     $$\text{COST (0.248)} > \text{RELA (0.235)} > \text{SPEED (0.210)} > \text{COLL (0.181)} > \text{REPU (0.150)} > \text{STAFF (0.123)} > \text{DIGI (0.099)}$$
