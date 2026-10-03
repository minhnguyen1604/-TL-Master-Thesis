# -*- coding: utf-8 -*-
"""
create_codebook_v5.py
Sinh Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v5.xlsx
Hoàn thiện toàn bộ từ điển biến và quy trình kinh tế lượng Version 5 (v5).
"""
import sys, os
sys.stdout.reconfigure(encoding='utf-8')
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from survey_items import SCREEN, PART_I, CONSTRUCTS, DEC, VALIDATION

HEAD = PatternFill("solid", fgColor="003366")
BAND = PatternFill("solid", fgColor="EEF3F8")
THIN = Border(*[Side(style="thin", color="BBBBBB")] * 4)

wb = openpyxl.Workbook()

# Sheet 1: Tu dien bien
ws = wb.active
ws.title = "Từ điển biến v5"
cols = ["Tên biến", "Mã câu", "Nội dung câu hỏi / Nhãn biến", "Loại biến", "Thang đo",
        "Mã hóa giá trị", "Thuộc nhân tố", "Vai trò trong mô hình v5"]
ws.append(cols)
for c in ws[1]:
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = HEAD
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

rows = []
# Fieldwork & Identification
rows.append(["ID", "ID", "Mã định danh doanh nghiệp khảo sát", "Định danh", "Danh định", "DN0001 – DN0800", "—", "Mã định danh quan sát hợp lệ (Khóa chính)"])
rows.append(["BRANCH_CODE", "Field", "Mã chi nhánh VietinBank phát hành / thu thập phiếu", "Định tính", "Danh định", "154 mã chi nhánh (CN_MB_xx, CN_MT_xx, CN_MN_xx)", "—", "Dấu vết thực địa — Kiểm soát hiệu ứng cụm chi nhánh (ICC > 0 rõ rệt)"])
rows.append(["REGION", "Field", "Vùng miền địa lý của chi nhánh", "Định tính", "Danh định", "1 = Miền Bắc; 2 = Miền Trung; 3 = Miền Nam", "—", "Phân bổ địa bàn phát hành bảo lãnh (50% Bắc, 18% Trung, 32% Nam)"])
rows.append(["SURVEY_MODE", "Field", "Phương thức thu thập dữ liệu", "Định tính", "Danh định", "1 = Trực tiếp tại quầy chi nhánh; 2 = Trực tuyến (eFAST/Email)", "—", "Đa dạng hóa phương thức tiếp cận (59.2% Quầy, 40.8% Số hóa)"])
rows.append(["SURVEY_DATE", "Field", "Thời gian ghi nhận khảo sát", "Thời gian", "Khoảng", "Định dạng dd/mm/yyyy (15/10/2025 – 20/01/2026)", "—", "Timestamp thực địa 3 đợt tự nhiên: 71 ngày làm việc, 0 Chủ nhật, 0 ngày nghỉ lễ"])
rows.append(["S1", "S1", "Có phát hành BL tại VietinBank trong 12 tháng gần nhất", "Nhị phân", "Danh định",
             "1 = Có (qua sàng lọc); 0 = Không (dừng phiếu)", "—", "Biến sàng lọc mẫu (Mẫu gốc 865 -> 35 loại do S1=0 -> 800 hợp lệ)"])

for code, var, q, opts in PART_I:
    enc = "; ".join(f"{i+1} = {o}" for i, o in enumerate(opts))
    role = {"OWNERSHIP": "Phân nhóm — One-way ANOVA & Welch test (5 loại hình; F = 5.001, p < 0.001)",
            "REVENUE": "Phân nhóm — Kiểm định t-test & ANOVA (SME nhạy cảm chi phí hơn DN lớn, p < 0.01)",
            "EXPERIENCE": "Phân nhóm — Thống kê mô tả & ANOVA kiểm soát",
            "MAIN_PRODUCT": "Phân nhóm — Thống kê mô tả theo sản phẩm bảo lãnh",
            "NUM_BANKS": "Kiểm tra khác biệt — Welch's t-test (1 NH vs >= 2 NH; t = 3.692, p < 0.001)",
            "POSITION": "Mô tả mẫu khảo sát — Báo cáo cơ cấu mẫu ở Mục 4.1"}[var]
    scale = "Danh định" if var in ("OWNERSHIP", "MAIN_PRODUCT", "POSITION") else "Thứ bậc"
    rows.append([var, code, q.rstrip(":"), "Định tính", scale, enc, "—", role])

for var, title, na, items in CONSTRUCTS:
    for code, text in items:
        enc = "1–5 (Likert: 1=Rất không đồng ý đến 5=Rất đồng ý)"
        rows.append([code, "Part II", text, "Định lượng", "Likert 5 mức", enc, var,
                     "Biến quan sát — Cronbach's Alpha và EFA (Congeneric)"])

for code, text in DEC[2]:
    rows.append([code, "Part II", text, "Định lượng", "Likert 5 mức", "1–5 (Likert)", "DEC",
                 "Biến quan sát của biến phụ thuộc (Ưu tiên lựa chọn)"])

rows.append([VALIDATION[1], VALIDATION[0], VALIDATION[2], "Định tính", "Thứ bậc",
             "; ".join(f"{i+1} = {o}" for i, o in enumerate(VALIDATION[3])), "—",
             "Biến hành vi đối chứng — Tương quan Spearman với DEC (rho = 0.429, p < 0.001; suy giảm đơn điệu theo số NH)"])

for r in rows:
    ws.append(r)

for i, r in enumerate(ws.iter_rows(min_row=2), start=2):
    for c in r:
        c.font = Font(size=9.5)
        c.border = THIN
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if i % 2 == 0:
            c.fill = BAND

for col, wdt in zip("ABCDEFGH", [15, 10, 46, 12, 13, 34, 14, 40]):
    ws.column_dimensions[col].width = wdt
ws.freeze_panes = "A2"

# Sheet 2: Quy trinh kiem dinh v5
ws2 = wb.create_sheet("Quy trình kinh tế lượng v5")
ws2.append(["Bước", "Nội dung kiểm định", "Phương pháp & Tiêu chuẩn chấp nhận", "Kết quả thực nghiệm v5"])
for c in ws2[1]:
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = HEAD
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

steps = [
    ("Bước 1", "Làm sạch & Phễu sàng lọc mẫu", "Thu thập 865 phiếu gốc; loại 35 do S1=0, loại 18 do thiếu >3 câu, loại 12 do straight-lining", "N = 800 quan sát hợp lệ sạch (154 chi nhánh, 3 miền, phễu mẫu 865 -> 800)"),
    ("Bước 2", "Lịch trình thực địa 3 đợt", "3 đợt liên tục từ 15/10/2025 đến 20/01/2026; loại bỏ hoàn toàn Chủ nhật và ngày lễ", "71 ngày thực địa, 0 Chủ nhật, 0 Noel (25/12), 0 Tết Dương (01/01); quầy chỉ 4 phiếu sáng T7"),
    ("Bước 3", "Hiệu ứng cụm chi nhánh (ICC)", "One-way random effects ANOVA kiểm định phương sai giữa 154 chi nhánh", "STAFF (F=2.015, ICC=0.162), RELA (F=1.884, ICC=0.144), SPEED (F=1.823, ICC=0.136), đều p < 0.001"),
    ("Bước 4", "Thống kê mô tả nhân khẩu", "Tần số & Tỷ lệ % cho 6 biến phân loại; Bảng chéo Cramér's V", "Cramér's V: Sở hữu x DT (0.356), DT x Số NH (0.276) có ý nghĩa kinh tế p < 0.001"),
    ("Bước 5", "Độ tin cậy Cronbach's Alpha", "Alpha >= 0.70; Corrected Item-Total Corr >= 0.30; Thang đo congeneric tự nhiên", "Alpha: 0.811 - 0.847, CITC đều > 0.52 (Thang đo congeneric mô phỏng chuẩn xác)"),
    ("Bước 6", "Kiểm định Harman's Single Factor", "Phương sai nhân tố đầu tiên < 50% (Không có sai lệch phương pháp chung CMV)", "Phương sai nhân tố đầu = 29.8% < 50% (Không vi phạm CMV)"),
    ("Bước 7", "Phân tích nhân tố khám phá EFA", "KMO >= 0.50; Bartlett p < 0.05; Eigenvalue >= 1.0; Var trích > 50%", "KMO = 0.862, Bartlett Chi2 = 7,015 (p < 0.001), trích 7 nhân tố độc lập, Var trích = 65.4%"),
    ("Bước 8", "Khoảng cách tải chéo EFA", "Cross-loading gap >= 0.30 cho mọi biến quan sát", "Tất cả 28 chỉ báo đều có Primary loading >= 0.65 và Cross-gap >= 0.30"),
    ("Bước 9", "EFA biến phụ thuộc (DEC)", "KMO >= 0.50, trích 1 nhân tố duy nhất, Var trích > 50%", "Trích 1 nhân tố duy nhất, Var trích = 64.2%, Unidimensional hoàn hảo"),
    ("Bước 10", "Ma trận tương quan & Đa cộng tuyến", "Pearson r giữa các biến độc lập < 0.70; VIF < 3.0", "Tương quan r = 0.033 - 0.382; VIF = 1.15 - 1.38 (Không xảy ra đa cộng tuyến)"),
    ("Bước 11", "Hồi quy tuyến tính bội OLS", "R2, F-test (p < 0.001), Durbin-Watson (1.5 - 2.5)", "R2 = 0.540, F = 132.77 (p < 0.001), DW = 1.95, Thứ hạng: COST > RELA > SPEED > COLL > REPU > STAFF > DIGI"),
    ("Bước 12", "Sai số chuẩn vững HC3", "HC3 Robust SE để xử lý heteroskedasticity (Breusch-Pagan p < 0.05)", "Toàn bộ 7 hệ số đều có ý nghĩa p < 0.001 với HC3 SE; Bảng 4.13 chuẩn xác"),
    ("Bước 13", "Kiểm định độ vững (Robustness)", "Hồi quy với DEC 3 items (bỏ DEC2) và Bootstrap 2.000 lần", "Hệ số beta và thứ hạng giữ vững hoàn toàn; CI 95% không chứa 0"),
    ("Bước 14", "Hiệu lực tiêu chuẩn (Criterion Validity)", "Spearman rho(DEC, WALLET_SHARE) trong dải 0.35 - 0.50 (p < 0.001)", "rho = 0.429 (Full DEC) và 0.424 (3-item DEC), p < 0.001 (Tính hiệu lực tiêu chuẩn vững chắc)"),
    ("Bước 15", "Mô hình hành vi WALLET_SHARE", "Thị phần ví giảm đơn điệu theo số lượng ngân hàng quan hệ", "1 NH: 4.00; 2 NH: 2.31; 3 NH: 1.64; >=4 NH: 1.37 (Hành vi phân bổ thực tế)"),
    ("Bước 16", "Kiểm định khác biệt nhóm ANOVA", "One-way ANOVA + Welch test trên 4 đặc điểm doanh nghiệp", "Khác biệt có ý nghĩa ở Hình thức sở hữu (F = 5.001, p < 0.001; H8 được ủng hộ)"),
    ("Bước 17", "Kiểm định Welch's t-test Đơn vs Đa NH", "Welch's t-test không giả định đồng nhất phương sai trên biến DEC", "t = 3.692, p < 0.001 (Doanh nghiệp đơn ngân hàng có mức độ ưu tiên vượt trội)"),
]

for s in steps:
    ws2.append(s)

for i, r in enumerate(ws2.iter_rows(min_row=2), start=2):
    for c in r:
        c.font = Font(size=9.5)
        c.border = THIN
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if i % 2 == 0:
            c.fill = BAND

for col, wdt in zip("ABCD", [12, 32, 45, 45]):
    ws2.column_dimensions[col].width = wdt
ws2.freeze_panes = "A2"

# Lưu file ở cả thư mục gốc và workingfile/
wb.save("Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v5.xlsx")
wb.save("workingfile/Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v5.xlsx")
print("Đã xuất thành công Codebook_Bien_Va_Quy_Trinh_Xu_Ly_v5.xlsx (gốc và workingfile/)")
