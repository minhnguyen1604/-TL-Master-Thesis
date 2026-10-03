# -*- coding: utf-8 -*-
"""
generate_figure_4_1_v5.py
Tạo Sơ đồ 4.1: Mô hình hồi quy thực nghiệm với các hệ số đường dẫn chuẩn hóa (Beta)
và giá trị kiểm định t-statistic cho Version 5 (v5).
Độ phân giải cao 300 DPI, chuẩn học thuật xuất bản luận văn Thạc sĩ.
"""
from PIL import Image, ImageDraw, ImageFont
import os, sys

sys.stdout.reconfigure(encoding='utf-8')

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Exact v5 regression summary and coefficients
    reg = {
        'R2': 0.540,
        'Adj_R2': 0.536,
        'F': 132.77,
        'df_res': 792
    }

    # Variables sorted by standardized Beta descending (v5 exact values)
    var_meta = [
        ("Năng lực cạnh tranh chi phí & phí BL", "COST_COMP", 0.248, 9.35),
        ("Mối quan hệ & Hạn mức tín dụng", "RELATIONSHIP", 0.235, 9.05),
        ("Tốc độ & Thời gian phát hành bảo lãnh", "PROC_SPEED", 0.210, 7.90),
        ("Chính sách tài sản bảo đảm linh hoạt", "COLL_POLICY", 0.181, 6.94),
        ("Uy tín thương hiệu & Quy mô ngân hàng", "BANK_REP", 0.150, 5.57),
        ("Trình độ & Năng lực cán bộ bảo lãnh", "STAFF_QUAL", 0.123, 4.84),
        ("Tiện ích nền tảng bảo lãnh số (eFAST)", "DIGITAL_CONV", 0.099, 3.73),
    ]

    W, H = 1400, 750
    img = Image.new("RGB", (W, H), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

    try:
        font_title = ImageFont.truetype("times.ttf", 26)
        font_box_title = ImageFont.truetype("timesbd.ttf", 17)
        font_box_sub = ImageFont.truetype("times.ttf", 15)
        font_beta = ImageFont.truetype("timesbd.ttf", 17)
        font_t = ImageFont.truetype("timesi.ttf", 13)
    except:
        font_title = ImageFont.load_default()
        font_box_title = ImageFont.load_default()
        font_box_sub = ImageFont.load_default()
        font_beta = ImageFont.load_default()
        font_t = ImageFont.load_default()

    NAVY = (0, 51, 102)
    DARK_GREY = (50, 50, 50)
    LIGHT_BLUE = (240, 245, 250)
    BORDER_BLUE = (0, 51, 102)
    ACCENT_GREEN = (0, 102, 51)
    WHITE = (255, 255, 255)

    # Box for Dependent Variable (DEC) on the right
    dec_x1, dec_y1, dec_x2, dec_y2 = 960, 220, 1340, 530
    draw.rectangle([dec_x1, dec_y1, dec_x2, dec_y2], fill=(235, 243, 250), outline=NAVY, width=3)

    # Text inside Dependent Variable box
    draw.text((dec_x1 + 25, dec_y1 + 35), "BIẾN PHỤ THUỘC (DEPENDENT)", font=font_box_sub, fill=(100, 100, 100))
    draw.text((dec_x1 + 25, dec_y1 + 65), "Quyết định ưu tiên lựa chọn", font=font_box_title, fill=NAVY)
    draw.text((dec_x1 + 25, dec_y1 + 95), "dịch vụ bảo lãnh (DEC)", font=font_box_title, fill=NAVY)
    draw.line([(dec_x1 + 25, dec_y1 + 145), (dec_x2 - 25, dec_y1 + 145)], fill=(180, 200, 220), width=2)

    draw.text((dec_x1 + 25, dec_y1 + 165), "Độ thích hợp mô hình thực nghiệm (v5):", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 200), f"R² = {reg['R2']:.3f} ({reg['R2']*100:.1f}%)", font=font_box_title, fill=ACCENT_GREEN)
    draw.text((dec_x1 + 25, dec_y1 + 235), f"R² hiệu chỉnh = {reg['Adj_R2']:.3f}", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 265), f"F(7, {reg['df_res']}) = {reg['F']:.2f}***", font=font_box_sub, fill=DARK_GREY)

    # Draw 7 independent variable boxes on the left with connecting arrows
    box_w = 345
    box_h = 68
    start_y = 50
    gap_y = 28

    for i, (title, code, beta_val, t_val) in enumerate(var_meta):
        bx1 = 60
        by1 = start_y + i * (box_h + gap_y)
        bx2 = bx1 + box_w
        by2 = by1 + box_h

        beta_str = f"β = +{beta_val:.3f}***"
        t_str = f"t = {t_val:.2f}"

        # Independent variable box
        draw.rectangle([bx1, by1, bx2, by2], fill=LIGHT_BLUE, outline=BORDER_BLUE, width=2)
        draw.text((bx1 + 16, by1 + 12), title, font=font_box_title, fill=NAVY)
        draw.text((bx1 + 16, by1 + 38), f"({code})", font=font_box_sub, fill=(80, 80, 80))

        # Connecting arrow from box right edge to DEC box left edge
        arrow_start = (bx2, by1 + box_h // 2)
        arrow_end = (dec_x1, dec_y1 + 35 + i * 36)

        # Draw connecting line
        draw.line([arrow_start, (arrow_start[0] + 120, arrow_start[1]), 
                   (dec_x1 - 40, arrow_end[1]), arrow_end], fill=NAVY, width=2)

        # Arrowhead
        ax, ay = arrow_end
        draw.polygon([(ax, ay), (ax - 10, ay - 5), (ax - 10, ay + 5)], fill=NAVY)

        # Text over arrow
        mid_x = (bx2 + dec_x1) // 2 - 35
        mid_y = (arrow_start[1] + arrow_end[1]) // 2 - 12
        draw.rectangle([mid_x - 12, mid_y - 12, mid_x + 130, mid_y + 26], fill=WHITE, outline=(220, 220, 220), width=1)
        draw.text((mid_x, mid_y - 10), beta_str, font=font_beta, fill=NAVY)
        draw.text((mid_x, mid_y + 10), f"({t_str})", font=font_t, fill=(100, 100, 100))

    # Save to workingfile and root
    out1 = os.path.join(script_dir, "..", "figure_4_1_empirical_model_v5.png")
    out2 = os.path.join(script_dir, "..", "..", "figure_4_1_empirical_model_v5.png")
    img.save(out1, dpi=(300, 300))
    img.save(out2, dpi=(300, 300))
    print(f"Sơ đồ 4.1 v5 đã được tạo thành công: {out1} và {out2}")

if __name__ == "__main__":
    main()
