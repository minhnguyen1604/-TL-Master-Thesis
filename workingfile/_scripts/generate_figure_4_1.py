# -*- coding: utf-8 -*-
"""
Tao anh Figure 4.1: Empirical Path Model bang PIL
Do phan giai cao, chuan hoc thuat quoc te.
"""
from PIL import Image, ImageDraw, ImageFont
import os

W, H = 1400, 750
img = Image.new("RGB", (W, H), color=(255, 255, 255))
draw = ImageDraw.Draw(img)

# Co gang load font he thong Arial / Times New Roman
try:
    font_title = ImageFont.truetype("times.ttf", 26)
    font_box_title = ImageFont.truetype("timesbd.ttf", 18)
    font_box_sub = ImageFont.truetype("times.ttf", 16)
    font_beta = ImageFont.truetype("timesbd.ttf", 18)
    font_t = ImageFont.truetype("timesi.ttf", 14)
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

# Danh sach 7 bien doc lap
indep_vars = [
    ("Price Competitiveness", "COST_COMP", "β = +0.266***", "t = 10.26"),
    ("Relationship & Limits", "RELATIONSHIP", "β = +0.211***", "t = 8.49"),
    ("Bank Reputation", "BANK_REP", "β = +0.191***", "t = 7.49"),
    ("Processing Speed", "PROC_SPEED", "β = +0.171***", "t = 6.82"),
    ("Digital eFAST Convenience", "DIGITAL_CONV", "β = +0.167***", "t = 6.73"),
    ("Collateral Policy", "COLL_POLICY", "β = +0.137***", "t = 5.57"),
    ("Staff Professionalism", "STAFF_QUAL", "β = +0.101***", "t = 4.09"),
]

# Hop bien phu thuoc DEC ben phai
dec_x1, dec_y1, dec_x2, dec_y2 = 960, 220, 1340, 530
draw.rectangle([dec_x1, dec_y1, dec_x2, dec_y2], fill=(235, 243, 250), outline=NAVY, width=3)

# Text bien phu thuoc
draw.text((dec_x1 + 25, dec_y1 + 40), "DEPENDENT VARIABLE", font=font_box_sub, fill=(100, 100, 100))
draw.text((dec_x1 + 25, dec_y1 + 75), "Selection Priority &", font=font_box_title, fill=NAVY)
draw.text((dec_x1 + 25, dec_y1 + 105), "Patronage Intention (DEC)", font=font_box_title, fill=NAVY)
draw.line([(dec_x1 + 25, dec_y1 + 155), (dec_x2 - 25, dec_y1 + 155)], fill=(180, 200, 220), width=2)

draw.text((dec_x1 + 25, dec_y1 + 175), "Model Explanatory Power:", font=font_box_sub, fill=DARK_GREY)
draw.text((dec_x1 + 25, dec_y1 + 210), "R² = 0.596 (59.6%)", font=font_box_title, fill=ACCENT_GREEN)
draw.text((dec_x1 + 25, dec_y1 + 245), "Adjusted R² = 0.593", font=font_box_sub, fill=DARK_GREY)
draw.text((dec_x1 + 25, dec_y1 + 275), "F(7, 792) = 167.16***", font=font_box_sub, fill=DARK_GREY)

# Ve 7 hop bien doc lap ben trai va mui ten
box_w = 340
box_h = 68
start_y = 50
gap_y = 28

for i, (title, code, beta_str, t_str) in enumerate(indep_vars):
    bx1 = 60
    by1 = start_y + i * (box_h + gap_y)
    bx2 = bx1 + box_w
    by2 = by1 + box_h
    
    # Hop bien doc lap
    draw.rectangle([bx1, by1, bx2, by2], fill=LIGHT_BLUE, outline=BORDER_BLUE, width=2)
    draw.text((bx1 + 18, by1 + 12), title, font=font_box_title, fill=NAVY)
    draw.text((bx1 + 18, by1 + 38), f"({code})", font=font_box_sub, fill=(80, 80, 80))
    
    # Mui ten tu hop trai sang DEC
    arrow_start = (bx2, by1 + box_h // 2)
    arrow_end = (dec_x1, dec_y1 + 35 + i * 36)
    
    # Ve duong noi
    draw.line([arrow_start, (arrow_start[0] + 120, arrow_start[1]), 
               (dec_x1 - 40, arrow_end[1]), arrow_end], fill=NAVY, width=2)
    # Ve dau mui ten
    ax, ay = arrow_end
    draw.polygon([(ax, ay), (ax - 10, ay - 5), (ax - 10, ay + 5)], fill=NAVY)
    
    # Nhan he so Beta tren duong noi
    lbl_x = bx2 + 140
    lbl_y = arrow_start[1] + (arrow_end[1] - arrow_start[1]) * 0.45 - 12
    draw.text((lbl_x, lbl_y - 8), beta_str, font=font_beta, fill=NAVY)
    draw.text((lbl_x, lbl_y + 14), f"({t_str})", font=font_t, fill=(100, 100, 100))

# Chu thich o duoi
draw.text((60, 715), "*** Significant at p < 0.001 level. All path coefficients represent standardized OLS regression weights (β).", 
          font=font_t, fill=(80, 80, 80))

out_path = "workingfile/figure_4_1_empirical_model.png"
img.save(out_path, "PNG", dpi=(300, 300))
print(f"Da tao thanh cong anh mo hinh thuc nghiem: {out_path}")
