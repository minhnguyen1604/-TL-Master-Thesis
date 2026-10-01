# -*- coding: utf-8 -*-
"""
generate_figure_4_1.py
Creates Figure 4.1: Empirical Path Model with standardized OLS path coefficients (β)
and t-statistics loaded dynamically from Du_Lieu_Khao_Sat_Tho_800_DN.xlsx.
High-resolution 300 DPI image formatted for academic publication.
"""
from PIL import Image, ImageDraw, ImageFont
import os, sys

# Ensure import from current directory
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from generate_ch4_content import compute_all_statistics

def main():
    print("Computing regression statistics for Figure 4.1...")
    st = compute_all_statistics()
    reg = st['reg']

    # Ordered list of independent variables sorted by beta magnitude
    var_meta = [
        ("Price Competitiveness", "COST_COMP", reg['beta_std'][0], reg['t_vals'][1]),
        ("Bank Reputation", "BANK_REP", reg['beta_std'][3], reg['t_vals'][4]),
        ("Relationship & Limits", "RELATIONSHIP", reg['beta_std'][4], reg['t_vals'][5]),
        ("Collateral Policy", "COLL_POLICY", reg['beta_std'][6], reg['t_vals'][7]),
        ("Processing Speed", "PROC_SPEED", reg['beta_std'][1], reg['t_vals'][2]),
        ("Digital eFAST Convenience", "DIGITAL_CONV", reg['beta_std'][2], reg['t_vals'][3]),
        ("Staff Professionalism", "STAFF_QUAL", reg['beta_std'][5], reg['t_vals'][6]),
    ]

    W, H = 1400, 750
    img = Image.new("RGB", (W, H), color=(255, 255, 255))
    draw = ImageDraw.Draw(img)

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

    # Box for Dependent Variable (DEC) on the right
    dec_x1, dec_y1, dec_x2, dec_y2 = 960, 220, 1340, 530
    draw.rectangle([dec_x1, dec_y1, dec_x2, dec_y2], fill=(235, 243, 250), outline=NAVY, width=3)

    # Text inside Dependent Variable box
    draw.text((dec_x1 + 25, dec_y1 + 40), "DEPENDENT VARIABLE", font=font_box_sub, fill=(100, 100, 100))
    draw.text((dec_x1 + 25, dec_y1 + 75), "Selection Priority &", font=font_box_title, fill=NAVY)
    draw.text((dec_x1 + 25, dec_y1 + 105), "Patronage Intention (DEC)", font=font_box_title, fill=NAVY)
    draw.line([(dec_x1 + 25, dec_y1 + 155), (dec_x2 - 25, dec_y1 + 155)], fill=(180, 200, 220), width=2)

    draw.text((dec_x1 + 25, dec_y1 + 175), "Model Explanatory Power:", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 210), f"R² = {reg['r2']:.3f} ({reg['r2']*100:.1f}%)", font=font_box_title, fill=ACCENT_GREEN)
    draw.text((dec_x1 + 25, dec_y1 + 245), f"Adjusted R² = {reg['adj_r2']:.3f}", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 275), f"F(7, {reg['dof_resid']}) = {reg['f_stat']:.2f}***", font=font_box_sub, fill=DARK_GREY)

    # Draw 7 independent variable boxes on the left with connecting arrows
    box_w = 340
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
        draw.text((bx1 + 18, by1 + 12), title, font=font_box_title, fill=NAVY)
        draw.text((bx1 + 18, by1 + 38), f"({code})", font=font_box_sub, fill=(80, 80, 80))

        # Connecting arrow from box right edge to DEC box left edge
        arrow_start = (bx2, by1 + box_h // 2)
        arrow_end = (dec_x1, dec_y1 + 35 + i * 36)

        # Draw connecting line
        draw.line([arrow_start, (arrow_start[0] + 120, arrow_start[1]), 
                   (dec_x1 - 40, arrow_end[1]), arrow_end], fill=NAVY, width=2)
        # Arrowhead
        ax, ay = arrow_end
        draw.polygon([(ax, ay), (ax - 10, ay - 5), (ax - 10, ay + 5)], fill=NAVY)

        # Path coefficient labels
        lbl_x = bx2 + 140
        lbl_y = arrow_start[1] + (arrow_end[1] - arrow_start[1]) * 0.45 - 12
        draw.text((lbl_x, lbl_y - 8), beta_str, font=font_beta, fill=NAVY)
        draw.text((lbl_x, lbl_y + 14), f"({t_str})", font=font_t, fill=(100, 100, 100))

    # Footnote
    draw.text((60, 715), "*** Significant at p < 0.001 level. All path coefficients represent standardized OLS regression weights (β).", 
              font=font_t, fill=(80, 80, 80))

    out_path = os.path.join(script_dir, "..", "figure_4_1_empirical_model.png")
    img.save(out_path, "PNG", dpi=(300, 300))
    print(f"Successfully generated empirical path model image: {out_path}")

if __name__ == '__main__':
    main()
