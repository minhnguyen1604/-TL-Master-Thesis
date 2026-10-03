# -*- coding: utf-8 -*-
"""
generate_figure_4_1_v4.py
Creates Figure 4.1: Empirical Path Model with standardized OLS path coefficients (β)
and t-statistics for Version 4.
High-resolution 300 DPI image formatted for academic publication.
"""
from PIL import Image, ImageDraw, ImageFont
import os, sys

sys.stdout.reconfigure(encoding='utf-8')

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Exact v4 regression summary and coefficients
    reg = {
        'R2': 0.549,
        'Adj_R2': 0.545,
        'F': 137.97,
        'df_res': 792
    }

    # Variables sorted by standardized Beta descending (v4 exact values)
    var_meta = [
        ("Price Competitiveness", "COST_COMP", 0.237, 9.08),
        ("Relationship & Limits", "RELATIONSHIP", 0.219, 7.88),
        ("Processing Speed", "PROC_SPEED", 0.190, 7.37),
        ("Collateral Policy", "COLL_POLICY", 0.169, 6.34),
        ("Bank Reputation", "BANK_REP", 0.159, 6.16),
        ("Staff Professionalism", "STAFF_QUAL", 0.153, 6.19),
        ("Digital e-Guarantee Convenience", "DIGITAL_CONV", 0.115, 4.58),
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
    draw.text((dec_x1 + 25, dec_y1 + 75), "Bank Guarantee Selection", font=font_box_title, fill=NAVY)
    draw.text((dec_x1 + 25, dec_y1 + 105), "& Patronage Intention (DEC)", font=font_box_title, fill=NAVY)
    draw.line([(dec_x1 + 25, dec_y1 + 155), (dec_x2 - 25, dec_y1 + 155)], fill=(180, 200, 220), width=2)

    draw.text((dec_x1 + 25, dec_y1 + 175), "Model Explanatory Power (v4):", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 210), f"R² = {reg['R2']:.3f} ({reg['R2']*100:.1f}%)", font=font_box_title, fill=ACCENT_GREEN)
    draw.text((dec_x1 + 25, dec_y1 + 245), f"Adjusted R² = {reg['Adj_R2']:.3f}", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 275), f"F(7, {reg['df_res']}) = {reg['F']:.2f}***", font=font_box_sub, fill=DARK_GREY)

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

        # Text over arrow
        mid_x = (bx2 + dec_x1) // 2 - 35
        mid_y = (arrow_start[1] + arrow_end[1]) // 2 - 12
        draw.rectangle([mid_x - 12, mid_y - 12, mid_x + 130, mid_y + 26], fill=WHITE, outline=(220, 220, 220), width=1)
        draw.text((mid_x, mid_y - 10), beta_str, font=font_beta, fill=NAVY)
        draw.text((mid_x, mid_y + 10), f"({t_str})", font=font_t, fill=(100, 100, 100))

    # Save to workingfile and root
    out1 = os.path.join(script_dir, "..", "figure_4_1_empirical_model_v4.png")
    out2 = os.path.join(script_dir, "..", "..", "figure_4_1_empirical_model_v4.png")
    img.save(out1, dpi=(300, 300))
    img.save(out2, dpi=(300, 300))
    print(f"Figure 4.1 v4 saved successfully: {out1}")

if __name__ == "__main__":
    main()
