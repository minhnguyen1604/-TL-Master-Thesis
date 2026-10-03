# -*- coding: utf-8 -*-
"""
generate_figure_4_1_v3.py
Creates Figure 4.1: Empirical Path Model with standardized OLS path coefficients (β)
and t-statistics loaded directly from workingfile/v3_empirical_results.json.
High-resolution 300 DPI image formatted for academic publication.
"""
from PIL import Image, ImageDraw, ImageFont
import os, sys, json

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    json_path = os.path.join(script_dir, "..", "v3_empirical_results.json")
    with open(json_path, 'r', encoding='utf-8') as f:
        res = json.load(f)

    reg = res['regression']['summary']
    coeffs = {c['var']: c for c in res['regression']['coefficients']}

    # Variables sorted by standardized Beta descending
    var_meta = [
        ("Price Competitiveness", "COST_COMP", coeffs['COST_COMP']['Beta'], coeffs['COST_COMP']['t_HC3']),
        ("Relationship & Limits", "RELATIONSHIP", coeffs['RELATIONSHIP']['Beta'], coeffs['RELATIONSHIP']['t_HC3']),
        ("Processing Speed", "PROC_SPEED", coeffs['PROC_SPEED']['Beta'], coeffs['PROC_SPEED']['t_HC3']),
        ("Collateral Policy", "COLL_POLICY", coeffs['COLL_POLICY']['Beta'], coeffs['COLL_POLICY']['t_HC3']),
        ("Bank Reputation", "BANK_REP", coeffs['BANK_REP']['Beta'], coeffs['BANK_REP']['t_HC3']),
        ("Staff Professionalism", "STAFF_QUAL", coeffs['STAFF_QUAL']['Beta'], coeffs['STAFF_QUAL']['t_HC3']),
        ("Digital e-Guarantee Convenience", "DIGITAL_CONV", coeffs['DIGITAL_CONV']['Beta'], coeffs['DIGITAL_CONV']['t_HC3']),
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

    draw.text((dec_x1 + 25, dec_y1 + 175), "Model Explanatory Power:", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 210), f"R² = {reg['R2']:.3f} ({reg['R2']*100:.1f}%)", font=font_box_title, fill=ACCENT_GREEN)
    draw.text((dec_x1 + 25, dec_y1 + 245), f"Adjusted R² = {reg['Adj_R2']:.3f}", font=font_box_sub, fill=DARK_GREY)
    draw.text((dec_x1 + 25, dec_y1 + 275), f"F(7, {reg['ANOVA']['df_res']}) = {reg['F']:.2f}***", font=font_box_sub, fill=DARK_GREY)

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
    draw.text((60, 715), "*** Significant at p < 0.001 level. All path coefficients represent standardized OLS regression weights (β) with HC3 robust t-statistics.", 
              font=font_t, fill=(80, 80, 80))

    out_path = os.path.join(script_dir, "..", "figure_4_1_empirical_model_v3.png")
    img.save(out_path, "PNG", dpi=(300, 300))
    # Also save to root
    img.save("figure_4_1_empirical_model_v3.png", "PNG", dpi=(300, 300))
    print("Successfully generated empirical path model image v3.")

if __name__ == '__main__':
    main()
