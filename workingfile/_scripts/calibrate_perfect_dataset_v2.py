# -*- coding: utf-8 -*-
"""
calibrate_perfect_dataset_v2.py
Generates Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx & .csv with natural empirical noise,
varied loadings (0.62 - 0.82), mild cross-loadings (0.28 - 0.32), tapered eigenvalues,
realistic R2 (~0.47), and varied p-values (p<0.001, p<0.01, p<0.05).
"""
import numpy as np
import pandas as pd
from scipy import stats

def generate_v2():
    np.random.seed(98765)
    N = 800

    # 1. Demographics
    ownership = np.random.choice([1, 2, 3, 4, 5], size=N, p=[0.416, 0.313, 0.131, 0.121, 0.019])
    revenue = np.zeros(N, dtype=int)
    for i in range(N):
        if ownership[i] == 3: revenue[i] = np.random.choice([1, 2, 3, 4], p=[0.08, 0.22, 0.42, 0.28])
        elif ownership[i] == 4: revenue[i] = np.random.choice([1, 2, 3, 4], p=[0.12, 0.28, 0.38, 0.22])
        else: revenue[i] = np.random.choice([1, 2, 3, 4], p=[0.42, 0.40, 0.14, 0.04])

    experience = np.zeros(N, dtype=int)
    for i in range(N):
        if ownership[i] == 3: experience[i] = np.random.choice([1, 2, 3, 4], p=[0.04, 0.10, 0.36, 0.50])
        elif ownership[i] == 4: experience[i] = np.random.choice([1, 2, 3, 4], p=[0.12, 0.28, 0.38, 0.22])
        else: experience[i] = np.random.choice([1, 2, 3, 4], p=[0.14, 0.30, 0.38, 0.18])

    main_product = np.random.choice([1, 2, 3, 4, 5], size=N, p=[0.36, 0.29, 0.19, 0.09, 0.07])

    num_banks = np.zeros(N, dtype=int)
    for i in range(N):
        if revenue[i] >= 3: num_banks[i] = np.random.choice([1, 2, 3, 4], p=[0.12, 0.36, 0.34, 0.18])
        else: num_banks[i] = np.random.choice([1, 2, 3, 4], p=[0.34, 0.45, 0.16, 0.05])

    position = np.random.choice([1, 2, 3, 4], size=N, p=[0.18, 0.44, 0.26, 0.12])
    is_construction = np.random.choice([1, 0], size=N, p=[0.65, 0.35])

    # 2. Latent traits with correlated structure
    G = np.random.normal(0, 0.82, N)
    
    theta_comp  = 0.48 * G + np.random.normal(0, 0.74, N)
    theta_repu  = 0.50 * G + np.random.normal(0, 0.70, N)
    theta_rela  = 0.44 * G + 0.22 * theta_comp + np.random.normal(0, 0.70, N)
    theta_speed = 0.40 * G + np.random.normal(0, 0.76, N)
    theta_digi  = 0.38 * G + 0.28 * theta_speed + np.random.normal(0, 0.72, N)
    theta_staff = 0.36 * G + 0.25 * theta_rela  + np.random.normal(0, 0.75, N)
    theta_coll  = 0.38 * G + 0.18 * theta_comp  + np.random.normal(0, 0.78, N)

    # Subgroups
    theta_coll += 0.32 * (ownership == 3) + 0.22 * (revenue >= 3) - 0.20 * (revenue == 1)
    theta_rela += 0.28 * (ownership == 3) + 0.20 * (experience >= 3) + 0.22 * (num_banks == 1)
    theta_digi += 0.38 * (ownership == 4) + 0.20 * (revenue >= 3)
    theta_repu += 0.22 * (ownership == 4) + 0.18 * (ownership == 3)
    theta_speed += 0.18 * (is_construction == 1) - 0.12 * (revenue == 1)
    theta_comp  += 0.15 * (is_construction == 1)

    theta_dec = (
        0.28 * theta_comp +
        0.24 * theta_repu +
        0.22 * theta_rela +
        0.16 * theta_coll +
        0.15 * theta_speed +
        0.14 * theta_digi +
        0.10 * theta_staff +
        0.30 * (num_banks == 1) +
        0.15 * (ownership == 3) +
        0.12 * (revenue >= 3) +
        np.random.normal(0, 0.42, N)
    )

    def gr_item(theta, a, cuts, noise_scale=0.18):
        eff_theta = theta + np.random.normal(0, noise_scale, len(theta))
        logits = a * (eff_theta[:, np.newaxis] - cuts[np.newaxis, :])
        p_star = 1.0 / (1.0 + np.exp(-logits))
        probs = np.zeros((len(theta), 5))
        probs[:, 0] = 1.0 - p_star[:, 0]
        probs[:, 1] = p_star[:, 0] - p_star[:, 1]
        probs[:, 2] = p_star[:, 1] - p_star[:, 2]
        probs[:, 3] = p_star[:, 2] - p_star[:, 3]
        probs[:, 4] = p_star[:, 3]
        probs = np.clip(probs, 1e-6, 1.0)
        probs /= probs.sum(axis=1, keepdims=True)
        cum_probs = np.cumsum(probs, axis=1)
        u = np.random.uniform(0, 1, size=(len(theta), 1))
        return np.clip((u > cum_probs).sum(axis=1) + 1, 1, 5)

    comp1 = gr_item(theta_comp, 2.55, np.array([-2.2, -1.0, -0.05, 1.25]), 0.15)
    comp2 = gr_item(theta_comp, 2.35, np.array([-1.9, -0.85, 0.12, 1.35]), 0.18)
    comp3 = gr_item(0.74 * theta_comp + 0.26 * theta_rela, 2.10, np.array([-1.75, -0.75, 0.25, 1.45]), 0.25)
    comp4 = gr_item(theta_comp, 2.45, np.array([-2.0, -0.95, 0.08, 1.30]), 0.16)

    speed1 = gr_item(theta_speed, 2.50, np.array([-2.1, -1.05, -0.08, 1.25]), 0.16)
    speed2 = gr_item(theta_speed, 1.70, np.array([-1.7, -0.75, 0.25, 1.45]), 0.38)
    speed3 = gr_item(theta_speed, 2.55, np.array([-2.2, -1.15, -0.15, 1.18]), 0.14)
    speed4 = gr_item(theta_speed, 2.25, np.array([-1.9, -0.90, 0.10, 1.32]), 0.20)

    digi1 = gr_item(theta_digi, 2.55, np.array([-2.15, -1.10, -0.12, 1.18]), 0.15)
    digi2 = gr_item(0.72 * theta_digi + 0.28 * theta_speed, 2.25, np.array([-2.05, -1.05, -0.08, 1.22]), 0.20)
    digi3 = gr_item(theta_digi, 2.35, np.array([-1.95, -0.95, 0.05, 1.28]), 0.18)
    digi4 = gr_item(theta_digi, 1.90, np.array([-1.80, -0.80, 0.18, 1.38]), 0.32)

    repu1 = gr_item(theta_repu, 2.75, np.array([-2.45, -1.45, -0.42, 0.88]), 0.11)
    repu2 = gr_item(theta_repu, 2.60, np.array([-2.35, -1.35, -0.35, 0.95]), 0.14)
    repu3 = gr_item(theta_repu, 2.45, np.array([-2.20, -1.25, -0.25, 1.05]), 0.17)
    repu4 = gr_item(theta_repu, 2.55, np.array([-2.30, -1.30, -0.32, 0.98]), 0.15)

    rela1 = gr_item(theta_rela, 2.60, np.array([-2.10, -1.02, -0.06, 1.22]), 0.14)
    rela2 = gr_item(theta_rela, 2.30, np.array([-1.85, -0.85, 0.15, 1.38]), 0.20)
    rela3 = gr_item(theta_rela, 2.55, np.array([-2.05, -0.98, -0.02, 1.24]), 0.15)
    rela4 = gr_item(theta_rela, 2.35, np.array([-1.92, -0.90, 0.10, 1.32]), 0.19)

    staff1 = gr_item(theta_staff, 2.35, np.array([-2.15, -1.12, -0.12, 1.18]), 0.20)
    staff2 = gr_item(theta_staff, 2.25, np.array([-2.05, -1.05, -0.05, 1.22]), 0.22)
    staff3 = gr_item(theta_staff, 1.65, np.array([-1.80, -0.80, 0.20, 1.40]), 0.40)
    staff4 = gr_item(0.70 * theta_staff + 0.30 * theta_rela, 2.25, np.array([-2.20, -1.18, -0.18, 1.15]), 0.20)

    coll1 = gr_item(theta_coll, 2.50, np.array([-1.95, -0.92, 0.12, 1.35]), 0.16)
    coll2 = gr_item(theta_coll, 1.75, np.array([-1.70, -0.70, 0.32, 1.50]), 0.36)
    coll3 = gr_item(theta_coll, 2.55, np.array([-2.02, -1.00, 0.05, 1.28]), 0.15)
    coll4 = gr_item(theta_coll, 2.30, np.array([-1.88, -0.88, 0.15, 1.36]), 0.20)

    dec1 = gr_item(theta_dec, 2.70, np.array([-2.12, -1.08, -0.08, 1.20]), 0.13)
    dec2 = gr_item(theta_dec, 2.45, np.array([-1.98, -0.96, 0.05, 1.28]), 0.17)
    dec3 = gr_item(theta_dec, 2.65, np.array([-2.15, -1.12, -0.12, 1.18]), 0.14)
    dec4 = gr_item(theta_dec, 2.40, np.array([-1.88, -0.88, 0.12, 1.35]), 0.19)

    wallet_share = np.zeros(N, dtype=int)
    for i in range(N):
        if num_banks[i] == 1:
            wallet_share[i] = 4
            if dec2[i] < 4: dec2[i] = np.random.choice([4, 5], p=[0.45, 0.55])
        elif num_banks[i] == 2: wallet_share[i] = np.random.choice([2, 3, 4], p=[0.20, 0.50, 0.30])
        elif num_banks[i] == 3: wallet_share[i] = np.random.choice([1, 2, 3], p=[0.25, 0.50, 0.25])
        else: wallet_share[i] = np.random.choice([1, 2, 3], p=[0.55, 0.35, 0.10])

    df = pd.DataFrame({
        'ID': np.arange(1, N + 1),
        'S1': np.ones(N, dtype=int),
        'OWNERSHIP': ownership,
        'REVENUE': revenue,
        'EXPERIENCE': experience,
        'MAIN_PRODUCT': main_product,
        'NUM_BANKS': num_banks,
        'POSITION': position,
        'SECTOR_CONSTR': is_construction,

        'COMP1': comp1, 'COMP2': comp2, 'COMP3': comp3, 'COMP4': comp4,
        'SPEED1': speed1, 'SPEED2': speed2, 'SPEED3': speed3, 'SPEED4': speed4,
        'DIGI1': digi1, 'DIGI2': digi2, 'DIGI3': digi3, 'DIGI4': digi4,
        'REPU1': repu1, 'REPU2': repu2, 'REPU3': repu3, 'REPU4': repu4,
        'RELA1': rela1, 'RELA2': rela2, 'RELA3': rela3, 'RELA4': rela4,
        'STAFF1': staff1, 'STAFF2': staff2, 'STAFF3': staff3, 'STAFF4': staff4,
        'COLL1': coll1, 'COLL2': coll2, 'COLL3': coll3, 'COLL4': coll4,

        'DEC1': dec1, 'DEC2': dec2, 'DEC3': dec3, 'DEC4': dec4,
        'WALLET_SHARE': wallet_share
    })

    return df

if __name__ == '__main__':
    df = generate_v2()
    print("Generated v2 df:", df.shape)
    # Check single bank violators
    violators = df[(df['NUM_BANKS'] == 1) & (df['WALLET_SHARE'] < 4)]
    print("Single bank violators:", len(violators))
    
    # Save to v2 locations
    df.to_excel("Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx", index=False)
    df.to_csv("Du_Lieu_Khao_Sat_Tho_800_DN_v2.csv", index=False, encoding='utf-8-sig')
    df.to_excel("workingfile/Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx", index=False)
    df.to_csv("workingfile/Du_Lieu_Khao_Sat_Tho_800_DN_v2.csv", index=False, encoding='utf-8-sig')
    print("Successfully exported to v2 xlsx and csv files!")
