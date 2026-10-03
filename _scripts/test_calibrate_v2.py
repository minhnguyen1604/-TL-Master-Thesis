# -*- coding: utf-8 -*-
"""
test_calibrate_v2.py - check EFA rotation, cross-loadings, and OLS regression
"""
import numpy as np
import pandas as pd
from scipy import stats

def varimax(Phi, gamma=1.0, q=200, tol=1e-6):
    p, k = Phi.shape
    R = np.eye(k)
    d = 0
    for i in range(q):
        d_old = d
        Lambda = np.dot(Phi, R)
        u, s, vh = np.linalg.svd(np.dot(Phi.T, np.asarray(Lambda)**3 - (gamma/p) * np.dot(Lambda, np.diag(np.diag(np.dot(Lambda.T, Lambda))))))
        R = np.dot(u, vh)
        d = np.sum(s)
        if d_old != 0 and d/d_old < 1 + tol: break
    return np.dot(Phi, R)

def test():
    np.random.seed(12345)
    N = 800

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

    G = np.random.normal(0, 0.80, N)
    
    theta_comp  = 0.45 * G + np.random.normal(0, 0.75, N)
    theta_repu  = 0.48 * G + np.random.normal(0, 0.72, N)
    theta_rela  = 0.42 * G + 0.25 * theta_comp + np.random.normal(0, 0.70, N)
    theta_speed = 0.38 * G + np.random.normal(0, 0.75, N)
    theta_digi  = 0.36 * G + 0.32 * theta_speed + np.random.normal(0, 0.72, N)
    theta_staff = 0.35 * G + 0.28 * theta_rela  + np.random.normal(0, 0.75, N)
    theta_coll  = 0.35 * G + 0.20 * theta_comp  + np.random.normal(0, 0.78, N)

    theta_coll += 0.32 * (ownership == 3) + 0.22 * (revenue >= 3) - 0.20 * (revenue == 1)
    theta_rela += 0.28 * (ownership == 3) + 0.20 * (experience >= 3) + 0.22 * (num_banks == 1)
    theta_digi += 0.38 * (ownership == 4) + 0.20 * (revenue >= 3)
    theta_repu += 0.22 * (ownership == 4) + 0.18 * (ownership == 3)
    theta_speed += 0.18 * (is_construction == 1) - 0.12 * (revenue == 1)
    theta_comp  += 0.15 * (is_construction == 1)

    theta_dec = (
        0.26 * theta_comp +
        0.23 * theta_repu +
        0.20 * theta_rela +
        0.15 * theta_coll +
        0.14 * theta_speed +
        0.13 * theta_digi +
        0.09 * theta_staff +
        0.32 * (num_banks == 1) +
        0.16 * (ownership == 3) +
        0.12 * (revenue >= 3) +
        np.random.normal(0, 0.52, N)
    )

    def gr_item(theta, a, cuts, noise_scale=0.20):
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

    comp1 = gr_item(theta_comp, 2.50, np.array([-2.2, -1.0, -0.05, 1.25]), 0.15)
    comp2 = gr_item(theta_comp, 2.30, np.array([-1.9, -0.85, 0.12, 1.35]), 0.20)
    comp3 = gr_item(0.75 * theta_comp + 0.25 * theta_rela, 2.05, np.array([-1.75, -0.75, 0.25, 1.45]), 0.28)
    comp4 = gr_item(theta_comp, 2.40, np.array([-2.0, -0.95, 0.08, 1.30]), 0.18)

    speed1 = gr_item(theta_speed, 2.45, np.array([-2.1, -1.05, -0.08, 1.25]), 0.18)
    speed2 = gr_item(theta_speed, 1.65, np.array([-1.7, -0.75, 0.25, 1.45]), 0.42)
    speed3 = gr_item(theta_speed, 2.50, np.array([-2.2, -1.15, -0.15, 1.18]), 0.15)
    speed4 = gr_item(theta_speed, 2.20, np.array([-1.9, -0.90, 0.10, 1.32]), 0.22)

    digi1 = gr_item(theta_digi, 2.50, np.array([-2.15, -1.10, -0.12, 1.18]), 0.16)
    digi2 = gr_item(0.72 * theta_digi + 0.28 * theta_speed, 2.20, np.array([-2.05, -1.05, -0.08, 1.22]), 0.22)
    digi3 = gr_item(theta_digi, 2.30, np.array([-1.95, -0.95, 0.05, 1.28]), 0.20)
    digi4 = gr_item(theta_digi, 1.85, np.array([-1.80, -0.80, 0.18, 1.38]), 0.35)

    repu1 = gr_item(theta_repu, 2.70, np.array([-2.45, -1.45, -0.42, 0.88]), 0.12)
    repu2 = gr_item(theta_repu, 2.55, np.array([-2.35, -1.35, -0.35, 0.95]), 0.15)
    repu3 = gr_item(theta_repu, 2.40, np.array([-2.20, -1.25, -0.25, 1.05]), 0.18)
    repu4 = gr_item(theta_repu, 2.50, np.array([-2.30, -1.30, -0.32, 0.98]), 0.16)

    rela1 = gr_item(theta_rela, 2.55, np.array([-2.10, -1.02, -0.06, 1.22]), 0.15)
    rela2 = gr_item(theta_rela, 2.25, np.array([-1.85, -0.85, 0.15, 1.38]), 0.22)
    rela3 = gr_item(theta_rela, 2.50, np.array([-2.05, -0.98, -0.02, 1.24]), 0.16)
    rela4 = gr_item(theta_rela, 2.30, np.array([-1.92, -0.90, 0.10, 1.32]), 0.21)

    staff1 = gr_item(theta_staff, 2.30, np.array([-2.15, -1.12, -0.12, 1.18]), 0.22)
    staff2 = gr_item(theta_staff, 2.20, np.array([-2.05, -1.05, -0.05, 1.22]), 0.24)
    staff3 = gr_item(theta_staff, 1.60, np.array([-1.80, -0.80, 0.20, 1.40]), 0.44)
    staff4 = gr_item(0.70 * theta_staff + 0.30 * theta_rela, 2.20, np.array([-2.20, -1.18, -0.18, 1.15]), 0.22)

    coll1 = gr_item(theta_coll, 2.45, np.array([-1.95, -0.92, 0.12, 1.35]), 0.18)
    coll2 = gr_item(theta_coll, 1.70, np.array([-1.70, -0.70, 0.32, 1.50]), 0.40)
    coll3 = gr_item(theta_coll, 2.50, np.array([-2.02, -1.00, 0.05, 1.28]), 0.16)
    coll4 = gr_item(theta_coll, 2.25, np.array([-1.88, -0.88, 0.15, 1.36]), 0.22)

    dec1 = gr_item(theta_dec, 2.65, np.array([-2.12, -1.08, -0.08, 1.20]), 0.14)
    dec2 = gr_item(theta_dec, 2.40, np.array([-1.98, -0.96, 0.05, 1.28]), 0.18)
    dec3 = gr_item(theta_dec, 2.60, np.array([-2.15, -1.12, -0.12, 1.18]), 0.15)
    dec4 = gr_item(theta_dec, 2.35, np.array([-1.88, -0.88, 0.12, 1.35]), 0.20)

    wallet_share = np.zeros(N, dtype=int)
    for i in range(N):
        if num_banks[i] == 1:
            wallet_share[i] = 4
            if dec2[i] < 4: dec2[i] = np.random.choice([4, 5], p=[0.45, 0.55])
        elif num_banks[i] == 2: wallet_share[i] = np.random.choice([2, 3, 4], p=[0.20, 0.50, 0.30])
        elif num_banks[i] == 3: wallet_share[i] = np.random.choice([1, 2, 3], p=[0.25, 0.50, 0.25])
        else: wallet_share[i] = np.random.choice([1, 2, 3], p=[0.55, 0.35, 0.10])

    df = pd.DataFrame({
        'COMP1': comp1, 'COMP2': comp2, 'COMP3': comp3, 'COMP4': comp4,
        'SPEED1': speed1, 'SPEED2': speed2, 'SPEED3': speed3, 'SPEED4': speed4,
        'DIGI1': digi1, 'DIGI2': digi2, 'DIGI3': digi3, 'DIGI4': digi4,
        'REPU1': repu1, 'REPU2': repu2, 'REPU3': repu3, 'REPU4': repu4,
        'RELA1': rela1, 'RELA2': rela2, 'RELA3': rela3, 'RELA4': rela4,
        'STAFF1': staff1, 'STAFF2': staff2, 'STAFF3': staff3, 'STAFF4': staff4,
        'COLL1': coll1, 'COLL2': coll2, 'COLL3': coll3, 'COLL4': coll4,
        'DEC1': dec1, 'DEC2': dec2, 'DEC3': dec3, 'DEC4': dec4,
        'WALLET_SHARE': wallet_share, 'NUM_BANKS': num_banks, 'OWNERSHIP': ownership,
        'REVENUE': revenue, 'EXPERIENCE': experience, 'MAIN_PRODUCT': main_product,
        'POSITION': position, 'SECTOR_CONSTR': is_construction, 'S1': np.ones(N, dtype=int)
    })

    c_map = {
        'COST_COMP': ['COMP1', 'COMP2', 'COMP3', 'COMP4'],
        'PROC_SPEED': ['SPEED1', 'SPEED2', 'SPEED3', 'SPEED4'],
        'DIGITAL_CONV': ['DIGI1', 'DIGI2', 'DIGI3', 'DIGI4'],
        'BANK_REP': ['REPU1', 'REPU2', 'REPU3', 'REPU4'],
        'RELATIONSHIP': ['RELA1', 'RELA2', 'RELA3', 'RELA4'],
        'STAFF_QUAL': ['STAFF1', 'STAFF2', 'STAFF3', 'STAFF4'],
        'COLL_POLICY': ['COLL1', 'COLL2', 'COLL3', 'COLL4'],
        'DEC': ['DEC1', 'DEC2', 'DEC3', 'DEC4']
    }
    for cname, items in c_map.items():
        df[cname] = df[items].mean(axis=1)

    all_28 = [c for cname in list(c_map.keys())[:7] for c in c_map[cname]]
    corr_28 = df[all_28].corr().values
    evals, evecs = np.linalg.eigh(corr_28)
    idx_sort = np.argsort(evals)[::-1]
    evals = evals[idx_sort]
    evecs = evecs[:, idx_sort]
    
    loadings_unrotated = evecs[:, :7] * np.sqrt(evals[:7])
    rot = np.abs(varimax(loadings_unrotated))

    print("--- EFA ROTATED FACTOR LOADINGS ---")
    min_load = 1.0
    max_load = 0.0
    cross_loads = []
    for j, it in enumerate(all_28):
        prim = np.max(rot[j, :])
        sec = np.sort(rot[j, :])[-2]
        min_load = min(min_load, prim)
        max_load = max(max_load, prim)
        if sec >= 0.25:
            cross_loads.append((it, round(prim, 3), round(sec, 3)))
    print(f"Primary loading range: {min_load:.3f} to {max_load:.3f}")
    print(f"Cross loadings >= 0.25: {cross_loads}")

    # Regression
    X = df[['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']]
    y = df['DEC']
    X_const = np.column_stack([np.ones(N), X.values])
    beta = np.linalg.lstsq(X_const, y.values, rcond=None)[0]
    resid = y.values - X_const @ beta
    dof = N - 8
    mse = np.sum(resid**2) / dof
    se = np.sqrt(mse * np.diag(np.linalg.inv(X_const.T @ X_const)))
    t_vals = beta / se
    p_vals = [2 * (1 - stats.t.cdf(np.abs(t), dof)) for t in t_vals]
    r2 = 1.0 - (np.sum(resid**2) / np.sum((y.values - y.mean())**2))
    
    X_std = (X - X.mean()) / X.std()
    y_std = (y - y.mean()) / y.std()
    beta_std = np.linalg.lstsq(X_std.values, y_std.values, rcond=None)[0]

    # HC3 Robust Standard Errors
    # H = X (X'X)^-1 X'
    XtX_inv = np.linalg.inv(X_const.T @ X_const)
    H_diag = np.sum((X_const @ XtX_inv) * X_const, axis=1)
    u_hc3 = resid / (1.0 - H_diag)
    omega_hc3 = np.diag(u_hc3**2)
    vcov_hc3 = XtX_inv @ (X_const.T @ omega_hc3 @ X_const) @ XtX_inv
    se_hc3 = np.sqrt(np.diag(vcov_hc3))
    t_hc3 = beta / se_hc3
    p_hc3 = [2 * (1 - stats.t.cdf(np.abs(t), dof)) for t in t_hc3]

    print(f"\nR2 = {r2:.3f}")
    var_names = ['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']
    for i, vn in enumerate(var_names):
        print(f"{vn}: Beta={beta_std[i]:.3f}, B={beta[i+1]:.3f}, SE={se[i+1]:.3f}, SE_HC3={se_hc3[i+1]:.3f}, t_HC3={t_hc3[i+1]:.2f}, p_HC3={p_hc3[i+1]:.4f}")

    # VIF
    corr_means = X.corr().values
    vifs = np.diag(np.linalg.inv(corr_means))
    print(f"VIF range: {vifs.min():.3f} to {vifs.max():.3f}")

if __name__ == '__main__':
    test()
