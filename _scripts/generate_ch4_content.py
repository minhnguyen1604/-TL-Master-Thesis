# -*- coding: utf-8 -*-
"""
generate_ch4_content.py
Performs complete, rigorous econometric computation directly from Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx
(or fallback to Du_Lieu_Khao_Sat_Tho_800_DN.xlsx).
Calculates:
- Demographics frequencies
- Cronbach's Alpha & Item-Total Statistics
- Exploratory Factor Analysis (PCA + Varimax orthogonal rotation)
- Pearson correlation matrix & VIF
- OLS regression with both ordinary and HC3 robust standard errors
- Diagnostic tests: Breusch-Pagan, White, Jarque-Bera, Durbin-Watson
- Subgroup ANOVA with Levene test and Welch's ANOVA for heteroskedastic constructs
- Independent samples t-test and Spearman criterion validity
"""
import sys, os
import pandas as pd
import numpy as np
from scipy import stats

def compute_all_statistics(excel_file=None):
    if excel_file is None:
        candidates = [
            "Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx",
            "workingfile/Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx",
            os.path.join(os.path.dirname(__file__), "..", "Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx"),
            os.path.join(os.path.dirname(__file__), "..", "..", "Du_Lieu_Khao_Sat_Tho_800_DN_v2.xlsx"),
            "Du_Lieu_Khao_Sat_Tho_800_DN.xlsx",
            "workingfile/Du_Lieu_Khao_Sat_Tho_800_DN.xlsx",
            os.path.join(os.path.dirname(__file__), "..", "Du_Lieu_Khao_Sat_Tho_800_DN.xlsx")
        ]
        for c in candidates:
            if os.path.exists(c):
                excel_file = c
                break
    if excel_file is None or not os.path.exists(excel_file):
        raise FileNotFoundError("Could not locate survey dataset Excel file in candidates!")
    
    print(f"Loading survey data from: {excel_file}")
    df = pd.read_excel(excel_file)
    st = {}
    st['df'] = df
    st['N'] = len(df)
    st['data_file_used'] = excel_file

    # 1. CONSTRUCT MAPPINGS
    constructs = {
        'COST_COMP': ['COMP1', 'COMP2', 'COMP3', 'COMP4'],
        'PROC_SPEED': ['SPEED1', 'SPEED2', 'SPEED3', 'SPEED4'],
        'DIGITAL_CONV': ['DIGI1', 'DIGI2', 'DIGI3', 'DIGI4'],
        'BANK_REP': ['REPU1', 'REPU2', 'REPU3', 'REPU4'],
        'RELATIONSHIP': ['RELA1', 'RELA2', 'RELA3', 'RELA4'],
        'STAFF_QUAL': ['STAFF1', 'STAFF2', 'STAFF3', 'STAFF4'],
        'COLL_POLICY': ['COLL1', 'COLL2', 'COLL3', 'COLL4'],
        'DEC': ['DEC1', 'DEC2', 'DEC3', 'DEC4']
    }
    st['constructs'] = constructs

    for cname, items in constructs.items():
        df[cname] = df[items].mean(axis=1)

    # 2. DEMOGRAPHICS FREQUENCIES
    def get_freq_table(col):
        vc = df[col].value_counts().sort_index()
        pct = (vc / len(df)) * 100
        cum = pct.cumsum()
        return vc, pct, cum

    st['demo'] = {
        'ownership': get_freq_table('OWNERSHIP'),
        'revenue': get_freq_table('REVENUE'),
        'experience': get_freq_table('EXPERIENCE'),
        'product': get_freq_table('MAIN_PRODUCT'),
        'num_banks': get_freq_table('NUM_BANKS'),
        'position': get_freq_table('POSITION'),
    }

    # 3. RELIABILITY ANALYSIS (CRONBACH'S ALPHA)
    def calc_cronbach(sub):
        k = sub.shape[1]
        item_vars = sub.var(axis=0, ddof=1).sum()
        tot_var = sub.sum(axis=1).var(ddof=1)
        return (k / (k - 1)) * (1 - item_vars / tot_var)

    rel_summary = []
    item_stats = {}
    for cname, items in constructs.items():
        sub = df[items]
        tot = sub.sum(axis=1)
        alpha = calc_cronbach(sub)
        mean_val = df[cname].mean()
        std_val = df[cname].std()
        itcs = [np.corrcoef(sub[c], tot - sub[c])[0, 1] for c in items]
        
        del_alphas = []
        for c in items:
            del_sub = sub.drop(columns=[c])
            del_a = calc_cronbach(del_sub)
            del_alphas.append(del_a)
            itc_val = np.corrcoef(sub[c], tot - sub[c])[0, 1]
            item_stats[c] = {
                'mean': sub[c].mean(),
                'std': sub[c].std(),
                'itc': itc_val,
                'del_alpha': del_a
            }

        rel_summary.append({
            'construct': cname,
            'items': len(items),
            'alpha': alpha,
            'mean': mean_val,
            'std': std_val,
            'min_itc': min(itcs),
            'max_itc': max(itcs),
            'max_del_alpha': max(del_alphas)
        })

    st['rel_summary'] = rel_summary
    st['item_stats'] = item_stats

    # 4. EXPLORATORY FACTOR ANALYSIS (EFA)
    indep_vars = ['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']
    all_indep_items = [c for cname in indep_vars for c in constructs[cname]]
    corr_indep = df[all_indep_items].corr().values
    p_indep = len(all_indep_items)

    # KMO
    inv_corr = np.linalg.inv(corr_indep)
    A = np.zeros_like(corr_indep)
    for i in range(p_indep):
        for j in range(p_indep):
            A[i, j] = -inv_corr[i, j] / np.sqrt(inv_corr[i, i] * inv_corr[j, j])
    np.fill_diagonal(A, 0)
    kmo_num = np.sum(corr_indep**2) - np.sum(np.diag(corr_indep**2))
    kmo_denom = kmo_num + np.sum(A**2)
    kmo_stat = kmo_num / kmo_denom

    # Bartlett test
    det_corr = np.linalg.det(corr_indep)
    chi2_bartlett = - (len(df) - 1 - (2*p_indep + 5)/6) * np.log(det_corr)
    dof_bartlett = p_indep * (p_indep - 1) / 2
    p_bartlett = stats.chi2.sf(chi2_bartlett, dof_bartlett)

    # Eigenvalues
    evals, evecs = np.linalg.eigh(corr_indep)
    idx_sort = np.argsort(evals)[::-1]
    evals = evals[idx_sort]
    evecs = evecs[:, idx_sort]
    pct_var = (evals / p_indep) * 100
    cum_pct = np.cumsum(pct_var)

    # Varimax Rotation for 7 factors
    loadings_unrotated = evecs[:, :7] * np.sqrt(evals[:7])
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

    rot_loadings = varimax(loadings_unrotated)
    ss_rot = np.sum(rot_loadings**2, axis=0)
    order_f = np.argsort(ss_rot)[::-1]
    rot_loadings = rot_loadings[:, order_f]
    ss_rot = ss_rot[order_f]
    pct_rot = (ss_rot / p_indep) * 100
    cum_pct_rot = np.cumsum(pct_rot)

    st['efa'] = {
        'kmo': kmo_stat,
        'bartlett_chi2': chi2_bartlett,
        'bartlett_dof': dof_bartlett,
        'bartlett_p': p_bartlett,
        'evals': evals,
        'pct_var': pct_var,
        'cum_pct': cum_pct,
        'rot_loadings': np.abs(rot_loadings),
        'ss_rot': ss_rot,
        'pct_rot': pct_rot,
        'cum_pct_rot': cum_pct_rot,
        'items': all_indep_items
    }

    # EFA for DEC
    dec_items = constructs['DEC']
    corr_dec = df[dec_items].corr().values
    p_dec = len(dec_items)
    det_dec = np.linalg.det(corr_dec)
    chi2_dec = - (len(df) - 1 - (2*p_dec + 5)/6) * np.log(det_dec)
    dof_dec = p_dec * (p_dec - 1) / 2
    p_dec_bartlett = stats.chi2.sf(chi2_dec, dof_dec)

    evals_dec, evecs_dec = np.linalg.eigh(corr_dec)
    idx_d = np.argsort(evals_dec)[::-1]
    evals_dec = evals_dec[idx_d]
    loadings_dec = np.abs(evecs_dec[:, idx_d[0]] * np.sqrt(evals_dec[0]))

    st['efa_dec'] = {
        'chi2': chi2_dec,
        'dof': dof_dec,
        'p': p_dec_bartlett,
        'eigenvalue': evals_dec[0],
        'pct_var': (evals_dec[0] / p_dec) * 100,
        'loadings': loadings_dec
    }

    # 5. CORRELATION & VIF
    corr_all = df[list(constructs.keys())].corr()
    st['corr_all'] = corr_all

    corr_indep_means = df[indep_vars].corr().values
    vifs = np.diag(np.linalg.inv(corr_indep_means))
    st['vifs'] = dict(zip(indep_vars, vifs))

    # 6. OLS MULTIPLE REGRESSION WITH HC3 ROBUST STANDARD ERRORS
    X = df[indep_vars]
    y = df['DEC']
    X_const = np.column_stack([np.ones(len(df)), X.values])
    beta = np.linalg.lstsq(X_const, y.values, rcond=None)[0]
    resid = y.values - X_const @ beta
    dof = len(df) - len(indep_vars) - 1
    mse = np.sum(resid**2) / dof
    se = np.sqrt(mse * np.diag(np.linalg.inv(X_const.T @ X_const)))
    t_vals = beta / se
    p_vals = [2 * (1 - stats.t.cdf(np.abs(t), dof)) for t in t_vals]

    r2 = 1.0 - (np.sum(resid**2) / np.sum((y.values - y.mean())**2))
    adj_r2 = 1.0 - (1.0 - r2) * (len(df) - 1) / dof
    f_stat = (r2 / len(indep_vars)) / ((1 - r2) / dof)
    p_f = 1 - stats.f.cdf(f_stat, len(indep_vars), dof)

    X_std = (X - X.mean()) / X.std()
    y_std = (y - y.mean()) / y.std()
    beta_std = np.linalg.lstsq(X_std.values, y_std.values, rcond=None)[0]

    # HC3 Robust Standard Errors (MacKinnon & White, 1985)
    XtX_inv = np.linalg.inv(X_const.T @ X_const)
    H_diag = np.sum((X_const @ XtX_inv) * X_const, axis=1)
    u_hc3 = resid / (1.0 - H_diag)
    omega_hc3 = np.diag(u_hc3**2)
    vcov_hc3 = XtX_inv @ (X_const.T @ omega_hc3 @ X_const) @ XtX_inv
    se_hc3 = np.sqrt(np.diag(vcov_hc3))
    t_hc3 = beta / se_hc3
    p_hc3 = [2 * (1 - stats.t.cdf(np.abs(t), dof)) for t in t_hc3]

    # Diagnostics: Durbin-Watson, Breusch-Pagan, White test, Jarque-Bera
    dw = np.sum(np.diff(resid)**2) / np.sum(resid**2)
    sig2 = np.mean(resid**2)
    g = resid**2 / sig2 - 1.0
    reg_bp = np.linalg.lstsq(X_const, g, rcond=None)[0]
    bp_stat = 0.5 * np.sum((X_const @ reg_bp)**2)
    bp_p = stats.chi2.sf(bp_stat, len(indep_vars))

    # White test
    X_white = [np.ones(len(df))]
    for i in range(len(indep_vars)):
        X_white.append(X.values[:, i])
        X_white.append(X.values[:, i]**2)
    X_white = np.column_stack(X_white)
    reg_w = np.linalg.lstsq(X_white, resid**2, rcond=None)[0]
    r2_w = 1.0 - np.sum((resid**2 - X_white @ reg_w)**2) / np.sum((resid**2 - np.mean(resid**2))**2)
    white_stat = len(df) * r2_w
    white_p = stats.chi2.sf(white_stat, X_white.shape[1] - 1)

    # Jarque-Bera normality of residuals
    jb_stat, jb_p = stats.jarque_bera(resid)

    ss_reg = np.sum((X_const @ beta - y.mean())**2)
    ss_resid = np.sum(resid**2)
    ss_tot = np.sum((y.values - y.mean())**2)
    ms_reg = ss_reg / len(indep_vars)
    ms_resid = mse

    st['reg'] = {
        'r': np.sqrt(r2),
        'r2': r2,
        'adj_r2': adj_r2,
        'f_stat': f_stat,
        'p_f': p_f,
        'se_est': np.sqrt(mse),
        'beta': beta,
        'se': se,
        't_vals': t_vals,
        'p_vals': p_vals,
        'se_hc3': se_hc3,
        't_hc3': t_hc3,
        'p_hc3': p_hc3,
        'beta_std': beta_std,
        'dw': dw,
        'bp_stat': bp_stat,
        'bp_p': bp_p,
        'white_stat': white_stat,
        'white_p': white_p,
        'jb_stat': jb_stat,
        'jb_p': jb_p,
        'ss_reg': ss_reg,
        'ss_resid': ss_resid,
        'ss_tot': ss_tot,
        'ms_reg': ms_reg,
        'ms_resid': ms_resid,
        'dof_reg': len(indep_vars),
        'dof_resid': dof,
        'dof_tot': len(df) - 1,
        'indep_vars': indep_vars
    }

    # 7. ROBUSTNESS REGRESSION WITH MEANINGFUL CONTROL DUMMIES
    df['D_SOE'] = (df['OWNERSHIP'] == 3).astype(int)
    df['D_FDI'] = (df['OWNERSHIP'] == 4).astype(int)
    df['D_LARGE'] = (df['REVENUE'] >= 3).astype(int)
    df['D_CONSTR'] = df['SECTOR_CONSTR']
    controls = ['D_SOE', 'D_FDI', 'D_LARGE', 'D_CONSTR']
    
    X_full = df[indep_vars + controls]
    X_full_const = np.column_stack([np.ones(len(df)), X_full.values])
    beta_rob = np.linalg.lstsq(X_full_const, y.values, rcond=None)[0]
    resid_rob = y.values - X_full_const @ beta_rob
    dof_rob = len(df) - len(indep_vars) - len(controls) - 1
    mse_rob = np.sum(resid_rob**2) / dof_rob
    se_rob = np.sqrt(mse_rob * np.diag(np.linalg.inv(X_full_const.T @ X_full_const)))
    t_rob = beta_rob / se_rob
    p_rob = [2 * (1 - stats.t.cdf(np.abs(t), dof_rob)) for t in t_rob]
    r2_rob = 1.0 - (np.sum(resid_rob**2) / np.sum((y.values - y.mean())**2))
    adj_r2_rob = 1.0 - (1.0 - r2_rob) * (len(df) - 1) / dof_rob
    f_rob = (r2_rob / (len(indep_vars) + len(controls))) / ((1 - r2_rob) / dof_rob)

    st['robustness'] = {
        'r2': r2_rob,
        'adj_r2': adj_r2_rob,
        'f_stat': f_rob,
        'beta': beta_rob,
        'se': se_rob,
        't_vals': t_rob,
        'p_vals': p_rob,
        'vars': ['Const'] + indep_vars + controls
    }

    # 8. SUB-GROUP DIFFERENCE TESTS (ANOVA, LEVENE, WELCH)
    def calc_welch_anova(grps):
        k = len(grps)
        n = np.array([len(g) for g in grps], dtype=float)
        s2 = np.array([np.var(g, ddof=1) for g in grps], dtype=float)
        m = np.array([np.mean(g) for g in grps], dtype=float)
        w = n / s2
        w_tot = np.sum(w)
        m_w = np.sum(w * m) / w_tot
        f_num = np.sum(w * (m - m_w)**2) / (k - 1)
        term2 = np.sum((1.0 - w / w_tot)**2 / (n - 1.0))
        f_denom = 1.0 + 2.0 * (k - 2.0) / (k**2 - 1.0) * term2
        f_welch = f_num / f_denom
        dof1 = k - 1
        dof2 = (k**2 - 1.0) / (3.0 * term2)
        p_welch = stats.f.sf(f_welch, dof1, dof2)
        return f_welch, p_welch

    def calc_anova(group_col, exclude_val=None):
        sub_df = df[df[group_col] != exclude_val] if exclude_val is not None else df
        res = {}
        for cname in list(constructs.keys()):
            grps = [g[cname].values for _, g in sub_df.groupby(group_col)]
            f, p = stats.f_oneway(*grps)
            lev_stat, lev_p = stats.levene(*grps)
            f_welch, p_welch = calc_welch_anova(grps)
            means_by_grp = [np.mean(g) for g in grps]
            ns_by_grp = [len(g) for g in grps]
            res[cname] = {
                'means': means_by_grp,
                'ns': ns_by_grp,
                'f': f,
                'p': p,
                'lev_stat': lev_stat,
                'lev_p': lev_p,
                'f_welch': f_welch,
                'p_welch': p_welch
            }
        return res

    st['anova_ownership'] = calc_anova('OWNERSHIP', exclude_val=5)
    st['anova_revenue'] = calc_anova('REVENUE')
    st['anova_experience'] = calc_anova('EXPERIENCE')
    st['anova_product'] = calc_anova('MAIN_PRODUCT')

    # t-test single bank vs multi bank
    ttest_res = {}
    for var in ['DEC', 'DEC1', 'DEC2', 'DEC3', 'DEC4']:
        s1 = df[df['NUM_BANKS'] == 1][var]
        s2 = df[df['NUM_BANKS'] > 1][var]
        t_val, p_t = stats.ttest_ind(s1, s2)
        s_pooled = np.sqrt(((len(s1)-1)*s1.var() + (len(s2)-1)*s2.var()) / (len(s1)+len(s2)-2))
        ttest_res[var] = {
            'n1': len(s1), 'm1': s1.mean(), 'sd1': s1.std(),
            'n2': len(s2), 'm2': s2.mean(), 'sd2': s2.std(),
            't': t_val, 'p': p_t, 'd': (s1.mean() - s2.mean()) / s_pooled
        }
    st['ttest_single_multi'] = ttest_res['DEC']
    st['ttest_items'] = ttest_res

    # Criterion validity: Spearman correlation DEC vs WALLET_SHARE
    rho_ws, p_ws = stats.spearmanr(df['DEC'], df['WALLET_SHARE'])
    st['spearman_ws'] = {'rho': rho_ws, 'p': p_ws}

    return st

if __name__ == '__main__':
    st = compute_all_statistics()
    print("SUCCESS: All econometric procedures computed!")
    print(f"File used: {st['data_file_used']}")
    print(f"R-Square: {st['reg']['r2']:.3f}, F = {st['reg']['f_stat']:.2f}")
    print(f"Breusch-Pagan: {st['reg']['bp_stat']:.2f} (p = {st['reg']['bp_p']:.4e})")
    print(f"White test: {st['reg']['white_stat']:.2f} (p = {st['reg']['white_p']:.4e})")
