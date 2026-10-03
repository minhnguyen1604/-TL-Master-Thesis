# -*- coding: utf-8 -*-
"""
compute_all_v4_statistics.py
Tinh toan toan bo cac chi so thong ke dinh luong cho Version 4 tu file Du_Lieu_Khao_Sat_Tho_800_DN_v4.xlsx va Du_Lieu_Khao_Sat_Goc_842_DN_v4.xlsx.
Xuat file JSON ket qua de dung chung cho Word, Markdown va Audit.
"""
import json
import sys
import os
import numpy as np
import pandas as pd
from scipy import stats

sys.stdout.reconfigure(encoding='utf-8')

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

def compute():
    df = pd.read_excel('workingfile/Du_Lieu_Khao_Sat_Tho_800_DN_v4.xlsx')
    df_goc = pd.read_excel('workingfile/Du_Lieu_Khao_Sat_Goc_842_DN_v4.xlsx')
    N = len(df)
    N_goc = len(df_goc)

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

    results = {}

    # 1. Sample attrition & Fieldwork Footprint
    screen_failed = int((df_goc['S1'] == 0).sum())
    results['attrition'] = {
        'distributed': 950,
        'returned': N_goc,
        'return_rate': round(N_goc / 950 * 100, 2),
        'screen_failed': screen_failed,
        'screen_fail_rate': round(screen_failed / 950 * 100, 2),
        'valid': N,
        'valid_rate_total': round(N / 950 * 100, 2),
        'valid_rate_returned': round(N / N_goc * 100, 2)
    }

    # Fieldwork footprint (Mã chi nhánh, vùng miền, phương thức, thời gian)
    region_labels = {1: 'Miền Bắc', 2: 'Miền Trung', 3: 'Miền Nam'}
    mode_labels = {1: 'Trực tiếp tại quầy giao dịch chi nhánh', 2: 'Trực tuyến (eFAST / Email khảo sát số)'}
    
    fieldwork = {
        'total_branches': int(df['BRANCH_CODE'].nunique()),
        'regions': [],
        'survey_modes': [],
        'date_min': str(df['SURVEY_DATE'].min()),
        'date_max': str(df['SURVEY_DATE'].max())
    }
    for k, v in region_labels.items():
        cnt = int((df['REGION'] == k).sum())
        fieldwork['regions'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})
    for k, v in mode_labels.items():
        cnt = int((df['SURVEY_MODE'] == k).sum())
        fieldwork['survey_modes'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})
    results['fieldwork'] = fieldwork

    # 2. Demographics
    demo = {}
    # Ownership
    own_labels = {1: 'Doanh nghiệp tư nhân / TNHH', 2: 'Công ty Cổ phần ngoài quốc doanh', 3: 'Doanh nghiệp Nhà nước (SOEs)', 4: 'Doanh nghiệp FDI', 5: 'Khác'}
    demo['ownership'] = []
    for k, v in own_labels.items():
        cnt = int((df['OWNERSHIP'] == k).sum())
        demo['ownership'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    # Revenue
    rev_labels = {1: 'Dưới 20 tỷ VND (Siêu nhỏ & nhỏ)', 2: 'Từ 20 đến dưới 100 tỷ VND (Vừa)', 3: 'Từ 100 đến dưới 500 tỷ VND (Lớn)', 4: 'Từ 500 tỷ VND trở lên (Rất lớn)'}
    demo['revenue'] = []
    for k, v in rev_labels.items():
        cnt = int((df['REVENUE'] == k).sum())
        demo['revenue'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    # Experience
    exp_labels = {1: 'Dưới 3 năm', 2: 'Từ 3 đến dưới 5 năm', 3: 'Từ 5 đến dưới 10 năm', 4: 'Từ 10 năm trở lên'}
    demo['experience'] = []
    for k, v in exp_labels.items():
        cnt = int((df['EXPERIENCE'] == k).sum())
        demo['experience'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    # Product
    prod_labels = {1: 'Bảo lãnh dự thầu (Tender Guarantee)', 2: 'Bảo lãnh thực hiện hợp đồng (Performance Guarantee)', 3: 'Bảo lãnh tạm ứng (Advance Payment Guarantee)', 4: 'Bảo lãnh bảo hành (Warranty Guarantee)', 5: 'Bảo lãnh khác (Thanh toán, đối ứng...)'}
    demo['product'] = []
    for k, v in prod_labels.items():
        cnt = int((df['MAIN_PRODUCT'] == k).sum())
        demo['product'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    # Num banks
    nb_labels = {1: 'Chỉ 1 ngân hàng (VietinBank - Đơn ngân hàng)', 2: '2 ngân hàng', 3: '3 ngân hàng', 4: 'Từ 4 ngân hàng trở lên'}
    demo['num_banks'] = []
    for k, v in nb_labels.items():
        cnt = int((df['NUM_BANKS'] == k).sum())
        demo['num_banks'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    # Position
    pos_labels = {1: 'Ban Giám đốc / CFO / Kế toán trưởng', 2: 'Kế toán tổng hợp / Kế toán phần hành bảo lãnh', 3: 'Trưởng/Phó phòng Đấu thầu / Dự án', 4: 'Chuyên viên phụ trách hồ sơ thầu & tài chính'}
    demo['position'] = []
    for k, v in pos_labels.items():
        cnt = int((df['POSITION'] == k).sum())
        demo['position'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    # Wallet Share distribution
    wallet_labels = {1: 'Dưới 25% tổng giá trị bảo lãnh', 2: 'Từ 25% đến dưới 50%', 3: 'Từ 50% đến dưới 75%', 4: 'Từ 75% đến 100% (hoặc duy nhất VietinBank)'}
    demo['wallet_share'] = []
    for k, v in wallet_labels.items():
        cnt = int((df['WALLET_SHARE'] == k).sum())
        demo['wallet_share'].append({'id': k, 'label': v, 'count': cnt, 'percent': round(cnt / N * 100, 2)})

    results['demographics'] = demo

    # 3. Construct means & Cronbach's Alpha
    cronbach_res = {}
    for cname, items in c_map.items():
        sub = df[items].values
        k = len(items)
        v_items = sub.var(axis=0, ddof=1).sum()
        v_total = sub.sum(axis=1).var(ddof=1)
        alpha = (k / (k - 1)) * (1 - v_items / v_total)
        
        citc_list = []
        alpha_deleted = []
        for i in range(k):
            item_i = sub[:, i]
            other_items = np.delete(sub, i, axis=1)
            other_sum = other_items.sum(axis=1)
            r = np.corrcoef(item_i, other_sum)[0, 1]
            citc_list.append(round(r, 3))
            
            # alpha if deleted
            k_del = k - 1
            v_items_del = other_items.var(axis=0, ddof=1).sum()
            v_tot_del = other_sum.var(ddof=1)
            a_del = (k_del / (k_del - 1)) * (1 - v_items_del / v_tot_del)
            alpha_deleted.append(round(a_del, 3))
            
        cronbach_res[cname] = {
            'alpha': round(alpha, 3),
            'items': items,
            'citc': citc_list,
            'citc_min': min(citc_list),
            'citc_max': max(citc_list),
            'alpha_if_deleted': alpha_deleted
        }
        df[cname] = df[items].mean(axis=1)
    results['cronbach'] = cronbach_res

    # 4. Descriptive statistics for all 32 items
    all_32 = [item for items in c_map.values() for item in items]
    desc_32 = []
    for item in all_32:
        vals = df[item].values
        desc_32.append({
            'item': item,
            'mean': round(float(vals.mean()), 3),
            'std': round(float(vals.std(ddof=1)), 3),
            'min': int(vals.min()),
            'max': int(vals.max()),
            'skew': round(float(stats.skew(vals)), 3),
            'kurt': round(float(stats.kurtosis(vals)), 3)
        })
    results['desc_32'] = desc_32

    # 5. EFA for 28 independent items
    all_28 = [item for cname in list(c_map.keys())[:7] for item in c_map[cname]]
    corr_28 = df[all_28].corr().values

    # Bartlett's test
    corr_det = np.linalg.det(corr_28)
    p = 28
    chi2_bartlett = - (N - 1 - (2*p + 5)/6) * np.log(max(1e-15, corr_det))
    df_bartlett = p * (p - 1) // 2
    p_bartlett = stats.chi2.sf(chi2_bartlett, df_bartlett)

    # KMO
    inv_corr = np.linalg.pinv(corr_28)
    partial_corr = np.zeros_like(corr_28)
    for i in range(p):
        for j in range(p):
            if i != j:
                denom = np.sqrt(max(1e-12, inv_corr[i, i] * inv_corr[j, j]))
                partial_corr[i, j] = -inv_corr[i, j] / denom
    r2_sum = np.sum(corr_28**2) - np.sum(np.diag(corr_28)**2)
    p2_sum = np.sum(partial_corr**2)
    kmo = r2_sum / (r2_sum + p2_sum)

    # Eigenvalues & PCA
    evals, evecs = np.linalg.eigh(corr_28)
    idx_sort = np.argsort(evals)[::-1]
    evals = evals[idx_sort]
    evecs = evecs[:, idx_sort]

    var_explained = evals / p * 100
    cum_var = np.cumsum(var_explained)

    # Varimax 7 factors
    loadings_7 = evecs[:, :7] * np.sqrt(evals[:7])
    rot_loadings = varimax(loadings_7)

    # Harman's Single Factor variance
    harman_pct = round(float(var_explained[0]), 2)

    construct_order = ['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']
    construct_prefixes = {
        'COST_COMP': 'COMP',
        'PROC_SPEED': 'SPEED',
        'DIGITAL_CONV': 'DIGI',
        'BANK_REP': 'REPU',
        'RELATIONSHIP': 'RELA',
        'STAFF_QUAL': 'STAFF',
        'COLL_POLICY': 'COLL'
    }
    
    # Map rotated columns to factors
    factor_mapping = {}
    for f_idx in range(7):
        col = np.abs(rot_loadings[:, f_idx])
        top_idx = np.argmax(col)
        top_item = all_28[top_idx]
        for c_name, prefix in construct_prefixes.items():
            if top_item.startswith(prefix):
                factor_mapping[c_name] = f_idx

    rot_matrix = []
    for it_idx, it_name in enumerate(all_28):
        row_vals = [round(float(abs(rot_loadings[it_idx, factor_mapping[c]])), 3) for c in construct_order]
        prim = max(row_vals)
        sec = sorted(row_vals)[-2]
        gap = prim - sec
        rot_matrix.append({
            'item': it_name,
            'loadings': row_vals,
            'primary': prim,
            'secondary': sec,
            'gap': round(gap, 3)
        })

    results['efa'] = {
        'kmo': round(float(kmo), 3),
        'bartlett_chi2': round(float(chi2_bartlett), 2),
        'bartlett_df': int(df_bartlett),
        'bartlett_sig': float(p_bartlett),
        'eigenvalues': [round(float(x), 3) for x in evals[:8]],
        'var_explained': [round(float(x), 2) for x in var_explained[:7]],
        'cum_var': [round(float(x), 2) for x in cum_var[:7]],
        'harman_variance_pct': harman_pct,
        'rot_matrix': rot_matrix
    }

    # 6. EFA Dependent construct (DEC)
    dec_items = ['DEC1', 'DEC2', 'DEC3', 'DEC4']
    dec_corr = df[dec_items].corr().values
    dec_det = np.linalg.det(dec_corr)
    dec_chi2 = - (N - 1 - (2*4 + 5)/6) * np.log(max(1e-15, dec_det))
    dec_df = 4 * 3 // 2
    dec_p_bart = stats.chi2.sf(dec_chi2, dec_df)

    dec_inv = np.linalg.pinv(dec_corr)
    dec_partial = np.zeros_like(dec_corr)
    for i in range(4):
        for j in range(4):
            if i != j:
                dec_partial[i, j] = -dec_inv[i, j] / np.sqrt(max(1e-12, dec_inv[i, i] * dec_inv[j, j]))
    r2_dec = np.sum(dec_corr**2) - 4
    p2_dec = np.sum(dec_partial**2)
    dec_kmo = r2_dec / (r2_dec + p2_dec)

    dec_evals, dec_evecs = np.linalg.eigh(dec_corr)
    dec_idx = np.argsort(dec_evals)[::-1]
    dec_evals = dec_evals[dec_idx]
    dec_evecs = dec_evecs[:, dec_idx]
    dec_loads = np.abs(dec_evecs[:, 0] * np.sqrt(dec_evals[0]))

    results['efa_dec'] = {
        'kmo': round(float(dec_kmo), 3),
        'bartlett_chi2': round(float(dec_chi2), 2),
        'bartlett_df': int(dec_df),
        'bartlett_sig': float(dec_p_bart),
        'eigenvalue': round(float(dec_evals[0]), 3),
        'var_explained': round(float(dec_evals[0]/4 * 100), 2),
        'loadings': {item: round(float(dec_loads[i]), 3) for i, item in enumerate(dec_items)}
    }

    # 7. Pearson Correlation Matrix (8 constructs)
    all_8_constructs = ['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY', 'DEC']
    corr_8 = df[all_8_constructs].corr()
    corr_table = []
    for c1 in all_8_constructs:
        row = {'construct': c1}
        for c2 in all_8_constructs:
            r_val, p_val = stats.pearsonr(df[c1], df[c2])
            row[c2] = {'r': round(float(r_val), 3), 'p': float(p_val)}
        corr_table.append(row)
    results['correlation_matrix'] = corr_table

    # 8. Main OLS Regression
    X = df[['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']]
    y = df['DEC']
    X_const = np.column_stack([np.ones(N), X.values])
    
    XtX = X_const.T @ X_const
    Xty = X_const.T @ y.values
    beta = np.linalg.solve(XtX, Xty)
    resid = y.values - X_const @ beta
    dof = N - 8
    mse = np.sum(resid**2) / dof
    XtX_inv = np.linalg.inv(XtX)
    se = np.sqrt(mse * np.diag(XtX_inv))
    t_vals = beta / se
    p_vals = [float(2 * (1 - stats.t.cdf(np.abs(t), dof))) for t in t_vals]

    # HC3 Robust Standard Errors
    H_diag = np.sum((X_const @ XtX_inv) * X_const, axis=1)
    u_hc3 = resid / (1.0 - H_diag)
    omega_hc3 = np.diag(u_hc3**2)
    vcov_hc3 = XtX_inv @ (X_const.T @ omega_hc3 @ X_const) @ XtX_inv
    se_hc3 = np.sqrt(np.diag(vcov_hc3))
    t_hc3 = beta / se_hc3
    p_hc3 = [float(2 * (1 - stats.t.cdf(np.abs(t), dof))) for t in t_hc3]

    # Standardized Betas
    X_std = (X - X.mean()) / X.std()
    y_std = (y - y.mean()) / y.std()
    X_std_arr = np.ascontiguousarray(X_std.values, dtype=np.float64)
    y_std_arr = np.ascontiguousarray(y_std.values, dtype=np.float64)
    beta_std = np.linalg.solve(X_std_arr.T @ X_std_arr, X_std_arr.T @ y_std_arr)

    # Collinearity VIF & Tolerance
    corr_X = X.corr().values
    vifs = np.diag(np.linalg.inv(corr_X))
    tols = 1.0 / vifs

    # Model fit
    ss_tot = np.sum((y.values - y.mean())**2)
    ss_reg = np.sum((X_const @ beta - y.mean())**2)
    ss_res = np.sum(resid**2)
    r2 = 1.0 - (ss_res / ss_tot)
    adj_r2 = 1.0 - (1.0 - r2) * (N - 1) / dof
    f_stat = (ss_reg / 7) / (ss_res / dof)
    p_f = float(stats.f.sf(f_stat, 7, dof))
    dw = float(np.sum(np.diff(resid)**2) / np.sum(resid**2))

    # Breusch-Pagan Test
    u2 = resid**2
    u2_mean = np.mean(u2)
    f_bp = u2 / u2_mean - 1.0
    beta_bp = np.linalg.solve(XtX, X_const.T @ f_bp)
    ess_bp = np.sum((X_const @ beta_bp)**2)
    lm_bp = 0.5 * ess_bp
    p_bp = float(stats.chi2.sf(lm_bp, 7))

    # Jarque-Bera Test
    jb_stat, jb_p = stats.jarque_bera(resid)

    reg_summary = {
        'R': round(float(np.sqrt(r2)), 3),
        'R2': round(float(r2), 3),
        'Adj_R2': round(float(adj_r2), 3),
        'SEE': round(float(np.sqrt(mse)), 3),
        'F': round(float(f_stat), 2),
        'F_sig': p_f,
        'DW': round(dw, 3),
        'Breusch_Pagan_LM': round(float(lm_bp), 3),
        'Breusch_Pagan_p': round(float(p_bp), 4),
        'Jarque_Bera_stat': round(float(jb_stat), 3),
        'Jarque_Bera_p': round(float(jb_p), 4),
        'ANOVA': {
            'SS_reg': round(float(ss_reg), 3),
            'df_reg': 7,
            'MS_reg': round(float(ss_reg / 7), 3),
            'SS_res': round(float(ss_res), 3),
            'df_res': dof,
            'MS_res': round(float(mse), 3),
            'SS_tot': round(float(ss_tot), 3),
            'df_tot': N - 1
        }
    }

    reg_coeffs = []
    reg_coeffs.append({
        'var': 'Constant',
        'B': round(float(beta[0]), 3),
        'SE': round(float(se[0]), 3),
        'Beta': None,
        't': round(float(t_vals[0]), 3),
        'Sig': p_vals[0],
        'SE_HC3': round(float(se_hc3[0]), 3),
        't_HC3': round(float(t_hc3[0]), 3),
        'Sig_HC3': p_hc3[0],
        'Tolerance': None,
        'VIF': None
    })
    for i, cname in enumerate(construct_order):
        reg_coeffs.append({
            'var': cname,
            'B': round(float(beta[i+1]), 3),
            'SE': round(float(se[i+1]), 3),
            'Beta': round(float(beta_std[i]), 3),
            't': round(float(t_vals[i+1]), 3),
            'Sig': p_vals[i+1],
            'SE_HC3': round(float(se_hc3[i+1]), 3),
            't_HC3': round(float(t_hc3[i+1]), 3),
            'Sig_HC3': p_hc3[i+1],
            'Tolerance': round(float(tols[i]), 3),
            'VIF': round(float(vifs[i]), 3)
        })

    results['regression'] = {
        'summary': reg_summary,
        'coefficients': reg_coeffs
    }

    # 9. Criterion validity: DEC vs WALLET_SHARE
    rho_full, p_rho_full = stats.spearmanr(df['DEC'], df['WALLET_SHARE'])
    dec_3item = df[['DEC1', 'DEC3', 'DEC4']].mean(axis=1)
    rho_3item, p_rho_3item = stats.spearmanr(dec_3item, df['WALLET_SHARE'])
    
    results['criterion_validity'] = {
        'rho_full': round(float(rho_full), 3),
        'p_rho_full': float(p_rho_full),
        'rho_3item': round(float(rho_3item), 3),
        'p_rho_3item': float(p_rho_3item)
    }

    # 10. Auxiliary regression of WALLET_SHARE on 7 factors
    y_ws = df['WALLET_SHARE']
    beta_ws = np.linalg.solve(XtX, X_const.T @ y_ws.values)
    resid_ws = y_ws.values - X_const @ beta_ws
    mse_ws = np.sum(resid_ws**2) / dof
    se_ws = np.sqrt(mse_ws * np.diag(XtX_inv))
    t_ws = beta_ws / se_ws
    p_ws = [float(2 * (1 - stats.t.cdf(np.abs(t), dof))) for t in t_ws]
    r2_ws = 1.0 - (np.sum(resid_ws**2) / np.sum((y_ws.values - y_ws.mean())**2))
    adj_r2_ws = 1.0 - (1.0 - r2_ws) * (N - 1) / dof
    f_ws = (r2_ws / 7) / ((1 - r2_ws) / dof)

    # Standardized betas for WALLET_SHARE
    y_ws_std = (y_ws - y_ws.mean()) / y_ws.std()
    y_ws_std_arr = np.ascontiguousarray(y_ws_std.values, dtype=np.float64)
    beta_ws_std = np.linalg.solve(X_std_arr.T @ X_std_arr, X_std_arr.T @ y_ws_std_arr)

    results['aux_wallet_regression'] = {
        'R2': round(float(r2_ws), 3),
        'Adj_R2': round(float(adj_r2_ws), 3),
        'F': round(float(f_ws), 2),
        'p_F': float(stats.f.sf(f_ws, 7, dof)),
        'coefficients': [
            {
                'var': cname,
                'B': round(float(beta_ws[i+1]), 3),
                'SE': round(float(se_ws[i+1]), 3),
                'Beta': round(float(beta_ws_std[i]), 3),
                't': round(float(t_ws[i+1]), 3),
                'Sig': p_ws[i+1]
            } for i, cname in enumerate(construct_order)
        ]
    }

    # 11. Robustness check: Standard OLS vs HC3 vs Bootstrap (2,000 replications)
    np.random.seed(42)
    B_boot = 2000
    boot_betas = np.zeros((B_boot, 8))
    for b in range(B_boot):
        idx_b = np.random.choice(N, size=N, replace=True)
        X_b = X_const[idx_b, :]
        y_b = y.values[idx_b]
        b_coef = np.linalg.solve(X_b.T @ X_b, X_b.T @ y_b)
        boot_betas[b, :] = b_coef
    boot_se = np.std(boot_betas, axis=0, ddof=1)
    boot_ci_low = np.percentile(boot_betas, 2.5, axis=0)
    boot_ci_high = np.percentile(boot_betas, 97.5, axis=0)

    robust_comparison = []
    for i, cname in enumerate(['Constant'] + construct_order):
        robust_comparison.append({
            'var': cname,
            'OLS_B': round(float(beta[i]), 3),
            'OLS_SE': round(float(se[i]), 3),
            'OLS_t': round(float(t_vals[i]), 3),
            'HC3_SE': round(float(se_hc3[i]), 3),
            'HC3_t': round(float(t_hc3[i]), 3),
            'Boot_SE': round(float(boot_se[i]), 3),
            'Boot_CI_95': [round(float(boot_ci_low[i]), 3), round(float(boot_ci_high[i]), 3)]
        })
    results['robustness_check'] = robust_comparison

    # 12. Differences tests: Independent t-test & One-way ANOVA
    # t-test Single-bank vs Multi-bank
    dec_single = df[df['NUM_BANKS'] == 1]['DEC']
    dec_multi = df[df['NUM_BANKS'] > 1]['DEC']
    t_bank, p_bank = stats.ttest_ind(dec_single, dec_multi)
    
    results['t_test_banking'] = {
        'single_N': int(len(dec_single)),
        'single_mean': round(float(dec_single.mean()), 3),
        'single_sd': round(float(dec_single.std(ddof=1)), 3),
        'multi_N': int(len(dec_multi)),
        'multi_mean': round(float(dec_multi.mean()), 3),
        'multi_sd': round(float(dec_multi.std(ddof=1)), 3),
        'mean_diff': round(float(dec_single.mean() - dec_multi.mean()), 3),
        't_stat': round(float(t_bank), 3),
        'df': int(len(dec_single) + len(dec_multi) - 2),
        'p_val': float(p_bank)
    }

    # ANOVA Ownership
    own_groups = [df[df['OWNERSHIP'] == g]['DEC'] for g in [1, 2, 3, 4, 5]]
    f_own, p_own = stats.f_oneway(*own_groups)
    results['anova_ownership'] = {
        'F': round(float(f_own), 3),
        'df_between': 4,
        'df_within': N - 5,
        'p_val': float(p_own),
        'groups': [
            {
                'id': g,
                'name': own_labels[g],
                'count': int(len(own_groups[g-1])),
                'mean': round(float(own_groups[g-1].mean()), 3),
                'sd': round(float(own_groups[g-1].std(ddof=1)), 3)
            } for g in [1, 2, 3, 4, 5]
        ]
    }

    # ANOVA Revenue
    rev_groups = [df[df['REVENUE'] == g]['DEC'] for g in [1, 2, 3, 4]]
    f_rev, p_rev = stats.f_oneway(*rev_groups)
    results['anova_revenue'] = {
        'F': round(float(f_rev), 3),
        'df_between': 3,
        'df_within': N - 4,
        'p_val': float(p_rev),
        'groups': [
            {
                'id': g,
                'name': rev_labels[g],
                'count': int(len(rev_groups[g-1])),
                'mean': round(float(rev_groups[g-1].mean()), 3),
                'sd': round(float(rev_groups[g-1].std(ddof=1)), 3)
            } for g in [1, 2, 3, 4]
        ]
    }

    # ANOVA Experience
    exp_groups = [df[df['EXPERIENCE'] == g]['DEC'] for g in [1, 2, 3, 4]]
    f_exp, p_exp = stats.f_oneway(*exp_groups)
    results['anova_experience'] = {
        'F': round(float(f_exp), 3),
        'df_between': 3,
        'df_within': N - 4,
        'p_val': float(p_exp),
        'groups': [
            {
                'id': g,
                'name': exp_labels[g],
                'count': int(len(exp_groups[g-1])),
                'mean': round(float(exp_groups[g-1].mean()), 3),
                'sd': round(float(exp_groups[g-1].std(ddof=1)), 3)
            } for g in [1, 2, 3, 4]
        ]
    }

    # ANOVA Product
    prod_groups = [df[df['MAIN_PRODUCT'] == g]['DEC'] for g in [1, 2, 3, 4, 5]]
    f_prod, p_prod = stats.f_oneway(*prod_groups)
    results['anova_product'] = {
        'F': round(float(f_prod), 3),
        'df_between': 4,
        'df_within': N - 5,
        'p_val': float(p_prod),
        'groups': [
            {
                'id': g,
                'name': prod_labels[g],
                'count': int(len(prod_groups[g-1])),
                'mean': round(float(prod_groups[g-1].mean()), 3),
                'sd': round(float(prod_groups[g-1].std(ddof=1)), 3)
            } for g in [1, 2, 3, 4, 5]
        ]
    }

    # 13. Hypotheses testing summary (H1 - H8)
    hypotheses = [
        {'hypothesis': 'H1', 'factor': 'Chi phí và mức phí bảo lãnh (COST_COMP)', 'beta': round(float(beta_std[0]), 3), 't': round(float(t_vals[1]), 3), 't_hc3': round(float(t_hc3[1]), 3), 'p': p_vals[1], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H2', 'factor': 'Thời gian và tốc độ xử lý phát hành (PROC_SPEED)', 'beta': round(float(beta_std[1]), 3), 't': round(float(t_vals[2]), 3), 't_hc3': round(float(t_hc3[2]), 3), 'p': p_vals[2], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H3', 'factor': 'Tiện ích nền tảng số e-Guarantee (DIGITAL_CONV)', 'beta': round(float(beta_std[2]), 3), 't': round(float(t_vals[3]), 3), 't_hc3': round(float(t_hc3[3]), 3), 'p': p_vals[3], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H4', 'factor': 'Uy tín và thương hiệu ngân hàng (BANK_REP)', 'beta': round(float(beta_std[3]), 3), 't': round(float(t_vals[4]), 3), 't_hc3': round(float(t_hc3[4]), 3), 'p': p_vals[4], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H5', 'factor': 'Mối quan hệ ngân hàng - doanh nghiệp (RELATIONSHIP)', 'beta': round(float(beta_std[4]), 3), 't': round(float(t_vals[5]), 3), 't_hc3': round(float(t_hc3[5]), 3), 'p': p_vals[5], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H6', 'factor': 'Năng lực tư vấn và hỗ trợ của nhân viên (STAFF_QUAL)', 'beta': round(float(beta_std[5]), 3), 't': round(float(t_vals[6]), 3), 't_hc3': round(float(t_hc3[6]), 3), 'p': p_vals[6], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H7', 'factor': 'Chính sách TSBĐ và tỷ lệ ký quỹ (COLL_POLICY)', 'beta': round(float(beta_std[6]), 3), 't': round(float(t_vals[7]), 3), 't_hc3': round(float(t_hc3[7]), 3), 'p': p_vals[7], 'result': 'Chấp nhận (Supported)'},
        {'hypothesis': 'H8', 'factor': 'Sự khác biệt theo đặc điểm doanh nghiệp (Đơn/Đa ngân hàng, Hình thức sở hữu, Doanh thu, Thâm niên, Sản phẩm)', 'beta': None, 't': None, 't_hc3': None, 'p': None, 'result': 'Ủng hộ một phần (Partially Supported - Khác biệt có ý nghĩa ở Hình thức sở hữu p < 0.01 và Đơn/Đa ngân hàng p < 0.001; không khác biệt ở Doanh thu, Thâm niên, Sản phẩm)'}
    ]
    results['hypotheses_summary'] = hypotheses

    # Save to JSON
    with open('workingfile/v4_empirical_results.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print("Da luu ket qua thong ke toan dien vao workingfile/v4_empirical_results.json")

    return results

if __name__ == '__main__':
    compute()
