# -*- coding: utf-8 -*-
"""
generate_ch4_content.py
Chua toan bo du lieu tinh toan tu Excel va ham sinh noi dung hoc thuat chuan muc.
"""
import sys, os
import pandas as pd
import numpy as np
from scipy import stats

def compute_all_statistics(excel_file):
    df = pd.read_excel(excel_file)
    stats_dict = {}

    # 1. Demographics
    stats_dict['df'] = df
    stats_dict['N'] = len(df)
    
    # Construct mappings
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
    stats_dict['constructs'] = constructs

    # Construct summated/mean scores
    for cname, items in constructs.items():
        df[cname] = df[items].mean(axis=1)

    # 2. Reliability (Cronbach's Alpha)
    def calc_alpha(sub):
        k = sub.shape[1]
        item_vars = sub.var(axis=0, ddof=1).sum()
        tot_var = sub.sum(axis=1).var(ddof=1)
        return (k / (k - 1)) * (1 - item_vars / tot_var)

    rel_summary = []
    item_stats = {}
    for cname, items in constructs.items():
        sub = df[items]
        tot = sub.sum(axis=1)
        alpha = calc_alpha(sub)
        mean_val = sub.mean().mean()
        std_val = sub.std().mean()
        itcs = [np.corrcoef(sub[c], tot - sub[c])[0, 1] for c in items]
        min_itc = min(itcs)
        
        # item deleted alphas
        del_alphas = []
        for c in items:
            del_sub = sub.drop(columns=[c])
            del_alphas.append(calc_alpha(del_sub))
            item_stats[c] = {
                'mean': sub[c].mean(),
                'std': sub[c].std(),
                'itc': np.corrcoef(sub[c], tot - sub[c])[0, 1],
                'del_alpha': calc_alpha(del_sub)
            }
        
        rel_summary.append({
            'construct': cname,
            'items': len(items),
            'alpha': alpha,
            'mean': mean_val,
            'std': std_val,
            'min_itc': min_itc,
            'max_del_alpha': max(del_alphas)
        })
    stats_dict['rel_summary'] = rel_summary
    stats_dict['item_stats'] = item_stats

    # 3. Correlation & VIF
    indep_vars = ['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']
    corr_all = df[list(constructs.keys())].corr()
    stats_dict['corr_all'] = corr_all

    corr_indep = df[indep_vars].corr().values
    vifs = np.diag(np.linalg.inv(corr_indep))
    stats_dict['vifs'] = dict(zip(indep_vars, vifs))

    # 4. Regression
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

    stats_dict['reg'] = {
        'r2': r2,
        'adj_r2': adj_r2,
        'f_stat': f_stat,
        'p_f': p_f,
        'se_est': np.sqrt(mse),
        'beta': beta,
        'se': se,
        't_vals': t_vals,
        'p_vals': p_vals,
        'beta_std': beta_std,
        'indep_vars': indep_vars
    }

    # 5. Robustness check regression
    df['D_SOE'] = (df['OWNERSHIP'] == 3).astype(int)
    df['D_FDI'] = (df['OWNERSHIP'] == 4).astype(int)
    df['D_LARGE'] = (df['REVENUE'] >= 3).astype(int)
    df['D_SINGLE'] = (df['NUM_BANKS'] == 1).astype(int)
    controls = ['D_SOE', 'D_FDI', 'D_LARGE', 'D_SINGLE']
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

    stats_dict['robustness'] = {
        'r2': r2_rob,
        'adj_r2': adj_r2_rob,
        'f_stat': f_rob,
        'beta': beta_rob,
        'se': se_rob,
        't_vals': t_rob,
        'p_vals': p_rob,
        'vars': ['Const'] + indep_vars + controls
    }

    # 6. Group difference tests
    # t-test single vs multi
    s1 = df[df['NUM_BANKS'] == 1]['DEC']
    s2 = df[df['NUM_BANKS'] > 1]['DEC']
    t_val, p_t = stats.ttest_ind(s1, s2)
    stats_dict['ttest_single_multi'] = {
        'n1': len(s1), 'm1': s1.mean(), 'sd1': s1.std(),
        'n2': len(s2), 'm2': s2.mean(), 'sd2': s2.std(),
        't': t_val, 'p': p_t, 'd': (s1.mean() - s2.mean()) / np.sqrt(((len(s1)-1)*s1.var() + (len(s2)-1)*s2.var()) / (len(s1)+len(s2)-2))
    }

    return stats_dict

if __name__ == '__main__':
    st = compute_all_statistics('Du_Lieu_Khao_Sat_Tho_800_DN.xlsx')
    print('All statistics successfully calculated from raw Excel dataset!')
    print('R2:', st['reg']['r2'], 'F:', st['reg']['f_stat'])
