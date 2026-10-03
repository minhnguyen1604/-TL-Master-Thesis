# -*- coding: utf-8 -*-
import json

with open('workingfile/v3_empirical_results.json', 'r', encoding='utf-8') as f:
    d = json.load(f)

print('=== 1. CRONBACH ALPHAS ===')
for k, v in d['cronbach'].items():
    print(f"{k:12s}: alpha={v['alpha']:.3f}, CITC min={v['citc_min']:.3f}, max={v['citc_max']:.3f}")

print('\n=== 2. EFA SUMMARY ===')
print(f"KMO: {d['efa']['kmo']:.3f}, Bartlett: chi2={d['efa']['bartlett_chi2']}, df={d['efa']['bartlett_df']}, p={d['efa']['bartlett_sig']:.4e}")
print(f"Cumulative variance: {d['efa']['cum_var'][-1]}%")
print(f"Harman Single Factor: {d['efa']['harman_variance_pct']}%")

print('\n=== 3. REGRESSION SUMMARY ===')
s = d['regression']['summary']
print(f"R={s['R']}, R2={s['R2']}, Adj R2={s['Adj_R2']}, F={s['F']} (p={s['F_sig']:.4e}), DW={s['DW']}")
print('\nCoefficients:')
for c in d['regression']['coefficients']:
    if c['var'] == 'Constant':
        print(f"{c['var']:12s}: B={c['B']:.3f}, SE={c['SE']:.3f}, t={c['t']:.2f}, p={c['Sig']:.4e}, SE_HC3={c['SE_HC3']:.3f}, t_HC3={c['t_HC3']:.2f}")
    else:
        print(f"{c['var']:12s}: Beta={c['Beta']:.3f}, B={c['B']:.3f}, SE={c['SE']:.3f}, t={c['t']:.2f}, p={c['Sig']:.4e}, SE_HC3={c['SE_HC3']:.3f}, t_HC3={c['t_HC3']:.2f}, VIF={c['VIF']:.3f}")

print('\n=== 4. CRITERION VALIDITY ===')
cv = d['criterion_validity']
print(f"Spearman DEC - WALLET_SHARE (full): rho={cv['rho_full']:.3f}, p={cv['p_rho_full']:.4e}")
print(f"Spearman DEC - WALLET_SHARE (3-item): rho={cv['rho_3item']:.3f}, p={cv['p_rho_3item']:.4e}")

print('\n=== 5. DIFFERENCES TESTS ===')
tb = d['t_test_banking']
print(f"t-test Single vs Multi: t={tb['t_stat']}, p={tb['p_val']:.4e}, Single Mean={tb['single_mean']}, Multi Mean={tb['multi_mean']}")
ao = d['anova_ownership']
print(f"ANOVA Ownership: F={ao['F']}, p={ao['p_val']:.4e}")
ar = d['anova_revenue']
print(f"ANOVA Revenue: F={ar['F']}, p={ar['p_val']:.4e}")
ae = d['anova_experience']
print(f"ANOVA Experience: F={ae['F']}, p={ae['p_val']:.4e}")
ap = d['anova_product']
print(f"ANOVA Product: F={ap['F']}, p={ap['p_val']:.4e}")

print('\n=== 6. STAFF4 FACTOR LOADINGS CHECK ===')
for row in d['efa']['rot_matrix']:
    if row['item'] == 'STAFF4':
        print(f"STAFF4: Primary={row['primary']}, Secondary={row['secondary']}, Gap={row['gap']}")
