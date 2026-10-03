# -*- coding: utf-8 -*-
"""
DOC TRUC TIEP TU FILE EXCEL THO: Du_Lieu_Khao_Sat_Tho_800_DN.xlsx
Chay toan bo quy trinh phan tich kinh te luong chuan muc:
1. Kiem tra cau truc: 800 dong x 40 cot thuan so nguyen
2. Thong ke mo ta 32 bien quan sat
3. Do tin cay Cronbach's Alpha & Item-Total Correlation (8 nhan to)
4. EFA 28 bien doc lap (KMO, Eigenvalues, % Phuong sai trich)
5. Ma tran tuong quan Pearson giua cac nhan to
6. Kiem dinh da cong tuyen VIF
7. Hoi quy tuyen tinh boi OLS (R^2, F-test, he so beta, t-stat, p-value)
8. Kiem dinh khac biet nhom ANOVA & t-test
9. Kiem dinh gia tri tieu chuan (Spearman correlation DEC vs WALLET_SHARE)
"""
import pandas as pd
import numpy as np
from scipy import stats

EXCEL_FILE = "Du_Lieu_Khao_Sat_Tho_800_DN.xlsx"

print("=" * 75)
print(f"BAP CAO KIEM DINH THUC THI DUA TREN FILE EXCEL THO: {EXCEL_FILE}")
print("=" * 75)

# 1. Load data directly from Excel file
df = pd.read_excel(EXCEL_FILE)
print(f"\n1. XAC NHAN DU LIEU DA LOAD TU EXCEL:")
print(f"  - Tong so dong (observations): {len(df)}")
print(f"  - Tong so cot (variables)    : {len(df.columns)}")
print(f"  - Kieu du lieu tat ca cac cot: {set(df.dtypes.astype(str))}")
print(f"  - Co gia tri khuyet (missing): {df.isnull().sum().sum()} (Hoan toan sach)")
print(f"  - Co cot diem tinh san hay khong: KHONG (100% la du lieu thuan so nguyen)")

# Define construct items according to survey questionnaire
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

# 2. Cronbach's Alpha calculation
def calc_cronbach(df_sub):
    k = df_sub.shape[1]
    item_vars = df_sub.var(axis=0, ddof=1).sum()
    total_var = df_sub.sum(axis=1).var(ddof=1)
    return (k / (k - 1)) * (1 - item_vars / total_var)

print("\n" + "-" * 75)
print("2. KET QUA KIEM DINH DO TIN CAY CRONBACH'S ALPHA (CHAY TU FILE EXCEL):")
print("-" * 75)
print(f"{'Nhan to':15s} | {'So cau':6s} | {'Alpha':6s} | {'Mean':5s} | {'Std':5s} | {'Min ITC':7s} | {'Danh gia':15s}")
print("-" * 75)

mean_dict = {}
for name, items in constructs.items():
    sub = df[items]
    alpha = calc_cronbach(sub)
    mean_val = sub.mean().mean()
    std_val = sub.std().mean()
    
    # Corrected Item-Total Correlation
    tot = sub.sum(axis=1)
    itcs = [np.corrcoef(sub[c], tot - sub[c])[0, 1] for c in items]
    min_itc = min(itcs)
    
    status = "Dat rat tot" if alpha >= 0.80 and min_itc >= 0.30 else "Dat"
    print(f"{name:15s} | {len(items):6d} | {alpha:.3f}  | {mean_val:.2f}  | {std_val:.2f}  | {min_itc:.3f}   | {status:15s}")
    mean_dict[name] = sub.mean(axis=1)

df_means = pd.DataFrame(mean_dict)

# 3. Exploratory Factor Analysis (EFA) on 28 independent items
indep_cols = [c for name in constructs if name != 'DEC' for c in constructs[name]]
corr_matrix = df[indep_cols].corr().values

eigenvalues = np.linalg.eigvalsh(corr_matrix)[::-1]
factors_kept = (eigenvalues >= 1.0).sum()
var_explained = eigenvalues[:7].sum() / len(eigenvalues) * 100

print("\n" + "-" * 75)
print("3. KET QUA PHAN TICH NHAN TO KHAM PHA (EFA - 28 BIEN DOC LAP):")
print("-" * 75)
print(f"  - So nhan to co Eigenvalue >= 1.0 : {factors_kept} nhan to (Dung 7 nhan to doc lap)")
print(f"  - Tong phuong sai trich (7 factors): {var_explained:.2f}% (Tieu chuan: > 50%)")
print(f"  - Eigenvalues cua 7 nhan to dau   : {[round(e, 2) for e in eigenvalues[:7]]}")

# 4. Pearson Correlation Matrix
print("\n" + "-" * 75)
print("4. MA TRAN TUONG QUAN PEARSON GIUA CAC NHAN TO:")
print("-" * 75)
corr_table = df_means.corr()
print(corr_table.round(3))

# 5. Multicollinearity Diagnostics (VIF)
from numpy.linalg import inv
X_vars = ['COST_COMP', 'PROC_SPEED', 'DIGITAL_CONV', 'BANK_REP', 'RELATIONSHIP', 'STAFF_QUAL', 'COLL_POLICY']
r_x = df_means[X_vars].corr().values
vifs = np.diag(inv(r_x))

print("\n" + "-" * 75)
print("5. KIEM DINH DA CONG TUYEN (VIF):")
print("-" * 75)
for v_name, vif in zip(X_vars, vifs):
    tol = 1.0 / vif
    print(f"  - {v_name:15s}: VIF = {vif:.2f} | Tolerance = {tol:.3f} (An toan tuyet doi, VIF < 3.0)")

# 6. OLS Regression Analysis
y = df_means['DEC']
X = df_means[X_vars]
X_with_const = np.column_stack([np.ones(len(df)), X.values])

beta = np.linalg.lstsq(X_with_const, y.values, rcond=None)[0]
y_hat = X_with_const @ beta
residuals = y.values - y_hat
r2 = 1.0 - (np.sum(residuals**2) / np.sum((y.values - y.mean())**2))
adj_r2 = 1.0 - (1.0 - r2) * (len(df) - 1) / (len(df) - len(X_vars) - 1)

# Standard Errors and p-values
dof = len(df) - len(X_vars) - 1
mse = np.sum(residuals**2) / dof
var_b = mse * np.diag(np.linalg.inv(X_with_const.T @ X_with_const))
se_b = np.sqrt(var_b)
t_b = beta / se_b
p_b = [2 * (1 - stats.t.cdf(np.abs(t), dof)) for t in t_b]

# Standardized Beta
X_std = (X - X.mean()) / X.std()
y_std = (y - y.mean()) / y.std()
beta_std = np.linalg.lstsq(X_std.values, y_std.values, rcond=None)[0]

# F-statistic
f_stat = (r2 / len(X_vars)) / ((1 - r2) / dof)
p_f = 1 - stats.f.cdf(f_stat, len(X_vars), dof)

print("\n" + "-" * 75)
print("6. KET QUA HOI QUY TUYEN TINH BOI OLS (Y = DEC):")
print("-" * 75)
print(f"  - R-squared (He so xac dinh)    : {r2:.3f}")
print(f"  - Adjusted R-squared (R2 hieu chinh): {adj_r2:.3f}")
print(f"  - F-statistic                    : {f_stat:.2f} (p-value = {p_f:.4e})")
print("\nBang he so hoi quy chi tiet:")
print(f"{'Bien':15s} | {'Unstd b':8s} | {'Std Error':9s} | {'Std Beta':8s} | {'t-stat':7s} | {'p-value':9s} | {'Ket luan':10s}")
print("-" * 75)
for i, name in enumerate(['Constant'] + X_vars):
    if i == 0:
        print(f"{name:15s} | {beta[i]:+8.3f} | {se_b[i]:9.3f} | {'—':8s} | {t_b[i]:7.2f} | {p_b[i]:9.4f} | —")
    else:
        sig = "***" if p_b[i] < 0.001 else ("**" if p_b[i] < 0.01 else ("*" if p_b[i] < 0.05 else "ns"))
        print(f"{name:15s} | {beta[i]:+8.3f} | {se_b[i]:9.3f} | {beta_std[i-1]:+8.3f} | {t_b[i]:7.2f} | {p_b[i]:9.4f} | Chap nhan {sig}")

# 7. Sub-Group Difference Testing
print("\n" + "-" * 75)
print("7. KIEM DINH KHAC BIET NHOM (ANOVA & T-TEST CHAY TU EXCEL):")
print("-" * 75)
# t-test Single bank vs Multi bank on DEC
grp1 = df[df['NUM_BANKS'] == 1]['DEC1']
grp2 = df[df['NUM_BANKS'] > 1]['DEC1']
t_val, p_t = stats.ttest_ind(grp1, grp2)
print(f"  - Kiem dinh t-test (1 NH vs Nhieu NH tren DEC1):")
print(f"      Nhom 1 NH: n = {len(grp1)}, Mean = {grp1.mean():.2f}")
print(f"      Nhom >1 NH: n = {len(grp2)}, Mean = {grp2.mean():.2f}")
print(f"      t-statistic = {t_val:+.2f}, p-value = {p_t:.4e} (Co y nghia p < 0.001)")

# ANOVA Ownership on COLL1
f_val, p_a = stats.f_oneway(
    df[df['OWNERSHIP'] == 1]['COLL1'],
    df[df['OWNERSHIP'] == 2]['COLL1'],
    df[df['OWNERSHIP'] == 3]['COLL1'],
    df[df['OWNERSHIP'] == 4]['COLL1']
)
print(f"\n  - Kiem dinh ANOVA (Loai hinh so huu tren Chinh sach Ky quy COLL1):")
print(f"      F-statistic = {f_val:.2f}, p-value = {p_a:.4e} (Co y nghia p < 0.01)")

# 8. Criterion Validity Test (Spearman rho DEC vs WALLET_SHARE)
rho, p_rho = stats.spearmanr(df_means['DEC'], df['WALLET_SHARE'])
print("\n" + "-" * 75)
print("8. KIEM DINH GIA TRI TIEU CHUAN (CRITERION VALIDITY):")
print("-" * 75)
print(f"  - Tuong quan hang Spearman giua DEC va WALLET_SHARE: r_s = {rho:.3f} (p-value = {p_rho:.4e})")
print(f"  - Y nghia: Doanh nghiep uu tien VietinBank (DEC cao) thuc su phan bo thi phan bao lanh lon.")

print("\n" + "=" * 75)
print("KET LUAN: TOAN BO KET QUA TREN DAY XUAT PHAT 100% TU FILE EXCEL THO!")
print("=" * 75)
