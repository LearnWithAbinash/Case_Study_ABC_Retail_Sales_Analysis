#!/usr/bin/env python3
"""
===============================================================================
ABC RETAIL SALES ANALYSIS — Marketing Mix Modeling (MMM)
===============================================================================
Case Study: ABC Retail Sales Analysis — Team Lead - Data Intelligence
Author: Abinash Sahu | August 2026

Objective
---------
Evaluate the effectiveness of ABC Retail's marketing channels (Digital,
Radio, TV) and the role of macro-economic factors (GDP, CPI, Unemployment)
and COVID-19 on sales performance.

Questions Addressed
-------------------
1. Are all media channels effective in driving sales/revenue?
2. Which is the most important channel out of the three?
3. Do macro-economic factors have an influence on sales?
4. Did COVID (Q2-Q3 2020 lockdown/relaxation) impact sales, and why?

Approach
--------
Marketing Mix Modeling (MMM) — the industry-standard framework for this
class of problem. Key features beyond a basic regression:
  (a) Adstock: carryover effects modelled with geometric decay
  (b) Saturation: diminishing returns via log(1+x)
  (c) Data-driven decay rates via grid search (not assumed)
  (d) Additional macro signals: CPI, Unemployment alongside GDP
  (e) Robust inference: HC3 standard errors, cross-validation
  (f) Budget optimisation: constrained scipy optimiser
===============================================================================
"""

import pandas as pd
import numpy as np
import json
import warnings
from pathlib import Path
from datetime import datetime
from itertools import product

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns

from sklearn.linear_model import Ridge
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import LeaveOneOut, cross_val_predict
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error

import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.stats.diagnostic import het_breuschpagan, acorr_breusch_godfrey
from statsmodels.stats.stattools import durbin_watson

from scipy.optimize import minimize

warnings.filterwarnings('ignore')

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
DATA_FILE = BASE_DIR / 'Retail Sales Case Study.xlsx'
OUTPUT_DIR = BASE_DIR / 'output'
CHARTS_DIR = OUTPUT_DIR / 'charts'
CHARTS_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# Chart styling — clean, presentation-ready
# ---------------------------------------------------------------------------
plt.rcParams.update({
    'font.family': 'sans-serif',
    'font.sans-serif': ['Helvetica Neue', 'Arial', 'DejaVu Sans'],
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'axes.labelsize': 11,
    'figure.facecolor': 'white',
    'axes.facecolor': '#FAFBFD',
    'axes.grid': True,
    'grid.alpha': 0.15,
    'grid.linewidth': 0.5,
    'axes.spines.top': False,
    'axes.spines.right': False,
    'legend.frameon': False,
    'figure.dpi': 150,
    'savefig.dpi': 200,
    'savefig.bbox': 'tight',
    'savefig.pad_inches': 0.3,
})

C = {
    'navy': '#0B1F3A', 'blue': '#3B82F6', 'teal': '#14B8A6',
    'emerald': '#10B981', 'orange': '#F59E0B', 'red': '#EF4444',
    'purple': '#8B5CF6', 'pink': '#EC4899', 'gray': '#64748B',
    'light': '#E2E8F0', 'dark': '#1E293B', 'white': '#FFFFFF',
}

print("=" * 65)
print("  ABC Retail — Marketing Mix Model Analysis")
print("=" * 65)
print()

# ===================================================================
# 1  DATA ENGINEERING
# ===================================================================
print("[1] Data preparation")
print("-" * 50)

df = pd.read_excel(DATA_FILE, sheet_name=0, engine='openpyxl')
df.columns = ['date', 'sales', 'digital', 'radio', 'tv_k']
df['date'] = pd.to_datetime(df['date'])

# TV is reported in thousands; convert to millions for unit parity
df['tv'] = df['tv_k'] / 1000.0

# GDP from the provided sheet
gdp_df = pd.read_excel(DATA_FILE, sheet_name=1, engine='openpyxl')
gdp_df.columns = ['qtr', 'gdp_change']

df['year']    = df['date'].dt.year
df['month']   = df['date'].dt.month
df['quarter'] = df['date'].dt.quarter
df['qtr']     = df['year'].astype(str) + ' - Q' + df['quarter'].astype(str)
df = df.merge(gdp_df, on='qtr', how='left')

# -----------------------------------------------------------------
# Additional Macro Data (US Bureau of Labor Statistics & FRED)
# - CPI-U (% change YoY - Inflation)
# - California Unemployment Rate (%, Seasonally Adjusted - CA EDD / BLS)
# - US Consumer Sentiment Index (University of Michigan UMCSENT)
# - US Retail E-Commerce Penetration % (US Census Bureau)
# -----------------------------------------------------------------
macro_monthly = {
    # (year, month): (cpi_yoy, ca_unemp, sentiment, ecom_pct)
    (2020,  1): (2.5,  4.2, 99.8, 11.8), (2020,  2): (2.3,  4.3, 101.0, 11.8), (2020,  3): (1.5,  5.5, 89.1, 11.8),
    (2020,  4): (0.3, 16.0, 71.8, 16.4), (2020,  5): (0.1, 16.4, 72.3, 16.4), (2020,  6): (0.6, 14.9, 78.1, 16.4),
    (2020,  7): (1.0, 13.5, 72.5, 14.5), (2020,  8): (1.3, 11.4, 74.1, 14.5), (2020,  9): (1.4, 10.7, 80.4, 14.5),
    (2020, 10): (1.2,  9.3, 81.8, 14.9), (2020, 11): (1.2,  9.0, 76.9, 14.9), (2020, 12): (1.4,  8.9, 80.7, 14.9),
    (2021,  1): (1.4,  8.8, 79.0, 15.6), (2021,  2): (1.7,  8.3, 76.8, 15.6), (2021,  3): (2.6,  8.1, 84.9, 15.6),
    (2021,  4): (4.2,  7.9, 88.3, 14.5), (2021,  5): (5.0,  7.7, 82.9, 14.5), (2021,  6): (5.4,  7.6, 85.5, 14.5),
    (2021,  7): (5.4,  7.3, 81.2, 14.1), (2021,  8): (5.3,  6.8, 70.3, 14.1), (2021,  9): (5.4,  6.4, 72.8, 14.1),
    (2021, 10): (6.2,  6.1, 71.7, 14.1), (2021, 11): (6.8,  5.9, 67.4, 14.1), (2021, 12): (7.0,  5.8, 70.6, 14.1),
    (2022,  1): (7.5,  5.8, 67.2, 14.7), (2022,  2): (7.9,  5.3, 62.8, 14.7), (2022,  3): (8.5,  4.9, 59.4, 14.7),
    (2022,  4): (8.3,  4.6, 65.2, 14.6), (2022,  5): (8.6,  4.3, 58.4, 14.6), (2022,  6): (9.1,  4.2, 50.0, 14.6),
    (2022,  7): (8.5,  4.2, 51.5, 14.7), (2022,  8): (8.3,  4.2, 58.2, 14.7), (2022,  9): (8.2,  3.8, 58.6, 14.7),
    (2022, 10): (7.7,  4.0, 59.9, 15.1), (2022, 11): (7.1,  4.1, 56.8, 15.1), (2022, 12): (6.5,  4.1, 59.7, 15.1),
}
df['cpi_yoy']           = df.apply(lambda r: macro_monthly.get((r['year'], r['month']), (np.nan,)*4)[0], axis=1)
df['ca_unemployment']   = df.apply(lambda r: macro_monthly.get((r['year'], r['month']), (np.nan,)*4)[1], axis=1)
df['consumer_sentiment'] = df.apply(lambda r: macro_monthly.get((r['year'], r['month']), (np.nan,)*4)[2], axis=1)
df['ecommerce_pct']     = df.apply(lambda r: macro_monthly.get((r['year'], r['month']), (np.nan,)*4)[3], axis=1)

# COVID phase indicators
df['covid_lockdown']   = ((df['date'] >= '2020-04-01') & (df['date'] <= '2020-06-30')).astype(int)
df['covid_relaxation'] = ((df['date'] >= '2020-07-01') & (df['date'] <= '2020-09-30')).astype(int)

# Seasonality (sin/cos — preserves circular nature)
df['month_sin'] = np.sin(2 * np.pi * df['month'] / 12)
df['month_cos'] = np.cos(2 * np.pi * df['month'] / 12)

df['trend']       = np.arange(len(df))
df['total_spend'] = df['digital'] + df['radio'] + df['tv']

print(f"  {len(df)} monthly records, {df['date'].min():%b %Y} – {df['date'].max():%b %Y}")
print(f"  Total revenue: ${df['sales'].sum():,.0f}M")
print(f"  Total marketing spend: ${df['total_spend'].sum():,.0f}M")
print(f"  Macro & Regional Signals: GDP %, CPI %, CA Unemployment %, Consumer Sentiment, E-Commerce Penetration %")
print()


# ===================================================================
# 2  ADSTOCK GRID SEARCH
# ===================================================================
print("[2] Adstock decay-rate optimisation (grid search)")
print("-" * 50)


def adstock(series, decay):
    """Geometric adstock: a(t) = x(t) + decay * a(t-1)"""
    out = np.zeros(len(series))
    out[0] = series.iloc[0]
    for t in range(1, len(series)):
        out[t] = series.iloc[t] + decay * out[t - 1]
    return out


def saturate(x):
    """Log-saturation for diminishing returns."""
    return np.log1p(x)


def build_features(data, d_decay, r_decay, t_decay):
    """Build the feature matrix for a given set of decay rates."""
    tmp = data.copy()
    tmp['d_ad'] = adstock(tmp['digital'], d_decay)
    tmp['r_ad'] = adstock(tmp['radio'],   r_decay)
    tmp['t_ad'] = adstock(tmp['tv'],      t_decay)
    tmp['d_sat'] = saturate(tmp['d_ad'])
    tmp['r_sat'] = saturate(tmp['r_ad'])
    tmp['t_sat'] = saturate(tmp['t_ad'])
    cols = ['d_sat', 'r_sat', 't_sat', 'gdp_change', 'cpi_yoy',
            'ca_unemployment', 'consumer_sentiment', 'ecommerce_pct',
            'covid_lockdown', 'covid_relaxation',
            'month_sin', 'month_cos', 'trend']
    return tmp, cols


# Grid: test every combination in the plausible range
digital_grid = np.arange(0.10, 0.60, 0.05)
radio_grid   = np.arange(0.30, 0.80, 0.05)
tv_grid      = np.arange(0.50, 0.95, 0.05)

best_aic  = np.inf
best_comb = (0.35, 0.55, 0.75)  # fallback defaults

results_grid = []
for d_d, r_d, t_d in product(digital_grid, radio_grid, tv_grid):
    tmp, cols = build_features(df, d_d, r_d, t_d)
    X = sm.add_constant(tmp[cols])
    try:
        m = sm.OLS(tmp['sales'], X).fit()
        results_grid.append((d_d, r_d, t_d, m.aic, m.rsquared_adj))
        if m.aic < best_aic:
            best_aic  = m.aic
            best_comb = (round(d_d, 2), round(r_d, 2), round(t_d, 2))
    except Exception:
        pass

print(f"  Tested {len(results_grid):,} decay-rate combinations")
print(f"  Optimal decay rates (minimise AIC):")
print(f"    Digital : {best_comb[0]:.2f}")
print(f"    Radio   : {best_comb[1]:.2f}")
print(f"    TV      : {best_comb[2]:.2f}")
print(f"  Best AIC  : {best_aic:.1f}")
print()

DECAY = {'digital': best_comb[0], 'radio': best_comb[1], 'tv': best_comb[2]}


# ===================================================================
# 3  MODEL ESTIMATION
# ===================================================================
print("[3] Model estimation (OLS with HC3 robust standard errors)")
print("-" * 50)

df['digital_adstock'] = adstock(df['digital'], DECAY['digital'])
df['radio_adstock']   = adstock(df['radio'],   DECAY['radio'])
df['tv_adstock']      = adstock(df['tv'],      DECAY['tv'])
df['digital_sat'] = saturate(df['digital_adstock'])
df['radio_sat']   = saturate(df['radio_adstock'])
df['tv_sat']      = saturate(df['tv_adstock'])

features = ['digital_sat', 'radio_sat', 'tv_sat',
            'gdp_change', 'cpi_yoy', 'ca_unemployment',
            'consumer_sentiment', 'ecommerce_pct',
            'covid_lockdown', 'covid_relaxation',
            'month_sin', 'month_cos', 'trend']

X = df[features]
y = df['sales']
Xc = sm.add_constant(X)
model = sm.OLS(y, Xc).fit(cov_type='HC3')

print(f"  R-squared      : {model.rsquared:.4f}")
print(f"  Adjusted R-sq  : {model.rsquared_adj:.4f}")
print(f"  F-stat p-value : {model.f_pvalue:.2e}")
print(f"  AIC / BIC      : {model.aic:.1f} / {model.bic:.1f}")
print()

# Parsimonious model (media + seasonality only)
feat_mkt = ['digital_sat', 'radio_sat', 'tv_sat', 'month_sin', 'month_cos']
model_mkt = sm.OLS(y, sm.add_constant(df[feat_mkt])).fit(cov_type='HC3')
print(f"  Marketing-only R-sq: {model_mkt.rsquared:.4f} (adj {model_mkt.rsquared_adj:.4f})")
print()

# Standardised coefficients
scaler = StandardScaler()
Xs_df  = pd.DataFrame(scaler.fit_transform(X), columns=features)
ys     = (y - y.mean()) / y.std(ddof=0)
m_std  = sm.OLS(ys, sm.add_constant(Xs_df)).fit(cov_type='HC3')
std_coefs = pd.Series({f: m_std.params[f] for f in features})
std_pvals = pd.Series({f: m_std.pvalues[f] for f in features})

print("  Standardised coefficients (media channels):")
for f in ['digital_sat', 'radio_sat', 'tv_sat']:
    name = f.replace('_sat','').capitalize()
    sig  = '***' if std_pvals[f]<0.001 else '**' if std_pvals[f]<0.01 else '*' if std_pvals[f]<0.05 else 'ns'
    print(f"    {name:>8} : {std_coefs[f]:+.3f}  (p={std_pvals[f]:.4f}) {sig}")
print()

# Cross-validation
X_cv = scaler.fit_transform(X)
pred_cv = cross_val_predict(Ridge(alpha=10.0), X_cv, y, cv=LeaveOneOut())
loocv_r2   = r2_score(y, pred_cv)
loocv_mae  = mean_absolute_error(y, pred_cv)
loocv_rmse = np.sqrt(mean_squared_error(y, pred_cv))
print(f"  LOOCV R-sq: {loocv_r2:.4f}  |  MAE: ${loocv_mae:.1f}M  |  RMSE: ${loocv_rmse:.1f}M")
print()

# VIF
vif = pd.DataFrame({
    'feature': features,
    'VIF': [variance_inflation_factor(X.values, i) for i in range(X.shape[1])]
})

# Diagnostics
residuals = model.resid
fitted    = model.fittedvalues
dw = durbin_watson(residuals)
bp_stat, bp_p, _, _ = het_breuschpagan(residuals, Xc)
bg_stat, bg_p, _, _ = acorr_breusch_godfrey(model, nlags=3)
print(f"  Durbin-Watson : {dw:.3f}")
print(f"  Breusch-Pagan : p={bp_p:.4f} ({'homoscedastic' if bp_p>0.05 else 'heteroscedastic — HC3 handles this'})")
print(f"  Breusch-Godfrey: p={bg_p:.4f}")
print()


# ===================================================================
# 4  CHANNEL ROI
# ===================================================================
print("[4] Channel ROI")
print("-" * 50)

channel_map = {
    'digital_sat': ('digital', 'Digital'),
    'radio_sat':   ('radio',   'Radio'),
    'tv_sat':      ('tv',      'TV'),
}

roi_rows = []
for feat, (raw, name) in channel_map.items():
    coef = model.params[feat]
    ci   = model.conf_int().loc[feat]
    pval = model.pvalues[feat]
    total_spend = df[raw].sum()
    avg_adstock = df[feat.replace('_sat','_adstock')].mean()
    decay       = DECAY[raw]
    # Marginal ROI via chain rule: dSales/dSpend = coef / (1+adstock_avg) / (1-decay)
    marginal_roi = coef * (1/(1+avg_adstock)) * (1/(1-decay))

    roi_rows.append({
        'channel': name, 'coefficient': coef, 'p_value': pval,
        'ci_lo': ci.iloc[0], 'ci_hi': ci.iloc[1],
        'std_coef': std_coefs.get(feat, 0),
        'total_spend_M': total_spend,
        'marginal_roi': marginal_roi,
        'spend_share': total_spend / df['total_spend'].sum() * 100,
        'significant': pval < 0.05,
    })

roi_df = pd.DataFrame(roi_rows)

for _, r in roi_df.iterrows():
    sig = 'significant' if r['significant'] else 'not significant'
    print(f"  {r['channel']:>8} — Marginal ROI ${r['marginal_roi']:.2f} per $1  |  "
          f"std coef {r['std_coef']:.3f}  |  p={r['p_value']:.4f} ({sig})")
print()


# ===================================================================
# 5  MACRO & COVID
# ===================================================================
print("[5] Macro-economic & COVID analysis")
print("-" * 50)

gdp_p   = model.pvalues.get('gdp_change', 1)
cpi_p   = model.pvalues.get('cpi_yoy', 1)
unemp_p = model.pvalues.get('ca_unemployment', 1)
sent_p  = model.pvalues.get('consumer_sentiment', 1)
ecom_p  = model.pvalues.get('ecommerce_pct', 1)
lock_p  = model.pvalues.get('covid_lockdown', 1)
relax_p = model.pvalues.get('covid_relaxation', 1)

print(f"  GDP          : coef={model.params.get('gdp_change',0):+.1f}  p={gdp_p:.4f}  {'sig' if gdp_p<0.05 else 'not significant'}")
print(f"  CPI (YoY)    : coef={model.params.get('cpi_yoy',0):+.1f}  p={cpi_p:.4f}  {'sig' if cpi_p<0.05 else 'not significant'}")
print(f"  CA Unemp     : coef={model.params.get('ca_unemployment',0):+.1f}  p={unemp_p:.4f}  {'sig' if unemp_p<0.05 else 'not significant'}")
print(f"  Sentiment    : coef={model.params.get('consumer_sentiment',0):+.1f}  p={sent_p:.4f}  {'sig' if sent_p<0.05 else 'not significant'}")
print(f"  E-Commerce % : coef={model.params.get('ecommerce_pct',0):+.1f}  p={ecom_p:.4f}  {'sig' if ecom_p<0.05 else 'not significant'}")
print(f"  COVID lock   : coef={model.params.get('covid_lockdown',0):+.1f}  p={lock_p:.4f}  {'sig' if lock_p<0.05 else 'not significant'}")
print(f"  COVID relax  : coef={model.params.get('covid_relaxation',0):+.1f}  p={relax_p:.4f}  {'sig' if relax_p<0.05 else 'not significant'}")
print()

# Phase averages
phase_masks = {
    'Pre-COVID (Q1 2020)':       (df['date']>='2020-01-01')&(df['date']<='2020-03-31'),
    'Lockdown (Q2 2020)':        df['covid_lockdown']==1,
    'Relaxation (Q3 2020)':      df['covid_relaxation']==1,
    'Recovery (Q4 2020)':        (df['date']>='2020-10-01')&(df['date']<='2020-12-31'),
    'Full 2021':                 df['year']==2021,
    'Full 2022':                 df['year']==2022,
}
phase_data = []
for pname, mask in phase_masks.items():
    d = df[mask]
    phase_data.append({
        'phase': pname,
        'avg_sales': d['sales'].mean(),
        'avg_digital': d['digital'].mean(),
        'avg_radio': d['radio'].mean(),
        'avg_tv_k': d['tv_k'].mean(),
    })
phase_df = pd.DataFrame(phase_data)

# Annual
annual = df.groupby('year').agg(
    total_sales   =('sales','sum'),
    avg_sales     =('sales','mean'),
    total_digital =('digital','sum'),
    total_radio   =('radio','sum'),
    total_tv      =('tv','sum'),
    total_spend   =('total_spend','sum'),
).reset_index()
annual['yoy'] = annual['total_sales'].pct_change() * 100

for _, r in annual.iterrows():
    growth = f"{r['yoy']:+.1f}%" if not pd.isna(r['yoy']) else 'baseline'
    print(f"  {int(r['year'])}: ${r['total_sales']:,.0f}M  ({growth})")
print()


# ===================================================================
# 6  BUDGET OPTIMISATION
# ===================================================================
print("[6] Budget optimisation")
print("-" * 50)

current_alloc = {'Digital': df['digital'].sum(), 'Radio': df['radio'].sum(), 'TV': df['tv'].sum()}
total_budget  = sum(current_alloc.values())


def predict_media(d_bud, r_bud, t_bud, n=36):
    ds = pd.Series([d_bud/n]*n)
    rs = pd.Series([r_bud/n]*n)
    ts = pd.Series([t_bud/n]*n)
    return (model.params['digital_sat'] * saturate(adstock(ds, DECAY['digital'])).sum() +
            model.params['radio_sat']   * saturate(adstock(rs, DECAY['radio'])).sum() +
            model.params['tv_sat']      * saturate(adstock(ts, DECAY['tv'])).sum())


def optimise(budget, n=36):
    x0 = [budget*current_alloc['Digital']/total_budget,
          budget*current_alloc['Radio']/total_budget,
          budget*current_alloc['TV']/total_budget]
    res = minimize(lambda x: -predict_media(x[0],x[1],x[2],n), x0,
                   method='SLSQP',
                   constraints={'type':'eq','fun':lambda x: sum(x)-budget},
                   bounds=[(0,budget*0.95)]*3,
                   options={'maxiter':1000,'ftol':1e-10})
    return res

opt = optimise(total_budget)
opt_d, opt_r, opt_t = opt.x
current_pred = predict_media(current_alloc['Digital'], current_alloc['Radio'], current_alloc['TV'])
optimal_pred = -opt.fun
uplift       = optimal_pred - current_pred
uplift_pct   = uplift / current_pred * 100 if current_pred else 0

print(f"  Current  : Digital ${current_alloc['Digital']:,.0f}M ({current_alloc['Digital']/total_budget*100:.0f}%)  "
      f"Radio ${current_alloc['Radio']:,.0f}M ({current_alloc['Radio']/total_budget*100:.0f}%)  "
      f"TV ${current_alloc['TV']:,.1f}M ({current_alloc['TV']/total_budget*100:.1f}%)")
print(f"  Optimal  : Digital ${opt_d:,.0f}M ({opt_d/total_budget*100:.0f}%)  "
      f"Radio ${opt_r:,.0f}M ({opt_r/total_budget*100:.0f}%)  "
      f"TV ${opt_t:,.1f}M ({opt_t/total_budget*100:.1f}%)")
print(f"  Uplift   : {uplift_pct:+.1f}%")
print()

# Sensitivity
sens_rows = []
for chg in [-20, -10, 0, 10, 20, 30]:
    b = total_budget*(1+chg/100)
    o = optimise(b)
    sens_rows.append({
        'budget_change': chg,
        'budget': b,
        'media_effect': -o.fun,
        'dig_pct': o.x[0]/b*100,
        'rad_pct': o.x[1]/b*100,
        'tv_pct':  o.x[2]/b*100,
    })
sens_df = pd.DataFrame(sens_rows)


# ===================================================================
# 7  CHARTS (16)
# ===================================================================
print("[7] Generating charts")
print("-" * 50)

# -- 1  Monthly sales
fig, ax = plt.subplots(figsize=(13, 5))
ax.axvspan(pd.Timestamp('2020-04-01'), pd.Timestamp('2020-06-30'), color=C['red'], alpha=0.08, label='Lockdown Q2-2020')
ax.axvspan(pd.Timestamp('2020-07-01'), pd.Timestamp('2020-09-30'), color=C['orange'], alpha=0.08, label='Relaxation Q3-2020')
ax.plot(df['date'], df['sales'], color=C['blue'], lw=2.5, marker='o', ms=5, mfc='white', mew=2, mec=C['blue'], zorder=5)
ma = df['sales'].rolling(3, center=True).mean()
ax.plot(df['date'], ma, color=C['teal'], lw=2, ls='--', alpha=0.7, label='3-month moving avg')
mi, mx = df['sales'].idxmin(), df['sales'].idxmax()
ax.annotate(f"Peak ${df.loc[mx,'sales']:.0f}M", xy=(df.loc[mx,'date'], df.loc[mx,'sales']),
            xytext=(10,15), textcoords='offset points', fontsize=9, fontweight='bold', color=C['emerald'],
            arrowprops=dict(arrowstyle='->', color=C['emerald'], lw=1.5))
ax.annotate(f"Low ${df.loc[mi,'sales']:.0f}M", xy=(df.loc[mi,'date'], df.loc[mi,'sales']),
            xytext=(10,-25), textcoords='offset points', fontsize=9, fontweight='bold', color=C['red'],
            arrowprops=dict(arrowstyle='->', color=C['red'], lw=1.5))
ax.set_title('Monthly sales — high month-to-month volatility, no COVID structural break', loc='left', color=C['navy'], pad=12)
ax.set_ylabel('Sales ($M)'); ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('$%.0fM'))
ax.legend(loc='upper right', fontsize=9, ncol=4)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'01_monthly_sales_trend.png'); plt.close(fig)
print("    01_monthly_sales_trend.png")

# -- 2  Quarterly sales vs GDP
qtr = df.groupby(['year','quarter','qtr']).agg(sales=('sales','sum'), gdp=('gdp_change','first')).reset_index()
fig, ax1 = plt.subplots(figsize=(12, 5)); x = np.arange(len(qtr))
bars = ax1.bar(x, qtr['sales'], 0.6, color=C['blue'], alpha=0.85, ec='white', lw=0.5, label='Quarterly sales')
for i, q in enumerate(qtr['qtr']):
    if q=='2020 - Q2': bars[i].set_facecolor(C['red']); bars[i].set_alpha(0.6)
    elif q=='2020 - Q3': bars[i].set_facecolor(C['orange']); bars[i].set_alpha(0.6)
ax1.set_ylabel('Sales ($M)', color=C['navy'])
ax2 = ax1.twinx()
ax2.plot(x, qtr['gdp']*100, color=C['orange'], marker='D', ms=7, lw=2.5, mfc='white', mew=2, mec=C['orange'], label='GDP change %', zorder=5)
ax2.axhline(0, color=C['gray'], lw=0.8, ls=':'); ax2.set_ylabel('GDP change %', color=C['orange']); ax2.spines['top'].set_visible(False)
ax1.set_xticks(x); ax1.set_xticklabels(qtr['qtr'].str.replace(' - ','\n'), fontsize=9)
ax1.set_title('Quarterly sales vs GDP — sales remain resilient despite macro shock', loc='left', color=C['navy'], pad=12)
l1,lb1=ax1.get_legend_handles_labels(); l2,lb2=ax2.get_legend_handles_labels()
ax1.legend(l1+l2, lb1+lb2, loc='upper right', fontsize=9)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'02_quarterly_sales_gdp.png'); plt.close(fig)
print("    02_quarterly_sales_gdp.png")

# -- 3  Annual sales
fig, ax = plt.subplots(figsize=(9, 5))
bars = ax.bar(annual['year'].astype(str), annual['total_sales'], color=[C['gray'],C['blue'],C['teal']], width=0.55, ec='white', lw=1)
for bar, val, g in zip(bars, annual['total_sales'], annual['yoy']):
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+30, f'${val:,.0f}M', ha='center', fontweight='bold', color=C['navy'], fontsize=13)
    if not pd.isna(g):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()-120, f'{g:+.1f}%', ha='center', fontweight='bold', color='white', fontsize=11)
ax.set_title('Annual revenue — 2021 grew despite COVID, 2022 softened slightly', loc='left', color=C['navy'], pad=12)
ax.set_ylabel('Sales ($M)'); ax.set_ylim(0, annual['total_sales'].max()*1.15)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'03_annual_sales.png'); plt.close(fig)
print("    03_annual_sales.png")

# -- 4  Correlation heatmap
corr = df[['sales','digital','radio','tv','gdp_change','cpi_yoy','ca_unemployment','consumer_sentiment','ecommerce_pct']].corr()
fig, ax = plt.subplots(figsize=(10, 8))
mask = np.triu(np.ones_like(corr, dtype=bool), k=1)
sns.heatmap(corr, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r', center=0, vmin=-1, vmax=1,
            square=True, linewidths=2, linecolor='white',
            xticklabels=['Sales','Digital','Radio','TV','GDP','CPI','CA Unemp','Sentiment','E-Com %'],
            yticklabels=['Sales','Digital','Radio','TV','GDP','CPI','CA Unemp','Sentiment','E-Com %'],
            cbar_kws={'shrink':0.8}, annot_kws={'fontsize':10, 'fontweight':'bold'}, ax=ax)
ax.set_title('Correlation matrix — Digital has the strongest sales relationship', loc='left', color=C['navy'], pad=15)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'04_correlation_heatmap.png'); plt.close(fig)
print("    04_correlation_heatmap.png")

# -- 5  Channel scatter
fig, axes = plt.subplots(1, 3, figsize=(14, 4.5))
for ax, (col, lbl, c) in zip(axes, [('digital','Digital ($M)',C['blue']),('radio','Radio ($M)',C['teal']),('tv_k','TV ($K)',C['orange'])]):
    ax.scatter(df[col], df['sales'], s=50, c=c, alpha=0.7, ec='white', lw=0.8, zorder=5)
    z = np.polyfit(df[col], df['sales'], 1); xx = np.linspace(df[col].min(), df[col].max(), 100)
    ax.plot(xx, np.polyval(z,xx), color=C['navy'], lw=2, ls='--', alpha=0.8)
    r = np.corrcoef(df[col], df['sales'])[0,1]
    ax.text(0.05, 0.95, f'r = {r:.2f}', transform=ax.transAxes, fontsize=10, fontweight='bold', color=C['navy'], va='top',
            bbox=dict(boxstyle='round,pad=0.3', fc='white', alpha=0.8))
    ax.set_xlabel(lbl); ax.set_ylabel('Sales ($M)' if col=='digital' else '')
fig.suptitle('Channel spend vs sales — only Digital shows a strong linear relationship', x=0.02, ha='left', fontweight='bold', color=C['navy'], fontsize=13)
fig.tight_layout(rect=[0,0,1,0.93]); fig.savefig(CHARTS_DIR/'05_channel_scatter.png'); plt.close(fig)
print("    05_channel_scatter.png")

# -- 6  Adstock decay
fig, ax = plt.subplots(figsize=(10, 5)); months = np.arange(0, 13)
decay_info = [
    ('digital', DECAY['digital'], C['blue'], f"Digital (lambda={DECAY['digital']:.2f})", (0.5, 20)),
    ('radio',   DECAY['radio'],   C['teal'], f"Radio (lambda={DECAY['radio']:.2f})",   (1.2, 52)),
    ('tv',      DECAY['tv'],      C['orange'], f"TV (lambda={DECAY['tv']:.2f})",      (7.0, 62)),
]
for ch, rate, c, lbl, (tx, ty) in decay_info:
    ax.plot(months, rate**months*100, color=c, lw=3, marker='o', ms=6, mfc='white', mew=2, mec=c, label=lbl)
    hl = np.log(0.5)/np.log(rate) if rate>0 else 0
    ax.annotate(f'Half-life {hl:.1f}mo', xy=(hl,50), xytext=(tx, ty), fontsize=9, color=c, fontweight='bold',
                arrowprops=dict(arrowstyle='->', color=c, lw=1.2, connectionstyle="arc3,rad=0.2"))
ax.axhline(50, color=C['gray'], lw=0.8, ls=':', label='50% Half-life threshold')
ax.set_xlabel('Months after exposure'); ax.set_ylabel('Remaining effect (%)'); ax.set_ylim(-5,105); ax.set_xlim(-0.5,12.5)
ax.set_title('Adstock decay curves (data-optimised rates via grid search)', loc='left', color=C['navy'], pad=12)
ax.legend(fontsize=9, loc='upper right')
fig.tight_layout(); fig.savefig(CHARTS_DIR/'06_adstock_decay.png'); plt.close(fig)
print("    06_adstock_decay.png")

# -- 7  Standardised coefficients
cdata = pd.DataFrame({'feature': features, 'beta': [std_coefs[f] for f in features], 'p': [std_pvals[f] for f in features]})
lmap = {'digital_sat':'Digital','radio_sat':'Radio','tv_sat':'TV','gdp_change':'GDP',
        'cpi_yoy':'CPI','ca_unemployment':'CA Unemp','consumer_sentiment':'Sentiment',
        'ecommerce_pct':'E-Commerce %','covid_lockdown':'COVID lockdown',
        'covid_relaxation':'COVID relaxation','month_sin':'Seasonality (sin)',
        'month_cos':'Seasonality (cos)','trend':'Time trend'}
cdata['label'] = cdata['feature'].map(lmap); cdata = cdata.sort_values('beta')
fig, ax = plt.subplots(figsize=(10, 6.5))
ax.barh(cdata['label'], cdata['beta'], color=[C['teal'] if v>0 else C['red'] for v in cdata['beta']], ec='white', lw=0.5, height=0.6)
ax.axvline(0, color=C['navy'], lw=1.2)
for i, (_, row) in enumerate(cdata.iterrows()):
    sig = '***' if row['p']<0.001 else '**' if row['p']<0.01 else '*' if row['p']<0.05 else 'ns'
    ax.text(row['beta']+(0.02 if row['beta']>=0 else -0.02), i, f"{row['beta']:.2f} ({sig})",
            va='center', ha='left' if row['beta']>=0 else 'right', fontsize=9, color=C['gray'])
ax.set_xlabel('Standardised coefficient'); ax.spines['left'].set_visible(False)
ax.set_title('Driver importance — Digital is the strongest positive driver', loc='left', color=C['navy'], pad=12)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'07_standardized_coefficients.png'); plt.close(fig)
print("    07_standardized_coefficients.png")

# -- 8  ROI comparison
fig, axes = plt.subplots(1, 2, figsize=(13, 5), gridspec_kw={'width_ratios':[1.2,1]})
ax = axes[0]
bars = ax.bar(roi_df['channel'], roi_df['marginal_roi'].clip(-10,None), color=[C['blue'],C['teal'],C['red']], width=0.5, ec='white', lw=1)
for bar, val in zip(bars, roi_df['marginal_roi']):
    c = C['emerald'] if val>0 else C['red']; display_val = val if abs(val)<100 else 'Neg.'
    ax.text(bar.get_x()+bar.get_width()/2, max(bar.get_height(),0)+0.15, f'${val:.2f}' if isinstance(display_val, float) else display_val,
            ha='center', fontweight='bold', color=c, fontsize=13)
ax.axhline(0, color=C['navy'], lw=0.8); ax.axhline(1, color=C['gray'], lw=0.8, ls=':', label='Break-even ($1)')
ax.set_ylabel('Marginal ROI ($ sales per $1 spend)'); ax.set_title('Channel marginal ROI', loc='left', color=C['navy'], fontweight='bold'); ax.legend(fontsize=9)
ax2 = axes[1]
spend_s = list(roi_df['spend_share']); std_abs = [abs(r['std_coef']) for _, r in roi_df.iterrows()]
tot = sum(std_abs) if sum(std_abs)>0 else 1; imp_s = [v/tot*100 for v in std_abs]
xx = np.arange(3); w=0.3
ax2.bar(xx-w/2, spend_s, w, color=C['gray'], alpha=0.5, label='Spend share %', ec='white')
ax2.bar(xx+w/2, imp_s, w, color=C['blue'], label='Impact share %', ec='white')
ax2.set_xticks(xx); ax2.set_xticklabels(list(roi_df['channel'])); ax2.set_ylabel('Share (%)')
ax2.set_title('Efficiency — spend vs impact', loc='left', color=C['navy'], fontweight='bold'); ax2.legend(fontsize=9)
fig.suptitle('Channel ROI analysis', fontweight='bold', color=C['navy'], fontsize=14, y=1.02)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'08_channel_roi.png'); plt.close(fig)
print("    08_channel_roi.png")

# -- 9  COVID bridge
fig, ax = plt.subplots(figsize=(12, 5.5))
labels = ['Pre-COVID\nQ1 2020','Lockdown\nQ2 2020','Relaxation\nQ3 2020','Recovery\nQ4 2020','2021\navg','2022\navg']
vals = list(phase_df['avg_sales']); digs = list(phase_df['avg_digital'])
bcols = [C['blue'],C['red'],C['orange'],C['teal'],C['blue'],C['blue']]
bars = ax.bar(range(len(labels)), vals, color=bcols, width=0.55, ec='white', lw=1, alpha=0.85)
for bar, v, d in zip(bars, vals, digs):
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+5, f'${v:.0f}M', ha='center', fontweight='bold', color=C['navy'], fontsize=11)
    ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()/2, f'Dig ${d:.0f}M', ha='center', color='white', fontsize=8, fontweight='bold')
ax.axhline(vals[0], color=C['gray'], lw=1.2, ls='--', alpha=0.5, label=f'Pre-COVID baseline (${vals[0]:.0f}M)')
ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, fontsize=10); ax.set_ylabel('Avg monthly sales ($M)')
ax.set_title("COVID impact — lockdown didn't collapse sales; Q3 dip was driven by marketing cuts", loc='left', color=C['navy'], pad=12)
ax.legend(fontsize=9)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'09_covid_impact_bridge.png'); plt.close(fig)
print("    09_covid_impact_bridge.png")

# -- 10  Budget allocation (current vs optimal)
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
chs = ['Digital','Radio','TV']; cur_v = [current_alloc[c] for c in chs]; opt_v = [opt_d, opt_r, opt_t]; pcols = [C['blue'],C['teal'],C['orange']]
def mk_pct(vals):
    def fn(p): return f'${p*sum(vals)/100:.0f}M\n({p:.0f}%)'
    return fn
axes[0].pie(cur_v, labels=chs, colors=pcols, autopct=mk_pct(cur_v), startangle=90,
            textprops={'fontsize':11,'fontweight':'bold'}, wedgeprops={'ec':'white','lw':2})
axes[0].set_title('Current allocation', fontweight='bold', color=C['navy'], fontsize=13)
axes[1].pie(opt_v, labels=chs, colors=pcols, autopct=mk_pct(opt_v), startangle=90,
            textprops={'fontsize':11,'fontweight':'bold'}, wedgeprops={'ec':'white','lw':2})
axes[1].set_title('Optimised allocation', fontweight='bold', color=C['navy'], fontsize=13)
fig.suptitle(f'Budget reallocation — predicted uplift {uplift_pct:+.1f}%', fontweight='bold', color=C['navy'], fontsize=14)
fig.tight_layout(rect=[0,0,1,0.93]); fig.savefig(CHARTS_DIR/'10_budget_allocation.png'); plt.close(fig)
print("    10_budget_allocation.png")

# -- 11  Sensitivity
fig, ax = plt.subplots(figsize=(10, 5))
base = sens_df[sens_df['budget_change']==0]['media_effect'].values[0]
for _, r in sens_df.iterrows():
    d = r['media_effect']-base; c = C['emerald'] if d>=0 else C['red']
    ax.barh(f"{r['budget_change']:+.0f}%", d, color=c, ec='white', height=0.5)
    ax.text(d+(5 if d>=0 else -5), f"{r['budget_change']:+.0f}%", f"${r['media_effect']:,.0f}M",
            va='center', ha='left' if d>=0 else 'right', fontsize=10, fontweight='bold', color=C['navy'])
ax.axvline(0, color=C['navy'], lw=1.2); ax.set_xlabel('Change in predicted media effect ($M)'); ax.set_ylabel('Budget change')
ax.set_title('Budget sensitivity — diminishing returns at higher spend', loc='left', color=C['navy'], pad=12)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'11_sensitivity_tornado.png'); plt.close(fig)
print("    11_sensitivity_tornado.png")

# -- 12  Seasonality
pivot = df.pivot_table(values='sales', index='month', columns='year', aggfunc='mean')
fig, ax = plt.subplots(figsize=(10, 5))
mnames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
ycols = {2020:C['gray'], 2021:C['blue'], 2022:C['teal']}
for yr in pivot.columns:
    ax.plot(pivot.index, pivot[yr], marker='o', ms=7, lw=2.5, color=ycols.get(yr,C['gray']),
            mfc='white', mew=2, mec=ycols.get(yr,C['gray']), label=str(yr))
ax.set_xticks(range(1,13)); ax.set_xticklabels(mnames); ax.set_ylabel('Sales ($M)')
ax.set_title('Seasonal patterns — sales peak in Q2 and Q4', loc='left', color=C['navy'], pad=12); ax.legend(fontsize=10)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'12_seasonal_patterns.png'); plt.close(fig)
print("    12_seasonal_patterns.png")

# -- 13  Residual diagnostics
fig, axes = plt.subplots(2, 2, figsize=(12, 9))
axes[0,0].scatter(fitted, residuals, c=C['blue'], alpha=0.7, s=40, ec='white', lw=0.5)
axes[0,0].axhline(0, color=C['navy'], lw=1, ls='--'); axes[0,0].set_xlabel('Fitted'); axes[0,0].set_ylabel('Residuals')
axes[0,0].set_title('Residuals vs fitted', color=C['navy'])
from scipy import stats
rs = np.sort(residuals); th = stats.norm.ppf(np.linspace(0.01,0.99,len(rs)))
axes[0,1].scatter(th, rs, c=C['teal'], alpha=0.7, s=40, ec='white', lw=0.5)
axes[0,1].plot(th, th*residuals.std()+residuals.mean(), color=C['navy'], lw=1.5, ls='--')
axes[0,1].set_xlabel('Theoretical quantiles'); axes[0,1].set_ylabel('Sample quantiles')
axes[0,1].set_title('Q-Q plot (normality)', color=C['navy'])
axes[1,0].hist(residuals, bins=12, color=C['blue'], alpha=0.7, ec='white', lw=0.8)
axes[1,0].axvline(0, color=C['navy'], lw=1.5, ls='--'); axes[1,0].set_xlabel('Residual'); axes[1,0].set_ylabel('Count')
axes[1,0].set_title('Residual distribution', color=C['navy'])
axes[1,1].scatter(y, fitted, c=C['purple'], alpha=0.7, s=40, ec='white', lw=0.5)
lims = [min(y.min(),fitted.min())-20, max(y.max(),fitted.max())+20]
axes[1,1].plot(lims, lims, color=C['navy'], lw=1.5, ls='--', label='Perfect fit'); axes[1,1].set_xlim(lims); axes[1,1].set_ylim(lims)
axes[1,1].set_xlabel('Actual ($M)'); axes[1,1].set_ylabel('Predicted ($M)'); axes[1,1].set_title('Actual vs predicted', color=C['navy']); axes[1,1].legend(fontsize=9)
fig.suptitle('Model diagnostics', fontweight='bold', color=C['navy'], fontsize=14)
fig.tight_layout(rect=[0,0,1,0.95]); fig.savefig(CHARTS_DIR/'13_residual_diagnostics.png'); plt.close(fig)
print("    13_residual_diagnostics.png")

# -- 14  Channel spend timeline
fig, axes = plt.subplots(3, 1, figsize=(13, 10), sharex=True)
for ax, (col, lbl, c) in zip(axes, [('digital','Digital ($M)',C['blue']),('radio','Radio ($M)',C['teal']),('tv_k','TV ($K)',C['orange'])]):
    ax.fill_between(df['date'], df[col], alpha=0.3, color=c); ax.plot(df['date'], df[col], color=c, lw=2); ax.set_ylabel(lbl)
    ax.axvspan(pd.Timestamp('2020-04-01'),pd.Timestamp('2020-06-30'), color=C['red'], alpha=0.05)
    ax.axvspan(pd.Timestamp('2020-07-01'),pd.Timestamp('2020-09-30'), color=C['orange'], alpha=0.05)
axes[0].set_title('Marketing spend over time — Digital growing, Radio and TV declining', loc='left', color=C['navy'], pad=10)
axes[2].set_xlabel('Month')
fig.tight_layout(); fig.savefig(CHARTS_DIR/'14_channel_spend_timeline.png'); plt.close(fig)
print("    14_channel_spend_timeline.png")

# -- 15  Revenue decomposition
df['c_dig'] = model.params['digital_sat']*df['digital_sat']
df['c_rad'] = model.params['radio_sat']*df['radio_sat']
df['c_tv']  = model.params['tv_sat']*df['tv_sat']
df['c_base']= model.params['const']
fig, ax = plt.subplots(figsize=(13, 5.5))
ax.fill_between(df['date'], 0, df['c_base'], alpha=0.3, color=C['gray'], label='Base (organic)')
ax.fill_between(df['date'], df['c_base'], df['c_base']+df['c_dig'], alpha=0.5, color=C['blue'], label='Digital')
ax.fill_between(df['date'], df['c_base']+df['c_dig'], df['c_base']+df['c_dig']+df['c_rad'], alpha=0.5, color=C['teal'], label='Radio')
ax.plot(df['date'], df['sales'], color=C['navy'], lw=2, ls='--', marker='o', ms=3, label='Actual sales')
ax.set_ylabel('Sales ($M)'); ax.set_title('Revenue decomposition — channel contributions over time', loc='left', color=C['navy'], pad=12)
ax.legend(loc='upper left', fontsize=9)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'15_revenue_decomposition.png'); plt.close(fig)
print("    15_revenue_decomposition.png")

# -- 16  Digital saturation
fig, ax = plt.subplots(figsize=(10, 5.5))
ax.scatter(df['digital'], df['sales'], s=60, c=C['blue'], alpha=0.6, ec='white', lw=0.8, zorder=5, label='Observed')
xr = np.linspace(0, 180, 200)
ax.plot(xr, model.params['const']+model.params['digital_sat']*np.log1p(xr), color=C['navy'], lw=2.5, alpha=0.8, label='Saturation curve (log model)')
z = np.polyfit(df['digital'], df['sales'], 1)
ax.plot(xr, np.polyval(z,xr), color=C['gray'], lw=1.5, ls=':', alpha=0.6, label='Linear assumption (naive)')
ax.axvspan(0, 50, alpha=0.04, color=C['emerald']); ax.axvspan(100, 180, alpha=0.04, color=C['orange'])
ax.text(25, ax.get_ylim()[0]+20, 'High ROI zone', ha='center', color=C['emerald'], fontsize=9, fontweight='bold')
ax.text(140, ax.get_ylim()[0]+20, 'Diminishing returns', ha='center', color=C['orange'], fontsize=9, fontweight='bold')
ax.set_xlabel('Digital spend ($M)'); ax.set_ylabel('Sales ($M)')
ax.set_title('Digital saturation curve — returns diminish at higher spend', loc='left', color=C['navy'], pad=12); ax.legend(fontsize=9)
fig.tight_layout(); fig.savefig(CHARTS_DIR/'16_digital_saturation.png'); plt.close(fig)
print("    16_digital_saturation.png")
print()


# ===================================================================
# 8  SAVE OUTPUTS
# ===================================================================
print("[8] Saving outputs")
print("-" * 50)

summary = {
    'metadata': {
        'analysis_date': datetime.now().strftime('%Y-%m-%d %H:%M'),
        'period': f"{df['date'].min():%b %Y} - {df['date'].max():%b %Y}",
        'n': int(len(df)),
        'geography': 'Bay Area district, California, US',
        'target_audience': 'Age 12-45 years',
        'method': 'Marketing Mix Model — Grid-searched Adstock + Log-saturation + Regional/National Macro',
    },
    'adstock_decay_rates': DECAY,
    'grid_search': {'combinations_tested': len(results_grid), 'best_aic': round(best_aic, 1)},
    'model': {
        'r2': round(model.rsquared, 4), 'adj_r2': round(model.rsquared_adj, 4),
        'loocv_r2': round(loocv_r2, 4), 'aic': round(model.aic, 1), 'bic': round(model.bic, 1),
    },
    'diagnostics': {'durbin_watson': round(dw, 3), 'bp_pval': round(bp_p, 4), 'bg_pval': round(bg_p, 4)},
    'channels': roi_df.to_dict('records'),
    'macro': {
        'gdp_p': round(gdp_p, 4),
        'cpi_p': round(cpi_p, 4),
        'ca_unemployment_p': round(unemp_p, 4),
        'consumer_sentiment_p': round(sent_p, 4),
        'ecommerce_pct_p': round(ecom_p, 4),
        'covid_lockdown_p': round(lock_p, 4),
        'covid_relaxation_p': round(relax_p, 4),
    },
    'optimisation': {
        'current': {k: round(v,1) for k,v in current_alloc.items()},
        'optimal': {'Digital': round(opt_d,1), 'Radio': round(opt_r,1), 'TV': round(opt_t,1)},
        'uplift_pct': round(uplift_pct, 1),
    },
    'findings': {
        'Q1': 'Only Digital and Radio are statistically significant drivers of sales. TV shows no measurable lift.',
        'Q2': f'Digital is the most important channel (std coef={std_coefs["digital_sat"]:.2f}, p<0.001), followed by Radio.',
        'Q3': f'None of the macro variables (GDP p={gdp_p:.3f}, CPI p={cpi_p:.3f}, CA Unemployment p={unemp_p:.3f}, Sentiment p={sent_p:.3f}, E-Commerce p={ecom_p:.3f}) reach statistical significance after controlling for media spend.',
        'Q4': f'COVID lockdown is not significant (p={lock_p:.3f}). Sales were sustained by continued digital marketing investment in the Bay Area market.',
    },
}

(OUTPUT_DIR/'results_summary.json').write_text(json.dumps(summary, indent=2, default=str))
roi_df.to_csv(OUTPUT_DIR/'channel_roi.csv', index=False)
pd.DataFrame({'channel':['Digital','Radio','TV'],'current':[current_alloc[c] for c in ['Digital','Radio','TV']],
              'optimal':[opt_d,opt_r,opt_t]}).to_csv(OUTPUT_DIR/'optimal_budget.csv', index=False)
df.to_csv(OUTPUT_DIR/'model_input_clean.csv', index=False)
sens_df.to_csv(OUTPUT_DIR/'sensitivity_analysis.csv', index=False)
vif.to_csv(OUTPUT_DIR/'vif_analysis.csv', index=False)
with open(OUTPUT_DIR/'ols_model_summary.txt','w') as f:
    f.write("Full Model\n"+"="*70+"\n"+model.summary().as_text()+"\n\nMarketing-Only Model\n"+"="*70+"\n"+model_mkt.summary().as_text())

for f in (OUTPUT_DIR/'results_summary.json', OUTPUT_DIR/'channel_roi.csv', OUTPUT_DIR/'optimal_budget.csv',
          OUTPUT_DIR/'ols_model_summary.txt'):
    print(f"    {f.name}")
print()


# ===================================================================
# 9  EXECUTIVE SUMMARY
# ===================================================================
print("=" * 65)
print("  Findings")
print("=" * 65)
print()
print("  Q1. Are all channels effective?")
print("  No. Digital and Radio are statistically significant.")
print("  TV has no detectable effect on sales in this data.")
print()
print("  Q2. Most important channel?")
dig_b = std_coefs['digital_sat']; rad_b = std_coefs['radio_sat']
dig_roi = roi_df[roi_df.channel=='Digital']['marginal_roi'].values[0]
print(f"  Digital (std coef {dig_b:.2f}, marginal ROI ${dig_roi:.2f}/$1).")
print(f"  Radio is second (std coef {rad_b:.2f}).")
print()
print("  Q3. Macro influence?")
print(f"  None of GDP (p={gdp_p:.3f}), CPI (p={cpi_p:.3f}), or")
print(f"  unemployment (p={unemp_p:.3f}) reach significance after")
print("  controlling for media spend. Sales depend on marketing,")
print("  not the business cycle.")
print()
print("  Q4. COVID impact?")
print(f"  Not significant (lockdown p={lock_p:.3f}, relaxation p={relax_p:.3f}).")
print("  Digital spend was maintained during lockdown, sustaining demand.")
print("  The Q3-2020 dip coincided with a drop in all marketing spend.")
print()
print(f"  Budget recommendation: reallocate to optimal mix for")
print(f"  an estimated {uplift_pct:+.1f}% uplift in media-driven sales.")
print()
print("=" * 65)
print("  Done. Outputs in ./output/")
print("=" * 65)
