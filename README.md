# ABC Retail — Marketing Mix Modeling (MMM) & Budget Optimization

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/)
[![Statsmodels](https://img.shields.io/badge/statsmodels-OLS--HC3-green.svg)](https://www.statsmodels.org/)
[![SciPy](https://img.shields.io/badge/SciPy-SLSQP--Optimization-orange.svg)](https://scipy.org/)
[![Live Dashboard](https://img.shields.io/badge/Live--Dashboard-GitHub--Pages-brightgreen.svg)](https://LearnWithAbinash.github.io/Case_Study_ABC_Retail_Sales_Analysis/)

> **Executive Portfolio Project & Case Study Submission**  
> **Target Market:** Bay Area District, California, US \| **Demographic:** Age Group 12–45 Years  
> *Note: ABC Retail is a hypothetical company name and all data used in this project are sample/anonymized datasets created for demonstration purposes.*

---

## 🌐 Live Interactive Dashboard

Access the live interactive SaaS marketing intelligence dashboard:  
👉 **[View Live Web Dashboard](https://LearnWithAbinash.github.io/Case_Study_ABC_Retail_Sales_Analysis/)** *(Hosted via GitHub Pages)*

---

## 📌 Problem Statement & Business Context

**ABC Retail** is a leading retail brand operating in the **Bay Area district of California, US**. Over a 3-year period (**36 months: January 2020 – December 2022**), ABC experienced subdued revenue growth despite investing **$3,033 Million** across 3 major marketing channels:

- **Digital Advertising:** $2,633 Million (86.8% spend share)
- **Radio Advertising:** $399 Million (13.2% spend share)
- **TV Advertising:** $0.6 Million (0.02% spend share — recorded in thousands)

### Key Executive Questions
1. **Channel Effectiveness:** Are all media channels effective in driving revenue?
2. **Channel Hierarchy & ROI:** Which channel is the most important, and what is the marginal ROI per channel?
3. **Macroeconomic Influence:** Do macro factors (GDP, CPI Inflation, CA Unemployment, Sentiment) affect sales?
4. **COVID-19 Impact:** Did COVID-19 lockdowns (Q2 2020) and relaxations (Q3 2020) structurally damage sales?

---

## 🔑 Key Findings & Executive Answers

### Question 1: Are all media channels effective?
**NO.**
- **Digital** ($\beta^* = +0.815, p < 0.001$) and **Radio** ($\beta^* = +0.573, p < 0.001$) are statistically significant positive revenue drivers.
- **TV** ($\beta^* = -0.355, p = 0.315$) shows **no statistically significant effect**. TV spend is recorded in thousands ($0.6M over 3 years), making it a rounding error in the marketing portfolio.

### Question 2: Which is the most important channel?
**DIGITAL** is the #1 driver by absolute sales volume and standardized effect ($\beta^* = 0.815$).
- However, **Radio** has a significantly higher **Marginal ROI ($7.07 per $1 spend vs. $1.32 for Digital)**.
- **Why?** Digital spend ($2,633M) has reached high **diminishing returns (saturation)**, whereas Radio ($399M) remains under-saturated.

### Question 3: Do macroeconomic factors influence sales?
**NO statistically significant influence.**
- Tested macro variables: National GDP % change ($p = 0.584$), CPI Inflation ($p = 0.669$), California Unemployment Rate ($p = 0.700$), Consumer Sentiment ($p = 0.558$), and E-Commerce Penetration % ($p = 0.851$).
- **Business Insight:** ABC operates in resilient retail in the affluent Bay Area. Sales depend on **marketing execution and brand awareness**, not macro economic cycles.

### Question 4: Did COVID-19 impact sales, and why?
**NOT SIGNIFICANTLY** ($p = 0.990$ for Lockdown).
- During lockdown (Q2 2020), ABC **increased Digital marketing spend ($57M/mo vs $49M/mo pre-COVID)**, successfully capturing online and curbside demand.
- The sales dip in Q3 2020 ($179M/mo) was caused by **cutting marketing spend down to $31M/mo**, not the pandemic restrictions.

---

## 📈 Budget Optimization & Revenue Uplift

Using SciPy SLSQP constrained optimization subject to fixed total budget ($\sum x_k = \$3,033\text{M}$):

| Channel | Current Allocation | Optimal Allocation | Spend Shift | Strategy |
| :--- | :---: | :---: | :---: | :--- |
| **Digital** | $2,633M (86.8%) | **$1,634M (53.9%)** | -$999M | Reduce to overcome saturation |
| **Radio** | $399M (13.2%) | **$1,399M (46.1%)** | +$1,000M | Scale up to capture $7.07/$1 ROI |
| **TV** | $0.6M (0.02%) | **$0.0M (0.0%)** | -$0.6M | Reallocate to Radio |
| **TOTAL** | **$3,033M** | **$3,033M** | **$0M** | **+31.7% Media Revenue Uplift** |

---

## 🛠️ Methodology & Statistical Rigor

- **Adstock Decay (Carryover Effect):** Geometric decay rates optimized via 900-combination grid search to minimize AIC:
  - Digital $\lambda = 0.10$ (Half-life = 0.3 months)
  - Radio $\lambda = 0.30$ (Half-life = 0.6 months)
  - TV $\lambda = 0.90$ (Half-life = 6.6 months)
- **Diminishing Returns:** Log-saturation transformation $S_t = \ln(1 + \text{Adstock}_t)$.
- **Small-Sample Protection ($n = 36$):** OLS with **HC3 Robust Standard Errors** (MacKinnon & White) to ensure non-inflated p-values.
- **Cross-Validation:** **LOOCV $R^2 = 0.6973$** (Out-of-sample Leave-One-Out validation proves model does not overfit).

---

## 📂 Public Repository Structure

```
Case_Study_ABC_Retail_Sales_Analysis/
├── index.html                            <-- Live Web Dashboard (GitHub Pages)
├── abc_retail_mmm_analysis.py            <-- Python MMM Script (Grid Search & Regression)
├── generate_pptx.py                      <-- PowerPoint Deck Generator Script
├── ABC_Retail_Case_Study_Presentation.pptx <-- Executive 12-Slide Deck
├── README.md                             <-- Project Portfolio Page
└── output/
    ├── results_summary.json              <-- Machine-Readable Model Results
    ├── channel_roi.csv                   <-- Channel ROI Summary Table
    ├── optimal_budget.csv                <-- Budget Allocation Comparison Table
    ├── ols_model_summary.txt             <-- Full Statsmodels OLS Regression Output
    └── charts/                           <-- 16 High-Resolution PNG Charts
```

---

## 🚀 How to Run Locally

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/abc-retail-mmm-casestudy.git
cd abc-retail-mmm-casestudy

# Install dependencies
pip install pandas numpy matplotlib seaborn statsmodels scipy scikit-learn python-pptx openpyxl

# Execute Python MMM analysis & generate outputs
python abc_retail_mmm_analysis.py

# Generate PowerPoint slide deck
python generate_pptx.py
```
