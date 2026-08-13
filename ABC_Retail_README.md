# ABC Retail Sales Analysis Package

## Deliverables
- `ABC_Retail_Marketing_Intelligence.pptx`: executive presentation.
- `ABC_Marketing_Intelligence_Prototype.html`: standalone clickable SaaS-style prototype. Open in a browser.
- `analysis.py`: reproducible Python analysis and chart generation.
- `analysis_results.json`: machine-readable model outputs.
- `model_input_clean.csv`: cleaned monthly modeling table.
- `model_coefficients.csv`: coefficient, confidence interval and significance outputs.

## Analytical approach
1. Standardized the supplied units, including converting TV from thousands to millions for modeling.
2. Joined quarterly GDP to monthly records.
3. Added Q2 2020 lockdown, Q3 2020 relaxation and time-trend features.
4. Estimated OLS with HC3 robust standard errors for inference.
5. Compared standardized coefficients for relative importance.
6. Used leave-one-out cross-validation with Ridge regression as a predictive sanity check.
7. Separated evidence from recommendation: association is not assumed to be causation.

## How to run
```bash
python analysis.py
```

## Key caveat
The dataset has only 36 monthly observations and no geography, price, promotion, seasonality, distribution, competitor, customer or experimental controls. Recommendations should be treated as test-and-learn priorities rather than final causal budget allocations.
