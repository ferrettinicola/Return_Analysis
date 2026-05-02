# 📈 Return Analysis

A collection of quantitative equity analyses built in Python — one stock at a time.  
Each analysis combines a Jupyter notebook for research and a Streamlit dashboard for interactive visualization.

---

## What this repo is about

This project applies standard quantitative finance tools to individual equities, starting with **NVIDIA (NVDA)**.  
The goal is to build a repeatable, modular framework that can be extended to any publicly traded stock.

For each company, the workflow is:

1. **Jupyter Notebook** — downloads price data via `yfinance` and Fama-French factors, computes excess returns, runs descriptive statistics, fits a CAPM regression via OLS, and exports results to Excel.
2. **Excel file** — the output of the notebook, used as the data source for the dashboard (no live internet connection required at runtime).
3. **Streamlit Dashboard** — loads the pre-computed Excel, renders interactive charts and performance metrics using Plotly.

---

## Current analyses

| Company | Ticker | Period | Notebook | Dashboard |
|---|---|---|---|---|
| NVIDIA Corporation | NVDA | Jan 2016 – Apr 2026 | `NVIDIA/NVIDIA.ipynb` | `NVIDIA/dashboard/nvda_dashboard.py` |

---

## Methodology

### Data
- **Prices**: adjusted close prices from Yahoo Finance via `yfinance`
- **Risk-free rate & market factor**: daily Fama-French factors (`Mkt-RF`, `RF`) from Kenneth French's data library

### Metrics computed
- **Distributional moments**: mean, standard deviation, skewness, excess kurtosis
- **Performance**: CAGR, annualized volatility, Sharpe ratio, Sortino ratio
- **Risk**: maximum drawdown, historical VaR at 5% and 1% (daily)
- **CAPM regression** (OLS): alpha, beta, R², t-statistics, p-values, Durbin-Watson
- **Rolling beta**: 252-day window to capture time-varying market sensitivity

All performance metrics are computed via [`quantstats`](https://github.com/ranaroussi/quantstats), ensuring consistency between the notebook output and the dashboard display.

---

## Repository structure

```
Return_Analysis/
└── NVIDIA/
    ├── NVIDIA.ipynb          # Research notebook
    ├── nvda_analysis.xlsx    # Output of the notebook (data source for dashboard)
    └── dashboard/
        ├── nvda_dashboard.py # Streamlit app
        ├── nvda_analysis.xlsx
        └── requirements.txt
```

---

## How to run

### 1 — Notebook
Open `NVIDIA/NVIDIA.ipynb` in Jupyter and run all cells.  
This will generate `nvda_analysis.xlsx` with four sheets: `NVDA_Prices`, `FamaFrench_3F`, `ER_Analysis`, `OLS_Results`.

### 2 — Dashboard (local)
```bash
pip install -r NVIDIA/dashboard/requirements.txt
streamlit run NVIDIA/dashboard/nvda_dashboard.py
```

### 2 — Dashboard (Streamlit Cloud)
The app is deployed on Streamlit Cloud. Dependencies are declared in `requirements.txt` and picked up automatically.  
The dashboard reads `nvda_analysis.xlsx` from its own folder — no live data download at runtime.

---

## Stack

`Python` · `pandas` · `numpy` · `yfinance` · `statsmodels` · `quantstats` · `scipy` · `plotly` · `streamlit` · `openpyxl`

---

## Roadmap

- [ ] Apple (AAPL)
- [ ] Microsoft (MSFT)
- [ ] Tesla (TSLA)
- [ ] Multi-stock comparison notebook
- [ ] Fama-French 3-factor and 5-factor extensions

---

## Author

Feel free to open an issue or a pull request if you want to contribute an analysis for another stock.
