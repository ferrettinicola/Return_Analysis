"""
NVIDIA (NVDA) — Investment Analysis Dashboard
Carica i dati da nvda_analysis.xlsx (generato dal notebook NVIDIA.ipynb)

Run with: streamlit run nvda_dashboard.py
"""

import streamlit as st
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from scipy import stats
from scipy.stats import norm
import statsmodels.api as sm
import quantstats as qs
import warnings
warnings.filterwarnings("ignore")

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="NVDA · Investment Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CUSTOM CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display:ital@0;1&family=DM+Sans:wght@300;400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');

/* ─── Base ─────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: #09090f;
    color: #e8e8ef;
}
.stApp { background: #09090f; }

/* ─── Header strip ──────────────────────────────────── */
.hero {
    background: linear-gradient(135deg, #0d0d1a 0%, #101028 50%, #0a1a2e 100%);
    border: 1px solid rgba(100,120,255,0.18);
    border-radius: 16px;
    padding: 40px 48px 36px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute; inset: 0;
    background: radial-gradient(ellipse at 80% 50%, rgba(118,85,255,0.12) 0%, transparent 65%);
}
.hero-tag {
    font-family: 'JetBrains Mono', monospace;
    font-size: 11px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #7655ff;
    margin-bottom: 10px;
}
.hero-title {
    font-family: 'DM Serif Display', serif;
    font-size: 52px;
    line-height: 1.05;
    margin: 0 0 8px;
    background: linear-gradient(90deg, #fff 0%, #a8b4ff 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
}
.hero-subtitle {
    color: #8888aa;
    font-size: 15px;
    font-weight: 300;
    letter-spacing: 0.3px;
}

/* ─── Section labels ────────────────────────────────── */
.section-label {
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px;
    letter-spacing: 3px;
    text-transform: uppercase;
    color: #7655ff;
    margin: 40px 0 14px;
}

/* ─── KPI cards ─────────────────────────────────────── */
.kpi-card {
    background: linear-gradient(145deg, #111120, #14142a);
    border: 1px solid rgba(118,85,255,0.15);
    border-radius: 14px;
    padding: 22px 24px;
    text-align: left;
    transition: border-color .25s;
}
.kpi-card:hover { border-color: rgba(118,85,255,0.4); }
.kpi-label {
    font-size: 11px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: #66668a;
    margin-bottom: 8px;
}
.kpi-value {
    font-family: 'DM Serif Display', serif;
    font-size: 32px;
    line-height: 1;
    color: #fff;
}
.kpi-value.positive { color: #4ade80; }
.kpi-value.negative { color: #f87171; }
.kpi-value.neutral  { color: #a78bfa; }
.kpi-sub {
    font-size: 11px;
    color: #55556a;
    margin-top: 6px;
}

/* ─── Regression table ──────────────────────────────── */
.reg-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 13px;
}
.reg-table th {
    background: rgba(118,85,255,0.12);
    color: #a78bfa;
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px;
    letter-spacing: 2px;
    text-transform: uppercase;
    padding: 10px 16px;
    text-align: left;
    border-bottom: 1px solid rgba(118,85,255,0.2);
}
.reg-table td {
    padding: 12px 16px;
    border-bottom: 1px solid rgba(255,255,255,0.05);
    color: #d0d0e8;
    font-family: 'JetBrains Mono', monospace;
}
.reg-table tr:last-child td { border-bottom: none; }
.reg-table tr:hover td { background: rgba(118,85,255,0.06); }
.sig { color: #4ade80; font-size: 11px; }

/* ─── Divider ───────────────────────────────────────── */
.fancy-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(118,85,255,0.4), transparent);
    margin: 40px 0;
}

/* ─── Info box ──────────────────────────────────────── */
.info-box {
    background: rgba(118,85,255,0.06);
    border-left: 3px solid #7655ff;
    border-radius: 0 10px 10px 0;
    padding: 16px 20px;
    font-size: 13px;
    color: #9999bb;
    line-height: 1.7;
}
</style>
""", unsafe_allow_html=True)


# ── DATA LOADING DA EXCEL ─────────────────────────────────────────────────────
from pathlib import Path
EXCEL_PATH = Path(__file__).parent / "nvda_analysis.xlsx"  # sempre nella cartella dello script

@st.cache_data(show_spinner=False)
def load_from_excel(path: str):
    """
    Carica i dati pre-calcolati dal notebook (nvda_analysis.xlsx).
    Foglio NVDA_Prices  → prezzi Adj Close
    Foglio ER_Analysis  → Mkt-Rf, ER, RF (excess returns già calcolati)
    """
    xls = pd.ExcelFile(path)

    # Prezzi
    prices = pd.read_excel(xls, sheet_name="NVDA_Prices", index_col=0, parse_dates=True)
    prices.columns = ["Adj Close"]          # rinomina per sicurezza

    # Excess returns + fattori
    df = pd.read_excel(xls, sheet_name="ER_Analysis", index_col=0, parse_dates=True)
    # Le colonne attese sono: Mkt-Rf, ER  (RF è opzionale ma presente)
    if "RF" not in df.columns:
        df["RF"] = 0.0

    return prices, df


# ── METRICS ──────────────────────────────────────────────────────────────────
def compute_metrics(er_series: pd.Series, rf_series: pd.Series):
    """
    Calcola le metriche usando quantstats, identico al notebook NVIDIA.ipynb.
    quantstats vuole i rendimenti in decimale (non in percentuale).
    """
    rets = er_series.dropna() / 100

    if len(rets) == 0:
        nan = float("nan")
        return dict(cagr=nan, vol=nan, skew=nan, kurt=nan,
                    sharpe=nan, sortino=nan, max_dd=nan,
                    var5=nan, var1=nan)

    return dict(
        cagr    = qs.stats.cagr(rets),
        vol     = qs.stats.volatility(rets),
        skew    = float(qs.stats.skew(rets)),
        kurt    = float(qs.stats.kurtosis(rets)),
        sharpe  = float(qs.stats.sharpe(rets)),
        sortino = float(qs.stats.sortino(rets)),
        max_dd  = qs.stats.max_drawdown(rets),
        var5    = qs.stats.value_at_risk(rets, sigma=0.05),
        var1    = qs.stats.value_at_risk(rets, sigma=0.01),
    )


def run_capm(df: pd.DataFrame):
    x = sm.add_constant(df["Mkt-Rf"])
    return sm.OLS(df["ER"], x).fit()


# ── PLOTLY THEME ──────────────────────────────────────────────────────────────
PLOT_BG   = "#09090f"
PAPER_BG  = "#09090f"
GRID_COL  = "rgba(255,255,255,0.05)"
ACCENT    = "#7655ff"
GREEN     = "#4ade80"
RED       = "#f87171"
FONT_FAM  = "DM Sans, sans-serif"

def base_layout(**kwargs):
    return dict(
        plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG,
        font=dict(family=FONT_FAM, color="#9999bb", size=11),
        xaxis=dict(gridcolor=GRID_COL, zeroline=False,
                   showline=False, tickfont=dict(size=10)),
        yaxis=dict(gridcolor=GRID_COL, zeroline=False,
                   showline=False, tickfont=dict(size=10)),
        margin=dict(l=16, r=16, t=48, b=16),
        **kwargs
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════

# ── HERO HEADER ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-tag">Equity Analysis Report · NASDAQ</div>
  <div class="hero-title">NVIDIA Corporation</div>
  <div class="hero-subtitle">
    Performance · Risk · CAPM Regression &nbsp;·&nbsp; Jan 2016 – Apr 2026
  </div>
</div>
""", unsafe_allow_html=True)

# ── LOAD DATA ─────────────────────────────────────────────────────────────────
try:
    with st.spinner("Caricamento dati da nvda_analysis.xlsx…"):
        prices, df = load_from_excel(EXCEL_PATH)
except FileNotFoundError:
    st.error(
        f"❌ File **{EXCEL_PATH}** non trovato. "
        "Esegui prima il notebook `NVIDIA.ipynb` per generarlo, "
        "poi metti `nvda_analysis.xlsx` nella stessa cartella di questo script."
    )
    st.stop()

if df.empty:
    st.error("Il file Excel è vuoto o corrotto. Riesegui il notebook.")
    st.stop()

metrics = compute_metrics(df["ER"], df["RF"])
model   = run_capm(df)
alpha   = model.params["const"]
beta    = model.params["Mkt-Rf"]
r2      = model.rsquared
p_alpha = model.pvalues["const"]
p_beta  = model.pvalues["Mkt-Rf"]

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 · KPI OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">01 · Performance Overview</div>',
            unsafe_allow_html=True)

kpi_data = [
    ("Expected Return (Ann.)", f"{metrics['cagr']*100:.1f}%", "positive",
     "CAGR degli excess return annualizzati"),
    ("Volatility (Ann.)",      f"{metrics['vol']*100:.1f}%",  "negative",
     "Deviazione standard annualizzata"),
    ("Sharpe Ratio",           f"{metrics['sharpe']:.3f}",    "neutral",
     "Rendimento aggiustato per il rischio"),
    ("Sortino Ratio",          f"{metrics['sortino']:.3f}",   "neutral",
     "Penalizza solo la volatilità negativa"),
    ("Max Drawdown",           f"{metrics['max_dd']*100:.1f}%", "negative",
     "Massima perdita dal picco"),
    ("VaR 5% (Daily)",         f"{metrics['var5']*100:.2f}%", "negative",
     "Perdita massima giornaliera al 95%"),
]

cols = st.columns(6)
for col, (label, value, cls, sub) in zip(cols, kpi_data):
    with col:
        st.markdown(f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value {cls}">{value}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

# ── Distribution moments ──────────────────────────────────────────────────────
st.markdown('<div class="section-label" style="margin-top:28px">02 · Momenti della Distribuzione</div>',
            unsafe_allow_html=True)

mom_cols = st.columns(4)
moments = [
    ("Mean (Daily)", f"{df['ER'].mean():.3f}%",   "positive", "Eccesso di rendimento medio"),
    ("Std Dev",      f"{df['ER'].std():.3f}%",    "neutral",  "Deviazione standard giornaliera"),
    ("Skewness",     f"{metrics['skew']:.3f}",    "neutral",  "Asimmetria positiva → code destre"),
    ("Excess Kurtosis", f"{metrics['kurt']:.3f}", "negative", "Code spesse → fat-tail risk"),
]
for col, (label, val, cls, sub) in zip(mom_cols, moments):
    with col:
        st.markdown(f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value {cls}">{val}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 · PRICE & RETURNS CHARTS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">03 · Andamento Prezzi e Rendimenti</div>',
            unsafe_allow_html=True)

c1, c2 = st.columns(2)

with c1:
    fig_price = go.Figure()
    fig_price.add_trace(go.Scatter(
        x=prices.index, y=prices["Adj Close"],
        line=dict(color=ACCENT, width=1.6),
        fill="tozeroy",
        fillcolor="rgba(118,85,255,0.08)",
        name="Adj Close",
    ))
    fig_price.update_layout(
        title=dict(text="Adj Close Price · NVDA (2016–2026)",
                   font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
        **base_layout(),
        yaxis_title="USD ($)",
        hovermode="x unified",
    )
    st.plotly_chart(fig_price, use_container_width=True)

with c2:
    colors = [GREEN if v >= 0 else RED for v in df["ER"]]
    fig_er = go.Figure()
    fig_er.add_trace(go.Bar(
        x=df.index, y=df["ER"],
        marker_color=colors, name="Excess Return",
        marker_line_width=0,
    ))
    fig_er.update_layout(
        title=dict(text="Excess Return Giornaliero · NVDA",
                   font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
        **base_layout(),
        yaxis_title="ER (%)",
        hovermode="x unified",
        bargap=0,
    )
    st.plotly_chart(fig_er, use_container_width=True)

# Drawdown chart ──────────────────────────────────────────────────────────────
rets_dec = df["ER"] / 100
cum      = (1 + rets_dec).cumprod()
roll_max = cum.cummax()
drawdown = (cum - roll_max) / roll_max * 100

fig_dd = go.Figure()
fig_dd.add_trace(go.Scatter(
    x=drawdown.index, y=drawdown,
    line=dict(color=RED, width=1.2),
    fill="tozeroy", fillcolor="rgba(248,113,113,0.1)",
    name="Drawdown",
))
fig_dd.update_layout(
    title=dict(text="Drawdown — NVDA (2016–2026)",
               font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
    **base_layout(),
    yaxis_title="Drawdown (%)",
    hovermode="x unified",
)
st.plotly_chart(fig_dd, use_container_width=True)

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 · DISTRIBUTION ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">04 · Analisi della Distribuzione</div>',
            unsafe_allow_html=True)

c3, c4 = st.columns(2)

with c3:
    er_vals = df["ER"].dropna().values
    x_axis  = np.linspace(er_vals.min(), er_vals.max(), 500)
    mu, sigma = er_vals.mean(), er_vals.std()
    normal_y  = norm.pdf(x_axis, mu, sigma)
    kde_obj   = stats.gaussian_kde(er_vals)
    kde_y     = kde_obj(x_axis)

    fig_hist = go.Figure()
    fig_hist.add_trace(go.Histogram(
        x=er_vals, histnorm="probability density",
        nbinsx=90,
        marker=dict(color="rgba(118,85,255,0.55)", line=dict(color=PLOT_BG, width=0.3)),
        name="ER dist.",
    ))
    fig_hist.add_trace(go.Scatter(
        x=x_axis, y=normal_y,
        line=dict(color=RED, width=2, dash="dash"),
        name="Normale teorica",
    ))
    fig_hist.add_trace(go.Scatter(
        x=x_axis, y=kde_y,
        line=dict(color=GREEN, width=2),
        name="KDE empirica",
    ))
    fig_hist.update_layout(
        title=dict(text="Distribuzione Excess Returns + KDE vs Normale",
                   font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
        **base_layout(),
        barmode="overlay",
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
        yaxis_title="Densità",
        xaxis_title="ER (%)",
    )
    st.plotly_chart(fig_hist, use_container_width=True)

with c4:
    (osm, osr), (slope, intercept, r) = stats.probplot(er_vals, dist="norm")
    qq_x = np.array([osm.min(), osm.max()])
    qq_y = slope * qq_x + intercept

    fig_qq = go.Figure()
    fig_qq.add_trace(go.Scatter(
        x=osm, y=osr,
        mode="markers",
        marker=dict(color=ACCENT, size=3, opacity=0.6),
        name="Quantili osservati",
    ))
    fig_qq.add_trace(go.Scatter(
        x=qq_x, y=qq_y,
        line=dict(color=RED, width=2, dash="dash"),
        name="Linea normale",
    ))
    fig_qq.update_layout(
        title=dict(text="Q-Q Plot — NVDA Excess Returns",
                   font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
        **base_layout(),
        xaxis_title="Quantili teorici (Normale)",
        yaxis_title="Quantili osservati",
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    )
    st.plotly_chart(fig_qq, use_container_width=True)

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 · CAPM REGRESSION
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">05 · CAPM Regression Analysis</div>',
            unsafe_allow_html=True)

col_scatter, col_results = st.columns([1, 1])

with col_scatter:
    sample = df.sample(min(len(df), 1000), random_state=42).sort_index()
    x_fit  = np.linspace(df["Mkt-Rf"].min(), df["Mkt-Rf"].max(), 300)
    y_fit  = alpha + beta * x_fit

    fig_sc = go.Figure()
    fig_sc.add_trace(go.Scatter(
        x=sample["Mkt-Rf"], y=sample["ER"],
        mode="markers",
        marker=dict(color=ACCENT, size=4, opacity=0.45,
                    line=dict(color="rgba(255,255,255,0.1)", width=0.5)),
        name="Osservazioni",
    ))
    fig_sc.add_trace(go.Scatter(
        x=x_fit, y=y_fit,
        line=dict(color=RED, width=2.5),
        name=f"SML fit (β={beta:.2f})",
    ))
    fig_sc.add_hline(y=0, line=dict(color="rgba(255,255,255,0.12)", dash="dot", width=1))
    fig_sc.add_vline(x=0, line=dict(color="rgba(255,255,255,0.12)", dash="dot", width=1))
    fig_sc.update_layout(
        title=dict(text="NVDA vs Mercato — Scatter CAPM",
                   font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
        **base_layout(),
        xaxis_title="Mkt-RF (Excess Return Mercato, %)",
        yaxis_title="ER NVDA (%)",
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    )
    st.plotly_chart(fig_sc, use_container_width=True)

with col_results:
    st.markdown("#### Risultati OLS — CAPM")

    def stars(p):
        if p < 0.001: return "***"
        if p < 0.01:  return "**"
        if p < 0.05:  return "*"
        return ""

    st.markdown(f"""
    <table class="reg-table">
      <thead>
        <tr>
          <th>Parametro</th><th>Coeff.</th><th>Std Err</th>
          <th>t-stat</th><th>P-value</th><th>Sig.</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td>Alpha (α)</td>
          <td>{model.params['const']:.4f}</td>
          <td>{model.bse['const']:.4f}</td>
          <td>{model.tvalues['const']:.3f}</td>
          <td>{model.pvalues['const']:.4f}</td>
          <td class="sig">{stars(model.pvalues['const'])}</td>
        </tr>
        <tr>
          <td>Beta (β)</td>
          <td>{model.params['Mkt-Rf']:.4f}</td>
          <td>{model.bse['Mkt-Rf']:.4f}</td>
          <td>{model.tvalues['Mkt-Rf']:.3f}</td>
          <td>{model.pvalues['Mkt-Rf']:.4f}</td>
          <td class="sig">{stars(model.pvalues['Mkt-Rf'])}</td>
        </tr>
      </tbody>
    </table>
    <br>
    <table class="reg-table">
      <thead><tr><th>Goodness-of-Fit</th><th>Valore</th></tr></thead>
      <tbody>
        <tr><td>R²</td><td>{model.rsquared:.4f}</td></tr>
        <tr><td>R² Adjusted</td><td>{model.rsquared_adj:.4f}</td></tr>
        <tr><td>F-statistic</td><td>{model.fvalue:.1f}</td></tr>
        <tr><td>Prob(F-stat)</td><td>{model.f_pvalue:.2e}</td></tr>
        <tr><td>Durbin-Watson</td><td>{sm.stats.stattools.durbin_watson(model.resid):.3f}</td></tr>
        <tr><td>N. Osservazioni</td><td>{int(model.nobs)}</td></tr>
      </tbody>
    </table>
    <br>
    <div style="font-size:11px;color:#55556a;">
      *** p&lt;0.001 &nbsp;·&nbsp; ** p&lt;0.01 &nbsp;·&nbsp; * p&lt;0.05
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    interp_alpha = "positivo e statisticamente significativo (α = {:.4f}, p = {:.4f}): NVDA ha generato un excess return sistematico rispetto al CAPM.".format(alpha, p_alpha)
    interp_beta  = "β = {:.2f} indica che NVDA è un titolo ad alta sensibilità al mercato: per ogni 1% di movimento del mercato, NVDA si muove mediamente del {:.1f}%.".format(beta, beta)
    st.markdown(f"""
    <div class="info-box">
      <b style="color:#a78bfa">Alpha:</b> {interp_alpha}<br><br>
      <b style="color:#a78bfa">Beta:</b> {interp_beta}<br><br>
      <b style="color:#a78bfa">R²:</b> Il modello CAPM spiega il {r2*100:.1f}% della varianza dei rendimenti di NVDA.
      Il restante {(1-r2)*100:.1f}% è attribuibile a fattori idiosincratici.
    </div>
    """, unsafe_allow_html=True)

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 · ROLLING BETA
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">06 · Rolling Beta (252 giorni)</div>',
            unsafe_allow_html=True)

window = 252
rolling_betas = []
for i in range(window, len(df) + 1):
    chunk = df.iloc[i - window: i]
    x_  = sm.add_constant(chunk["Mkt-Rf"])
    b_  = sm.OLS(chunk["ER"], x_).fit().params["Mkt-Rf"]
    rolling_betas.append((df.index[i - 1], b_))

rb_df = pd.DataFrame(rolling_betas, columns=["Date", "Beta"]).set_index("Date")

fig_rb = go.Figure()
fig_rb.add_hline(y=1, line=dict(color="rgba(255,255,255,0.2)", dash="dot", width=1))
fig_rb.add_trace(go.Scatter(
    x=rb_df.index, y=rb_df["Beta"],
    line=dict(color=ACCENT, width=1.8),
    fill="tozeroy", fillcolor="rgba(118,85,255,0.07)",
    name="Beta (rolling 252gg)",
))
fig_rb.update_layout(
    title=dict(text="Beta Rolling 1 Anno — NVDA vs Market",
               font=dict(family="DM Serif Display, serif", size=16, color="#d0d0e8")),
    **base_layout(),
    yaxis_title="Beta",
    hovermode="x unified",
)
st.plotly_chart(fig_rb, use_container_width=True)

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center; color:#33334a; font-size:11px; padding:24px 0 8px;
            font-family:'JetBrains Mono',monospace; letter-spacing:1px;">
  NVDA · INVESTMENT ANALYSIS DASHBOARD &nbsp;·&nbsp;
  FONTE DATI: NVIDIA.ipynb → nvda_analysis.xlsx &nbsp;·&nbsp;
  PERIODO 2016–2026
</div>
""", unsafe_allow_html=True)
