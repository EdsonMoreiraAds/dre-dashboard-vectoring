import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import pandas as pd
import openpyxl
import os

# ── Configuração da página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="DRE Dashboard — Vectoring",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS personalizado ─────────────────────────────────────────────────────────
st.markdown("""
<style>
  /* Fundo geral */
  .stApp { background-color: #0f1117; }
  [data-testid="stAppViewContainer"] { background-color: #0f1117; }
  [data-testid="stHeader"] { background: transparent; }

  /* KPI cards */
  .kpi-card {
    background: #1a1d27;
    border: 1px solid #2a2d3e;
    border-radius: 12px;
    padding: 18px 20px;
    margin-bottom: 4px;
    position: relative;
  }
  .kpi-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 3px;
    background: var(--accent, #6c63ff);
    border-radius: 12px 12px 0 0;
  }
  .kpi-label { font-size: 11px; color: #8892a4; text-transform: uppercase; letter-spacing: .8px; margin-bottom: 6px; }
  .kpi-value { font-size: 24px; font-weight: 800; color: #e2e8f0; line-height: 1; margin-bottom: 4px; }
  .kpi-sub   { font-size: 12px; }
  .badge-up   { color: #22c55e; }
  .badge-down { color: #ef4444; }
  .badge-neutral { color: #8892a4; }

  /* Section titles */
  .section-title {
    font-size: 11px; font-weight: 700; text-transform: uppercase;
    letter-spacing: 1.2px; color: #8892a4;
    border-bottom: 1px solid #2a2d3e;
    padding-bottom: 8px; margin-top: 32px; margin-bottom: 16px;
  }

  /* Insight box */
  .insight {
    background: #1a1d27;
    border: 1px solid #2a2d3e;
    border-left: 4px solid var(--c, #6c63ff);
    border-radius: 12px;
    padding: 16px 18px;
    margin-bottom: 8px;
  }
  .insight-title { font-size: 13px; font-weight: 700; color: #e2e8f0; margin-bottom: 6px; }
  .insight-body  { font-size: 12px; color: #8892a4; line-height: 1.6; }

  /* Header */
  .main-header {
    background: linear-gradient(135deg, #1a1d27 0%, #12141f 100%);
    border: 1px solid #2a2d3e;
    border-radius: 16px;
    padding: 24px 32px;
    margin-bottom: 24px;
    display: flex; align-items: center; gap: 16px;
  }
  .logo-icon {
    width: 48px; height: 48px; border-radius: 12px;
    background: linear-gradient(135deg, #6c63ff, #00d4aa);
    display: flex; align-items: center; justify-content: center;
    font-size: 22px; font-weight: 900; color: #fff;
    flex-shrink: 0;
  }
  .header-text h1 { font-size: 22px; font-weight: 800; color: #e2e8f0; margin: 0; }
  .header-text p  { font-size: 13px; color: #8892a4; margin: 0; }

  /* Streamlit overrides */
  .stDataFrame { background: #1a1d27 !important; }
  div[data-testid="metric-container"] {
    background: #1a1d27; border: 1px solid #2a2d3e; border-radius: 12px; padding: 16px;
  }
  [data-testid="stMetricValue"] { color: #e2e8f0 !important; }
  [data-testid="stMetricDelta"] svg { display: none; }
  .element-container .stMarkdown h3 { color: #e2e8f0; }
  .stTabs [data-baseweb="tab-list"] { background: #1a1d27; border-radius: 10px; padding: 4px; }
  .stTabs [data-baseweb="tab"] { color: #8892a4; }
  .stTabs [aria-selected="true"] { background: #6c63ff; color: #fff; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)


# ── Leitura dos dados ─────────────────────────────────────────────────────────
@st.cache_data
def carregar_dados():
    base_dir = os.path.dirname(__file__)
    arquivo  = os.path.join(base_dir, "DRE - Cliente Vectoring.xlsx")
    wb = openpyxl.load_workbook(arquivo, read_only=True, data_only=True)
    ws = wb["DRE"]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    def ext(row_idx, start=2, end=26):
        return [v if v is not None else 0 for v in rows[row_idx][start:end]]

    def base_val(row_idx):
        v = rows[row_idx][1]
        return v if v is not None else 0

    meses_ano1 = ["Dez","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    meses_ano2 = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    todos_meses = [f"ANO1/{m}" for m in meses_ano1] + [f"ANO2/{m}" for m in meses_ano2]

    return {
        "meses":        todos_meses,
        "meses_ano1":   [f"ANO1/{m}" for m in meses_ano1],
        "meses_ano2":   [f"ANO2/{m}" for m in meses_ano2],
        "receita":      ext(10),
        "investimento": ext(3),
        "roas":         ext(5),
        "pedidos":      ext(8),
        "ticket":       base_val(12),
        "custo_prod":   ext(16),
        "custo_var":    ext(19),
        "custos_fixos": ext(22),
        "margem_rs":    ext(20),
        "margem_pct":   [v * 100 for v in ext(21)],
        "ebitda_rs":    ext(32),
        "ebitda_pct":   [v * 100 for v in ext(33)],
        "base_receita": base_val(10),
        "base_ebitda":  base_val(32),
        "base_invest":  base_val(3),
        "base_roas":    base_val(5),
        "base_pedidos": base_val(8),
        "base_ticket":  base_val(12),
    }

d = carregar_dados()

# ── Helpers ───────────────────────────────────────────────────────────────────
def brl(v):
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def pct_delta(final, base):
    return ((final / base) - 1) * 100 if base else 0

LAYOUT = dict(
    paper_bgcolor="#1a1d27",
    plot_bgcolor="#1a1d27",
    font=dict(color="#8892a4", size=11),
    margin=dict(l=8, r=8, t=32, b=8),
    legend=dict(orientation="h", y=-0.18, font=dict(size=11)),
)

AXIS_X = dict(gridcolor="#2a2d3e", tickfont=dict(size=10))
AXIS_Y = dict(gridcolor="#2a2d3e", tickfont=dict(size=10))

def L(**kwargs):
    """Mescla LAYOUT base com parâmetros extras sem conflito de chaves."""
    return {**LAYOUT, "xaxis": AXIS_X, "yaxis": AXIS_Y, **kwargs}

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <div class="logo-icon">V</div>
  <div class="header-text">
    <h1>Vectoring — DRE Dashboard</h1>
    <p>Projeção ANO1 &amp; ANO2 · Performance de Tráfego Pago · 24 meses</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Indicadores-Chave — Projeção Final (ANO2/Dezembro)</div>', unsafe_allow_html=True)

receita_final  = d["receita"][-1]
ebitda_final   = d["ebitda_rs"][-1]
invest_final   = d["investimento"][-1]
roas_final     = d["roas"][-1]
pedidos_final  = d["pedidos"][-1]

cr_receita = pct_delta(receita_final, d["base_receita"])
cr_ebitda  = pct_delta(ebitda_final,  d["base_ebitda"])
cr_invest  = pct_delta(invest_final,  d["base_invest"])
cr_pedidos = pct_delta(pedidos_final, d["base_pedidos"])

kpis = [
    ("Receita Projetada",     brl(receita_final),   f"▲ {cr_receita:.0f}% vs Base",   "#6c63ff", "up"),
    ("EBITDA",                brl(ebitda_final),    f"▲ {cr_ebitda:.0f}% vs Base",    "#00d4aa", "up"),
    ("Margem EBITDA",         f"{d['ebitda_pct'][-1]:.1f}%", f"Base: {d['base_ebitda']/d['base_receita']*100:.1f}%", "#ffd166", "neutral"),
    ("Investimento Ads",      brl(invest_final),    f"▲ {cr_invest:.0f}% vs Base",    "#f97316", "up"),
    ("ROAS Final",            f"{roas_final:.1f}x", f"▼ Base: {d['base_roas']:.1f}x", "#ec4899", "down"),
    ("Pedidos/mês",           f"{pedidos_final:.0f}", f"▲ {cr_pedidos:.0f}% vs Base", "#22c55e", "up"),
    ("Ticket Médio",          brl(d["base_ticket"]), "Constante",                      "#a78bfa", "neutral"),
    ("Margem Contribuição",   "40,0%",              "Estável todo período",            "#38bdf8", "neutral"),
]

cols = st.columns(4)
for i, (label, value, sub, color, direction) in enumerate(kpis):
    badge_cls = {"up": "badge-up", "down": "badge-down", "neutral": "badge-neutral"}[direction]
    with cols[i % 4]:
        st.markdown(f"""
        <div class="kpi-card" style="--accent:{color}">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{value}</div>
          <div class="kpi-sub {badge_cls}">{sub}</div>
        </div>""", unsafe_allow_html=True)

# ── GRÁFICOS PRINCIPAIS ───────────────────────────────────────────────────────
st.markdown('<div class="section-title">Evolução da Receita & Resultados</div>', unsafe_allow_html=True)

# Gráfico 1 — Receita + Custos + EBITDA (wide)
fig_main = go.Figure()
fig_main.add_trace(go.Bar(
    name="Custo Variável", x=d["meses"], y=d["custo_var"],
    marker_color="rgba(239,68,68,.6)", offsetgroup="a"
))
fig_main.add_trace(go.Bar(
    name="Custo Fixo", x=d["meses"], y=d["custos_fixos"],
    marker_color="rgba(255,209,102,.6)", offsetgroup="a"
))
fig_main.add_trace(go.Bar(
    name="EBITDA", x=d["meses"], y=d["ebitda_rs"],
    marker_color="rgba(0,212,170,.6)", offsetgroup="a"
))
fig_main.add_trace(go.Scatter(
    name="Receita Bruta", x=d["meses"], y=d["receita"],
    mode="lines+markers", line=dict(color="#6c63ff", width=2.5),
    marker=dict(size=4), fill="tozeroy",
    fillcolor="rgba(108,99,255,.08)"
))
fig_main.update_layout(L(
    barmode="stack",
    title=dict(text="Receita Bruta vs Custos vs EBITDA (24 meses)", font=dict(color="#e2e8f0", size=13)),
    yaxis=dict(tickprefix="R$ ", gridcolor="#2a2d3e", tickfont=dict(size=10)),
    xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
    height=320,
))
st.plotly_chart(fig_main, use_container_width=True)

# Gráficos em 2 colunas
col1, col2 = st.columns(2)

with col1:
    fig_ebitda = go.Figure(go.Scatter(
        x=d["meses"], y=d["ebitda_rs"],
        fill="tozeroy", fillcolor="rgba(0,212,170,.15)",
        line=dict(color="#00d4aa", width=2.5),
        mode="lines+markers", marker=dict(size=4),
        name="EBITDA R$"
    ))
    fig_ebitda.update_layout(L(
        title=dict(text="EBITDA R$ — Evolução Mensal", font=dict(color="#e2e8f0", size=13)),
        yaxis=dict(tickprefix="R$ ", gridcolor="#2a2d3e", tickfont=dict(size=10)),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
        height=280, showlegend=False,
    ))
    st.plotly_chart(fig_ebitda, use_container_width=True)

with col2:
    fig_inv = make_subplots(specs=[[{"secondary_y": True}]])
    fig_inv.add_trace(go.Scatter(
        x=d["meses"], y=d["receita"],
        name="Receita", line=dict(color="#6c63ff", width=2.5),
        fill="tozeroy", fillcolor="rgba(108,99,255,.1)"
    ), secondary_y=False)
    fig_inv.add_trace(go.Scatter(
        x=d["meses"], y=d["investimento"],
        name="Investimento Ads", line=dict(color="#ffd166", width=2),
        fill="tozeroy", fillcolor="rgba(255,209,102,.08)"
    ), secondary_y=True)
    fig_inv.update_layout(L(
        title=dict(text="Investimento Ads vs Receita", font=dict(color="#e2e8f0", size=13)),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
        height=280,
    ))
    fig_inv.update_yaxes(tickprefix="R$ ", gridcolor="#2a2d3e", secondary_y=False)
    fig_inv.update_yaxes(tickprefix="R$ ", gridcolor="rgba(0,0,0,0)", secondary_y=True)
    st.plotly_chart(fig_inv, use_container_width=True)

# ── PERFORMANCE DE TRÁFEGO PAGO ───────────────────────────────────────────────
st.markdown('<div class="section-title">Performance de Tráfego Pago</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

with col1:
    fig_roas = go.Figure(go.Scatter(
        x=d["meses"], y=d["roas"],
        fill="tozeroy", fillcolor="rgba(236,72,153,.15)",
        line=dict(color="#ec4899", width=2.5),
        mode="lines+markers", marker=dict(size=4), name="ROAS"
    ))
    fig_roas.update_layout(L(
        title=dict(text="ROAS — Queda Planejada Trimestral", font=dict(color="#e2e8f0", size=13)),
        yaxis=dict(ticksuffix="x", gridcolor="#2a2d3e", tickfont=dict(size=10)),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
        height=260, showlegend=False,
    ))
    st.plotly_chart(fig_roas, use_container_width=True)

with col2:
    fig_ped = go.Figure(go.Bar(
        x=d["meses"], y=d["pedidos"],
        marker_color="rgba(108,99,255,.7)",
        marker_line_color="#6c63ff", marker_line_width=1,
        name="Pedidos"
    ))
    fig_ped.update_layout(L(
        title=dict(text="Pedidos Finalizados / Mês", font=dict(color="#e2e8f0", size=13)),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
        height=260, showlegend=False,
    ))
    st.plotly_chart(fig_ped, use_container_width=True)

with col3:
    fig_marg = go.Figure()
    fig_marg.add_trace(go.Scatter(
        x=d["meses"], y=d["ebitda_pct"],
        name="EBITDA %", line=dict(color="#00d4aa", width=2.5),
        fill="tozeroy", fillcolor="rgba(0,212,170,.1)"
    ))
    fig_marg.add_trace(go.Scatter(
        x=d["meses"], y=d["margem_pct"],
        name="Margem Contrib. %", line=dict(color="#ffd166", width=2, dash="dot"),
    ))
    fig_marg.update_layout(L(
        title=dict(text="Margens % — EBITDA & Contribuição", font=dict(color="#e2e8f0", size=13)),
        yaxis=dict(ticksuffix="%", gridcolor="#2a2d3e", tickfont=dict(size=10), range=[0, 55]),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
        height=260,
    ))
    st.plotly_chart(fig_marg, use_container_width=True)

# ── COMPARATIVO ANO1 vs ANO2 ─────────────────────────────────────────────────
st.markdown('<div class="section-title">Comparativo ANO1 vs ANO2</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    meses_curtos = ["Dez","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(
        x=meses_curtos, y=d["receita"][:12],
        name="ANO1", marker_color="rgba(108,99,255,.7)", marker_line_width=0
    ))
    fig_comp.add_trace(go.Bar(
        x=meses_curtos, y=d["receita"][12:],
        name="ANO2", marker_color="rgba(0,212,170,.7)", marker_line_width=0
    ))
    fig_comp.update_layout(L(
        barmode="group",
        title=dict(text="Receita — ANO1 vs ANO2 (mesmo mês)", font=dict(color="#e2e8f0", size=13)),
        yaxis=dict(tickprefix="R$ ", gridcolor="#2a2d3e", tickfont=dict(size=10)),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10)),
        height=280,
    ))
    st.plotly_chart(fig_comp, use_container_width=True)

with col2:
    fig_custos = go.Figure()
    fig_custos.add_trace(go.Bar(
        x=d["meses"], y=d["custo_var"],
        name="Custo Variável", marker_color="rgba(239,68,68,.65)",
    ))
    fig_custos.add_trace(go.Bar(
        x=d["meses"], y=d["custos_fixos"],
        name="Custo Fixo", marker_color="rgba(255,209,102,.65)",
    ))
    fig_custos.update_layout(L(
        barmode="stack",
        title=dict(text="Composição dos Custos — Fixos vs Variáveis", font=dict(color="#e2e8f0", size=13)),
        yaxis=dict(tickprefix="R$ ", gridcolor="#2a2d3e", tickfont=dict(size=10)),
        xaxis=dict(gridcolor="#2a2d3e", tickfont=dict(size=10), tickangle=45),
        height=280,
    ))
    st.plotly_chart(fig_custos, use_container_width=True)

# ── TABELA DRE DETALHADA ──────────────────────────────────────────────────────
st.markdown('<div class="section-title">DRE Detalhado — Dados Completos</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📅 ANO 1", "📅 ANO 2"])

meses_h_a1 = ["Dez","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
meses_h_a2 = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def build_df(start, end, meses_h):
    idx = slice(start, end)
    df = pd.DataFrame({
        "Indicador": [
            "Receita Bruta (R$)", "Investimento Ads (R$)", "ROAS",
            "Pedidos Finalizados", "Custo Produto (R$)", "Total Custo Variável (R$)",
            "Custos Fixos (R$)", "EBITDA (R$)", "EBITDA (%)"
        ],
    })
    for i, m in enumerate(meses_h):
        j = start + i
        df[m] = [
            brl(d["receita"][j]),
            brl(d["investimento"][j]),
            f"{d['roas'][j]:.1f}x",
            f"{d['pedidos'][j]:.0f}",
            brl(d["custo_prod"][j]),
            brl(d["custo_var"][j]),
            brl(d["custos_fixos"][j]),
            brl(d["ebitda_rs"][j]),
            f"{d['ebitda_pct'][j]:.1f}%",
        ]
    return df

with tab1:
    df1 = build_df(0, 12, meses_h_a1)
    st.dataframe(
        df1.style
           .set_properties(**{"background-color": "#1a1d27", "color": "#e2e8f0", "border-color": "#2a2d3e"})
           .apply(lambda x: ["background-color: rgba(0,212,170,.12); color:#e2e8f0" if x.name in [7,8]
                             else "" for i in x], axis=1),
        use_container_width=True, hide_index=True
    )

with tab2:
    df2 = build_df(12, 24, meses_h_a2)
    st.dataframe(
        df2.style
           .set_properties(**{"background-color": "#1a1d27", "color": "#e2e8f0", "border-color": "#2a2d3e"})
           .apply(lambda x: ["background-color: rgba(0,212,170,.12); color:#e2e8f0" if x.name in [7,8]
                             else "" for i in x], axis=1),
        use_container_width=True, hide_index=True
    )

# ── INSIGHTS ─────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Análise & Insights</div>', unsafe_allow_html=True)

insights = [
    ("#6c63ff", "📈 Crescimento de Receita",
     f"A receita projeta crescimento de <strong>{cr_receita:.0f}%</strong> da base ({brl(d['base_receita'])}) "
     f"até o final do ANO2 ({brl(receita_final)}), sustentado pelo aumento de {cr_invest:.0f}% no investimento em Ads."),
    ("#00d4aa", "💰 EBITDA Sólido",
     f"EBITDA cresce de <strong>{brl(d['base_ebitda'])}</strong> para <strong>{brl(ebitda_final)}</strong>, "
     f"mantendo margem estável entre <strong>35–36%</strong> durante todo o período projetado."),
    ("#ffd166", "⚠️ ROAS em Declínio",
     f"O ROAS cai de <strong>{d['base_roas']:.1f}x</strong> (base) para <strong>{roas_final:.1f}x</strong> (ANO2/Dez) — "
     f"queda planejada de 5% por trimestre. Acompanhar com otimizações contínuas de criativos e audiências."),
    ("#22c55e", "🛒 Volume de Pedidos",
     f"Pedidos saem de <strong>{d['base_pedidos']:.0f}/mês</strong> para <strong>{pedidos_final:.0f}/mês</strong>. "
     f"Com ticket médio fixo em <strong>{brl(d['base_ticket'])}</strong>, o crescimento é puramente volumétrico."),
    ("#f97316", "📦 Custos Variáveis",
     "Frete, Imposto e Comissão de Marketplace zerados na projeção. Qualquer ativação de Marketplace "
     "ou mudança tributária impacta diretamente os 40% de margem de contribuição."),
    ("#ec4899", "🏗️ Custos Fixos",
     f"Custos fixos reduzem de <strong>{brl(d['custos_fixos'][0])}</strong> para <strong>{brl(d['custos_fixos'][-1])}</strong>/mês "
     "após ajuste de plataforma no 3º trimestre do ANO1, melhorando o EBITDA a partir do ANO1/Jul."),
]

col1, col2 = st.columns(2)
for i, (color, title, body) in enumerate(insights):
    with (col1 if i % 2 == 0 else col2):
        st.markdown(f"""
        <div class="insight" style="--c:{color}">
          <div class="insight-title">{title}</div>
          <div class="insight-body">{body}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align:center;color:#2a2d3e;font-size:11px;'>Vectoring DRE Dashboard · Dados projetados · "
    "Gerado com Python + Streamlit + Plotly</p>",
    unsafe_allow_html=True
)
