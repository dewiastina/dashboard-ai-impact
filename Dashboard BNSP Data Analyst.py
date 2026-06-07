"""
Dashboard BI — Dampak AI Generatif terhadap Performa Akademik & Kesejahteraan Mahasiswa
BNSP Data Analyst · Gusti Ayu Dewi Astinasari
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io
import warnings
warnings.filterwarnings("ignore")

# ── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Impact — BNSP Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"], .stApp { font-family:'IBM Plex Sans',sans-serif; }
.stApp { background:#F4F6FA; }

section[data-testid="stSidebar"] { background:#0F1923; border-right:1px solid #1E2D3D; }
section[data-testid="stSidebar"] * { color:#CBD5E0 !important; }
section[data-testid="stSidebar"] label { color:#94A3B8 !important; font-size:11px !important;
  text-transform:uppercase; letter-spacing:.08em; font-weight:500; }

[data-testid="metric-container"] {
  background:#fff; border:1px solid #E2E8F0; border-radius:8px;
  padding:16px 18px; box-shadow:0 1px 4px rgba(0,0,0,.05);
}
[data-testid="metric-container"] [data-testid="stMetricLabel"] {
  font-size:10.5px !important; text-transform:uppercase; letter-spacing:.09em;
  color:#64748B !important; font-weight:600;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
  font-size:24px !important; font-weight:700; color:#0F172A !important;
  font-family:'IBM Plex Mono',monospace !important;
}

.stTabs [data-baseweb="tab-list"] {
  background:#fff; border-bottom:2px solid #E2E8F0; border-radius:8px 8px 0 0; padding:0 8px;
}
.stTabs [data-baseweb="tab"] { font-size:13px; font-weight:500; color:#64748B; padding:10px 18px; }
.stTabs [aria-selected="true"] { color:#1D4ED8 !important; border-bottom:2px solid #1D4ED8 !important; background:transparent !important; }

.sec-header {
  font-size:11px; font-weight:700; text-transform:uppercase; letter-spacing:.1em;
  color:#64748B; margin:20px 0 10px; padding-bottom:6px; border-bottom:1px solid #E2E8F0;
}
.bq { display:inline-block; background:#EFF6FF; color:#1D4ED8; border:1px solid #BFDBFE;
  border-radius:4px; padding:1px 7px; font-size:10.5px; font-weight:700;
  font-family:'IBM Plex Mono',monospace; margin-right:6px; }
.insight { background:#F0FDF4; border-left:3px solid #16A34A; border-radius:0 6px 6px 0;
  padding:11px 15px; margin:10px 0; font-size:12.5px; color:#14532D; line-height:1.65; }
.warn { background:#FFFBEB; border-left:3px solid #D97706; border-radius:0 6px 6px 0;
  padding:11px 15px; margin:10px 0; font-size:12.5px; color:#78350F; line-height:1.65; }
.note { background:#F8FAFC; border:1px solid #E2E8F0; border-radius:6px;
  padding:11px 16px; margin-top:18px; font-size:11.5px; color:#64748B; line-height:1.7; }
</style>
""", unsafe_allow_html=True)

# ── PALETTE ───────────────────────────────────────────────────────────────────
C = {
    "burnout":  {"High": "#DC2626", "Medium": "#F59E0B", "Low": "#16A34A"},
    "policy":   {"Allowed_With_Citation": "#1D4ED8", "Strict_Ban": "#DC2626", "Actively_Encouraged": "#16A34A"},
    "segment":  {"Light": "#06B6D4", "Moderate": "#F59E0B", "Heavy": "#EF4444"},
    "skill":    {"Beginner": "#94A3B8", "Intermediate": "#3B82F6", "Advanced": "#7C3AED"},
    "major":    {"STEM": "#1D4ED8", "Business": "#D97706", "Humanities": "#16A34A", "Medical": "#DC2626", "Arts": "#7C3AED"},
}
LAYOUT = dict(
    font_family="IBM Plex Sans", font_color="#374151",
    plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
    margin=dict(l=16, r=16, t=40, b=16),
    title_font_size=13, title_font_color="#0F172A",
    hoverlabel=dict(bgcolor="white", font_size=12,
                    font_family="IBM Plex Sans"),
)
POL_LABEL = {
    "Allowed_With_Citation": "Allowed + Citation",
    "Strict_Ban": "Strict Ban",
    "Actively_Encouraged": "Actively Encouraged",
}

# ── DATA LOADING ───────────────────────────────────────────────────────────────
DATA_PATH = "ai_student_impact_CLEAN (1).xlsx"


@st.cache_data(show_spinner="Memuat dataset…")
def load_data():
    df = pd.read_excel(DATA_PATH, engine="openpyxl")
    df = df.replace(r'^\s*$', np.nan, regex=True)

    # ── Cleaning sesuai Modul 4 ──
    valid_years = ['Freshman', 'Sophomore', 'Junior', 'Senior', 'Graduate']
    valid_burnout = ['Low', 'Medium', 'High']
    valid_policy = ['Allowed_With_Citation',
                    'Strict_Ban', 'Actively_Encouraged']

    df = df[df['Year_of_Study'].isin(valid_years)]
    df['Pre_Semester_GPA'] = pd.to_numeric(
        df['Pre_Semester_GPA'],  errors='coerce')
    df['Post_Semester_GPA'] = pd.to_numeric(
        df['Post_Semester_GPA'], errors='coerce')
    df = df[df['Pre_Semester_GPA'].between(0, 4.0, inclusive='both')]
    df = df[df['Post_Semester_GPA'].between(0, 4.0, inclusive='both')]
    df['Tool_Diversity'] = pd.to_numeric(
        df['Tool_Diversity'], errors='coerce').clip(upper=5)
    df = df.drop_duplicates(subset=['Student_ID'], keep='first')

    for c in ['Weekly_GenAI_Hours', 'Traditional_Study_Hours',
              'Perceived_AI_Dependency', 'Anxiety_Level_During_Exams', 'Skill_Retention_Score']:
        df[c] = pd.to_numeric(df[c], errors='coerce')

    df = df.dropna(subset=['Post_Semester_GPA'])

    # ── Derived (Modul 4.6 + 5.4) ──
    df['GPA_Delta'] = (df['Post_Semester_GPA'] -
                       df['Pre_Semester_GPA']).round(4)
    df['AI_Segment'] = pd.cut(
        df['Weekly_GenAI_Hours'], bins=[-0.01, 5, 15, 40], labels=['Light', 'Moderate', 'Heavy']
    ).astype(str)

    df['Risk_Profile'] = 'Low Dep + Low Burnout'
    df.loc[(df['Perceived_AI_Dependency'] >= 7) & (
        df['Burnout_Risk_Level'] != 'High'), 'Risk_Profile'] = 'High Dep + Low Burnout'
    df.loc[(df['Perceived_AI_Dependency'] < 7) & (
        df['Burnout_Risk_Level'] == 'High'), 'Risk_Profile'] = 'Low Dep + High Burnout'
    df.loc[(df['Perceived_AI_Dependency'] >= 7) & (
        df['Burnout_Risk_Level'] == 'High'), 'Risk_Profile'] = 'High Dep + High Burnout'

    return df.reset_index(drop=True)


try:
    df_full = load_data()
except Exception as e:
    st.error(
        f"**Gagal memuat file:** {e}\n\nPastikan file ada di:\n`{DATA_PATH}`")
    st.stop()

# ── HELPER ────────────────────────────────────────────────────────────────────


def fmt2(v): return f"{v:.2f}"
def fmtp(v): return f"{v:.2f}%"
def fmtd(v): return f"{v:+.2f}"


def apply_filters(df, maj, yr, pol, seg):
    m = pd.Series(True, index=df.index)
    if maj:
        m &= df['Major_Category'].isin(maj)
    if yr:
        m &= df['Year_of_Study'].isin(yr)
    if pol:
        m &= df['Institutional_Policy'].isin(pol)
    if seg:
        m &= df['AI_Segment'].isin(seg)
    return df[m].copy()


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dashboard BI")
    st.markdown(
        "**Dampak AI Generatif**  \nPerforma Akademik & Kesejahteraan Mahasiswa")
    st.markdown("---")
    st.markdown("### Global Filter")

    ALL_MAJORS = sorted(df_full['Major_Category'].dropna().unique())
    ALL_YEARS = [y for y in ['Freshman', 'Sophomore', 'Junior', 'Senior', 'Graduate']
                 if y in df_full['Year_of_Study'].unique()]
    ALL_POLICIES = sorted(df_full['Institutional_Policy'].dropna().unique())
    ALL_SEGS = [s for s in ['Light', 'Moderate', 'Heavy']
                if s in df_full['AI_Segment'].unique()]

    maj_sel = st.multiselect(
        "Bidang Studi",
        ALL_MAJORS,
        default=ALL_MAJORS
    )

    yr_sel = st.multiselect(
        "Jenjang Studi",
        ALL_YEARS,
        default=ALL_YEARS
    )

    pol_sel = st.multiselect(
        "Kebijakan Institusi",
        ALL_POLICIES,
        default=ALL_POLICIES
    )

    seg_sel = st.multiselect(
        "Segmen Penggunaan AI",
        ALL_SEGS,
        default=ALL_SEGS
    )

    st.markdown("---")
    st.caption(f"Dataset: **{len(df_full):,}** mahasiswa · 16 variabel")
    st.caption("⚠ Analisis bersifat *asosiatif*, bukan kausal.")

df = apply_filters(df_full, maj_sel, yr_sel, pol_sel, seg_sel)
N = len(df)

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:#0F1923;padding:22px 26px;border-radius:8px;margin-bottom:20px;">
  <div style="font-size:10.5px;letter-spacing:.12em;text-transform:uppercase;color:#64748B;font-weight:600;margin-bottom:5px;">
    BNSP Data Analyst · Divisi Riset &amp; Kebijakan
  </div>
  <div style="font-size:21px;font-weight:700;color:#F1F5F9;margin-bottom:3px;">
    Dampak AI Generatif terhadap Performa Akademik &amp; Kesejahteraan Mahasiswa
  </div>
  <div style="font-size:12.5px;color:#94A3B8;">AI Impact on Students — Academic &amp; Well-being Dataset</div>
</div>
""", unsafe_allow_html=True)

if N < len(df_full):
    st.info(
        f"Filter aktif: **{N:,}** dari {len(df_full):,} mahasiswa ({N/len(df_full)*100:.1f}%) ditampilkan.")

# ── KPI ROW ───────────────────────────────────────────────────────────────────
k1, k2, k3, k4, k5, k6 = st.columns(6)
avg_delta = df['GPA_Delta'].mean()
avg_post = df['Post_Semester_GPA'].mean()
avg_ret = df['Skill_Retention_Score'].mean()
pct_burn = (df['Burnout_Risk_Level'] == 'High').mean()*100
avg_ai_hrs = df['Weekly_GenAI_Hours'].mean()
avg_dep = df['Perceived_AI_Dependency'].mean()

k1.metric("Jumlah Mahasiswa",        f"{N:,}")
k2.metric("Rata-rata GPA Delta",     f"{avg_delta:+.2f}")
k3.metric("Rata-rata Post GPA",      f"{avg_post:.2f}")
k4.metric("Rata-rata Skill Retention", f"{avg_ret:.2f}")
k5.metric("Proporsi High Burnout",   f"{pct_burn:.2f}%")
k6.metric("Rata-rata Jam AI/Minggu", f"{avg_ai_hrs:.2f} jam")

st.markdown("<hr style='border:none;border-top:1px solid #E2E8F0;margin:16px 0'>",
            unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
SEG_ORDER = ['Light', 'Moderate', 'Heavy']
SKILL_ORD = ['Beginner', 'Intermediate', 'Advanced']
BURN_ORD = ['Low', 'Medium', 'High']

tab1, tab2, tab3, tab4 = st.tabs([
    "📘 BQ-01 & 02 — GPA & Kualitas Penggunaan AI",
    "📗 BQ-03 & 04 — Retensi Pengetahuan",
    "📕 BQ-05 & 06 — Burnout & Kecemasan",
    "📙 BQ-07 & 08 — Efektivitas Kebijakan",
])

# ──────────────────────────────────────────────────────────────────────────────
# TAB 1 — GPA
# ──────────────────────────────────────────────────────────────────────────────
with tab1:
    # ── BQ-01 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-01</span>Asosiasi Intensitas Penggunaan AI dengan GPA Delta</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1.4, 1])
    with c1:
        df_seg = df[df['AI_Segment'].isin(SEG_ORDER)]
        fig = px.box(df_seg, x='AI_Segment', y='GPA_Delta',
                     color='AI_Segment', color_discrete_map=C['segment'],
                     category_orders={'AI_Segment': SEG_ORDER},
                     title="Distribusi GPA Delta per Segmen Penggunaan AI",
                     labels={'AI_Segment': 'Segmen AI',
                             'GPA_Delta': 'GPA Delta (Post − Pre)'},
                     points='outliers')
        fig.add_hline(y=0, line_dash='dash',
                      line_color='#64748B', line_width=1.2)
        fig.update_layout(**LAYOUT, showlegend=False, height=330,
                          xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        seg_tbl = (df.groupby('AI_Segment')[['GPA_Delta', 'Post_Semester_GPA']]
                   .mean().round(2)
                   .reindex([s for s in SEG_ORDER if s in df['AI_Segment'].unique()]))
        n_seg = df.groupby('AI_Segment').size().rename('Jumlah Mhs')
        seg_tbl = seg_tbl.join(n_seg)
        seg_tbl.columns = ['Rata-rata GPA Δ',
                           'Rata-rata Post GPA', 'Jumlah Mhs']
        seg_tbl['Jumlah Mhs'] = seg_tbl['Jumlah Mhs'].map('{:,}'.format)
        st.markdown("**Ringkasan per Segmen AI**")
        st.dataframe(seg_tbl, use_container_width=True)

        corr_ai_gpa = df['Weekly_GenAI_Hours'].corr(df['GPA_Delta'])
        st.metric("Korelasi: Jam AI ↔ GPA Delta", f"r = {corr_ai_gpa:.2f}")
        st.markdown(f"""<div class="insight">
        <strong>Temuan BQ-01:</strong> Korelasi jam penggunaan AI dan GPA Delta sangat lemah 
        (r ≈ {corr_ai_gpa:.2f}). <em>Kuantitas</em> penggunaan AI bukan prediktor performa akademik yang andal 
        — ini mengarahkan analisis ke kualitas penggunaan (BQ-02).
        </div>""", unsafe_allow_html=True)

    # Binned mean GPA Delta vs AI Hours
    df_bin = df.copy()
    df_bin['AI_Hours_Bin'] = pd.cut(df_bin['Weekly_GenAI_Hours'],
                                    bins=[0, 5, 10, 15, 20, 25, 40],
                                    labels=['0–5', '5–10', '10–15', '15–20', '20–25', '25+'])
    ai_bin_gpa = (df_bin.groupby('AI_Hours_Bin')['GPA_Delta']
                  .agg(['mean', 'count']).round(2).reset_index())
    ai_bin_gpa.columns = ['Bin', 'Avg GPA Delta', 'Jumlah Mhs']
    fig_bin = px.bar(ai_bin_gpa, x='Bin', y='Avg GPA Delta',
                     text='Avg GPA Delta',
                     title="Rata-rata GPA Delta per Rentang Jam Penggunaan AI",
                     labels={'Bin': 'Jam AI/Minggu',
                             'Avg GPA Delta': 'Avg GPA Δ'},
                     color='Avg GPA Delta',
                     color_continuous_scale=['#DC2626', '#F59E0B', '#16A34A'])
    fig_bin.add_hline(y=0, line_dash='dash',
                      line_color='#64748B', line_width=1)
    fig_bin.update_traces(texttemplate='%{text:+.2f}', textposition='outside')
    fig_bin.update_layout(**LAYOUT, height=280, showlegend=False,
                          coloraxis_showscale=False,
                          xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
    st.plotly_chart(fig_bin, use_container_width=True)

    # ── BQ-02 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-02</span>Moderasi Prompt Skill & Tool Diversity terhadap GPA dan Retensi</div>', unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        skill_gpa = (df.groupby('Prompt_Engineering_Skill')['GPA_Delta']
                     .mean().round(2)
                     .reindex(SKILL_ORD).reset_index())
        skill_gpa.columns = ['Skill', 'Avg GPA Delta']
        fig_sk = px.bar(skill_gpa, x='Skill', y='Avg GPA Delta',
                        color='Skill', color_discrete_map=C['skill'],
                        text='Avg GPA Delta',
                        title="Rata-rata GPA Delta per Kemampuan Prompting",
                        labels={'Skill': 'Prompt Engineering Skill'})
        fig_sk.update_traces(
            texttemplate='%{text:+.2f}', textposition='outside')
        fig_sk.add_hline(y=0, line_dash='dash',
                         line_color='#64748B', line_width=1)
        fig_sk.update_layout(**LAYOUT, showlegend=False, height=300,
                             xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig_sk, use_container_width=True)

    with c4:
        tool_agg = (df.groupby('Tool_Diversity')[['GPA_Delta', 'Skill_Retention_Score']]
                    .mean().round(2).reset_index())
        fig_tool = go.Figure()
        fig_tool.add_trace(go.Scatter(x=tool_agg['Tool_Diversity'], y=tool_agg['GPA_Delta'],
                                      mode='lines+markers', name='GPA Delta',
                                      line=dict(color='#1D4ED8', width=2), marker=dict(size=8)))
        fig_tool.add_trace(go.Scatter(x=tool_agg['Tool_Diversity'],
                                      y=tool_agg['Skill_Retention_Score']/25,
                                      mode='lines+markers', name='Retensi (÷25)',
                                      line=dict(color='#16A34A', width=2, dash='dot'), marker=dict(size=8)))
        fig_tool.add_hline(y=0, line_dash='dash',
                           line_color='#64748B', line_width=1)
        fig_tool.update_layout(**LAYOUT, height=300,
                               title="GPA Delta & Retensi per Jumlah Tool AI",
                               xaxis=dict(title='Tool Diversity (1–5)',
                                          tickmode='linear', showgrid=False),
                               yaxis=dict(title='Nilai', gridcolor='#F1F5F9'),
                               legend=dict(orientation='h', y=-0.22))
        st.plotly_chart(fig_tool, use_container_width=True)

    # Heatmap Skill × Segment
    pivot = (df.groupby(['Prompt_Engineering_Skill', 'AI_Segment'])['GPA_Delta']
             .mean().round(2).unstack()
             .reindex(SKILL_ORD, axis=0)
             .reindex([s for s in SEG_ORDER if s in df['AI_Segment'].unique()], axis=1))
    fig_heat = px.imshow(pivot,
                         color_continuous_scale=[
                             '#DC2626', '#F9FAFB', '#16A34A'],
                         aspect='auto', text_auto=True,
                         title="Heatmap GPA Delta: Prompt Skill × Segmen AI",
                         labels=dict(x='Segmen AI', y='Prompt Skill', color='GPA Δ'))
    fig_heat.update_layout(**LAYOUT, height=220,
                           coloraxis_colorbar=dict(title='GPA Δ', len=0.85))
    st.plotly_chart(fig_heat, use_container_width=True)

    # ── Primary Use Case
    st.markdown('<div class="sec-header">Rata-rata GPA Delta per Tujuan Penggunaan AI (<em>Primary Use Case</em>)</div>', unsafe_allow_html=True)
    uc_gpa = (df.groupby('Primary_Use_Case')['GPA_Delta']
              .mean().round(2).sort_values().reset_index())
    uc_gpa.columns = ['Use Case', 'Avg GPA Delta']
    n_uc = df.groupby('Primary_Use_Case').size().rename(
        'Jumlah Mhs').reset_index()
    n_uc.columns = ['Use Case', 'Jumlah Mhs']
    uc_gpa = uc_gpa.merge(n_uc, on='Use Case')
    fig_uc = px.bar(uc_gpa, x='Avg GPA Delta', y='Use Case', orientation='h',
                    color='Avg GPA Delta', text='Avg GPA Delta',
                    color_continuous_scale=['#DC2626', '#F9FAFB', '#16A34A'],
                    title="Rata-rata GPA Delta per Primary Use Case",
                    labels={'Avg GPA Delta': 'Avg GPA Δ', 'Use Case': 'Tujuan Penggunaan AI'})
    fig_uc.update_traces(texttemplate='%{text:+.2f}', textposition='outside')
    fig_uc.add_vline(x=0, line_dash='dash', line_color='#64748B')
    fig_uc.update_layout(**LAYOUT, showlegend=False, height=260,
                         coloraxis_showscale=False,
                         xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
    st.plotly_chart(fig_uc, use_container_width=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 2 — RETENSI
# ──────────────────────────────────────────────────────────────────────────────
with tab2:
    # ── BQ-03 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-03</span>Ketergantungan AI dan Skill Retention</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1.3, 1])
    with c1:
        samp = df.sample(min(4000, N), random_state=42)
        fig_sc = px.scatter(samp, x='Perceived_AI_Dependency', y='Skill_Retention_Score',
                            color='AI_Segment', color_discrete_map=C['segment'],
                            opacity=0.4,
                            title="AI Dependency vs Skill Retention",
                            labels={'Perceived_AI_Dependency': 'AI Dependency Score (1–10)',
                                    'Skill_Retention_Score': 'Skill Retention Score (0–100)'})
        fig_sc.update_layout(**LAYOUT, height=340,
                             xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'),
                             legend=dict(orientation='h', y=-0.22))
        st.plotly_chart(fig_sc, use_container_width=True)

    with c2:
        corr_dep_ret = df['Perceived_AI_Dependency'].corr(
            df['Skill_Retention_Score'])
        corr_dep_gpa = df['Perceived_AI_Dependency'].corr(df['GPA_Delta'])
        corr_trad_ret = df['Traditional_Study_Hours'].corr(
            df['Skill_Retention_Score'])

        st.metric("Korelasi: Dependency ↔ Retention",
                  f"r = {corr_dep_ret:.2f}")
        st.metric("Korelasi: Dependency ↔ GPA Delta",
                  f"r = {corr_dep_gpa:.2f}")
        st.metric("Korelasi: Jam Tradisional ↔ Retention",
                  f"r = {corr_trad_ret:.2f}")

        df_dep = df.copy()
        df_dep['Dep_Bin'] = pd.cut(df_dep['Perceived_AI_Dependency'],
                                   bins=[0, 3, 6, 10], labels=['Rendah (1–3)', 'Sedang (4–6)', 'Tinggi (7–10)'])
        dep_ret = df_dep.groupby('Dep_Bin')[
            'Skill_Retention_Score'].mean().round(2).reset_index()
        dep_ret.columns = ['Dep Level', 'Avg Retention']
        fig_dep = px.bar(dep_ret, x='Dep Level', y='Avg Retention',
                         color='Dep Level',
                         color_discrete_sequence=[
                             '#16A34A', '#F59E0B', '#DC2626'],
                         text='Avg Retention',
                         title="Rata-rata Retensi per Tingkat Dependency",
                         labels={'Dep Level': 'Tingkat Dependency'})
        fig_dep.update_traces(
            texttemplate='%{text:.2f}', textposition='outside')
        fig_dep.update_layout(**LAYOUT, showlegend=False, height=270,
                              xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig_dep, use_container_width=True)

    st.markdown(f"""<div class="insight">
    <strong>Temuan BQ-03:</strong> Ketergantungan AI berkorelasi negatif dengan retensi pengetahuan 
    (r ≈ {corr_dep_ret:.2f}) dan dengan GPA Delta (r ≈ {corr_dep_gpa:.2f}). 
    Mahasiswa yang merasa sangat bergantung pada AI cenderung memiliki skor retensi lebih rendah, 
    mengindikasikan risiko degradasi kemampuan belajar mandiri.
    </div>""", unsafe_allow_html=True)

    # ── BQ-04 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-04</span>Segmentasi 2×2: Belajar Tradisional × AI — Implikasi terhadap GPA & Retensi</div>', unsafe_allow_html=True)

    med_trad = df['Traditional_Study_Hours'].median()
    med_ai = df['Weekly_GenAI_Hours'].median()
    df_q = df.copy()
    df_q['Kuadran'] = np.select(
        [(df_q['Traditional_Study_Hours'] >= med_trad) & (df_q['Weekly_GenAI_Hours'] < med_ai),
         (df_q['Traditional_Study_Hours'] >= med_trad) & (
             df_q['Weekly_GenAI_Hours'] >= med_ai),
         (df_q['Traditional_Study_Hours'] < med_trad) & (df_q['Weekly_GenAI_Hours'] >= med_ai)],
        ['Trad Tinggi / AI Rendah', 'Trad Tinggi / AI Tinggi',
            'Trad Rendah / AI Tinggi'],
        default='Trad Rendah / AI Rendah')

    quad_agg = df_q.groupby('Kuadran').agg(
        Jumlah_Mhs=('Student_ID', 'count'),
        Avg_GPA_Delta=('GPA_Delta', 'mean'),
        Avg_Retention=('Skill_Retention_Score', 'mean'),
        Pct_High_Burnout=('Burnout_Risk_Level',
                          lambda x: (x == 'High').mean()*100)
    ).round(2).reset_index()

    c3, c4 = st.columns(2)
    with c3:
        fig_q = px.scatter(quad_agg, x='Avg_GPA_Delta', y='Avg_Retention',
                           size='Jumlah_Mhs', color='Kuadran', text='Kuadran',
                           title="Kuadran Belajar: GPA Delta × Skill Retention",
                           labels={'Avg_GPA_Delta': 'Avg GPA Δ', 'Avg_Retention': 'Avg Skill Retention'})
        fig_q.update_traces(textposition='top center', textfont_size=10)
        fig_q.add_hline(y=df['Skill_Retention_Score'].mean(),
                        line_dash='dot', line_color='#94A3B8')
        fig_q.add_vline(x=0, line_dash='dot', line_color='#94A3B8')
        fig_q.update_layout(**LAYOUT, showlegend=False, height=340,
                            xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig_q, use_container_width=True)

    with c4:
        st.markdown("**Profil 4 Kuadran Belajar**")
        quad_disp = quad_agg.copy()
        quad_disp['Jumlah_Mhs'] = quad_disp['Jumlah_Mhs'].map('{:,}'.format)
        quad_disp['Avg_GPA_Delta'] = quad_disp['Avg_GPA_Delta'].map(
            '{:+.2f}'.format)
        quad_disp['Avg_Retention'] = quad_disp['Avg_Retention'].map(
            '{:.2f}'.format)
        quad_disp['Pct_High_Burnout'] = quad_disp['Pct_High_Burnout'].map(
            '{:.2f}%'.format)
        quad_disp.columns = ['Kuadran', 'Jumlah Mhs',
                             'GPA Δ', 'Retention', '% High Burnout']
        st.dataframe(quad_disp, use_container_width=True,
                     hide_index=True, height=190)

        # Jam tradisional vs retensi
        trad_bin = df.copy()
        trad_bin['Trad_Bin'] = pd.cut(trad_bin['Traditional_Study_Hours'],
                                      bins=[0, 7, 14, 21, 36],
                                      labels=['1–7 jam', '8–14 jam', '15–21 jam', '22+ jam'])
        trad_ret = trad_bin.groupby(
            'Trad_Bin')['Skill_Retention_Score'].mean().round(2).reset_index()
        trad_ret.columns = ['Bin', 'Avg Retention']
        fig_tr = px.line(trad_ret, x='Bin', y='Avg Retention', markers=True,
                         title="Retensi per Rentang Jam Belajar Tradisional",
                         labels={'Bin': 'Jam Belajar/Minggu', 'Avg Retention': 'Avg Skill Retention'})
        fig_tr.update_traces(line_color='#1D4ED8', marker_size=9)
        fig_tr.update_layout(**LAYOUT, height=230,
                             xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig_tr, use_container_width=True)

    st.markdown(f"""<div class="insight">
    <strong>Temuan BQ-04:</strong> Mahasiswa dengan jam belajar tradisional tinggi 
    (>{med_trad:.0f} jam/minggu) konsisten menunjukkan GPA Delta dan Skill Retention lebih baik, 
    terlepas dari intensitas penggunaan AI. Ini mendukung hipotesis bahwa AI paling efektif 
    sebagai <em>pelengkap</em>, bukan pengganti belajar mandiri.
    </div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 3 — BURNOUT & KECEMASAN
# ──────────────────────────────────────────────────────────────────────────────
with tab3:
    # ── BQ-05 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-05</span>Profil Karakteristik Mahasiswa dengan Burnout High</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        burn_dist = df['Burnout_Risk_Level'].value_counts().reset_index()
        burn_dist.columns = ['Level', 'Jumlah']
        burn_dist['Pct'] = (burn_dist['Jumlah'] / N * 100).round(2)
        burn_dist['Color'] = burn_dist['Level'].map(C['burnout'])
        fig_pie = go.Figure(go.Pie(
            labels=burn_dist['Level'], values=burn_dist['Jumlah'],
            marker_colors=burn_dist['Color'], hole=0.55,
            textinfo='percent+label', textfont_size=12))
        fig_pie.update_layout(**LAYOUT, title="Distribusi Burnout Risk Level",
                              showlegend=False, height=290)
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        # Burnout per AI Segment stacked
        ba = (df.groupby(['AI_Segment', 'Burnout_Risk_Level']).size()
              .reset_index(name='N'))
        tot = df.groupby('AI_Segment').size().rename('Total').reset_index()
        ba = ba.merge(tot, on='AI_Segment')
        ba['Pct'] = (ba['N']/ba['Total']*100).round(2)
        fig_ba = px.bar(ba[ba['AI_Segment'].isin(SEG_ORDER)],
                        x='AI_Segment', y='Pct', color='Burnout_Risk_Level',
                        color_discrete_map=C['burnout'],
                        category_orders={'AI_Segment': SEG_ORDER,
                                         'Burnout_Risk_Level': BURN_ORD},
                        barmode='stack', title="Proporsi Burnout per Segmen AI (%)",
                        labels={'AI_Segment': 'Segmen AI', 'Pct': '%', 'Burnout_Risk_Level': 'Burnout'})
        fig_ba.update_layout(**LAYOUT, height=290, xaxis=dict(showgrid=False),
                             yaxis=dict(gridcolor='#F1F5F9'), legend=dict(orientation='h', y=-0.3))
        st.plotly_chart(fig_ba, use_container_width=True)

    with c3:
        # Avg Dependency per Burnout
        dep_burn = (df.groupby('Burnout_Risk_Level')
                    .agg(Avg_Dependency=('Perceived_AI_Dependency', 'mean'),
                         Avg_Anxiety=('Anxiety_Level_During_Exams', 'mean'),
                         Avg_AI_Hours=('Weekly_GenAI_Hours', 'mean'))
                    .round(2).reindex(BURN_ORD))
        fig_dep_b = go.Figure()
        for col, name, color in [
            ('Avg_Dependency', 'Avg Dependency', '#1D4ED8'),
            ('Avg_Anxiety', 'Avg Anxiety', '#DC2626'),
        ]:
            fig_dep_b.add_trace(go.Bar(x=dep_burn.index, y=dep_burn[col],
                                       name=name, marker_color=color,
                                       text=dep_burn[col].map('{:.2f}'.format),
                                       textposition='outside'))
        fig_dep_b.update_layout(**LAYOUT, height=290, barmode='group',
                                title="Rata-rata Dependency & Anxiety per Burnout Level",
                                xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'),
                                legend=dict(orientation='h', y=-0.3))
        st.plotly_chart(fig_dep_b, use_container_width=True)

    # Risk Profile kuadran
    rp_pal = {'Low Dep + Low Burnout': '#16A34A', 'High Dep + Low Burnout': '#F59E0B',
              'Low Dep + High Burnout': '#F97316', 'High Dep + High Burnout': '#DC2626'}
    rp_c = df['Risk_Profile'].value_counts().reset_index()
    rp_c.columns = ['Profile', 'Jumlah']
    rp_c['Pct'] = (rp_c['Jumlah']/N*100).round(2)
    fig_rp = px.bar(rp_c.sort_values('Jumlah'), x='Jumlah', y='Profile', orientation='h',
                    color='Profile', color_discrete_map=rp_pal, text='Pct',
                    title="Risk Profile: Kuadran AI Dependency × Burnout",
                    labels={'Jumlah': 'Jumlah Mahasiswa', 'Profile': 'Profil Risiko'})
    fig_rp.update_traces(texttemplate='%{text:.2f}%', textposition='outside')
    fig_rp.update_layout(**LAYOUT, showlegend=False, height=240,
                         xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
    st.plotly_chart(fig_rp, use_container_width=True)

    high_dep_high_burn = (df['Risk_Profile'] ==
                          'High Dep + High Burnout').mean()*100
    st.markdown(f"""<div class="warn">
    <strong>Perhatian BQ-05:</strong> Mahasiswa <em>Heavy User</em> menunjukkan proporsi High Burnout tertinggi. 
    Segmen <strong>"High Dep + High Burnout"</strong> (ketergantungan AI ≥7 <em>dan</em> burnout tinggi) 
    mencakup <strong>{high_dep_high_burn:.2f}%</strong> dari total mahasiswa — 
    ini adalah target prioritas intervensi oleh Bagian Kemahasiswaan.
    </div>""", unsafe_allow_html=True)

    # ── BQ-06 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-06</span>Kebijakan Institusi dan Dampaknya terhadap Kecemasan & Burnout</div>', unsafe_allow_html=True)

    c4, c5 = st.columns(2)
    with c4:
        pol_anx = (df.groupby('Institutional_Policy')
                   .agg(Avg_Anxiety=('Anxiety_Level_During_Exams', 'mean'))
                   .round(2).reset_index())
        pol_anx['Label'] = pol_anx['Institutional_Policy'].map(POL_LABEL)
        pol_anx['Color'] = pol_anx['Institutional_Policy'].map(C['policy'])
        fig_anx = px.bar(pol_anx, x='Label', y='Avg_Anxiety',
                         color='Label',
                         color_discrete_map={v: C['policy'][k]
                                             for k, v in POL_LABEL.items()},
                         text='Avg_Anxiety',
                         title="Rata-rata Anxiety Level per Kebijakan",
                         labels={'Label': 'Kebijakan', 'Avg_Anxiety': 'Avg Anxiety (1–10)'})
        fig_anx.update_traces(
            texttemplate='%{text:.2f}', textposition='outside')
        fig_anx.add_hline(y=df['Anxiety_Level_During_Exams'].mean(),
                          line_dash='dash', line_color='#64748B',
                          annotation_text='Rata-rata keseluruhan')
        fig_anx.update_layout(**LAYOUT, showlegend=False, height=300,
                              xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig_anx, use_container_width=True)

    with c5:
        bp = (df.groupby('Institutional_Policy')['Burnout_Risk_Level']
              .value_counts(normalize=True).mul(100).round(2).reset_index())
        bp.columns = ['Policy', 'Burnout', 'Pct']
        bp['Label'] = bp['Policy'].map(POL_LABEL)
        fig_bp = px.bar(bp, x='Label', y='Pct', color='Burnout',
                        color_discrete_map=C['burnout'],
                        barmode='stack',
                        category_orders={'Burnout': BURN_ORD},
                        title="Distribusi Burnout per Kebijakan (%)",
                        labels={'Label': 'Kebijakan', 'Pct': '%', 'Burnout': 'Burnout Level'})
        fig_bp.update_layout(**LAYOUT, height=300, xaxis=dict(showgrid=False),
                             yaxis=dict(gridcolor='#F1F5F9'),
                             legend=dict(orientation='h', y=-0.25))
        st.plotly_chart(fig_bp, use_container_width=True)

    strict_anx = df[df['Institutional_Policy'] ==
                    'Strict_Ban']['Anxiety_Level_During_Exams'].mean()
    enc_anx = df[df['Institutional_Policy'] ==
                 'Actively_Encouraged']['Anxiety_Level_During_Exams'].mean()
    strict_burn = (df[df['Institutional_Policy'] == 'Strict_Ban']
                   ['Burnout_Risk_Level'] == 'High').mean()*100
    st.markdown(f"""<div class="warn">
    <strong>Temuan BQ-06:</strong> Kebijakan <em>Strict Ban</em> tidak mengurangi risiko burnout — 
    justru menunjukkan rata-rata kecemasan {strict_anx:.2f}/10 
    (vs Actively Encouraged: {enc_anx:.2f}/10) dengan proporsi High Burnout {strict_burn:.2f}%. 
    Pelarangan penuh AI berpotensi meningkatkan tekanan psikologis mahasiswa.
    </div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# TAB 4 — EFEKTIVITAS KEBIJAKAN
# ──────────────────────────────────────────────────────────────────────────────
with tab4:
    # ── BQ-07 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-07</span>Perbandingan Multi-KPI: Kebijakan Institusi Mana yang Terbaik?</div>', unsafe_allow_html=True)

    pol_agg = df.groupby('Institutional_Policy').agg(
        Jumlah_Mhs=('Student_ID', 'count'),
        Avg_Post_GPA=('Post_Semester_GPA', 'mean'),
        Avg_GPA_Delta=('GPA_Delta', 'mean'),
        Avg_Retention=('Skill_Retention_Score', 'mean'),
        Avg_Anxiety=('Anxiety_Level_During_Exams', 'mean'),
        Pct_High_Burnout=('Burnout_Risk_Level',
                          lambda x: (x == 'High').mean()*100)
    ).round(2).reset_index()
    pol_agg['Label'] = pol_agg['Institutional_Policy'].map(POL_LABEL)

    c1, c2 = st.columns([1.4, 1])
    with c1:
        # Radar multi-KPI
        cats = ['Post GPA', 'GPA Delta', 'Skill Retention',
                'Low Burnout %', 'Low Anxiety']

        def norm(s): return (s-s.min())/(s.max()-s.min()+1e-9)
        pol_agg['n_post_gpa'] = norm(pol_agg['Avg_Post_GPA'])
        pol_agg['n_gpa_delta'] = norm(pol_agg['Avg_GPA_Delta'])
        pol_agg['n_retention'] = norm(pol_agg['Avg_Retention'])
        pol_agg['n_low_burnout'] = norm(100 - pol_agg['Pct_High_Burnout'])
        pol_agg['n_low_anxiety'] = norm(10 - pol_agg['Avg_Anxiety'])

        fig_radar = go.Figure()
        norm_cols = ['n_post_gpa', 'n_gpa_delta',
                     'n_retention', 'n_low_burnout', 'n_low_anxiety']
        for _, row in pol_agg.iterrows():
            vals = [row[c] for c in norm_cols] + [row[norm_cols[0]]]
            cat = cats + [cats[0]]
            pol = row['Institutional_Policy']
            fig_radar.add_trace(go.Scatterpolar(
                r=vals, theta=cat, fill='toself', opacity=0.35,
                name=POL_LABEL.get(pol, pol),
                line_color=C['policy'].get(pol, '#94A3B8'),
                fillcolor=C['policy'].get(pol, '#94A3B8')))
        fig_radar.update_layout(
            **LAYOUT, height=380,
            polar=dict(radialaxis=dict(visible=True, range=[0, 1],
                                       gridcolor='#E2E8F0', tickfont_size=9)),
            title="Radar Multi-KPI per Kebijakan (normalisasi 0–1, lebih tinggi = lebih baik)",
            legend=dict(orientation='h', y=-0.15))
        st.plotly_chart(fig_radar, use_container_width=True)

    with c2:
        st.markdown("**Tabel Komparatif Multi-KPI**")
        disp = pol_agg[['Label', 'Jumlah_Mhs', 'Avg_Post_GPA', 'Avg_GPA_Delta',
                        'Avg_Retention', 'Pct_High_Burnout', 'Avg_Anxiety']].copy()
        disp['Jumlah_Mhs'] = disp['Jumlah_Mhs'].map('{:,}'.format)
        disp['Avg_GPA_Delta'] = disp['Avg_GPA_Delta'].map('{:+.2f}'.format)
        disp['Avg_Post_GPA'] = disp['Avg_Post_GPA'].map('{:.2f}'.format)
        disp['Avg_Retention'] = disp['Avg_Retention'].map('{:.2f}'.format)
        disp['Pct_High_Burnout'] = disp['Pct_High_Burnout'].map(
            '{:.2f}%'.format)
        disp['Avg_Anxiety'] = disp['Avg_Anxiety'].map('{:.2f}'.format)
        disp.columns = ['Kebijakan', 'Jumlah Mhs', 'Post GPA', 'GPA Δ',
                        'Skill Retention', '% High Burnout', 'Avg Anxiety']
        st.dataframe(disp, use_container_width=True,
                     hide_index=True, height=160)

        # Composite score
        pol_agg['Composite'] = (pol_agg['n_post_gpa']*0.35 +
                                pol_agg['n_gpa_delta']*0.20 +
                                pol_agg['n_retention']*0.25 +
                                pol_agg['n_low_burnout']*0.10 +
                                pol_agg['n_low_anxiety']*0.10)
        best_row = pol_agg.loc[pol_agg['Composite'].idxmax()]
        best_label = best_row['Label']
        st.markdown(f"""<div class="insight">
        <strong>Temuan BQ-07:</strong> Berdasarkan composite score (GPA 35% · Retensi 25% · 
        GPA Δ 20% · Low Burnout 10% · Low Anxiety 10%), kebijakan terbaik adalah 
        <strong>{best_label}</strong> — menghasilkan kombinasi performa akademik 
        tinggi dan risiko kesehatan mental lebih rendah.
        </div>""", unsafe_allow_html=True)

        # Composite bar
        pol_agg['Composite_pct'] = (pol_agg['Composite']*100).round(2)
        fig_cs = px.bar(pol_agg.sort_values('Composite_pct'),
                        x='Composite_pct', y='Label', orientation='h',
                        color='Institutional_Policy',
                        color_discrete_map=C['policy'], text='Composite_pct',
                        title="Composite Score per Kebijakan",
                        labels={'Composite_pct': 'Composite Score (%)', 'Label': 'Kebijakan'})
        fig_cs.update_traces(
            texttemplate='%{text:.2f}', textposition='outside')
        fig_cs.update_layout(**LAYOUT, showlegend=False, height=200,
                             xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
        st.plotly_chart(fig_cs, use_container_width=True)

    # Tren GPA per Jenjang × Kebijakan
    year_pol_gpa = (df.groupby(['Year_of_Study', 'Institutional_Policy'])['Post_Semester_GPA']
                    .mean().round(2).reset_index())
    year_pol_gpa['Label'] = year_pol_gpa['Institutional_Policy'].map(POL_LABEL)
    fig_yp = px.line(year_pol_gpa, x='Year_of_Study', y='Post_Semester_GPA',
                     color='Label', markers=True,
                     color_discrete_map={v: C['policy'][k]
                                         for k, v in POL_LABEL.items()},
                     title="Tren Post GPA per Jenjang Studi × Kebijakan",
                     labels={'Year_of_Study': 'Jenjang',
                             'Post_Semester_GPA': 'Avg Post GPA'},
                     category_orders={'Year_of_Study': ['Freshman', 'Sophomore', 'Junior', 'Senior', 'Graduate']})
    fig_yp.update_layout(**LAYOUT, height=280, xaxis=dict(showgrid=False),
                         yaxis=dict(gridcolor='#F1F5F9'),
                         legend=dict(orientation='h', y=-0.25, title='Kebijakan'))
    st.plotly_chart(fig_yp, use_container_width=True)

    # ── BQ-08 ──
    st.markdown('<div class="sec-header"><span class="bq">BQ-08</span>Heterogenitas Dampak AI per Bidang Studi — Implikasi Kebijakan Diferensial</div>', unsafe_allow_html=True)

    major_agg = df.groupby('Major_Category').agg(
        Jumlah_Mhs=('Student_ID', 'count'),
        Avg_GPA_Delta=('GPA_Delta', 'mean'),
        Avg_Post_GPA=('Post_Semester_GPA', 'mean'),
        Avg_Retention=('Skill_Retention_Score', 'mean'),
        Avg_AI_Hours=('Weekly_GenAI_Hours', 'mean'),
        Pct_High_Burnout=('Burnout_Risk_Level',
                          lambda x: (x == 'High').mean()*100)
    ).round(2).reset_index()

    c3, c4 = st.columns(2)
    with c3:
        fig_maj = px.bar(major_agg.sort_values('Avg_GPA_Delta'),
                         x='Avg_GPA_Delta', y='Major_Category', orientation='h',
                         color='Major_Category', color_discrete_map=C['major'],
                         text='Avg_GPA_Delta',
                         title="Rata-rata GPA Delta per Bidang Studi",
                         labels={'Avg_GPA_Delta': 'Avg GPA Δ', 'Major_Category': 'Bidang Studi'})
        fig_maj.update_traces(
            texttemplate='%{text:+.2f}', textposition='outside')
        fig_maj.add_vline(x=0, line_dash='dash', line_color='#64748B')
        fig_maj.update_layout(**LAYOUT, showlegend=False, height=300,
                              xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
        st.plotly_chart(fig_maj, use_container_width=True)

    with c4:
        fig_msc = px.scatter(major_agg,
                             x='Avg_GPA_Delta', y='Avg_Retention',
                             size='Jumlah_Mhs', color='Major_Category',
                             color_discrete_map=C['major'], text='Major_Category',
                             title="GPA Delta × Skill Retention per Major (ukuran = Jumlah Mhs)",
                             labels={'Avg_GPA_Delta': 'Avg GPA Δ', 'Avg_Retention': 'Avg Skill Retention'})
        fig_msc.update_traces(textposition='top center', textfont_size=10)
        fig_msc.add_hline(y=df['Skill_Retention_Score'].mean(),
                          line_dash='dot', line_color='#94A3B8')
        fig_msc.add_vline(x=0, line_dash='dot', line_color='#94A3B8')
        fig_msc.update_layout(**LAYOUT, showlegend=False, height=300,
                              xaxis=dict(showgrid=False), yaxis=dict(gridcolor='#F1F5F9'))
        st.plotly_chart(fig_msc, use_container_width=True)

    # Grouped: Major × Kebijakan → Post GPA
    mp_gpa = (df.groupby(['Major_Category', 'Institutional_Policy'])['Post_Semester_GPA']
              .mean().round(2).reset_index())
    mp_gpa['Kebijakan'] = mp_gpa['Institutional_Policy'].map(POL_LABEL)
    fig_mp = px.bar(mp_gpa, x='Major_Category', y='Post_Semester_GPA',
                    color='Kebijakan',
                    color_discrete_map={v: C['policy'][k]
                                        for k, v in POL_LABEL.items()},
                    barmode='group', text='Post_Semester_GPA',
                    title="Rata-rata Post GPA per Bidang Studi × Kebijakan Institusi",
                    labels={'Major_Category': 'Bidang Studi', 'Post_Semester_GPA': 'Avg Post GPA'})
    fig_mp.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig_mp.update_layout(**LAYOUT, height=310, xaxis=dict(showgrid=False),
                         yaxis=dict(gridcolor='#F1F5F9'),
                         legend=dict(orientation='h', y=-0.2, title='Kebijakan'))
    st.plotly_chart(fig_mp, use_container_width=True)

    # Summary major table
    st.markdown("**Ringkasan KPI per Bidang Studi**")
    ma_disp = major_agg.copy()
    ma_disp['Jumlah_Mhs'] = ma_disp['Jumlah_Mhs'].map('{:,}'.format)
    ma_disp['Avg_GPA_Delta'] = ma_disp['Avg_GPA_Delta'].map('{:+.2f}'.format)
    ma_disp['Avg_Post_GPA'] = ma_disp['Avg_Post_GPA'].map('{:.2f}'.format)
    ma_disp['Avg_Retention'] = ma_disp['Avg_Retention'].map('{:.2f}'.format)
    ma_disp['Avg_AI_Hours'] = ma_disp['Avg_AI_Hours'].map('{:.2f} jam'.format)
    ma_disp['Pct_High_Burnout'] = ma_disp['Pct_High_Burnout'].map(
        '{:.2f}%'.format)
    ma_disp.columns = ['Bidang Studi', 'Jumlah Mhs', 'GPA Δ', 'Post GPA',
                       'Skill Retention', 'Avg AI Hours', '% High Burnout']
    st.dataframe(ma_disp, use_container_width=True, hide_index=True)

    best_major = major_agg.loc[
        (major_agg['Avg_GPA_Delta']*0.5 + major_agg['Avg_Retention']*0.01).idxmax(), 'Major_Category']
    worst_major = major_agg.loc[
        (major_agg['Avg_GPA_Delta']*0.5 + major_agg['Avg_Retention']*0.01).idxmin(), 'Major_Category']
    st.markdown(f"""<div class="insight">
    <strong>Temuan BQ-08:</strong> Dampak AI signifikan berbeda antar bidang studi — 
    <strong>{best_major}</strong> menunjukkan kombinasi GPA Delta dan Retention terbaik, 
    sementara <strong>{worst_major}</strong> paling rentan. 
    Ini menunjukkan bahwa kebijakan AI <em>universal</em> kurang optimal; 
    pendekatan <em>diferensial per major</em> akan lebih efektif.
    </div>""", unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="note">
<strong>Catatan metodologis:</strong> Seluruh analisis bersifat <em>asosiatif</em>, bukan kausal. 
Data berasal dari survei satu semester (<em>cross-sectional</em>) dengan self-report measurement. 
Temuan tidak dapat secara langsung membuktikan bahwa penggunaan AI <em>menyebabkan</em> perubahan GPA atau burnout. 
Interpretasi kebijakan harus mempertimbangkan batasan ini secara eksplisit. 
AI Segment: Light ≤5 jam/mgg · Moderate 5–15 jam/mgg · Heavy &gt;15 jam/mgg. 
Risk Profile: tinggi jika Dependency ≥7 atau Burnout = High.
</div>
<div style="text-align:center;font-size:11px;color:#94A3B8;padding:12px 0 4px;">
BNSP Data Analyst · TIKGLOBAL/LSP-0039/SKM/2024 · Gusti Ayu Dewi Astinasari
</div>
""", unsafe_allow_html=True)
