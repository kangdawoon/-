# -*- coding: utf-8 -*-
"""
오쏘몰 마케팅 인텔리전스 대시보드 (리디자인 v2)

[실행 방법]
pip install streamlit pandas plotly openpyxl
streamlit run orthomol_dashboard.py

[필요 파일 - 같은 폴더]
naver_search_trend.csv, naver_seasonality_trend.csv, naver_content_summary.csv,
naver_content_recent_posts.csv, naver_demographics.csv, naver_purchase_intent.csv,
naver_shopping_category_trend.csv, google_trend_kr.csv, youtube_content.csv,
오쏘몰_키워드.xlsx, 연관어분석_20260828.xlsx
파일이 없는 항목은 화면에서 건너뛰거나 업로드 안내가 표시됩니다.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

st.set_page_config(page_title="오쏘몰 마케팅 인텔리전스", layout="wide", initial_sidebar_state="collapsed")

# ============================================================
# 디자인 토큰 — 모던 미니멀 (SaaS 대시보드 톤)
# ============================================================
INK = "#0F1115"         # 거의 블랙에 가까운 텍스트
AMBER = "#F0A83C"        # 시그니처 포인트 컬러 (단 하나의 강조색)
BG = "#FFFFFF"
SURFACE = "#F6F7F9"      # 카드/보조 배경
CARD = "#FFFFFF"
LINE = "#EAEBEF"
TEXT_MAIN = "#0F1115"
TEXT_SUB = "#8A8F98"
RUST = "#E5484D"         # 하락/경고 (선명한 레드)
SAGE = "#30A46C"         # 상승/긍정 (선명한 그린)
STEEL = "#6E7787"        # 보조 계열색 1 (뉴트럴 그레이블루)
PLUM = "#8E4EC6"         # 보조 계열색 2 (바이올렛, 포인트용)
SAND = "#D8C9A3"

BRAND_COLORS = {"오쏘몰": AMBER, "아임비타": PLUM, "센트룸": STEEL, "고려은단": SAGE}
GOLD = AMBER  # 하위 호환

PLOTLY_FONT = dict(family="Inter, sans-serif", color=TEXT_MAIN, size=12)

# ============================================================
# CSS 주입
# ============================================================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap');

html, body, [class*="css"] {{
    font-family: 'Inter', sans-serif;
}}
.stApp {{
    background-color: {BG};
}}
#MainMenu, footer, header {{visibility: hidden;}}
.block-container {{
    padding-top: 2.5rem;
    padding-bottom: 3rem;
    max-width: 1180px;
}}

.eyebrow {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 11.5px;
    font-weight: 600;
    letter-spacing: 0.14em;
    color: {AMBER};
    text-transform: uppercase;
    margin-bottom: 10px;
}}
.hero-title {{
    font-family: 'Inter', sans-serif;
    font-size: 34px;
    font-weight: 800;
    letter-spacing: -0.02em;
    color: {INK};
    margin: 0 0 8px 0;
    line-height: 1.2;
}}
.hero-sub {{
    font-family: 'Inter', sans-serif;
    font-size: 14px;
    font-weight: 400;
    color: {TEXT_SUB};
    margin-bottom: 32px;
}}

.kpi-strip {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 14px;
    margin-bottom: 36px;
}}
.kpi-card {{
    background: {CARD};
    border: 1px solid {LINE};
    border-radius: 16px;
    padding: 20px 22px;
    box-shadow: 0 1px 2px rgba(15,17,21,0.04), 0 8px 20px -12px rgba(15,17,21,0.08);
    transition: box-shadow .15s ease;
}}
.kpi-label {{
    font-family: 'Inter', sans-serif;
    font-size: 12px;
    font-weight: 500;
    color: {TEXT_SUB};
    margin-bottom: 10px;
    line-height: 1.4;
    min-height: 32px;
    display: flex;
    align-items: flex-end;
}}
.kpi-value {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 27px;
    font-weight: 600;
    color: {INK};
    letter-spacing: -0.02em;
}}
.kpi-delta {{
    font-family: 'Inter', sans-serif;
    font-size: 12px;
    font-weight: 600;
    margin-top: 6px;
    display: inline-block;
    padding: 4px 10px;
    border-radius: 8px;
    line-height: 1.5;
}}
.delta-down {{ color: {RUST}; background: rgba(229,72,77,0.08); }}
.delta-up {{ color: {SAGE}; background: rgba(48,164,108,0.08); }}
.delta-flat {{ color: {TEXT_SUB}; background: {SURFACE}; }}

.section-title {{
    font-family: 'Inter', sans-serif;
    font-size: 19px;
    font-weight: 700;
    letter-spacing: -0.01em;
    color: {INK};
    margin: 8px 0 2px 0;
}}
.section-desc {{
    font-family: 'Inter', sans-serif;
    font-size: 12.5px;
    color: {TEXT_SUB};
    margin-bottom: 16px;
}}

.chart-card {{
    background: {CARD};
    border: 1px solid {LINE};
    border-radius: 16px;
    padding: 20px 22px 8px 22px;
    margin-bottom: 16px;
    box-shadow: 0 1px 2px rgba(15,17,21,0.03);
}}
.chart-label {{
    font-family: 'Inter', sans-serif;
    font-size: 13px;
    font-weight: 600;
    color: {INK};
    margin-bottom: 2px;
}}
.chart-note {{
    font-family: 'Inter', sans-serif;
    font-size: 11.5px;
    color: {TEXT_SUB};
    margin-bottom: 12px;
}}

.insight-box {{
    background: {SURFACE};
    border: 1px solid {LINE};
    border-radius: 14px;
    padding: 16px 20px;
    margin: 10px 0 26px 0;
    font-family: 'Inter', sans-serif;
    font-size: 13px;
    color: {TEXT_MAIN};
    line-height: 1.75;
}}
.insight-box b {{ color: {INK}; font-weight: 700; }}

.stTabs [data-baseweb="tab-list"] {{
    gap: 6px;
    background: {SURFACE};
    border-radius: 12px;
    padding: 4px;
    border-bottom: none;
}}
.stTabs [data-baseweb="tab"] {{
    font-family: 'Inter', sans-serif;
    font-weight: 600;
    font-size: 13.5px;
    color: {TEXT_SUB};
    padding: 9px 16px;
    border-radius: 9px;
}}
.stTabs [aria-selected="true"] {{
    color: {INK} !important;
    background: {CARD} !important;
    box-shadow: 0 1px 2px rgba(15,17,21,0.08);
}}
[data-testid="stExpander"] {{
    border: 1px solid {LINE};
    border-radius: 12px;
    background: {CARD};
}}
</style>
""", unsafe_allow_html=True)


def base_layout(fig, height=340, legend=True):
    fig.update_layout(
        height=height,
        font=PLOTLY_FONT,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(t=10, b=10, l=10, r=10),
        legend=dict(orientation="h", yanchor="bottom", y=-0.3, font=dict(size=11)) if legend else dict(),
        showlegend=legend,
    )
    fig.update_xaxes(gridcolor=LINE, zeroline=False)
    fig.update_yaxes(gridcolor=LINE, zeroline=False)
    return fig


def chart_card_open(label, note=None):
    st.markdown(f'<div class="chart-card"><div class="chart-label">{label}</div>'
                + (f'<div class="chart-note">{note}</div>' if note else ""), unsafe_allow_html=True)


def chart_card_close():
    st.markdown("</div>", unsafe_allow_html=True)


def insight(html_text):
    st.markdown(f'<div class="insight-box">{html_text}</div>', unsafe_allow_html=True)


def kpi_html(label, value, delta_text, tone):
    tone_class = {"down": "delta-down", "up": "delta-up", "flat": "delta-flat"}[tone]
    arrow = {"down": "▼", "up": "▲", "flat": "—"}[tone]
    return f"""<div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-delta {tone_class}">{arrow} {delta_text}</div>
    </div>"""


@st.cache_data
def load_csv(path_or_buffer):
    return pd.read_csv(path_or_buffer, encoding="utf-8-sig")


@st.cache_data
def load_excel(path_or_buffer, sheet_name=0):
    return pd.read_excel(path_or_buffer, sheet_name=sheet_name)


def safe_csv(path):
    return load_csv(path) if os.path.exists(path) else None


def safe_excel(path):
    return load_excel(path) if os.path.exists(path) else None


def missing_note(*names):
    st.info(f"표시할 데이터가 없습니다 — {', '.join(names)} 파일을 같은 폴더에 넣어주세요.")


# ============================================================
# 데이터 로드
# ============================================================
brand_df = safe_csv("naver_search_trend.csv")
season_df = safe_csv("naver_seasonality_trend.csv")
content_summary_df = safe_csv("naver_content_summary.csv")
content_posts_df = safe_csv("naver_content_recent_posts.csv")
demo_df = safe_csv("naver_demographics.csv")
intent_df = safe_csv("naver_purchase_intent.csv")
shopping_df = safe_csv("naver_shopping_category_trend.csv")
google_df = safe_csv("google_trend_kr.csv")
youtube_df = safe_csv("youtube_content.csv")
news_keyword_df = safe_excel("오쏘몰_키워드.xlsx")
related_words_df = safe_excel("연관어분석_20260828.xlsx")
gift_tone_df = safe_csv("naver_gift_tone_trend.csv")
gift_recipient_df = safe_csv("naver_gift_recipient_demographics.csv")
gift_recommend_df = safe_csv("naver_gift_recommend_variants.csv")

if brand_df is not None:
    brand_df["기간"] = pd.to_datetime(brand_df["기간"])
if season_df is not None:
    season_df["기간"] = pd.to_datetime(season_df["기간"])
    season_pivot = season_df.pivot(index="기간", columns="키워드그룹", values="검색관심도(상대값)")
else:
    season_pivot = None
if intent_df is not None:
    intent_df["기간"] = pd.to_datetime(intent_df["기간"])
if google_df is not None:
    google_df["기간"] = pd.to_datetime(google_df["기간"])
if gift_tone_df is not None:
    gift_tone_df["기간"] = pd.to_datetime(gift_tone_df["기간"])

AGE_ORDER = ["0~12세", "13~18세", "19~24세", "25~29세", "30~34세",
             "35~39세", "40~44세", "45~49세", "50~54세", "55~60세", "60세 이상"]

# ============================================================
# 헤더
# ============================================================
st.markdown('<div class="eyebrow">ORTHOMOL · AI 검색 & 마케팅 인텔리전스</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">오쏘몰, 지금 시장에서 어떻게 보이고 있나</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-sub">네이버·구글·유튜브·뉴스 7개 공개 데이터 소스를 통합한 실데이터 기반 브랜드 진단</div>', unsafe_allow_html=True)

# ============================================================
# 핵심 지표 스트립
# ============================================================
kpi_cards_html = ""

if brand_df is not None:
    o = brand_df[brand_df["브랜드"] == "오쏘몰"].sort_values("기간")
    fv, lv = o["검색관심도(상대값)"].iloc[0], o["검색관심도(상대값)"].iloc[-1]
    pct = (lv - fv) / fv * 100
    kpi_cards_html += kpi_html("네이버 검색 관심도 (12개월)", f"{lv:.0f}", f"{pct:.1f}% ({fv:.0f}→{lv:.0f})", "down" if pct < 0 else "up")

if google_df is not None:
    go_ = google_df[google_df["브랜드"] == "오쏘몰"].sort_values("기간")["검색관심도(상대값)"]
    gf, gl = go_.iloc[:4].mean(), go_.iloc[-4:].mean()
    gpct = (gl - gf) / gf * 100
    kpi_cards_html += kpi_html("구글 검색 관심도 (교차검증)", f"{gl:.0f}", f"{gpct:.1f}% 동일 하락 추세", "down" if gpct < 0 else "up")

if season_pivot is not None and "명절 선물" in season_pivot.columns and "오쏘몰" in season_pivot.columns:
    match = season_pivot["명절 선물"].idxmax() == season_pivot["오쏘몰"].idxmax()
    kpi_cards_html += kpi_html("명절 시즌 상관관계", "일치" if match else "불일치", "명절선물 피크 주 일치" if match else "시기 어긋남", "up" if match else "flat")

if youtube_df is not None:
    med = youtube_df.groupby("브랜드")["조회수"].median()
    if "오쏘몰" in med.index:
        rank = int((med.rank(ascending=False))["오쏘몰"])
        kpi_cards_html += kpi_html("유튜브 조회수 순위 (4개 브랜드 중)", f"{rank}위", "중앙값 기준 최하위권", "down")

if kpi_cards_html:
    st.markdown(f'<div class="kpi-strip">{kpi_cards_html}</div>', unsafe_allow_html=True)

# ============================================================
# 기간 필터 — 검색 추이 차트 전용 (탭1 상단 차트 · 탭4 교차검증에만 적용, KPI 스트립은 항상 전체 기간)
# ============================================================
date_range = None
if brand_df is not None:
    min_d, max_d = brand_df["기간"].min().date(), brand_df["기간"].max().date()
    if min_d < max_d:
        fcol1, fcol2 = st.columns([1, 4])
        with fcol1:
            st.markdown(f'<div style="padding-top:10px; font-size:13px; font-weight:600; color:{TEXT_MAIN};">기간 필터</div>', unsafe_allow_html=True)
        with fcol2:
            date_range = st.slider("검색 추이 차트 기간", min_value=min_d, max_value=max_d, value=(min_d, max_d),
                                    format="YYYY-MM", label_visibility="collapsed")
        st.markdown('<div class="section-desc" style="margin-top:-4px;">브랜드별 검색 관심도 추이(탭1)·네이버 vs 구글 교차검증(탭4) 차트에만 적용됩니다.</div>', unsafe_allow_html=True)


def apply_date_range(df, col="기간"):
    if date_range is None or df is None:
        return df
    return df[(df[col].dt.date >= date_range[0]) & (df[col].dt.date <= date_range[1])]


# ============================================================
# 탭 구성
# ============================================================
tab0, tab1, tab2, tab3, tab4 = st.tabs(["🏢  자사분석", "🔍  검색 관심도", "📢  콘텐츠 & 도달", "🛒  쇼핑 행동", "✅  교차 검증"])

# ------------------------------------------------------------
# TAB 0. 자사분석
# ------------------------------------------------------------
with tab0:
    st.markdown('<div class="section-title">오쏘몰 자사분석</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">동아제약 공식 발표 및 언론 보도 기반 (2026년 8월 기준)</div>', unsafe_allow_html=True)

    chart_card_open("브랜드 개요")
    st.markdown(f"""
    <div style="font-size:13.5px; color:{TEXT_MAIN}; line-height:1.9;">
    &bull; <b>&ldquo;Orthomol&rdquo;</b> = ortho(올바르게) + molecular(분자단위)<br>
    &bull; 1991년 독일에서 시작, 독일 현지 100% 생산 — 동아제약이 2020년부터 국내 공식 단독 수입·유통<br>
    &bull; 담당 브랜드 매니저: 정인애 (동아제약 생활건강사업부)
    </div>
    """, unsafe_allow_html=True)
    chart_card_close()

    chart_card_open("5개년 매출 성장 추이")
    revenue_years = ["2020", "2021", "2022", "2023", "2024"]
    revenue_values = [87, 284, 655, 1204, 1302]
    f0 = go.Figure(go.Bar(x=revenue_years, y=revenue_values, marker_color=GOLD))
    f0.update_yaxes(title_text="매출액(억원)")
    st.plotly_chart(base_layout(f0, height=320, legend=False), use_container_width=True)
    st.markdown(f'<div style="font-size:11.5px; color:{TEXT_SUB}; margin-top:-4px;">2024년 기준 동아제약 내 매출 2위 제품(1위 박카스D 1,411억원) · 2023~2025 3년 연속 멀티비타민 판매 1위(유로모니터 인터내셔널 기준)</div>', unsafe_allow_html=True)
    chart_card_close()

    chart_card_open("제품·유통 전략 타임라인")
    timeline_items = [
        ("2020", "국내 정식 출시, 액상+정제 이중제형으로 시장 진입"),
        ("2025.11", "편의점(GS25·CU) 입점"),
        ("2026.05", "카카오톡 선물하기 전용 기획팩 5종 출시(가정의 달)"),
        ("2026(최근)", "6년 만의 신제형 &ldquo;이뮨 ODP&rdquo;(구강용해파우더) 출시"),
    ]
    timeline_html = "".join(
        f'<div style="display:flex; gap:14px; padding:8px 0; border-bottom:1px solid {LINE};">'
        f'<div style="min-width:84px; font-family:\'JetBrains Mono\', monospace; font-size:12px; font-weight:600; color:{AMBER};">{t}</div>'
        f'<div style="font-size:13px; color:{TEXT_MAIN};">{d}</div></div>'
        for t, d in timeline_items
    )
    st.markdown(f"<div>{timeline_html}</div>", unsafe_allow_html=True)
    chart_card_close()

    st.markdown('<div class="section-title" style="font-size:16px; margin-top:20px;">강점 / 약점</div>', unsafe_allow_html=True)
    swot_col1, swot_col2 = st.columns(2)
    with swot_col1:
        st.markdown(f"""
        <div style="background: rgba(48,164,108,0.08); border:1px solid rgba(48,164,108,0.25); border-radius:14px; padding:18px 20px; font-size:13px; color:{TEXT_MAIN}; line-height:1.85;">
        <div style="font-weight:700; color:{SAGE}; margin-bottom:8px;">강점 (Strength)</div>
        &bull; 압도적 매출 1위, 5년간 약 15배 성장<br>
        &bull; 독일 오리지널 + 이중제형이라는 명확한 차별화<br>
        &bull; 유통 확장 적극적(편의점, 카카오톡 선물하기)<br>
        &bull; 동아제약 내 2위 매출로 회사 차원의 투자 여력 확보
        </div>
        """, unsafe_allow_html=True)
    with swot_col2:
        st.markdown(f"""
        <div style="background: rgba(229,72,77,0.08); border:1px solid rgba(229,72,77,0.25); border-radius:14px; padding:18px 20px; font-size:13px; color:{TEXT_MAIN}; line-height:1.85;">
        <div style="font-weight:700; color:{RUST}; margin-bottom:8px;">약점 (Weakness)</div>
        &bull; 후발주자(대웅제약 에너씨슬, 일동제약 마이니부스터, 동국제약 마이핏V)의 디자인·가격 모방으로 차별화 압박<br>
        &bull; 검색 관심도는 하락 추세 (다른 탭 데이터와 연결되는 지점)<br>
        &bull; 신제형 출시가 6년 만 — 혁신 속도 상대적으로 느림<br>
        &bull; 프리미엄 가격대로 인한 가격 저항 존재
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    insight("오쏘몰은 5년간 매출이 약 15배 성장하며 명실상부한 국내 멀티비타민 1위 브랜드로 "
            "자리잡았지만, 최근 후발주자들의 유사 전략(디자인·가격 모방)과 검색 관심도 하락이 "
            "동시에 나타나고 있어 성장 둔화 신호로 해석할 수 있다. 편의점 입점, 카카오톡 선물하기 "
            "기획팩 등 유통·선물 채널 확장이 최근 성장 전략의 핵심 축으로 보인다.")

# ------------------------------------------------------------
# TAB 1. 검색 관심도
# ------------------------------------------------------------
with tab1:
    st.markdown('<div class="section-title">브랜드별 검색 관심도 추이</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">네이버 데이터랩 · 12개월 · 최고치=100 상대 지수</div>', unsafe_allow_html=True)

    if brand_df is not None:
        brand_df_f = apply_date_range(brand_df)
        chart_card_open("월간 검색 관심도")
        fig = go.Figure()
        for b in brand_df_f["브랜드"].unique():
            d = brand_df_f[brand_df_f["브랜드"] == b].sort_values("기간")
            fig.add_trace(go.Scatter(x=d["기간"], y=d["검색관심도(상대값)"], name=b,
                                      line=dict(width=4 if b == "오쏘몰" else 1.8, color=BRAND_COLORS.get(b)),
                                      mode="lines+markers", marker=dict(size=4)))
        st.plotly_chart(base_layout(fig), use_container_width=True)
        chart_card_close()

        o_f = brand_df_f[brand_df_f["브랜드"] == "오쏘몰"].sort_values("기간")["검색관심도(상대값)"]
        if len(o_f) >= 2:
            first_val, last_val = o_f.iloc[0], o_f.iloc[-1]
            insight(f"오쏘몰은 선택 기간 동안 <b>{(last_val-first_val)/first_val*100:.1f}%</b> {'하락' if last_val < first_val else '상승'}"
                    f"({first_val:.0f}→{last_val:.0f})한 반면, "
                    f"고려은단은 동기간 상승 흐름을 유지했습니다. 판매량 1위(별도 브랜드 리포트 기준)와 검색 관심도 하락이 동시에 나타나는 점은 "
                    f"신규 유입보다 기존 고객 재구매 위주 매출 구조일 가능성을 시사합니다.")
        else:
            st.info("선택한 기간에 데이터가 부족합니다 — 기간을 넓혀주세요.")
    else:
        missing_note("naver_search_trend.csv")

    st.markdown('<div class="section-title">왜 이 페르소나인가 — 선물 상황의 다양성</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">AI 검색 가시성 모니터링용으로 설계한 187개 질문 중, 상황·이벤트 카테고리(63개)의 내부 구성</div>', unsafe_allow_html=True)

    SITUATION_CATEGORIES = {
        "명절 시즌": 19,
        "기념일·축하": 18,
        "감사 인사 상황": 16,
        "답례 상황": 7,
        "첫 만남·미팅": 3,
    }
    col_s1, col_s2 = st.columns([1, 1])
    with col_s1:
        chart_card_open("선물이 필요한 상황 5가지", "질문 세트 설계 당시 정의한 상황 분류 (합계 63개)")
        f_sit = go.Figure(go.Pie(
            labels=list(SITUATION_CATEGORIES.keys()),
            values=list(SITUATION_CATEGORIES.values()),
            hole=0.55,
            marker=dict(colors=[GOLD, STEEL, PLUM, SAGE, RUST]),
            textinfo="label+percent",
            textfont=dict(size=11),
        ))
        st.plotly_chart(base_layout(f_sit, height=320), use_container_width=True)
        chart_card_close()
    with col_s2:
        st.markdown(f"""
        <div class="chart-card" style="padding-top:20px;">
            <div class="chart-label">해석</div>
            <div class="chart-note" style="margin-bottom:14px;">니즈의 크기가 아니라 니즈가 발생하는 상황의 개수·빈도가 핵심</div>
            <div style="font-size:13px; line-height:1.9; color:{TEXT_MAIN};">
            선물 니즈 자체는 특정 연령대에 국한되지 않고 전 연령대에 고르게 나타납니다
            (하단 &ldquo;연령·성별&rdquo; 데이터 참고). 대신 <b>5~8년차 직장인</b>은 상사·후배·거래처·
            고객·부모님 등 격식을 갖춰야 하는 관계가 한 사람 안에 동시다발적으로 존재하고,
            <b>명절·기념일·답례·감사인사·첫만남</b> 등 서로 다른 5가지 상황에서 반복적으로
            &ldquo;실패 없는 선물&rdquo;을 골라야 하는 빈도 자체가 특징적입니다.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="section-title">시즌·상황별 검증</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">명절 시즌 vs 직장인 선물 시즌 — 나란히 비교</div>', unsafe_allow_html=True)

    if season_pivot is not None:
        col1, col2 = st.columns(2)
        with col1:
            chart_card_open("명절 선물 검색 vs 오쏘몰", "추석·설날 전후 스파이크 비교")
            f2 = make_subplots(specs=[[{"secondary_y": True}]])
            if "명절 선물" in season_pivot.columns:
                f2.add_trace(go.Scatter(x=season_pivot.index, y=season_pivot["명절 선물"], name="명절 선물",
                                         line=dict(color=RUST, width=2), fill="tozeroy", fillcolor="rgba(181,80,46,0.08)"), secondary_y=False)
            if "오쏘몰" in season_pivot.columns:
                f2.add_trace(go.Scatter(x=season_pivot.index, y=season_pivot["오쏘몰"], name="오쏘몰",
                                         line=dict(color=GOLD, width=2.5)), secondary_y=True)
            st.plotly_chart(base_layout(f2, height=300), use_container_width=True)
            chart_card_close()

        with col2:
            chart_card_open("선물추천·직장인선물 vs 오쏘몰", "연말 시즌 갭 확인")
            f3 = make_subplots(specs=[[{"secondary_y": True}]])
            for kw, c in [("선물 추천", STEEL), ("직장인 선물", PLUM)]:
                if kw in season_pivot.columns:
                    f3.add_trace(go.Scatter(x=season_pivot.index, y=season_pivot[kw], name=kw, line=dict(color=c, width=1.8)), secondary_y=False)
            if "오쏘몰" in season_pivot.columns:
                f3.add_trace(go.Scatter(x=season_pivot.index, y=season_pivot["오쏘몰"], name="오쏘몰", line=dict(color=GOLD, width=2.5)), secondary_y=True)
            st.plotly_chart(base_layout(f3, height=300), use_container_width=True)
            chart_card_close()

        holiday_peak = season_pivot["명절 선물"].idxmax() if "명절 선물" in season_pivot.columns else None
        orthomol_peak = season_pivot["오쏘몰"].idxmax() if "오쏘몰" in season_pivot.columns else None
        match = holiday_peak == orthomol_peak

        # 명절 구간(추석 전후, 설날 전후) vs 비명절 구간 비중 계산
        holiday_windows = [
            (pd.Timestamp("2025-09-08"), pd.Timestamp("2025-10-06")),
            (pd.Timestamp("2026-01-26"), pd.Timestamp("2026-02-23")),
        ]
        is_holiday = pd.Series(False, index=season_pivot.index)
        for start, end in holiday_windows:
            is_holiday |= (season_pivot.index >= start) & (season_pivot.index <= end)
        holiday_share = season_pivot.loc[is_holiday, "오쏘몰"].sum() / season_pivot["오쏘몰"].sum() * 100 if "오쏘몰" in season_pivot.columns else None
        holiday_week_ratio = is_holiday.sum() / len(is_holiday) * 100

        insight(f"명절선물·오쏘몰 검색 <b>{'피크 시점 일치' if match else '피크 시점 불일치'}</b> — 명절 마케팅이 실제 수요와 맞물려 작동 중임을 시사합니다. "
                + (f"52주 중 명절 구간은 <b>{holiday_week_ratio:.0f}%(10주)</b>에 불과한데, 오쏘몰 연간 검색량의 <b>{holiday_share:.1f}%</b>가 이 구간에 몰려있습니다 — "
                   f"명절이라는 짧은 시기에 대한 의존도가 높다는 뜻입니다. " if holiday_share is not None else "")
                + f"반면 &ldquo;선물 추천&rdquo;·&ldquo;직장인 선물&rdquo; 검색은 대체로 연말(12월)에 몰리는데, 같은 시기 오쏘몰 검색은 연중 저점을 보여 "
                f"<b>&ldquo;명절&rdquo; 프레임에는 강하지만 &ldquo;연말·직장인 선물&rdquo; 등 평상시 프레임의 수요는 놓치고 있을 가능성</b>이 있습니다.")
    else:
        missing_note("naver_seasonality_trend.csv")

    st.markdown('<div class="section-title">선물 톤·격식 키워드 비교</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">캐주얼한 톤 vs 격식·목적성 있는 톤 — 어떤 표현이 실제로 더 검색되는지</div>', unsafe_allow_html=True)

    if gift_tone_df is not None:
        PRIMARY_AXIS = ["오쏘몰", "선물 추천", "건강 선물", "효도 선물", "비타민 선물", "명절 선물세트"]
        SECONDARY_AXIS = ["부담없는 선물", "고급 선물", "센스있는 선물", "가성비 선물", "감동 선물"]
        TONE_DOWN = "#D5D7DC"
        PRIMARY_EMPHASIS = {"비타민 선물", "건강 선물"}
        SECONDARY_EMPHASIS = {"센스있는 선물"}

        col_t1, col_t2 = st.columns(2)
        with col_t1:
            chart_card_open("주축 — 격식·목적성 키워드", "오쏘몰·선물 추천·건강/효도/비타민 선물·명절 선물세트")
            f_primary = go.Figure()
            for kw in PRIMARY_AXIS:
                gd = gift_tone_df[gift_tone_df["키워드그룹"] == kw].sort_values("기간")
                if gd.empty:
                    continue
                emphasized = kw in PRIMARY_EMPHASIS
                f_primary.add_trace(go.Scatter(x=gd["기간"], y=gd["검색관심도(상대값)"], name=kw,
                                                line=dict(color=SAGE if emphasized else TONE_DOWN, width=3.5 if emphasized else 1.2)))
            st.plotly_chart(base_layout(f_primary, height=340), use_container_width=True)
            chart_card_close()

        with col_t2:
            chart_card_open("보조축 — 캐주얼 톤 키워드", "부담없는·고급·센스있는·가성비·감동 선물")
            f_secondary = go.Figure()
            for kw in SECONDARY_AXIS:
                gd = gift_tone_df[gift_tone_df["키워드그룹"] == kw].sort_values("기간")
                if gd.empty:
                    continue
                emphasized = kw in SECONDARY_EMPHASIS
                f_secondary.add_trace(go.Scatter(x=gd["기간"], y=gd["검색관심도(상대값)"], name=kw,
                                                  line=dict(color=SAGE if emphasized else TONE_DOWN, width=3.5 if emphasized else 1.2)))
            st.plotly_chart(base_layout(f_secondary, height=340), use_container_width=True)
            chart_card_close()

        st.markdown(f'<div style="font-size:11.5px; color:{TEXT_SUB}; margin-top:-8px; margin-bottom:16px;">'
                    f'※ 두 차트는 서로 다른 시점에 수집되어 그룹 간 절대 크기 비교는 어려우며, 각 키워드의 시간에 따른 상대적 흐름'
                    f'(언제 오르고 내리는지)만 비교하는 용도입니다 — y축 스케일 차이(왼쪽 0~100대 · 오른쪽 0~4대)에 유의해주세요.</div>', unsafe_allow_html=True)

        insight("&ldquo;부담없는 선물&rdquo;·&ldquo;가성비 선물&rdquo;·&ldquo;감동 선물&rdquo;처럼 <b>캐주얼한 톤의 키워드는 검색량 자체가 매우 낮습니다</b> "
                "— 전체 검색 생태계에서 비중이 작은 표현이라는 뜻입니다. 반면 &ldquo;고급 선물&rdquo;·&ldquo;센스있는 선물&rdquo;처럼 "
                "<b>격식 있는 톤</b>과, &ldquo;효도 선물&rdquo;·&ldquo;건강 선물&rdquo;·&ldquo;비타민 선물&rdquo;처럼 <b>목적이 뚜렷한 키워드</b>가 "
                "상대적으로 더 많이 검색됩니다. 오쏘몰이 위치할 자리는 &ldquo;가성비&rdquo;보다 <b>&ldquo;격식·목적성&rdquo; 계열 키워드와 더 맞닿아 있다</b>는 "
                "해석이 가능합니다. &ldquo;비타민 선물&rdquo;·&ldquo;건강 선물&rdquo;·&ldquo;센스있는 선물&rdquo;(강조 표시)은 오쏘몰·명절선물세트처럼 "
                "특정 시즌에 급등락하는 다른 키워드들과 달리, <b>1년 내내 상대적으로 평탄한 흐름을 유지</b>합니다.")

        HOLIDAY_WINDOWS = [
            (pd.Timestamp("2025-09-08"), pd.Timestamp("2025-10-06")),
            (pd.Timestamp("2026-01-26"), pd.Timestamp("2026-02-23")),
        ]
        RATIO_KEYWORDS = ["오쏘몰", "명절 선물세트", "건강 선물", "비타민 선물", "센스있는 선물", "선물 추천"]

        def holiday_ratio(keyword):
            d = gift_tone_df[gift_tone_df["키워드그룹"] == keyword]
            if d.empty:
                return None
            is_holiday = pd.Series(False, index=d.index)
            for start, end in HOLIDAY_WINDOWS:
                is_holiday |= (d["기간"] >= start) & (d["기간"] <= end)
            hol_avg = d.loc[is_holiday, "검색관심도(상대값)"].mean()
            normal_avg = d.loc[~is_holiday, "검색관심도(상대값)"].mean()
            return hol_avg / normal_avg if normal_avg else None

        ratios = {kw: r for kw in RATIO_KEYWORDS if (r := holiday_ratio(kw)) is not None}

        if ratios:
            st.markdown('<div class="section-title" style="font-size:16px; margin-top:24px;">명절 시즌 쏠림 배율</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-desc">명절 구간(추석·설날 전후) 평균 ÷ 평시 구간 평균 — 키워드별 계절 쏠림 정도</div>', unsafe_allow_html=True)

            sorted_kw = sorted(ratios, key=ratios.get, reverse=True)
            top_kw = sorted_kw[0]
            rest_kw = [kw for kw in sorted_kw if kw != top_kw]

            st.markdown(f'<div style="background:{SURFACE}; border:1px solid {LINE}; border-radius:10px; '
                        f'padding:7px 14px; font-size:11.5px; color:{TEXT_SUB}; margin-bottom:10px; display:inline-block;">'
                        f'참고: &ldquo;{top_kw}&rdquo;만 <b style="color:{TEXT_MAIN};">{ratios[top_kw]:.1f}배</b> 극단적 쏠림 (비교 기준, 아래 차트에서는 제외)</div>',
                        unsafe_allow_html=True)

            chart_card_open("키워드별 명절/평시 검색 배율", f"&ldquo;{top_kw}&rdquo; 제외 — 나머지 {len(rest_kw)}개 키워드의 차이에 집중")
            f_ratio = go.Figure(go.Bar(
                x=rest_kw, y=[ratios[kw] for kw in rest_kw],
                marker_color=GOLD,
                text=[f"{ratios[kw]:.1f}배" for kw in rest_kw],
                textposition="outside",
            ))
            f_ratio.update_yaxes(title_text="명절/평시 배율(×)", range=[0, 2])
            st.plotly_chart(base_layout(f_ratio, height=320, legend=False), use_container_width=True)
            chart_card_close()

            rest_vals = [ratios[kw] for kw in rest_kw]
            insight(f"&ldquo;{top_kw}&rdquo;처럼 <b>명절이라는 단어가 박힌 키워드만 극단적으로 쏠리며({ratios[top_kw]:.1f}배)</b>, "
                    f"건강 선물·비타민 선물·센스있는 선물·오쏘몰 등 나머지 키워드는 <b>{min(rest_vals):.1f}~{max(rest_vals):.1f}배 사이의 완만한 계절성</b>을 보입니다 "
                    f"— 즉 건강·비타민 계열 선물은 명절이라는 특정 시즌에 갇혀있지 않습니다.")
    else:
        missing_note("naver_gift_tone_trend.csv")

    st.markdown('<div class="section-title">20~30대의 다양한 선물 대상</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">상사·거래처·부모님·동료·첫만남 — 대상별 선물 키워드를 누가 검색하나</div>', unsafe_allow_html=True)

    if gift_recipient_df is not None:
        RECIPIENT_AGES = ["19~24세", "25~29세", "30~34세", "35~39세"]
        RECIPIENT_KEYWORDS = ["상사 선물", "거래처 선물", "부모님 선물", "동료 선물", "첫만남 선물"]
        age_colors = {"19~24세": STEEL, "25~29세": PLUM, "30~34세": GOLD, "35~39세": SAGE}

        gr = gift_recipient_df[gift_recipient_df["연령대"].isin(RECIPIENT_AGES)]

        chart_card_open("연령대별 대상별 선물 키워드 검색 관심도", "0~12세 등 검색량이 극히 적은 구간은 제외")
        f_recipient = go.Figure()
        for age in RECIPIENT_AGES:
            gd = gr[gr["연령대"] == age].set_index("키워드").reindex(RECIPIENT_KEYWORDS)
            f_recipient.add_trace(go.Bar(x=RECIPIENT_KEYWORDS, y=gd["평균검색관심도"], name=age, marker_color=age_colors[age]))
        f_recipient.update_layout(barmode="group")
        st.plotly_chart(base_layout(f_recipient, height=360), use_container_width=True)
        chart_card_close()

        rmin, rmax = gr["평균검색관심도"].min(), gr["평균검색관심도"].max()
        insight(f"<b>{len(RECIPIENT_AGES)}개 연령대 x {len(RECIPIENT_KEYWORDS)}개 대상 = {len(RECIPIENT_AGES) * len(RECIPIENT_KEYWORDS)}개 구간 모두 "
                f"{rmin:.1f}~{rmax:.1f} 범위에 분포하며 0에 가까운 값이 없습니다</b> — 20~30대는 특정 관계에 치우치지 않고 "
                f"상사·거래처·부모님·동료·첫만남 등 다양한 관계에서 고르게 선물 검색 관심을 보입니다.")
        st.markdown('<div class="section-title">20~30대가 챙겨야 할 관계의 우선순위</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-desc">19~24세~35~39세 평균 — 어떤 관계가 상대적으로 더 신경 쓰이나</div>', unsafe_allow_html=True)

        recipient_means = gr.groupby("키워드")["평균검색관심도"].mean().reindex(RECIPIENT_KEYWORDS).sort_values(ascending=False)

        chart_card_open("대상별 평균 검색 관심도 (20~30대 평균)")
        f_rank = go.Figure(go.Bar(
            x=recipient_means.index, y=recipient_means.values,
            marker_color=GOLD,
            text=[f"{v:.1f}" for v in recipient_means.values],
            textposition="outside",
        ))
        st.plotly_chart(base_layout(f_rank, height=320, legend=False), use_container_width=True)
        chart_card_close()

        insight("20~30대가 신경 쓰는 관계는 가족(부모님)뿐 아니라 <b>새로운 인연(첫만남)과 직장 내 관계(상사·거래처·동료)까지 폭넓게 걸쳐 있습니다</b> "
                "— 학생 시기에는 크지 않았을 &ldquo;상사·거래처·첫만남&rdquo;이라는 관계가 사회 진출과 함께 새롭게 추가되며, "
                "<b>챙겨야 할 대상의 범위 자체가 넓어지고 있음</b>을 보여줍니다.")
    else:
        missing_note("naver_gift_recipient_demographics.csv")

    st.markdown('<div class="section-title">상황마다 매번 새로 찾아야 하는 선물추천</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">상황별 &ldquo;~선물추천&rdquo; 검색어를 20~30대가 얼마나 검색하나</div>', unsafe_allow_html=True)

    if gift_recommend_df is not None:
        RECOMMEND_AGES = ["19~24세", "25~29세", "30~34세", "35~39세"]
        RECOMMEND_KEYWORDS = ["부모님 선물추천", "상사 선물추천", "생일선물 추천", "첫만남 선물추천", "지인 선물추천"]
        recommend_age_colors = {"19~24세": STEEL, "25~29세": PLUM, "30~34세": GOLD, "35~39세": SAGE}

        grv = gift_recommend_df[gift_recommend_df["키워드"].isin(RECOMMEND_KEYWORDS) & gift_recommend_df["연령대"].isin(RECOMMEND_AGES)]

        chart_card_open("상황별 선물추천 키워드 x 연령대 검색 관심도")
        f_recommend = go.Figure()
        for age in RECOMMEND_AGES:
            gd = grv[grv["연령대"] == age].set_index("키워드").reindex(RECOMMEND_KEYWORDS)
            f_recommend.add_trace(go.Bar(x=RECOMMEND_KEYWORDS, y=gd["평균검색관심도"], name=age, marker_color=recommend_age_colors[age]))
        f_recommend.update_layout(barmode="group")
        st.plotly_chart(base_layout(f_recommend, height=360), use_container_width=True)
        st.markdown(f'<div style="font-size:11.5px; color:{TEXT_SUB}; margin-top:-4px;">'
                    f'※ &ldquo;센스있는 선물추천&rdquo;·&ldquo;적당한 선물추천&rdquo;은 표본 부족으로 제외</div>', unsafe_allow_html=True)
        chart_card_close()

        vmin, vmax = grv["평균검색관심도"].min(), grv["평균검색관심도"].max()
        insight(f"<b>5개의 서로 다른 상황별 &ldquo;선물추천&rdquo; 검색어 모두 20~30대 전 구간에서 {vmin:.1f}~{vmax:.1f}의 고른 검색량을 보입니다</b> "
                f"— 한 번의 검색으로 끝나지 않고, 상황이 바뀔 때마다(부모님/상사/생일/첫만남/지인) "
                f"<b>매번 새로운 추천을 다시 찾아봐야 하는 반복적인 의사결정 부담</b>을 보여줍니다.")
    else:
        missing_note("naver_gift_recommend_variants.csv")

    st.markdown('<div class="section-title">참고 — 연령·성별로 본 선물 니즈의 배경</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">선물 니즈 자체는 전 연령대에 고르게 나타남 — 아래는 배경 참고 데이터</div>', unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        if demo_df is not None:
            chart_card_open("오쏘몰 연령대 x 성별 검색 관심도", "값은 그룹별 12개월 내 상대적 안정성 지표 (그룹 간 절대 비교 아님)")
            od = demo_df[demo_df["브랜드"] == "오쏘몰"].copy()
            od["연령대"] = pd.Categorical(od["연령대"], categories=AGE_ORDER, ordered=True)
            od = od.sort_values("연령대")

            HIGHLIGHT_AGES = {"30~34세", "35~39세"}
            f5 = go.Figure()
            for gender, c in [("남성", GOLD), ("여성", PLUM)]:
                gd = od[od["성별"] == gender]
                colors = [RUST if age in HIGHLIGHT_AGES else c for age in gd["연령대"]]
                f5.add_trace(go.Bar(x=gd["연령대"], y=gd["평균검색관심도"], name=gender, marker_color=colors))

            overall_row = od.loc[od["평균검색관심도"].idxmax()]
            highlight_df = od[od["연령대"].isin(HIGHLIGHT_AGES)]
            thirties_row = highlight_df.loc[highlight_df["평균검색관심도"].idxmax()]
            gap = thirties_row["평균검색관심도"] - overall_row["평균검색관심도"]
            male_shift = -20  # 남성 막대는 그룹 내 왼쪽에 위치

            f5.add_annotation(x=overall_row["연령대"], y=overall_row["평균검색관심도"], text="최고 구간",
                               showarrow=False, yshift=16, xshift=male_shift if overall_row["성별"] == "남성" else 20,
                               font=dict(size=10, color=TEXT_SUB))
            f5.add_annotation(x=thirties_row["연령대"], y=thirties_row["평균검색관심도"], text=f"{gap:+.1f}p",
                               showarrow=False, yshift=16, xshift=male_shift if thirties_row["성별"] == "남성" else 20,
                               font=dict(size=10, color=RUST, family="'JetBrains Mono', monospace"))

            st.plotly_chart(base_layout(f5, height=320), use_container_width=True)
            chart_card_close()
        else:
            missing_note("naver_demographics.csv")

    with col4:
        if intent_df is not None:
            chart_card_open("구매의도 키워드 추이", "후기·가격·부작용 — 브랜드명 대비 규모 차이 있어 보조축 사용")
            f6 = make_subplots(specs=[[{"secondary_y": True}]])
            BRAND_G = "오쏘몰(브랜드명)"
            colors_map = {BRAND_G: GOLD, "오쏘몰 후기": STEEL, "오쏘몰 가격": RUST, "오쏘몰 부작용": PLUM, "오쏘몰 정품": SAGE}
            for g, c in colors_map.items():
                gd = intent_df[intent_df["키워드그룹"] == g].sort_values("기간")
                if gd.empty:
                    continue
                f6.add_trace(go.Scatter(x=gd["기간"], y=gd["검색관심도(상대값)"], name=g,
                                         line=dict(color=c, width=3 if g == BRAND_G else 1.6)),
                             secondary_y=(g == BRAND_G))
            st.plotly_chart(base_layout(f6, height=320), use_container_width=True)
            chart_card_close()
        else:
            missing_note("naver_purchase_intent.csv")

    if demo_df is not None or intent_df is not None:
        insight_parts = []

        if demo_df is not None:
            od = demo_df[demo_df["브랜드"] == "오쏘몰"].copy()
            male = od[od["성별"] == "남성"]
            female = od[od["성별"] == "여성"]
            rank_txt = ""
            if not male.empty and "35~39세" in male["연령대"].values:
                male_sorted = male.sort_values("평균검색관심도", ascending=False).reset_index(drop=True)
                rank_3539 = int(male_sorted.index[male_sorted["연령대"] == "35~39세"][0]) + 1
                higher = male_sorted.loc[: rank_3539 - 2, "연령대"].tolist()
                if higher:
                    rank_txt = (f"<b>35~39세(페르소나 구간)는 남성 중 {rank_3539}위</b>로 {'·'.join(higher)} 남성과 큰 차이가 없습니다 — "
                                f"즉 <b>선물 니즈는 특정 연령대에 국한되지 않고 전 연령대에 걸쳐 있다는 것</b>이 배경입니다. ")
            if not male.empty and not female.empty:
                insight_parts.append(f"오쏘몰 브랜드명 검색 자체는 남성 평균({male['평균검색관심도'].mean():.1f})이 여성({female['평균검색관심도'].mean():.1f})보다 높으며, {rank_txt}")

        insight_parts.append("5~8년차 직장인이 타겟으로 유효한 이유는 검색량 크기가 아니라, 위에서 확인한 <b>&ldquo;5가지 상황을 동시에 마주하는 빈도&rdquo;</b>에 있습니다.")

        if intent_df is not None:
            def trend_pct(group):
                d = intent_df[intent_df["키워드그룹"] == group].sort_values("기간")
                if len(d) < 6:
                    return None
                fv = d["검색관심도(상대값)"].iloc[:3].mean()
                lv = d["검색관심도(상대값)"].iloc[-3:].mean()
                return (lv - fv) / fv * 100 if fv else None

            review_pct = trend_pct("오쏘몰 후기")
            price_pct = trend_pct("오쏘몰 가격")
            if review_pct is not None and price_pct is not None:
                both_down = review_pct < 0 and price_pct < 0
                tail = ("브랜드 검색 관심도 하락과 함께 구매 직전 단계 키워드(후기·가격) 검색도 동반 감소해, "
                        "구매를 고려하는 잠재 고객 규모 자체가 줄고 있을 가능성을 시사합니다."
                        if both_down else
                        "브랜드 검색 관심도와 구매 직전 단계 키워드의 방향이 엇갈려, 추가 확인이 필요합니다.")
                insight_parts.append(
                    f"구매의도 측면에서는 <b>&ldquo;오쏘몰 후기&rdquo; 검색이 초기 3개월 대비 최근 3개월 평균 {review_pct:+.1f}%</b>, "
                    f"<b>&ldquo;오쏘몰 가격&rdquo;은 {price_pct:+.1f}%</b>로 나타났습니다 — {tail}"
                )

        insight(" ".join(insight_parts))

# ------------------------------------------------------------
# TAB 2. 콘텐츠 & 도달
# ------------------------------------------------------------
with tab2:
    st.markdown('<div class="section-title">콘텐츠 공급량</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">블로그·카페·뉴스에 쌓인 콘텐츠 총량과 생성 속도</div>', unsafe_allow_html=True)

    if content_summary_df is not None:
        col1, col2 = st.columns(2)
        with col1:
            chart_card_open("브랜드 x 채널 총 언급건수")
            f7 = go.Figure()
            for ch, c in [("블로그", INK), ("카페글", STEEL), ("뉴스", RUST)]:
                d = content_summary_df[content_summary_df["채널"] == ch]
                f7.add_trace(go.Bar(x=d["브랜드"], y=d["총_언급건수"], name=ch, marker_color=c))
            f7.update_layout(barmode="group")
            st.plotly_chart(base_layout(f7, height=320), use_container_width=True)
            chart_card_close()

        with col2:
            if content_posts_df is not None:
                chart_card_open("브랜드별 블로그 일평균 게시글 수")
                bd = content_posts_df[content_posts_df["채널"] == "블로그"].copy()
                bd["날짜"] = pd.to_datetime(bd["날짜"], format="%Y%m%d", errors="coerce")
                bd = bd.dropna(subset=["날짜"])
                rows = []
                for b, g in bd.groupby("브랜드"):
                    span = (g["날짜"].max() - g["날짜"].min()).days + 1
                    rows.append({"브랜드": b, "일평균": len(g) / span if span > 0 else 0})
                dv = pd.DataFrame(rows)
                f8 = go.Figure(go.Bar(x=dv["브랜드"], y=dv["일평균"], marker_color=[BRAND_COLORS.get(b, INK) for b in dv["브랜드"]]))
                st.plotly_chart(base_layout(f8, height=320, legend=False), use_container_width=True)
                chart_card_close()

        bc = content_summary_df[content_summary_df["채널"].isin(["블로그", "카페글"])].groupby("브랜드")["총_언급건수"].sum().sort_values(ascending=False)
        top_brand = bc.index[0] if not bc.empty else None

        gift_txt = ""
        if content_posts_df is not None and top_brand is not None:
            gift_ratio = content_posts_df.groupby("브랜드")["제목"].apply(lambda s: s.str.contains("선물", na=False).mean() * 100)
            if top_brand in gift_ratio.index and gift_ratio.notna().any():
                gift_rank = int(gift_ratio.rank(ascending=False)[top_brand])
                gift_txt = (f" 제목 키워드 분석 결과 <b>&ldquo;선물&rdquo; 프레이밍 비율도 4개 브랜드 중 {top_brand}이 {gift_rank}위"
                            f"({gift_ratio[top_brand]:.1f}%)</b>로, 프리미엄 선물 포지셔닝이 콘텐츠 레벨에서도 확인됩니다.")

        if top_brand is not None:
            insight(f"{top_brand}은 <b>블로그·카페 총 언급량 1위</b>지만, 검색 관심도(수요)는 하락 추세입니다 — 콘텐츠 공급과 실제 수요가 엇갈리는 지점입니다.{gift_txt}")
    else:
        missing_note("naver_content_summary.csv", "naver_content_recent_posts.csv")

    st.markdown('<div class="section-title">뉴스 언급 · 연관어</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">빅카인즈 뉴스빅데이터 분석</div>', unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        if news_keyword_df is not None:
            nd = news_keyword_df.copy()
            nd["기간"] = pd.to_datetime(nd["date"].astype(str), format="%Y%m")
            nd = nd.sort_values("기간")
            nd["label"] = nd["기간"].dt.strftime("%Y-%m")
            peak_label = nd.loc[nd["오쏘몰"].idxmax(), "label"]
            chart_card_open("월별 뉴스 언급 건수", f"{peak_label}은 노이즈 검증 필요 구간(하단 인사이트 참고)")
            colors = [RUST if lb == peak_label else INK for lb in nd["label"]]
            f9 = go.Figure(go.Bar(x=nd["label"], y=nd["오쏘몰"], marker_color=colors))
            st.plotly_chart(base_layout(f9, height=300, legend=False), use_container_width=True)
            chart_card_close()
        else:
            missing_note("오쏘몰_키워드.xlsx")

    with col4:
        if related_words_df is not None:
            chart_card_open("상위 연관어 (가중치)")
            top = related_words_df.sort_values("가중치", ascending=False).head(10).sort_values("가중치")
            f10 = go.Figure(go.Bar(x=top["가중치"], y=top["키워드"], orientation="h", marker_color=GOLD))
            st.plotly_chart(base_layout(f10, height=300, legend=False), use_container_width=True)
            chart_card_close()
        else:
            missing_note("연관어분석_20260828.xlsx")

    if news_keyword_df is not None or related_words_df is not None:
        insight_parts = ["3월(3년 연속 1위 발표)·4월(ODP 신제품 출시)처럼 <b>기업 발표 시점에 언급이 집중</b>되는 패턴이 뚜렷합니다."]
        if related_words_df is not None:
            top_words = related_words_df.sort_values("가중치", ascending=False).head(4)["키워드"].tolist()
            words_txt = "·".join(f"&ldquo;{w}&rdquo;" for w in top_words)
            insight_parts.append(f"연관어 상위에 {words_txt} 등이 나와 브랜드 포지셔닝(프리미엄·이중제형)과 일치합니다.")
        if news_keyword_df is not None:
            insight_parts.append(f"⚠️ <b>{peak_label} 스파이크는 상당수가 무관한 기사(경품 이벤트 등 스치는 언급)로 노이즈가 섞여있어 해석에 주의</b>가 필요합니다.")
        insight(" ".join(insight_parts))

    st.markdown('<div class="section-title">유튜브 콘텐츠</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">영상 조회수 · 인게이지먼트 비교</div>', unsafe_allow_html=True)

    if youtube_df is not None:
        col5, col6 = st.columns(2)
        with col5:
            chart_card_open("평균 vs 중앙값 조회수", "로그스케일 — 평균은 이상치에 왜곡될 수 있음")
            vs = youtube_df.groupby("브랜드")["조회수"].agg(["mean", "median"]).reset_index()
            f11 = go.Figure()
            f11.add_trace(go.Bar(x=vs["브랜드"], y=vs["mean"], name="평균", marker_color=INK))
            f11.add_trace(go.Bar(x=vs["브랜드"], y=vs["median"], name="중앙값", marker_color=GOLD))
            f11.update_layout(barmode="group", yaxis_type="log")
            st.plotly_chart(base_layout(f11, height=300), use_container_width=True)
            chart_card_close()

        with col6:
            chart_card_open("평균 좋아요 · 댓글수")
            es = youtube_df.groupby("브랜드")[["좋아요수", "댓글수"]].mean().reset_index()
            f12 = go.Figure()
            f12.add_trace(go.Bar(x=es["브랜드"], y=es["좋아요수"], name="좋아요", marker_color=INK))
            f12.add_trace(go.Bar(x=es["브랜드"], y=es["댓글수"], name="댓글", marker_color=GOLD))
            f12.update_layout(barmode="group")
            st.plotly_chart(base_layout(f12, height=300), use_container_width=True)
            chart_card_close()

        vs_idx = vs.set_index("브랜드")
        if "오쏘몰" in vs_idx.index and vs_idx.loc["오쏘몰", "median"] > 0 and len(vs_idx) > 1:
            top_brand = vs_idx["median"].idxmax()
            is_lowest = vs_idx["median"].idxmin() == "오쏘몰" and vs_idx["mean"].idxmin() == "오쏘몰"
            rank_txt = "최하위" if is_lowest else f"{int(vs_idx['median'].rank(ascending=False)['오쏘몰'])}위"
            ratio = vs_idx.loc[top_brand, "median"] / vs_idx.loc["오쏘몰", "median"]
            insight(f"오쏘몰은 {len(vs_idx)}개 브랜드 중 <b>중앙값 조회수 기준 {rank_txt}</b>({top_brand} 대비 약 <b>{ratio:.0f}배</b> 낮음)입니다. "
                    f"상위 영상 중 상당수가 자동생성형 &ldquo;랭킹&rdquo; 채널의 저품질 콘텐츠이며, 진짜 브랜드 콘텐츠는 공식 채널(동아쏘시오그룹) 소수에 그칩니다. "
                    f"텍스트 콘텐츠(블로그)는 강세였지만 <b>영상 콘텐츠는 뚜렷한 약점 영역</b>입니다.")
    else:
        missing_note("youtube_content.csv")

# ------------------------------------------------------------
# TAB 3. 쇼핑 행동
# ------------------------------------------------------------
with tab3:
    st.markdown('<div class="section-title">쇼핑 클릭 트렌드</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">네이버 쇼핑인사이트 · 비타민/미네랄 카테고리 전체</div>', unsafe_allow_html=True)

    if shopping_df is not None:
        chart_card_open("카테고리 전체 클릭 추이 vs 오쏘몰 검색 관심도")
        sd = shopping_df.copy()
        sd["기간"] = pd.to_datetime(sd["기간"])
        f13 = make_subplots(specs=[[{"secondary_y": True}]])
        f13.add_trace(go.Scatter(x=sd["기간"], y=sd["클릭비율(상대값)"], name="카테고리 전체 클릭",
                                  line=dict(color=STEEL, width=2), fill="tozeroy", fillcolor="rgba(76,110,145,0.08)"), secondary_y=False)
        if brand_df is not None:
            o = brand_df[brand_df["브랜드"] == "오쏘몰"].sort_values("기간")
            f13.add_trace(go.Scatter(x=o["기간"], y=o["검색관심도(상대값)"], name="오쏘몰 검색",
                                      line=dict(color=GOLD, width=2.5)), secondary_y=True)
        st.plotly_chart(base_layout(f13, height=360), use_container_width=True)
        chart_card_close()

        insight("브랜드명 단위 클릭 데이터는 표본이 희박해 집계되지 않아, 카테고리 전체 추이로 대신 확인했습니다. "
                "시장 전체가 위축되는 추세인지, 오쏘몰만의 하락인지 두 라인의 기울기를 비교해보세요.")
    else:
        missing_note("naver_shopping_category_trend.csv")

# ------------------------------------------------------------
# TAB 4. 교차 검증
# ------------------------------------------------------------
with tab4:
    st.markdown('<div class="section-title">네이버 vs 구글 교차 검증</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">서로 다른 검색엔진에서도 같은 추세가 확인되는지</div>', unsafe_allow_html=True)

    if brand_df is not None and google_df is not None:
        o_naver = apply_date_range(brand_df)
        o_naver = o_naver[o_naver["브랜드"] == "오쏘몰"].sort_values("기간")
        o_google = apply_date_range(google_df)
        o_google = o_google[o_google["브랜드"] == "오쏘몰"].sort_values("기간")

        chart_card_open("오쏘몰 검색 관심도: 네이버(월간) vs 구글(주간)")
        f14 = go.Figure()
        f14.add_trace(go.Scatter(x=o_naver["기간"], y=o_naver["검색관심도(상대값)"], name="네이버 (월간)",
                                  line=dict(color=GOLD, width=3), mode="lines+markers"))
        f14.add_trace(go.Scatter(x=o_google["기간"], y=o_google["검색관심도(상대값)"], name="구글 (주간)",
                                  line=dict(color=STEEL, width=2, dash="dot"), mode="lines+markers", marker=dict(size=4)))
        st.plotly_chart(base_layout(f14, height=380), use_container_width=True)
        chart_card_close()

        if len(o_naver) >= 2 and len(o_google) >= 2:
            nv = o_naver["검색관심도(상대값)"]
            n_pct = (nv.iloc[-1] - nv.iloc[0]) / nv.iloc[0] * 100
            gv = o_google["검색관심도(상대값)"]
            g_head = gv.iloc[: min(4, len(gv))].mean()
            g_tail = gv.iloc[-min(4, len(gv)):].mean()
            g_pct = (g_tail - g_head) / g_head * 100

            col1, col2 = st.columns(2)
            with col1:
                st.markdown(kpi_html("네이버 (첫달 대비 마지막달)", f"{nv.iloc[-1]:.1f}", f"{n_pct:+.1f}%", "down" if n_pct < 0 else "up"), unsafe_allow_html=True)
            with col2:
                st.markdown(kpi_html("구글 (초반 대비 최근 평균)", f"{g_tail:.1f}", f"{g_pct:+.1f}%", "down" if g_pct < 0 else "up"), unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            same_dir = (n_pct < 0) == (g_pct < 0)
            insight(f"네이버({n_pct:+.1f}%)와 구글({g_pct:+.1f}%) "
                    + (f"<b>두 검색엔진에서 방향성이 일치</b> — 오쏘몰 검색 관심도 변화가 특정 플랫폼의 우연이 아니라 <b>실제 시장 추세</b>임을 교차 검증합니다."
                       if same_dir else
                       f"<b>두 검색엔진에서 방향성이 엇갈립니다</b> — 선택한 기간에서는 플랫폼 간 차이가 있어 추가 확인이 필요합니다.")
                    + " 집계 주기(월간 vs 주간)가 달라 계산 방식은 다르지만, 같은 기간을 비교했습니다.")
        else:
            st.info("선택한 기간에 데이터가 부족합니다 — 기간을 넓혀주세요.")
    else:
        missing_note("naver_search_trend.csv", "google_trend_kr.csv")

# ============================================================
# 최종 메시지: 고민 없는 선물
# ============================================================
st.markdown('<div class="section-title" style="margin-top:32px;">최종 메시지: 고민 없는 선물</div>', unsafe_allow_html=True)
st.markdown(f"""
<div style="font-family:'Inter',sans-serif; font-size:26px; font-weight:800; letter-spacing:-0.02em;
            color:{INK}; line-height:1.4; margin:8px 0 24px 0;">
오래 고민하지 않아도 되는, 누구에게나 줘도 센스있고 적당한 선물
</div>
""", unsafe_allow_html=True)

REASON_FOR = {"프리미엄 이미지": 39, "효과·성분": 24, "구성·혜택": 22, "브랜드 신뢰": 16}
REASON_AGAINST = {"가격 부담": 5, "취향 안 맞음": 3, "고민됨·확신없음": 0}
MUTED = "#D5D7DC"

col_final1, col_final2 = st.columns(2)
with col_final1:
    chart_card_open("오쏘몰을 선물로 정한 이유")
    f_for = go.Figure(go.Bar(
        x=list(REASON_FOR.keys()), y=list(REASON_FOR.values()), marker_color=GOLD,
        text=[str(v) for v in REASON_FOR.values()], textposition="outside",
    ))
    st.plotly_chart(base_layout(f_for, height=300, legend=False), use_container_width=True)
    chart_card_close()

with col_final2:
    chart_card_open("오쏘몰을 꺼리는 이유")
    against_colors = [RUST if k == "가격 부담" else MUTED for k in REASON_AGAINST]
    f_against = go.Figure(go.Bar(
        x=list(REASON_AGAINST.keys()), y=list(REASON_AGAINST.values()), marker_color=against_colors,
        text=[str(v) for v in REASON_AGAINST.values()], textposition="outside",
    ))
    f_against.add_annotation(x="고민됨·확신없음", y=0, text="확신이 없어서가 아니라<br>가격이 유일한 걸림돌",
                              showarrow=False, yshift=45, font=dict(size=10.5, color=TEXT_SUB))
    st.plotly_chart(base_layout(f_against, height=300, legend=False), use_container_width=True)
    chart_card_close()

def mini_bar_fig(labels, values, colors):
    fig = go.Figure(go.Bar(x=labels, y=values, marker_color=colors))
    fig = base_layout(fig, height=150, legend=False)
    fig.update_xaxes(tickfont=dict(size=8.5), tickangle=-30, title=None)
    fig.update_yaxes(showticklabels=False, title=None, showgrid=False, zeroline=False)
    fig.update_layout(margin=dict(t=6, b=30, l=4, r=4))
    return fig


def tag_card_open(label):
    st.markdown(f'<div style="background:{CARD}; border:1px solid {LINE}; border-radius:14px; '
                f'padding:14px 12px 10px 12px; box-shadow:0 1px 2px rgba(15,17,21,0.03);">'
                f'<div style="font-family:\'Inter\',sans-serif; font-size:15px; font-weight:800; color:{AMBER}; margin-bottom:2px;">{label}</div>',
                unsafe_allow_html=True)


def tag_card_close(caption):
    st.markdown(f'<div style="font-size:11px; color:{TEXT_SUB}; text-align:center; margin-top:-4px;">{caption}</div></div>',
                unsafe_allow_html=True)


TAG_AGES = ["19~24세", "25~29세", "30~34세", "35~39세"]
tag_cols = st.columns(4)

with tag_cols[0]:
    tag_card_open("고민하지 않아도")
    if gift_recommend_df is not None:
        c1_map = {"부모님 선물추천": "부모님", "상사 선물추천": "상사", "생일선물 추천": "생일",
                  "첫만남 선물추천": "첫만남", "지인 선물추천": "지인"}
        c1 = gift_recommend_df[gift_recommend_df["키워드"].isin(c1_map) & gift_recommend_df["연령대"].isin(TAG_AGES)]
        c1_means = c1.groupby("키워드")["평균검색관심도"].mean().reindex(c1_map.keys())
        st.plotly_chart(mini_bar_fig([c1_map[k] for k in c1_means.index], c1_means.values, GOLD),
                         use_container_width=True, config={"displayModeBar": False})
    tag_card_close("5개 상황 모두 33~67 구간")

with tag_cols[1]:
    tag_card_open("누구에게나")
    if gift_recipient_df is not None:
        c2_map = {"상사 선물": "상사", "거래처 선물": "거래처", "부모님 선물": "부모님", "동료 선물": "동료", "첫만남 선물": "첫만남"}
        c2 = gift_recipient_df[gift_recipient_df["키워드"].isin(c2_map) & gift_recipient_df["연령대"].isin(TAG_AGES)]
        c2_means = c2.groupby("키워드")["평균검색관심도"].mean().reindex(c2_map.keys())
        st.plotly_chart(mini_bar_fig([c2_map[k] for k in c2_means.index], c2_means.values, STEEL),
                         use_container_width=True, config={"displayModeBar": False})
    tag_card_close("5개 관계 모두 22.8~61.8 구간")

with tag_cols[2]:
    tag_card_open("센스있고")
    c3_colors = [GOLD if k == "프리미엄 이미지" else MUTED for k in REASON_FOR]
    st.plotly_chart(mini_bar_fig(list(REASON_FOR.keys()), list(REASON_FOR.values()), c3_colors),
                     use_container_width=True, config={"displayModeBar": False})
    tag_card_close("프리미엄 이미지, 압도적 1위")

with tag_cols[3]:
    tag_card_open("적당한")
    c4_labels = ["긍정 요인", "가격 부담"]
    c4_values = [REASON_FOR["효과·성분"] + REASON_FOR["구성·혜택"], REASON_AGAINST["가격 부담"]]
    st.plotly_chart(mini_bar_fig(c4_labels, c4_values, [SAGE, RUST]),
                     use_container_width=True, config={"displayModeBar": False})
    tag_card_close(f"긍정 요인 {c4_values[0]}건 vs 가격부담 {c4_values[1]}건")

st.markdown(f'<div style="font-size:11px; color:{TEXT_SUB}; margin:16px 0 8px 0; line-height:1.6;">'
            f'※ 네이버 블로그·카페 게시물 텍스트 키워드 빈도 기반 간이분석(총 149건), 게시물 대상이 20~30대로 언급된 글 기준이며 '
            f'글쓴이 나이는 아님</div>', unsafe_allow_html=True)

st.markdown(f"""
<div style="margin-top:36px; padding-top:16px; border-top:1px solid {LINE}; font-family:'Inter',sans-serif;
            font-size:11.5px; color:{TEXT_SUB}; line-height:1.7;">
※ 검색 관심도는 각 데이터셋 내 최고치를 100으로 하는 상대 지수이며, 실제 검색량(절대값)이 아닙니다.<br>
※ 매출·구매 데이터는 포함되어 있지 않으며, 모든 해석은 공개 데이터 기반의 가설 수준으로 참고 바랍니다.
</div>
""", unsafe_allow_html=True)
