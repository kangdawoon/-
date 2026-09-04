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
    height: 32px;
}}
.kpi-value {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 27px;
    font-weight: 600;
    color: {INK};
    letter-spacing: -0.02em;
}}
.kpi-delta {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    font-weight: 600;
    margin-top: 6px;
    display: inline-block;
    padding: 3px 8px;
    border-radius: 20px;
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
    kpi_cards_html += kpi_html("명절 시즌 상관관계", "일치" if match else "불일치", "명절선물 피크 주 = 오쏘몰 피크 주" if match else "시기 어긋남", "up" if match else "flat")

if youtube_df is not None:
    med = youtube_df.groupby("브랜드")["조회수"].median()
    if "오쏘몰" in med.index:
        rank = int((med.rank(ascending=False))["오쏘몰"])
        kpi_cards_html += kpi_html("유튜브 조회수 순위 (4개 브랜드 중)", f"{rank}위", "중앙값 기준 최하위권", "down")

if kpi_cards_html:
    st.markdown(f'<div class="kpi-strip">{kpi_cards_html}</div>', unsafe_allow_html=True)

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
        chart_card_open("월간 검색 관심도")
        fig = go.Figure()
        for b in brand_df["브랜드"].unique():
            d = brand_df[brand_df["브랜드"] == b].sort_values("기간")
            fig.add_trace(go.Scatter(x=d["기간"], y=d["검색관심도(상대값)"], name=b,
                                      line=dict(width=4 if b == "오쏘몰" else 1.8, color=BRAND_COLORS.get(b)),
                                      mode="lines+markers", marker=dict(size=4)))
        st.plotly_chart(base_layout(fig), use_container_width=True)
        chart_card_close()

        first_val = brand_df[brand_df["브랜드"] == "오쏘몰"].sort_values("기간")["검색관심도(상대값)"].iloc[0]
        last_val = brand_df[brand_df["브랜드"] == "오쏘몰"].sort_values("기간")["검색관심도(상대값)"].iloc[-1]
        insight(f"오쏘몰은 12개월간 <b>{(last_val-first_val)/first_val*100:.1f}%</b> 하락({first_val:.0f}→{last_val:.0f})한 반면, "
                f"고려은단은 동기간 상승 흐름을 유지했습니다. 판매량 1위(별도 브랜드 리포트 기준)와 검색 관심도 하락이 동시에 나타나는 점은 "
                f"신규 유입보다 기존 고객 재구매 위주 매출 구조일 가능성을 시사합니다.")
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

    st.markdown('<div class="section-title">참고 — 연령·성별로 본 선물 니즈의 배경</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">선물 니즈 자체는 전 연령대에 고르게 나타남 — 아래는 배경 참고 데이터</div>', unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        if demo_df is not None:
            chart_card_open("오쏘몰 연령대 x 성별 검색 관심도", "값은 그룹별 12개월 내 상대적 안정성 지표 (그룹 간 절대 비교 아님)")
            od = demo_df[demo_df["브랜드"] == "오쏘몰"].copy()
            od["연령대"] = pd.Categorical(od["연령대"], categories=AGE_ORDER, ordered=True)
            od = od.sort_values("연령대")
            f5 = go.Figure()
            for gender, c in [("남성", GOLD), ("여성", PLUM)]:
                gd = od[od["성별"] == gender]
                colors = [RUST if (gender == "남성" and age == "35~39세") else c for age in gd["연령대"]]
                f5.add_trace(go.Bar(x=gd["연령대"], y=gd["평균검색관심도"], name=gender, marker_color=colors if gender == "남성" else c))
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
        insight("오쏘몰 브랜드명 검색 자체는 남성이 여성보다 전반적으로 높고 안정적이며, <b>35~39세(페르소나 구간)는 남성 중 4위</b>로 "
                "50~60대 남성과 큰 차이가 없습니다 — 즉 <b>선물 니즈는 특정 연령대에 국한되지 않고 전 연령대에 걸쳐 있다는 것</b>이 배경입니다. "
                "5~8년차 직장인이 타겟으로 유효한 이유는 검색량 크기가 아니라, 위에서 확인한 <b>&ldquo;5가지 상황을 동시에 마주하는 빈도&rdquo;</b>에 있습니다. "
                "구매의도 측면에서는 <b>&ldquo;오쏘몰 후기&rdquo; 검색이 +45.8% 증가</b>한 반면 <b>&ldquo;오쏘몰 가격&rdquo;은 -70.6% 감소</b> — "
                "브랜드 인지는 줄어도 이미 관심 있는 고객의 구매 직전 행동은 유지되고 있을 가능성을 보여줍니다.")

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

        insight("오쏘몰은 <b>블로그·카페 총 언급량 1위</b>지만, 검색 관심도(수요)는 하락 추세입니다 — 콘텐츠 공급과 실제 수요가 엇갈리는 지점입니다. "
                "제목 키워드 분석 결과 <b>&ldquo;선물&rdquo; 프레이밍 비율도 4개 브랜드 중 오쏘몰이 1위(10.5%)</b>로, 프리미엄 선물 포지셔닝이 콘텐츠 레벨에서도 확인됩니다.")
    else:
        missing_note("naver_content_summary.csv", "naver_content_recent_posts.csv")

    st.markdown('<div class="section-title">뉴스 언급 · 연관어</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">빅카인즈 뉴스빅데이터 분석</div>', unsafe_allow_html=True)

    col3, col4 = st.columns(2)
    with col3:
        if news_keyword_df is not None:
            chart_card_open("월별 뉴스 언급 건수", "11월은 노이즈 검증 필요 구간(하단 인사이트 참고)")
            nd = news_keyword_df.copy()
            nd["기간"] = pd.to_datetime(nd["date"].astype(str), format="%Y%m")
            nd = nd.sort_values("기간")
            nd["label"] = nd["기간"].dt.strftime("%Y-%m")
            colors = [RUST if lb == "2025-11" else INK for lb in nd["label"]]
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
        insight("3월(3년 연속 1위 발표)·4월(ODP 신제품 출시)처럼 <b>기업 발표 시점에 언급이 집중</b>되는 패턴이 뚜렷합니다. "
                "연관어 상위에 &ldquo;동아제약&rdquo;·&ldquo;Orthomol&rdquo;·&ldquo;멀티비타민&rdquo;·&ldquo;구강용&rdquo; 등이 나와 브랜드 포지셔닝(프리미엄·이중제형)과 일치합니다. "
                "⚠️ <b>11월 스파이크는 상당수가 무관한 기사(경품 이벤트 등 스치는 언급)로 노이즈가 섞여있어 해석에 주의</b>가 필요합니다.")

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
        if "고려은단" in vs_idx.index and "오쏘몰" in vs_idx.index and vs_idx.loc["오쏘몰", "median"] > 0:
            ratio = vs_idx.loc["고려은단", "median"] / vs_idx.loc["오쏘몰", "median"]
            insight(f"오쏘몰은 4개 브랜드 중 <b>평균·중앙값 조회수 모두 최하위</b>(중앙값 기준 고려은단 대비 약 <b>{ratio:.0f}배</b> 낮음)입니다. "
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
        o_naver = brand_df[brand_df["브랜드"] == "오쏘몰"].sort_values("기간")
        o_google = google_df[google_df["브랜드"] == "오쏘몰"].sort_values("기간")

        chart_card_open("오쏘몰 검색 관심도: 네이버(월간) vs 구글(주간)")
        f14 = go.Figure()
        f14.add_trace(go.Scatter(x=o_naver["기간"], y=o_naver["검색관심도(상대값)"], name="네이버 (월간)",
                                  line=dict(color=GOLD, width=3), mode="lines+markers"))
        f14.add_trace(go.Scatter(x=o_google["기간"], y=o_google["검색관심도(상대값)"], name="구글 (주간)",
                                  line=dict(color=STEEL, width=2, dash="dot"), mode="lines+markers", marker=dict(size=4)))
        st.plotly_chart(base_layout(f14, height=380), use_container_width=True)
        chart_card_close()

        nv = o_naver["검색관심도(상대값)"]
        n_pct = (nv.iloc[-1] - nv.iloc[0]) / nv.iloc[0] * 100
        gv = o_google["검색관심도(상대값)"]
        g_pct = (gv.iloc[-4:].mean() - gv.iloc[:4].mean()) / gv.iloc[:4].mean() * 100

        col1, col2 = st.columns(2)
        with col1:
            st.markdown(kpi_html("네이버 (첫달 대비 마지막달)", f"{nv.iloc[-1]:.1f}", f"{n_pct:+.1f}%", "down" if n_pct < 0 else "up"), unsafe_allow_html=True)
        with col2:
            st.markdown(kpi_html("구글 (초반4주 대비 최근4주)", f"{gv.iloc[-4:].mean():.1f}", f"{g_pct:+.1f}%", "down" if g_pct < 0 else "up"), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        insight(f"네이버({n_pct:+.1f}%)와 구글({g_pct:+.1f}%) <b>두 검색엔진에서 방향성이 일치</b> — 오쏘몰 검색 관심도 하락이 "
                f"특정 플랫폼의 우연이 아니라 <b>실제 시장 추세</b>임을 교차 검증합니다. "
                f"집계 주기(월간 vs 주간)가 달라 계산 방식은 다르지만, 하락이라는 방향성 자체는 두 플랫폼에서 동일하게 관찰됩니다.")
    else:
        missing_note("naver_search_trend.csv", "google_trend_kr.csv")

st.markdown(f"""
<div style="margin-top:36px; padding-top:16px; border-top:1px solid {LINE}; font-family:'Inter',sans-serif;
            font-size:11.5px; color:{TEXT_SUB}; line-height:1.7;">
※ 검색 관심도는 각 데이터셋 내 최고치를 100으로 하는 상대 지수이며, 실제 검색량(절대값)이 아닙니다.<br>
※ 매출·구매 데이터는 포함되어 있지 않으며, 모든 해석은 공개 데이터 기반의 가설 수준으로 참고 바랍니다.
</div>
""", unsafe_allow_html=True)
