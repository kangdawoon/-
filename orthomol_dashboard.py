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
tab1, tab2, tab3, tab4 = st.tabs(["🔍  검색 관심도", "📢  콘텐츠 & 도달", "🛒  쇼핑 행동", "✅  교차 검증"])

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
        insight(f"명절선물·오쏘몰 검색 <b>{'피크 시점 일치' if match else '피크 시점 불일치'}</b> — 명절 마케팅이 실제 수요와 맞물려 작동 중임을 시사합니다. "
                f"반면 &ldquo;선물 추천&rdquo;·&ldquo;직장인 선물&rdquo; 검색은 대체로 연말(12월)에 몰리는데, 같은 시기 오쏘몰 검색은 연중 저점을 보여 "
                f"<b>&ldquo;명절&rdquo; 프레임에는 강하지만 &ldquo;연말·직장인 선물&rdquo; 프레임의 수요는 놓치고 있을 가능성</b>이 있습니다.")
    else:
        missing_note("naver_seasonality_trend.csv")

    st.markdown('<div class="section-title">누가, 어떤 목적으로 검색하나</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-desc">연령·성별 분포와 구매 단계별 키워드 — 나란히 비교</div>', unsafe_allow_html=True)

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
        insight_parts = []

        if demo_df is not None:
            od = demo_df[demo_df["브랜드"] == "오쏘몰"].copy()
            male = od[od["성별"] == "남성"]
            female = od[od["성별"] == "여성"]
            male_rank_txt = ""
            if not male.empty and "35~39세" in male["연령대"].values:
                male_sorted = male.sort_values("평균검색관심도", ascending=False).reset_index(drop=True)
                rank_3539 = int(male_sorted.index[male_sorted["연령대"] == "35~39세"][0]) + 1
                higher = male_sorted.loc[: rank_3539 - 2, "연령대"].tolist()
                if higher:
                    male_rank_txt = (f"<b>35~39세(페르소나 구간)는 남성 중 {rank_3539}위</b>로 오히려 {'·'.join(higher)} 남성이 근소하게 더 높습니다 — "
                                      f"타겟을 &ldquo;35세 단독&rdquo;보다 &ldquo;남성 3060 전반&rdquo;으로 넓게 해석할 여지가 있습니다. ")
            if not male.empty and not female.empty:
                insight_parts.append(f"남성 평균 검색 관심도({male['평균검색관심도'].mean():.1f})가 여성({female['평균검색관심도'].mean():.1f})보다 높으며, {male_rank_txt}")

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

        if insight_parts:
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

st.markdown(f"""
<div style="margin-top:36px; padding-top:16px; border-top:1px solid {LINE}; font-family:'Inter',sans-serif;
            font-size:11.5px; color:{TEXT_SUB}; line-height:1.7;">
※ 검색 관심도는 각 데이터셋 내 최고치를 100으로 하는 상대 지수이며, 실제 검색량(절대값)이 아닙니다.<br>
※ 매출·구매 데이터는 포함되어 있지 않으며, 모든 해석은 공개 데이터 기반의 가설 수준으로 참고 바랍니다.
</div>
""", unsafe_allow_html=True)
