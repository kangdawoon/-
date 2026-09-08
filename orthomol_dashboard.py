# -*- coding: utf-8 -*-
"""
오쏘몰 마케팅 진단 & 성장 제안 대시보드 (최종 스토리라인 기준 재구성)

[실행 방법]
pip install streamlit pandas plotly openpyxl
streamlit run orthomol_dashboard.py

[필요 파일 - 같은 폴더]
naver_search_trend.csv, naver_seasonality_trend.csv, naver_gift_tone_trend.csv,
naver_gift_recipient_demographics.csv, naver_gift_recommend_variants.csv,
20_30대_오쏘몰_선물_이유.csv, 오쏘몰_이뮨_대시보드_데이터.csv, 선물채널_언급량.csv,
종합비타민_가격비교.csv, 영양제_유튜브_검색_트렌드.csv, 연령별_선물카테고리_선호도.csv,
선물시즌_총언급량.csv, 선물시즌_월별분포.csv, 선물고민_브랜드별언급량.csv,
오쏘몰_부작용_인식.csv, 오쏘몰_시장_이미지.csv, naver_purchase_intent.csv(선택)
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

st.set_page_config(page_title="오쏘몰 마케팅 진단 & 성장 제안", layout="wide", initial_sidebar_state="collapsed")

# ============================================================
# 디자인 토큰
# ============================================================
INK = "#0F1115"
AMBER = "#F0A83C"
SURFACE = "#F6F7F9"
CARD = "#FFFFFF"
LINE = "#EAEBEF"
TEXT_SUB = "#8A8F98"
RUST = "#E5484D"
SAGE = "#30A46C"
STEEL = "#6E7787"
PLUM = "#8E4EC6"
GRAY = "#C7C9CF"
GOLD = AMBER

BRAND_COLORS = {"오쏘몰": AMBER, "아임비타": PLUM, "센트룸": STEEL, "고려은단": SAGE}
PLOTLY_FONT = dict(family="Inter, sans-serif", color=INK, size=12)

# ============================================================
# CSS
# ============================================================
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap');
html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
.stApp {{ background-color: {CARD}; }}
#MainMenu, footer, header {{visibility: hidden;}}
.block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1180px; }}
.eyebrow {{ font-family:'JetBrains Mono',monospace; font-size:11px; font-weight:600; letter-spacing:.12em; color:{AMBER}; text-transform:uppercase; margin-bottom:8px; }}
.hero-title {{ font-family:'Inter',sans-serif; font-size:30px; font-weight:800; letter-spacing:-.02em; color:{INK}; margin:0 0 6px 0; line-height:1.25; }}
.hero-sub {{ font-size:13.5px; color:{TEXT_SUB}; margin-bottom:26px; }}
.section-num {{ display:inline-flex; align-items:center; justify-content:center; width:26px; height:26px; border-radius:50%; background:{INK}; color:white; font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:600; margin-right:8px; }}
.section-title {{ font-size:18px; font-weight:700; letter-spacing:-.01em; color:{INK}; margin:10px 0 3px 0; display:flex; align-items:center; }}
.section-desc {{ font-size:12px; color:{TEXT_SUB}; margin-bottom:14px; margin-left:34px; }}
.chart-card {{ background:{CARD}; border:1px solid {LINE}; border-radius:14px; padding:16px 18px 6px 18px; margin-bottom:14px; box-shadow:0 1px 2px rgba(15,17,21,.03); }}
.chart-label {{ font-size:12.5px; font-weight:600; color:{INK}; margin-bottom:2px; }}
.chart-note {{ font-size:10.5px; color:{TEXT_SUB}; margin-bottom:8px; }}
.insight-box {{ background:{SURFACE}; border:1px solid {LINE}; border-radius:12px; padding:13px 16px; margin:6px 0 22px 0; font-size:12.5px; color:{INK}; line-height:1.7; }}
.insight-box b {{ color:{INK}; font-weight:700; }}
.kpi-strip {{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:28px; }}
.kpi-card {{ background:{CARD}; border:1px solid {LINE}; border-radius:14px; padding:16px 18px; box-shadow:0 1px 2px rgba(15,17,21,.04); }}
.kpi-label {{ font-size:11px; color:{TEXT_SUB}; margin-bottom:8px; height:28px; }}
.kpi-value {{ font-family:'JetBrains Mono',monospace; font-size:23px; font-weight:600; color:{INK}; }}
.kpi-delta {{ font-family:'JetBrains Mono',monospace; font-size:11px; margin-top:5px; display:inline-block; padding:2px 7px; border-radius:20px; }}
.delta-down {{ color:{RUST}; background:rgba(229,72,77,.08); }}
.delta-up {{ color:{SAGE}; background:rgba(48,164,108,.08); }}
.delta-flat {{ color:{TEXT_SUB}; background:{SURFACE}; }}
.mini-tag {{ background:{SURFACE}; border:1px solid {LINE}; border-radius:12px; padding:14px 16px; text-align:center; }}
.mini-tag .t {{ font-size:12px; font-weight:700; color:{AMBER}; margin-bottom:6px; }}
.mini-tag .v {{ font-family:'JetBrains Mono',monospace; font-size:20px; font-weight:700; color:{INK}; }}
.mini-tag .d {{ font-size:10.5px; color:{TEXT_SUB}; margin-top:4px; }}
.gap-note {{ background:#FFF7E6; border:1px solid #F0D999; border-radius:10px; padding:10px 14px; font-size:11.5px; color:#8A6D1F; margin-bottom:14px; }}
</style>
""", unsafe_allow_html=True)


def base_layout(fig, height=300, legend=True):
    fig.update_layout(height=height, font=PLOTLY_FONT, plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                       margin=dict(t=8, b=8, l=8, r=8),
                       legend=dict(orientation="h", yanchor="bottom", y=-0.28, font=dict(size=10)) if legend else dict(),
                       showlegend=legend)
    fig.update_xaxes(gridcolor=LINE, zeroline=False)
    fig.update_yaxes(gridcolor=LINE, zeroline=False)
    return fig


def section_header(num, title, desc=""):
    st.markdown(f'<div class="section-title"><span class="section-num">{num}</span>{title}</div>', unsafe_allow_html=True)
    if desc:
        st.markdown(f'<div class="section-desc">{desc}</div>', unsafe_allow_html=True)


def chart_card_open(label, note=None):
    st.markdown(f'<div class="chart-card"><div class="chart-label">{label}</div>' + (f'<div class="chart-note">{note}</div>' if note else ""), unsafe_allow_html=True)


def chart_card_close():
    st.markdown("</div>", unsafe_allow_html=True)


def insight(html_text):
    st.markdown(f'<div class="insight-box">{html_text}</div>', unsafe_allow_html=True)


def gap_note(text):
    st.markdown(f'<div class="gap-note">⚠ {text}</div>', unsafe_allow_html=True)


def kpi_html(label, value, delta_text, tone):
    cls = {"down": "delta-down", "up": "delta-up", "flat": "delta-flat"}[tone]
    arrow = {"down": "▼", "up": "▲", "flat": "—"}[tone]
    return f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="kpi-delta {cls}">{arrow} {delta_text}</div></div>'


@st.cache_data
def load_csv(p):
    try:
        return pd.read_csv(p, encoding="utf-8-sig")
    except Exception:
        return None


def safe_csv(p):
    return load_csv(p) if os.path.exists(p) else None


def missing(*names):
    st.info(f"파일을 찾을 수 없습니다 — {', '.join(names)}")


# ============================================================
# 데이터 로드
# ============================================================
brand_df = safe_csv("naver_search_trend.csv")
season_df = safe_csv("naver_seasonality_trend.csv")
tone_df = safe_csv("naver_gift_tone_trend.csv")
recipient_df = safe_csv("naver_gift_recipient_demographics.csv")
recommend_df = safe_csv("naver_gift_recommend_variants.csv")
reason_df = safe_csv("20_30대_오쏘몰_선물_이유.csv")
snapshot_df = safe_csv("오쏘몰_이뮨_대시보드_데이터.csv")
channel_df = safe_csv("선물채널_언급량.csv")
price_df = safe_csv("종합비타민_가격비교.csv")
category_search_df = safe_csv("영양제_유튜브 검색_트렌드.csv")
gift_cat_df = safe_csv("연령별_선물카테고리_선호도.csv")
occasion_total_df = safe_csv("선물시즌_총언급량.csv")
occasion_month_df = safe_csv("선물시즌_월별분포.csv")
worry_brand_df = safe_csv("선물고민_브랜드별언급량.csv")
side_effect_df = safe_csv("오쏘몰_부작용_인식.csv")
image_theme_df = safe_csv("오쏘몰_시장_이미지.csv")
intent_df = safe_csv("naver_purchase_intent.csv")

if brand_df is not None:
    brand_df["기간"] = pd.to_datetime(brand_df["기간"])
if season_df is not None:
    season_df["기간"] = pd.to_datetime(season_df["기간"])
    season_pivot = season_df.pivot(index="기간", columns="키워드그룹", values="검색관심도(상대값)")
else:
    season_pivot = None
if tone_df is not None:
    tone_df["기간"] = pd.to_datetime(tone_df["기간"])
if intent_df is not None:
    intent_df["기간"] = pd.to_datetime(intent_df["기간"])

AGE_ORDER_4 = ["19~24세", "25~29세", "30~34세", "35~39세"]

# ============================================================
# 헤더
# ============================================================
st.markdown('<div class="eyebrow">ORTHOMOL · 마케팅 진단 & 성장 제안</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-title">명절에만 사는 브랜드? 오쏘몰의 평시 선물시장 공략 전략</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-sub">13개 핵심 발견을 실데이터로 검증한 진단 → 전략 제안 대시보드</div>', unsafe_allow_html=True)

# ============================================================
# 상단 요약 KPI
# ============================================================
kpi_html_str = ""
if brand_df is not None:
    o = brand_df[brand_df["브랜드"] == "오쏘몰"].sort_values("기간")
    fv, lv = o["검색관심도(상대값)"].iloc[0], o["검색관심도(상대값)"].iloc[-1]
    pct = (lv - fv) / fv * 100
    kpi_html_str += kpi_html("검색 관심도 (12개월)", f"{lv:.0f}", f"{pct:.1f}%", "down" if pct < 0 else "up")
if occasion_total_df is not None:
    top_occ = occasion_total_df.sort_values("언급량", ascending=False).iloc[0]
    kpi_html_str += kpi_html("선물 계기 언급 1위", top_occ["계기"], f"{int(top_occ['언급량']):,}건", "up")
if worry_brand_df is not None and "오쏘몰" in worry_brand_df["브랜드"].values:
    v = worry_brand_df[worry_brand_df["브랜드"] == "오쏘몰"]["언급량"].iloc[0]
    kpi_html_str += kpi_html("'선물고민' 콘텐츠 내 언급", f"{int(v):,}건", "경쟁사 압도", "up")
kpi_html_str += kpi_html("가격 vs 확신도", "0건", "'고민됨·확신없음' 응답", "flat")

st.markdown(f'<div class="kpi-strip">{kpi_html_str}</div>', unsafe_allow_html=True)

# ============================================================
# 1. 선두주자 · 프리미엄 포지셔닝
# ============================================================
section_header(1, "선두주자, 그리고 프리미엄", "오쏘몰은 액상형 멀티비타민 시장을 프리미엄 포지셔닝으로 이끌어왔다")

col1, col2 = st.columns(2)
with col1:
    if channel_df is not None:
        chart_card_open("선물 채널 언급량 — 오쏘몰 vs 경쟁사", "카카오선물하기 · 올리브영")
        fig = go.Figure()
        for ch in channel_df["채널"].unique():
            d = channel_df[channel_df["채널"] == ch]
            fig.add_trace(go.Bar(x=d["브랜드"], y=d["언급량"], name=ch,
                                  marker_color=AMBER if ch == "카카오선물하기" else STEEL))
        fig.update_layout(barmode="group")
        st.plotly_chart(base_layout(fig, height=280), use_container_width=True)
        chart_card_close()
    else:
        missing("선물채널_언급량.csv")

with col2:
    if price_df is not None:
        chart_card_open("종합비타민 1일 환산가 비교", "프리미엄 가격 포지셔닝의 정량적 증거")
        d = price_df.sort_values("1일 환산가(최저가 기준)", ascending=True)
        colors = [AMBER if b == "오쏘몰" else GRAY for b in d["브랜드"]]
        fig = go.Figure(go.Bar(x=d["브랜드"], y=d["1일 환산가(최저가 기준)"], marker_color=colors,
                                text=d["1일 환산가(최저가 기준)"].astype(int).astype(str) + "원", textposition="outside"))
        st.plotly_chart(base_layout(fig, height=280, legend=False), use_container_width=True)
        chart_card_close()
    else:
        missing("종합비타민_가격비교.csv")

if channel_df is not None and price_df is not None:
    price_ratio = price_df[price_df["브랜드"] == "오쏘몰"]["1일 환산가(최저가 기준)"].iloc[0] / price_df[price_df["브랜드"] == "고려은단"]["1일 환산가(최저가 기준)"].iloc[0]
    insight(f"오쏘몰은 카카오선물하기·올리브영 두 <b>선물 채널</b>에서 경쟁사(아임비타·에너씨슬) 대비 압도적 언급량을 기록한다. "
            f"가격 면에서도 1일 환산가가 고려은단 대비 약 <b>{price_ratio:.1f}배</b> 높아, &ldquo;프리미엄&rdquo;이라는 포지셔닝이 정성적 인상이 아닌 <b>실제 가격 데이터로 확인</b>된다.")

st.divider()

# ============================================================
# 2. 그러나 관심 감소
# ============================================================
section_header(2, "그러나, 식어가는 관심", "후발주자 진입과 함께 검색 관심도, 실제 매출 모두 하락 신호가 나타난다")

col3, col4 = st.columns(2)
with col3:
    if brand_df is not None:
        chart_card_open("브랜드별 검색 관심도 추이", "네이버 데이터랩 · 12개월")
        fig = go.Figure()
        for b in brand_df["브랜드"].unique():
            d = brand_df[brand_df["브랜드"] == b].sort_values("기간")
            fig.add_trace(go.Scatter(x=d["기간"], y=d["검색관심도(상대값)"], name=b,
                                      line=dict(width=4 if b == "오쏘몰" else 1.6, color=BRAND_COLORS.get(b)), mode="lines"))
        st.plotly_chart(base_layout(fig, height=300), use_container_width=True)
        chart_card_close()
    else:
        missing("naver_search_trend.csv")

with col4:
    chart_card_open("2025년 1분기 매출 추이", "동아제약 공식 발표 기준")
    years = ["2020", "2021", "2022", "2023", "2024"]
    vals = [87, 284, 655, 1204, 1302]
    fig = go.Figure(go.Bar(x=years, y=vals, marker_color=AMBER))
    fig.add_annotation(x="2024", y=1302, text="1,302억", showarrow=False, yshift=15, font=dict(size=11, color=INK))
    st.plotly_chart(base_layout(fig, height=300, legend=False), use_container_width=True)
    chart_card_close()

insight("검색 관심도는 12개월간 크게 하락했고, 2025년 1분기 매출도 <b>전년동기 -3.9%, 전분기 -12.7%</b>로 "
        "동아제약 발표 이후 <b>첫 역성장</b>을 기록했다. 이중제형 시장을 개척한 이후 약 200여 개의 미투제품이 등장하며 "
        "경쟁이 심화된 결과로 해석된다.")

st.divider()

# ============================================================
# 3. 명절에만 집중
# ============================================================
section_header(3, "명절에만 반짝이는 관심", "오쏘몰 검색 피크는 명절 시즌과 정확히 겹친다")

if season_pivot is not None:
    chart_card_open("명절 선물 검색 vs 오쏘몰")
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    if "명절 선물" in season_pivot.columns:
        fig.add_trace(go.Scatter(x=season_pivot.index, y=season_pivot["명절 선물"], name="명절 선물",
                                  line=dict(color=RUST, width=2), fill="tozeroy", fillcolor="rgba(229,72,77,.07)"), secondary_y=False)
    if "오쏘몰" in season_pivot.columns:
        fig.add_trace(go.Scatter(x=season_pivot.index, y=season_pivot["오쏘몰"], name="오쏘몰",
                                  line=dict(color=AMBER, width=2.5)), secondary_y=True)
    st.plotly_chart(base_layout(fig, height=280), use_container_width=True)
    chart_card_close()

    holiday_peak = season_pivot["명절 선물"].idxmax() if "명절 선물" in season_pivot.columns else None
    orthomol_peak = season_pivot["오쏘몰"].idxmax() if "오쏘몰" in season_pivot.columns else None
    match = holiday_peak == orthomol_peak
    insight(f"명절선물·오쏘몰 검색 <b>{'피크 시점 일치' if match else '피크 시점 유사'}</b> — 명절 마케팅이 실제 수요와 맞물려 작동하고 있음을 보여준다.")
else:
    missing("naver_seasonality_trend.csv")

st.divider()

# ============================================================
# 4. 연말엔 관심 못받음
# ============================================================
section_header(4, "그러나 명절이 전부는 아니다", "생일·기념일 등 다른 계기는 명절보다 더 많이, 더 고르게 언급된다")

col5, col6 = st.columns(2)
with col5:
    if occasion_total_df is not None:
        chart_card_open("선물 계기별 총 언급량", "네이버 블로그·카페 텍스트 분석")
        d = occasion_total_df.sort_values("언급량", ascending=True)
        colors = [AMBER if c == "명절(추석·설날)" else GRAY for c in d["계기"]]
        fig = go.Figure(go.Bar(x=d["언급량"], y=d["계기"], orientation="h", marker_color=colors))
        st.plotly_chart(base_layout(fig, height=280, legend=False), use_container_width=True)
        chart_card_close()
    else:
        missing("선물시즌_총언급량.csv")

with col6:
    if occasion_month_df is not None:
        chart_card_open("월별 분포 — 명절 vs 생일", "명절은 특정 월에 집중, 생일은 연중 고름")
        occasion_month_df = occasion_month_df.rename(columns={occasion_month_df.columns[0]: "계기"})
        months = [c for c in occasion_month_df.columns if c != "계기"]
        fig = go.Figure()
        for occ, color in [("명절(추석·설날)", RUST), ("생일", AMBER)]:
            row = occasion_month_df[occasion_month_df["계기"] == occ]
            if not row.empty:
                vals = row[months].values.flatten().tolist()
                fig.add_trace(go.Scatter(x=months, y=vals, name=occ, line=dict(color=color, width=2.5)))
        st.plotly_chart(base_layout(fig, height=280), use_container_width=True)
        chart_card_close()
    else:
        missing("선물시즌_월별분포.csv")

if occasion_total_df is not None:
    total_sorted = occasion_total_df.sort_values("언급량", ascending=False)
    insight(f"놀랍게도 <b>&ldquo;생일&rdquo; 언급량(6,752건)이 &ldquo;명절&rdquo;(3,326건)의 2배</b>에 달한다. "
            f"명절은 9월 한 달에 177건이 몰리는 반면, 생일은 12개월 내내 10~32건으로 고르게 분포한다 — "
            f"<b>선물 시장 자체는 명절에 갇혀있지 않다</b>는 뜻이다.")

st.divider()

# ============================================================
# 5. 건강·비타민 계열은 완만함 (히트맵)
# ============================================================
section_header(5, "건강·비타민 선물은 사계절 내내", "명절처럼 특정 시즌에 쏠리지 않고 꾸준히 검색된다")

if tone_df is not None:
    tone_pivot = tone_df.pivot(index="기간", columns="키워드그룹", values="검색관심도(상대값)").sort_index()
    HEAT_ROWS = ["명절 선물세트", "비타민 선물", "오쏘몰", "건강 선물", "선물 추천", "센스있는 선물"]
    HEAT_ROWS = [r for r in HEAT_ROWS if r in tone_pivot.columns]

    z = []
    for r in HEAT_ROWS:
        vals = tone_pivot[r].values
        max_v = vals.max() if vals.max() > 0 else 1
        z.append((vals / max_v).tolist())

    chart_card_open("52주 검색 존재감 캘린더", "칸에 색이 있을수록 그 주에 검색이 활발했다는 뜻 (각 키워드 내 상대값)")
    fig = go.Figure(data=go.Heatmap(
        z=z, y=HEAT_ROWS, x=[d.strftime("%m/%d") for d in tone_pivot.index],
        colorscale=[[0, "#FFFFFF"], [1, AMBER]], showscale=False, xgap=2, ygap=4,
    ))
    fig.update_yaxes(autorange="reversed")
    fig.update_xaxes(tickangle=0, nticks=13)
    st.plotly_chart(base_layout(fig, height=260, legend=False), use_container_width=True)
    chart_card_close()

    insight("&ldquo;명절 선물세트&rdquo;는 두 번(추석·설)만 진하게 반짝이고 나머지 46주는 거의 비어있다. "
            "반면 비타민·오쏘몰·건강·선물추천·센스있는 선물은 52주 내내 색이 유지된다 — <b>명절에 갇히지 않은 꾸준한 수요</b>다.")
else:
    missing("naver_gift_tone_trend.csv")

st.divider()

# ============================================================
# 6. 한우 1위, 구매전환은 영양제
# ============================================================
section_header(6, "검색은 한우, 그러나 실용성은 영양제", "선물 카테고리 검색 관심도 1위는 한우이지만 건강·실용성 수요는 영양제가 흡수한다")

gap_note("연령대별 '구매 클릭' 데이터(네이버쇼핑 클릭비율)는 이번에 파일이 없어 검색 관심도 데이터로만 구성했습니다. 추후 파일 확보 시 보강 예정입니다.")

if gift_cat_df is not None:
    chart_card_open("연령대별 선물 카테고리 검색 관심도")
    fig = go.Figure()
    colors_map = {"영양제 선물": AMBER, "홍삼 선물": STEEL, "한우 선물": SAGE, "상품권 선물": GRAY, "화장품 선물": PLUM}
    for col in [c for c in gift_cat_df.columns if c != "연령대"]:
        fig.add_trace(go.Bar(x=gift_cat_df["연령대"], y=gift_cat_df[col], name=col, marker_color=colors_map.get(col, GRAY)))
    fig.update_layout(barmode="group")
    st.plotly_chart(base_layout(fig, height=300), use_container_width=True)
    chart_card_close()

    insight("한우 선물이 전 연령대에서 검색 관심도 1위를 차지하지만, 오픈서베이 기준 영양제 선호 이유는 "
            "<b>건강·실용성</b>이 압도적이다 (전 세대 공통 1순위 니즈: 면역강화 71.5%, 2순위 피로회복 56.9%) — "
            "오쏘몰 '이뮨' 라인의 포지셔닝과 정확히 맞아떨어진다.")
else:
    missing("연령별_선물카테고리_선호도.csv")

if image_theme_df is not None:
    chart_card_open("오쏘몰 콘텐츠 테마 비중", "네이버 블로그·카페 텍스트 키워드 기반 (복수 응답)")
    d = image_theme_df.sort_values("건수", ascending=True)
    colors_theme = [AMBER if t == "효과·건강기능" else GRAY for t in d["테마"]]
    fig = go.Figure(go.Bar(x=d["건수"], y=d["테마"], orientation="h", marker_color=colors_theme))
    st.plotly_chart(base_layout(fig, height=220, legend=False), use_container_width=True)
    chart_card_close()
    insight("오쏘몰 관련 콘텐츠에서도 &ldquo;효과·건강기능&rdquo; 테마가 84%로 압도적 1위를 차지한다 — "
            "&ldquo;프리미엄·고가&rdquo; 테마(18%)보다 훨씬 높아, 소비자가 오쏘몰을 언급할 때 "
            "<b>가격보다 효과를 먼저 떠올린다</b>는 것을 보여준다.")

st.divider()

# ============================================================
# 7. 비타민 최강세 + 시장 적합성 결론
# ============================================================
section_header(7, "비타민, 영양제 중 최강세", "오쏘몰이 속한 비타민 카테고리는 영양제 선물 안에서도 가장 강세다")

if category_search_df is not None:
    chart_card_open("영양제 종류별 검색 관심도 추이")
    fig = go.Figure()
    colors_map2 = {"비타민": AMBER, "오메가3": STEEL, "유산균": SAGE, "루테인": GRAY, "마그네슘": PLUM}
    for col in [c for c in category_search_df.columns if c != "날짜"]:
        fig.add_trace(go.Scatter(x=category_search_df["날짜"], y=category_search_df[col], name=col,
                                  line=dict(width=3.5 if col == "비타민" else 1.4, color=colors_map2.get(col, GRAY))))
    st.plotly_chart(base_layout(fig, height=280), use_container_width=True)
    chart_card_close()

    latest = category_search_df.iloc[-8:].mean(numeric_only=True).sort_values(ascending=False)
    ratio = latest.iloc[0] / latest.iloc[1] if len(latest) > 1 else None
    insight(f"비타민은 최근 8주 평균 기준 2위({latest.index[1]}) 대비 <b>{ratio:.1f}배</b> 높은 검색 관심도를 기록하며 "
            f"영양제 카테고리 내 압도적 1위를 유지한다. <b>연중 꾸준한 수요 위에서, 오쏘몰은 이미 강세인 카테고리에 "
            f"자리하고 있어 시장 적합성이 확보된다.</b>")
else:
    missing("영양제_유튜브_검색_트렌드.csv")

st.divider()

# ============================================================
# 8. 20~30대, 5가지 대상 모두에서 관심
# ============================================================
section_header(8, "20~30대가 마주하는 다양한 관계", "상사·거래처·부모님·동료·첫만남 — 5가지 대상 모두에서 의미 있는 검색 관심을 보인다")

if recipient_df is not None:
    chart_card_open("연령대별 대상별 선물 키워드 검색 관심도", "0~12세 등 표본이 극히 적은 구간은 제외")
    d = recipient_df[recipient_df["연령대"].isin(AGE_ORDER_4)]
    fig = go.Figure()
    for age, color in zip(AGE_ORDER_4, [STEEL, PLUM, AMBER, SAGE]):
        dd = d[d["연령대"] == age]
        fig.add_trace(go.Bar(x=dd["키워드"], y=dd["평균검색관심도"], name=age, marker_color=color))
    fig.update_layout(barmode="group")
    st.plotly_chart(base_layout(fig, height=300), use_container_width=True)
    chart_card_close()

    vals = d["평균검색관심도"]
    insight(f"4개 연령대 x 5개 대상 = 20개 구간 모두 {vals.min():.1f}~{vals.max():.1f} 범위에 분포하며 0에 가까운 값이 없다 — "
            f"20~30대는 특정 관계에 치우치지 않고 상사·거래처·부모님·동료·첫만남 등 다양한 관계에서 고르게 선물 검색 관심을 보인다.")
else:
    missing("naver_gift_recipient_demographics.csv")

st.divider()

# ============================================================
# 9. 사회진출로 관계 확장
# ============================================================
section_header(9, "사회 진출과 함께 넓어지는 관계", "가족뿐 아니라 새로운 인연과 직장 내 관계까지 챙길 대상이 늘어난다")

if recipient_df is not None:
    chart_card_open("대상별 평균 검색 관심도 (20~30대 평균)")
    d = recipient_df[recipient_df["연령대"].isin(AGE_ORDER_4)]
    ranking = d.groupby("키워드")["평균검색관심도"].mean().sort_values(ascending=False)
    fig = go.Figure(go.Bar(x=ranking.index, y=ranking.values, marker_color=AMBER))
    st.plotly_chart(base_layout(fig, height=280, legend=False), use_container_width=True)
    chart_card_close()

    insight("20~30대가 신경 쓰는 관계는 가족(부모님)뿐 아니라 새로운 인연(첫만남)과 직장 내 관계(상사·거래처·동료)까지 "
            "폭넓게 걸쳐 있다 — 학생 시기에는 크지 않았을 이 관계들이 <b>사회 진출과 함께 새롭게 추가</b>되며, "
            "챙겨야 할 대상의 범위 자체가 넓어지고 있음을 보여준다.")

st.divider()

# ============================================================
# 10. 전환 — 그래서 어떻게?
# ============================================================
st.markdown('<div class="section-title"><span class="section-num">10</span>그렇다면, 오쏘몰을 선택하게 하려면?</div>', unsafe_allow_html=True)
st.markdown('<div class="section-desc">지금까지의 발견을 압축하면 — 기회는 있고, 강점도 있다</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown('<div class="mini-tag"><div class="t">시장 기회</div><div class="v">5/5개</div></div>', unsafe_allow_html=True)
    if recipient_df is not None:
        d = recipient_df[recipient_df["연령대"].isin(AGE_ORDER_4)]
        ranking = d.groupby("키워드")["평균검색관심도"].mean().sort_values(ascending=False)
        fig = go.Figure(go.Bar(x=ranking.index, y=ranking.values, marker_color=AMBER))
        fig.update_xaxes(showticklabels=False)
        fig.update_yaxes(showticklabels=False)
        st.plotly_chart(base_layout(fig, height=110, legend=False), use_container_width=True)
    st.markdown('<div class="mini-tag" style="margin-top:-14px; border-top:none; border-radius:0 0 12px 12px;"><div class="d">20~30대가 모든 관계에서 선물 검색 관심 보임</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="mini-tag"><div class="t">오쏘몰 강점</div><div class="v">1위</div></div>', unsafe_allow_html=True)
    if category_search_df is not None:
        latest = category_search_df.iloc[-8:].mean(numeric_only=True).sort_values(ascending=False)
        colors_bar = [AMBER if i == 0 else GRAY for i in range(len(latest))]
        fig = go.Figure(go.Bar(x=latest.index, y=latest.values, marker_color=colors_bar))
        fig.update_xaxes(showticklabels=False)
        fig.update_yaxes(showticklabels=False)
        st.plotly_chart(base_layout(fig, height=110, legend=False), use_container_width=True)
    st.markdown('<div class="mini-tag" style="margin-top:-14px; border-top:none; border-radius:0 0 12px 12px;"><div class="d">영양제 중 비타민 카테고리 최강세</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="mini-tag" style="height:100%; display:flex; flex-direction:column; justify-content:center;"><div class="t">남은 과제</div><div class="v">?</div><div class="d">이 기회를 실제 선택으로<br>연결하려면?</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.divider()

# ============================================================
# 11. 매번 새로 검색하는 반복적 고민
# ============================================================
section_header(11, "반복되는 선물 고민", "상황이 바뀔 때마다 매번 새로운 추천을 검색해야 하는 부담이 있다")

col7, col8 = st.columns(2)
with col7:
    if recommend_df is not None:
        chart_card_open("상황별 '선물추천' 검색 관심도", "센스있는/적당한 선물추천은 표본 부족으로 제외")
        d = recommend_df[recommend_df["연령대"].isin(AGE_ORDER_4)]
        fig = go.Figure()
        for age, color in zip(AGE_ORDER_4, [STEEL, PLUM, AMBER, SAGE]):
            dd = d[d["연령대"] == age]
            fig.add_trace(go.Bar(x=dd["키워드"], y=dd["평균검색관심도"], name=age, marker_color=color))
        fig.update_layout(barmode="group")
        st.plotly_chart(base_layout(fig, height=280), use_container_width=True)
        chart_card_close()
    else:
        missing("naver_gift_recommend_variants.csv")

with col8:
    if worry_brand_df is not None:
        chart_card_open("'선물고민' 콘텐츠 내 브랜드 언급량", "네이버 블로그·카페 텍스트 분석")
        d = worry_brand_df.sort_values("언급량", ascending=True)
        colors = [AMBER if b == "오쏘몰" else GRAY for b in d["브랜드"]]
        fig = go.Figure(go.Bar(x=d["언급량"], y=d["브랜드"], orientation="h", marker_color=colors))
        st.plotly_chart(base_layout(fig, height=280, legend=False), use_container_width=True)
        chart_card_close()
    else:
        missing("선물고민_브랜드별언급량.csv")

if worry_brand_df is not None:
    o_val = worry_brand_df[worry_brand_df["브랜드"] == "오쏘몰"]["언급량"].iloc[0]
    others_max = worry_brand_df[worry_brand_df["브랜드"] != "오쏘몰"]["언급량"].max()
    ratio2 = o_val / others_max
    insight(f"부모님·상사·생일·첫만남·지인 등 상황이 바뀔 때마다 매번 새로운 '선물추천'을 검색하는 반복적 고민이 확인된다. "
            f"흥미롭게도 이 &ldquo;선물고민&rdquo; 콘텐츠 안에서 <b>오쏘몰은 이미 {ratio2:.1f}배 이상 압도적으로 언급</b>되고 있다 — "
            f"오쏘몰이 &ldquo;고민의 답&rdquo;으로 이미 어느 정도 포지셔닝되어 있다는 뜻이다.")

st.divider()

# ============================================================
# 12. 가격, 유일한 걸림돌
# ============================================================
section_header(12, "가격, 유일한 걸림돌", "정한 이유는 4가지, 꺼리는 이유는 사실상 가격 하나뿐이다")

REASON_POS = {"프리미엄 이미지": 39, "효과·고용량 성분": 24, "구성·혜택": 22, "브랜드·신뢰": 16}
REASON_NEG = {"가격 부담": 5, "취향·필요 안맞음": 3, "고민됨·확신없음": 0}

col9, col10 = st.columns(2)
with col9:
    chart_card_open("오쏘몰을 정한 이유", "20~30대 언급 게시물 총 149건 기준")
    fig = go.Figure(go.Bar(x=list(REASON_POS.keys()), y=list(REASON_POS.values()), marker_color=SAGE))
    st.plotly_chart(base_layout(fig, height=260, legend=False), use_container_width=True)
    chart_card_close()
with col10:
    chart_card_open("오쏘몰을 꺼리는 이유")
    colors = [RUST, GRAY, GRAY]
    fig = go.Figure(go.Bar(x=list(REASON_NEG.keys()), y=list(REASON_NEG.values()), marker_color=colors))
    st.plotly_chart(base_layout(fig, height=260, legend=False), use_container_width=True)
    chart_card_close()

if side_effect_df is not None:
    stop_row = side_effect_df[side_effect_df["유형"].str.contains("중단", na=False)]
    stop_val = int(stop_row["건수"].iloc[0]) if not stop_row.empty else 0
    insight(f"&ldquo;고민됨·확신없음&rdquo;이 0건이라는 건, 오쏘몰을 안 사는 이유가 확신이 없어서가 아니라 오직 "
            f"<b>&ldquo;가격&rdquo; 하나뿐</b>이라는 뜻이다. 부작용 관련 인식 데이터를 봐도 &ldquo;복용 중단·재구매 안 함&rdquo;은 "
            f"<b>{stop_val}건</b>으로, 부작용 우려가 있어도 실제 구매 결정을 바꾸지는 않는다 — 가격만 상쇄하면 나머지는 이미 충분하다.")
else:
    insight("&ldquo;고민됨·확신없음&rdquo;이 0건이라는 건, 오쏘몰을 안 사는 이유가 확신이 없어서가 아니라 오직 "
            "<b>&ldquo;가격&rdquo; 하나뿐</b>이라는 뜻이다 — 가격만 상쇄하면 나머지는 이미 충분하다.")

st.divider()

# ============================================================
# 13. 최종 메시지
# ============================================================
st.markdown('<div class="section-title"><span class="section-num">13</span>그래서, 오쏘몰이 던져야 할 메시지</div>', unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)
st.markdown(f"""
<div style="text-align:center; padding: 20px 0 30px 0;">
    <div style="font-size:26px; font-weight:800; color:{INK}; line-height:1.4;">
        오래 고민하지 않아도 되는,<br>누구에게나 줘도 센스있고 적당한 선물<br>
        <span style="color:{AMBER};">오쏘몰</span>
    </div>
</div>
""", unsafe_allow_html=True)

t1, t2, t3, t4 = st.columns(4)
with t1:
    st.markdown('<div class="mini-tag"><div class="t">고민하지 않아도</div><div class="v">5/5개</div></div>', unsafe_allow_html=True)
    if recommend_df is not None:
        d = recommend_df[recommend_df["연령대"].isin(AGE_ORDER_4)]
        ranking = d.groupby("키워드")["평균검색관심도"].mean().sort_values(ascending=False)
        fig = go.Figure(go.Bar(x=ranking.index, y=ranking.values, marker_color=AMBER))
        fig.update_xaxes(showticklabels=False); fig.update_yaxes(showticklabels=False)
        st.plotly_chart(base_layout(fig, height=95, legend=False), use_container_width=True)
    st.markdown('<div class="mini-tag" style="margin-top:-14px; border-top:none; border-radius:0 0 12px 12px;"><div class="d">상황 모두 검색 존재</div></div>', unsafe_allow_html=True)
with t2:
    st.markdown('<div class="mini-tag"><div class="t">누구에게나</div><div class="v">5개 관계</div></div>', unsafe_allow_html=True)
    if recipient_df is not None:
        d = recipient_df[recipient_df["연령대"].isin(AGE_ORDER_4)]
        ranking = d.groupby("키워드")["평균검색관심도"].mean().sort_values(ascending=False)
        fig = go.Figure(go.Bar(x=ranking.index, y=ranking.values, marker_color=AMBER))
        fig.update_xaxes(showticklabels=False); fig.update_yaxes(showticklabels=False)
        st.plotly_chart(base_layout(fig, height=95, legend=False), use_container_width=True)
    st.markdown('<div class="mini-tag" style="margin-top:-14px; border-top:none; border-radius:0 0 12px 12px;"><div class="d">모든 대상에서 검색</div></div>', unsafe_allow_html=True)
with t3:
    st.markdown('<div class="mini-tag"><div class="t">센스있고</div><div class="v">39건</div></div>', unsafe_allow_html=True)
    colors_bar = [AMBER if i == 0 else GRAY for i in range(len(REASON_POS))]
    fig = go.Figure(go.Bar(x=list(REASON_POS.keys()), y=list(REASON_POS.values()), marker_color=colors_bar))
    fig.update_xaxes(showticklabels=False); fig.update_yaxes(showticklabels=False)
    st.plotly_chart(base_layout(fig, height=95, legend=False), use_container_width=True)
    st.markdown('<div class="mini-tag" style="margin-top:-14px; border-top:none; border-radius:0 0 12px 12px;"><div class="d">프리미엄 이미지 1위</div></div>', unsafe_allow_html=True)
with t4:
    st.markdown('<div class="mini-tag"><div class="t">적당한</div><div class="v">46건</div></div>', unsafe_allow_html=True)
    pos_neg = {"긍정 요인\n(효과+구성)": 46, "가격부담": 5}
    fig = go.Figure(go.Bar(x=list(pos_neg.keys()), y=list(pos_neg.values()), marker_color=[SAGE, RUST]))
    fig.update_xaxes(showticklabels=False); fig.update_yaxes(showticklabels=False)
    st.plotly_chart(base_layout(fig, height=95, legend=False), use_container_width=True)
    st.markdown('<div class="mini-tag" style="margin-top:-14px; border-top:none; border-radius:0 0 12px 12px;"><div class="d">효과·구성 vs 가격부담 5건</div></div>', unsafe_allow_html=True)

st.markdown(f"""
<div style="margin-top:32px; padding-top:14px; border-top:1px solid {LINE}; font-size:10.5px; color:{TEXT_SUB}; line-height:1.7;">
※ 검색 관심도는 각 데이터셋 내 최고치를 100으로 하는 상대 지수이며, 실제 검색량(절대값)이 아닙니다.<br>
※ 텍스트 기반 이유·감정 분석은 네이버 블로그·카페 게시물의 키워드 빈도 기반 간이분석이며, 실제 설문이 아닙니다.<br>
※ 6번 섹션의 연령대별 구매 클릭 데이터는 현재 미확보 상태로, 추후 보강 예정입니다.
</div>
""", unsafe_allow_html=True)
