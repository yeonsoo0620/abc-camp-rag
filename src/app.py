import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
import os
import json as _json

st.set_page_config(
    page_title="Yes24 IT 모바일 베스트셀러 대시보드",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem; }
    .sub-header { font-size: 1rem; color: #6c757d; margin-bottom: 1.5rem; }
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white; border-radius: 12px; padding: 1.2rem 1rem;
        text-align: center; box-shadow: 0 4px 15px rgba(102,126,234,0.3);
    }
    .metric-card h3 { margin: 0; font-size: 0.85rem; opacity: 0.9; }
    .metric-card p { margin: 0.3rem 0 0 0; font-size: 1.6rem; font-weight: 700; }
    .search-result { background: #f8f9fa; border-left: 4px solid #667eea;
        padding: 0.8rem 1rem; margin: 0.5rem 0; border-radius: 0 8px 8px 0; }
    .highlight { background: #fff3cd; padding: 0 2px; border-radius: 2px; }
    div[data-testid="stExpander"] { border: 1px solid #e0e0e0; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

DATA_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "yes24_it_mobile_bestsellers.csv")
CHROMA_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "chroma_db")
COLLECTION_NAME = "yes24_books"
EMBED_MODEL = "klue/bert-base"


@st.cache_data
def load_data():
    df = pd.read_csv(DATA_PATH, encoding="utf-8-sig")
    for col in ["판매가", "정가", "판매지수"]:
        df[col] = (
            df[col].astype(str).str.replace(",", "", regex=False).str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["할인율_num"] = pd.to_numeric(
        df["할인율"].str.replace("%", "", regex=False).str.strip(), errors="coerce"
    )
    df["출판일_연도"] = pd.to_numeric(
        df["출판일"].str.extract(r"(\d{4})", expand=False), errors="coerce"
    )
    df["출판일_월"] = df["출판일"].str.extract(r"(\d{2})월", expand=False)
    df["출판일_월"] = pd.to_numeric(df["출판일_월"], errors="coerce")
    df["가격대"] = pd.cut(
        df["판매가"],
        bins=[0, 15000, 20000, 25000, 30000, 50000, 100000],
        labels=["~1.5만", "1.5~2만", "2~2.5만", "2.5~3만", "3~5만", "5만~"],
    )
    df["정가_대비_할인"] = ((1 - df["판매가"] / df["정가"]) * 100).round(1)
    return df


df = load_data()

# ── Sidebar ──
with st.sidebar:
    st.markdown("## 🎛️ 필터")
    st.markdown("---")

    price_range = st.slider(
        "💰 판매가 범위 (원)",
        int(df["판매가"].min() or 0),
        int(df["판매가"].max() or 100000),
        (int(df["판매가"].min() or 0), int(df["판매가"].max() or 100000)),
        step=1000,
    )

    publishers = st.multiselect(
        "🏢 출판사",
        sorted(df["출판사"].dropna().unique().tolist()),
    )

    year_options = sorted(df["출판일_연도"].dropna().astype(int).unique().tolist())
    years = st.multiselect("📅 출판 연도", year_options)

    min_rank = st.number_input("🏅 최소 순위", 1, int(df["순위"].max() or 1000), 1)

    discount_range = st.slider(
        "🏷️ 할인율 범위 (%)",
        0, 100, (0, 100),
    )

    st.markdown("---")
    st.caption("데이터: yes24_it_mobile_bestsellers.csv")

    st.markdown("---")
    st.markdown("## 🤖 챗봇 설정")
    groq_api_key = st.text_input("Groq API Key", type="password", help="https://console.groq.com 에서 발급")
    groq_model = st.selectbox(
        "모델",
        ["llama-3.3-70b-versatile", "llama-3.1-8b-instant", "gemma2-9b-it"],
        index=0,
    )

filtered = df[
    (df["판매가"] >= price_range[0])
    & (df["판매가"] <= price_range[1])
    & (df["순위"] >= min_rank)
    & (df["할인율_num"].fillna(0) >= discount_range[0])
    & (df["할인율_num"].fillna(0) <= discount_range[1])
]
if publishers:
    filtered = filtered[filtered["출판사"].isin(publishers)]
if years:
    filtered = filtered[filtered["출판일_연도"].isin(years)]

# ── Header ──
st.markdown('<p class="main-header">📚 Yes24 IT 모바일 베스트셀러 대시보드</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">탐색적 데이터 분석 (EDA) & 키워드 검색</p>', unsafe_allow_html=True)

# ── KPI Cards ──
k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    st.markdown(f'<div class="metric-card"><h3>총 도서 수</h3><p>{len(filtered):,}</p></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="metric-card"><h3>평균 판매가</h3><p>{filtered["판매가"].mean():,.0f}원</p></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="metric-card"><h3>평균 판매지수</h3><p>{filtered["판매지수"].mean():,.0f}</p></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="metric-card"><h3>출판사 수</h3><p>{filtered["출판사"].nunique()}개</p></div>', unsafe_allow_html=True)
with k5:
    st.markdown(f'<div class="metric-card"><h3>평균 할인율</h3><p>{filtered["할인율_num"].mean():.1f}%</p></div>', unsafe_allow_html=True)
with k6:
    st.markdown(f'<div class="metric-card"><h3>최고 판매지수</h3><p>{filtered["판매지수"].max():,.0f}</p></div>', unsafe_allow_html=True)

st.markdown("---")

# ── Tabs ──
tab_overview, tab_price, tab_publisher, tab_discount, tab_sales, tab_trend, tab_search, tab_chat = st.tabs(
    ["📊 개요", "💰 가격 분석", "🏢 출판사 분석", "🏷️ 할인율 분석", "📈 판매지수 분석", "📅 트렌드 분석", "🔍 키워드 검색", "🤖 도서 추천 챗봇"]
)

# ═══════════════════════════════════════
# 1. 개요
# ═══════════════════════════════════════
with tab_overview:
    col_l, col_r = st.columns([1, 1])

    with col_l:
        st.subheader("연도별 출판 도서 수")
        year_counts = filtered["출판일_연도"].dropna().astype(int).value_counts().sort_index()
        fig_year = px.bar(
            x=year_counts.index.astype(str),
            y=year_counts.values,
            labels={"x": "출판 연도", "y": "도서 수"},
            text=year_counts.values,
            color=year_counts.values,
            color_continuous_scale="Viridis",
        )
        fig_year.update_traces(textposition="outside", showlegend=False)
        fig_year.update_layout(margin=dict(t=30, b=30), height=350)
        st.plotly_chart(fig_year, use_container_width=True)

    with col_r:
        st.subheader("가격대별 분포")
        price_dist = filtered["가격대"].value_counts().reindex(
            ["~1.5만", "1.5~2만", "2~2.5만", "2.5~3만", "3~5만", "5만~"]
        ).dropna()
        fig_pie = px.pie(
            values=price_dist.values,
            names=price_dist.index,
            hole=0.45,
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig_pie.update_traces(textposition="inside", textinfo="percent+label")
        fig_pie.update_layout(margin=dict(t=30, b=30), height=350, showlegend=False)
        st.plotly_chart(fig_pie, use_container_width=True)

    st.subheader("📋 데이터 미리보기")
    display_cols = ["순위", "도서명", "저자", "출판사", "출판일", "판매가", "정가", "할인율", "판매지수"]
    st.dataframe(
        filtered[display_cols].reset_index(drop=True),
        use_container_width=True,
        height=400,
    )

# ═══════════════════════════════════════
# 2. 가격 분석
# ═══════════════════════════════════════
with tab_price:
    col_a, col_b = st.columns([1, 1])

    with col_a:
        st.subheader("판매가 분포 (히스토그램)")
        fig_hist = px.histogram(
            filtered, x="판매가", nbins=40, marginal="box",
            color_discrete_sequence=["#667eea"],
        )
        fig_hist.update_layout(margin=dict(t=30), height=380)
        st.plotly_chart(fig_hist, use_container_width=True)

    with col_b:
        st.subheader("출판사별 평균 판매가 (상위 15)")
        pub_avg = filtered.groupby("출판사")["판매가"].mean().nlargest(15).sort_values()
        fig_pub_price = px.bar(
            x=pub_avg.values, y=pub_avg.index, orientation="h",
            labels={"x": "평균 판매가 (원)", "y": ""},
            color=pub_avg.values, color_continuous_scale="Blues",
        )
        fig_pub_price.update_layout(margin=dict(t=30), height=380, showlegend=False)
        st.plotly_chart(fig_pub_price, use_container_width=True)

    st.subheader("정가 vs 판매가 산점도")
    scatter_data = filtered.dropna(subset=["정가", "판매가"])
    fig_scatter = px.scatter(
        scatter_data, x="정가", y="판매가",
        hover_data=["도서명", "출판사", "순위"],
        opacity=0.5, color_discrete_sequence=["#667eea"],
    )
    max_val = max(scatter_data["정가"].max(), scatter_data["판매가"].max())
    fig_scatter.add_trace(go.Scatter(
        x=[0, max_val], y=[0, max_val], mode="lines",
        name="100% line", line=dict(dash="dash", color="red", width=1),
    ))
    fig_scatter.update_layout(margin=dict(t=30), height=400)
    st.plotly_chart(fig_scatter, use_container_width=True)

# ═══════════════════════════════════════
# 3. 출판사 분석
# ═══════════════════════════════════════
with tab_publisher:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("출판사별 도서 수 (상위 20)")
        pub_counts = filtered["출판사"].value_counts().nlargest(20)
        fig_pub = px.bar(
            x=pub_counts.index, y=pub_counts.values,
            labels={"x": "출판사", "y": "도서 수"},
            text=pub_counts.values,
            color=pub_counts.values, color_continuous_scale="Teal",
        )
        fig_pub.update_traces(textposition="outside")
        fig_pub.update_layout(margin=dict(t=30), height=450, xaxis_tickangle=-45, showlegend=False)
        st.plotly_chart(fig_pub, use_container_width=True)

    with col2:
        st.subheader("출판사별 평균 판매지수 (상위 15)")
        pub_sales = filtered.groupby("출판사")["판매지수"].mean().nlargest(15).sort_values()
        fig_pub_sales = px.bar(
            x=pub_sales.values, y=pub_sales.index, orientation="h",
            labels={"x": "평균 판매지수", "y": ""},
            color=pub_sales.values, color_continuous_scale="Oranges",
        )
        fig_pub_sales.update_layout(margin=dict(t=30), height=450, showlegend=False)
        st.plotly_chart(fig_pub_sales, use_container_width=True)

    st.subheader("출판사 종합 통계")
    pub_stats = (
        filtered.groupby("출판사")
        .agg(
            도서수=("도서명", "count"),
            평균판매가=("판매가", "mean"),
            평균판매지수=("판매지수", "mean"),
            총판매지수=("판매지수", "sum"),
            평균할인율=("할인율_num", "mean"),
        )
        .sort_values("총판매지수", ascending=False)
        .head(20)
    )
    pub_stats["평균판매가"] = pub_stats["평균판매가"].round(0).astype(int)
    pub_stats["평균판매지수"] = pub_stats["평균판매지수"].round(0).astype(int)
    pub_stats["총판매지수"] = pub_stats["총판매지수"].round(0).astype(int)
    pub_stats["평균할인율"] = pub_stats["평균할인율"].round(1)
    st.dataframe(pub_stats, use_container_width=True)

# ═══════════════════════════════════════
# 4. 할인율 분석
# ═══════════════════════════════════════
with tab_discount:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("할인율 분포")
        disc_counts = filtered["할인율"].dropna().value_counts().sort_index()
        fig_disc = px.bar(
            x=disc_counts.index, y=disc_counts.values,
            labels={"x": "할인율", "y": "도서 수"},
            text=disc_counts.values,
            color_discrete_sequence=["#FFC107"],
        )
        fig_disc.update_traces(textposition="outside")
        fig_disc.update_layout(margin=dict(t=30), height=350)
        st.plotly_chart(fig_disc, use_container_width=True)

    with col2:
        st.subheader("출판사별 평균 할인율 (상위 15)")
        pub_disc = filtered.groupby("출판사")["할인율_num"].mean().nlargest(15).sort_values()
        fig_pd = px.bar(
            x=pub_disc.values, y=pub_disc.index, orientation="h",
            labels={"x": "평균 할인율 (%)", "y": ""},
            color=pub_disc.values, color_continuous_scale="YlOrRd",
        )
        fig_pd.update_layout(margin=dict(t=30), height=350, showlegend=False)
        st.plotly_chart(fig_pd, use_container_width=True)

    st.subheader("할인율 vs 판매지수 관계")
    disc_sales = filtered.dropna(subset=["할인율_num", "판매지수"])
    fig_disc_scatter = px.scatter(
        disc_sales, x="할인율_num", y="판매지수",
        hover_data=["도서명", "출판사"],
        opacity=0.5, color_discrete_sequence=["#e74c3c"],
    )
    fig_disc_scatter.update_layout(margin=dict(t=30), height=350)
    st.plotly_chart(fig_disc_scatter, use_container_width=True)

# ═══════════════════════════════════════
# 5. 판매지수 분석
# ═══════════════════════════════════════
with tab_sales:
    st.subheader("🏆 판매지수 Top 20 도서")
    top_sales = filtered.nlargest(20, "판매지수").sort_values("판매지수")
    fig_top = px.bar(
        top_sales, x="판매지수", y="도서명", orientation="h",
        hover_data=["저자", "출판사", "판매가"],
        labels={"도서명": ""},
        color=top_sales["판매지수"], color_continuous_scale="Viridis",
    )
    fig_top.update_layout(margin=dict(t=30), height=550, showlegend=False)
    st.plotly_chart(fig_top, use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("출판사별 총 판매지수 (상위 15)")
        pub_total = filtered.groupby("출판사")["판매지수"].sum().nlargest(15).sort_values()
        fig_pt = px.bar(
            x=pub_total.values, y=pub_total.index, orientation="h",
            labels={"x": "총 판매지수", "y": ""},
            color=pub_total.values, color_continuous_scale="Purples",
        )
        fig_pt.update_layout(margin=dict(t=30), height=400, showlegend=False)
        st.plotly_chart(fig_pt, use_container_width=True)

    with col2:
        st.subheader("판매가 vs 판매지수")
        price_sales = filtered.dropna(subset=["판매가", "판매지수"])
        fig_ps = px.scatter(
            price_sales, x="판매가", y="판매지수",
            hover_data=["도서명", "출판사"],
            opacity=0.5, color_discrete_sequence=["#2ecc71"],
            trendline="ols",
        )
        fig_ps.update_layout(margin=dict(t=30), height=400)
        st.plotly_chart(fig_ps, use_container_width=True)

# ═══════════════════════════════════════
# 6. 트렌드 분석
# ═══════════════════════════════════════
with tab_trend:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("연도별 평균 판매가 추이")
        yearly_price = filtered.groupby("출판일_연도")["판매가"].mean().dropna()
        fig_yp = px.line(
            x=yearly_price.index.astype(int), y=yearly_price.values,
            labels={"x": "연도", "y": "평균 판매가 (원)"},
            markers=True,
        )
        fig_yp.update_traces(line_color="#667eea", line_width=3, marker_size=8)
        fig_yp.update_layout(margin=dict(t=30), height=350)
        st.plotly_chart(fig_yp, use_container_width=True)

    with col2:
        st.subheader("연도별 평균 판매지수 추이")
        yearly_sales = filtered.groupby("출판일_연도")["판매지수"].mean().dropna()
        fig_ys = px.line(
            x=yearly_sales.index.astype(int), y=yearly_sales.values,
            labels={"x": "연도", "y": "평균 판매지수"},
            markers=True,
        )
        fig_ys.update_traces(line_color="#e74c3c", line_width=3, marker_size=8)
        fig_ys.update_layout(margin=dict(t=30), height=350)
        st.plotly_chart(fig_ys, use_container_width=True)

    st.subheader("연도별 출판사 점유율 변화")
    top5_publishers = filtered["출판사"].value_counts().nlargest(5).index.tolist()
    trend_data = (
        filtered[filtered["출판사"].isin(top5_publishers)]
        .groupby(["출판일_연도", "출판사"])
        .size()
        .reset_index(name="도서 수")
    )
    fig_trend = px.area(
        trend_data, x="출판일_연도", y="도서 수", color="출판사",
        labels={"출판일_연도": "연도"},
    )
    fig_trend.update_layout(margin=dict(t=30), height=400)
    st.plotly_chart(fig_trend, use_container_width=True)

# ═══════════════════════════════════════
# 7. 키워드 검색
# ═══════════════════════════════════════
with tab_search:
    st.subheader("🔍 도서 검색")
    st.caption("제목, 저자, 출판사에서 키워드로 검색할 수 있습니다.")

    col_input, col_target = st.columns([3, 1])
    with col_input:
        keyword = st.text_input("검색 키워드 입력", placeholder="예: 바이브 코딩, 파이썬, 한빛미디어...")
    with col_target:
        search_target = st.radio("검색 대상", ["전체", "도서명", "저자", "출판사"], horizontal=True)

    if keyword:
        kw = keyword.lower()

        if search_target == "전체":
            mask = (
                filtered["도서명"].str.lower().str.contains(kw, na=False)
                | filtered["저자"].str.lower().str.contains(kw, na=False)
                | filtered["출판사"].str.lower().str.contains(kw, na=False)
            )
        elif search_target == "도서명":
            mask = filtered["도서명"].str.lower().str.contains(kw, na=False)
        elif search_target == "저자":
            mask = filtered["저자"].str.lower().str.contains(kw, na=False)
        else:
            mask = filtered["출판사"].str.lower().str.contains(kw, na=False)

        results = filtered[mask].sort_values("판매지수", ascending=False)

        st.info(f"'**{keyword}**' 검색 결과: **{len(results)}건**")

        if not results.empty:
            col_list, col_chart = st.columns([2, 1])

            with col_chart:
                st.subheader("검색 결과 통계")
                stat_c1, stat_c2 = st.columns(2)
                stat_c1.metric("검색 결과 수", f"{len(results)}건")
                stat_c2.metric("평균 판매지수", f"{results['판매지수'].mean():,.0f}")

                if len(results) > 1:
                    fig_search_price = px.histogram(
                        results, x="판매가", nbins=15,
                        title="검색 결과 가격 분포",
                        color_discrete_sequence=["#667eea"],
                    )
                    fig_search_price.update_layout(margin=dict(t=40, b=20), height=250)
                    st.plotly_chart(fig_search_price, use_container_width=True)

            with col_list:
                for _, row in results.head(30).iterrows():
                    rank = int(row["순위"]) if pd.notna(row["순위"]) else "-"
                    title = str(row["도서명"])
                    author = str(row["저자"])
                    publisher = str(row["출판사"])
                    price = int(row["판매가"]) if pd.notna(row["판매가"]) else 0
                    sales_idx = int(row["판매지수"]) if pd.notna(row["판매지수"]) else 0
                    disc = str(row["할인율"]) if pd.notna(row["할인율"]) else "-"
                    link = str(row["링크"]) if pd.notna(row["링크"]) else ""

                    # 하이라이트 처리
                    def highlight(text, kw):
                        return re.sub(
                            re.escape(kw),
                            f'<span class="highlight">{kw}</span>',
                            text,
                            flags=re.IGNORECASE,
                        )

                    title_html = highlight(title, keyword)
                    author_html = highlight(author, keyword)
                    publisher_html = highlight(publisher, keyword)

                    card_html = f"""
                    <div class="search-result">
                        <strong>#{rank}</strong> &nbsp;
                        <a href="{link}" target="_blank" style="color:#1a1a2e; text-decoration:none;">
                            {title_html}
                        </a>
                        <br/>
                        <small style="color:#6c757d;">
                            저자: {author_html} | 출판사: {publisher_html} |
                            판매가: {price:,}원 | 할인율: {disc} | 판매지수: {sales_idx:,}
                        </small>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)

                if len(results) > 30:
                    st.caption(f"상위 30건만 표시됩니다. 총 {len(results)}건의 검색 결과가 있습니다.")
        else:
            st.warning("검색 결과가 없습니다. 다른 키워드로 시도해 보세요.")
    else:
        st.info("💡 팁: '바이브 코딩', '클로드', '제미나이', 'ChatGPT', '인공지능' 등의 키워드로 검색해 보세요.")

        st.subheader("🔥 인기 키워드 (도서명 기반)")
        all_titles = " ".join(filtered["도서명"].dropna().tolist())
        keywords_extract = re.findall(r"[가-힣]{2,}|[A-Za-z]{3,}", all_titles)
        kw_freq = pd.Series(keywords_extract).value_counts().head(20)

        kw_cols = st.columns(5)
        for i, (word, count) in enumerate(kw_freq.items()):
            with kw_cols[i % 5]:
                if st.button(f"{word} ({count})", key=f"kw_{i}", use_container_width=True):
                    st.session_state["search_keyword"] = word
                    st.rerun()

        if "search_keyword" in st.session_state:
            del st.session_state["search_keyword"]

# ═══════════════════════════════════════
# 8. 도서 추천 챗봇 (RAG + Groq)
# ═══════════════════════════════════════
with tab_chat:
    st.subheader("🤖 AI 도서 추천 챗봇")
    st.caption("KLUE-BERT 임베딩 + ChromaDB 벡터검색 + Groq AI 기반 RAG 챗봇")

    if not groq_api_key:
        st.warning("좌측 사이드바에서 **Groq API Key**를 입력해 주세요. https://console.groq.com 에서 무료 발급 가능합니다.")
        st.stop()

    import chromadb
    from sentence_transformers import SentenceTransformer

    @st.cache_resource(show_spinner="ChromaDB + KLUE-BERT 임베딩 모델 로딩 중...")
    def load_rag():
        client = chromadb.PersistentClient(path=CHROMA_PATH)
        collection = client.get_collection(name=COLLECTION_NAME)
        encoder = SentenceTransformer(EMBED_MODEL, trust_remote_code=True)
        return collection, encoder

    try:
        rag_collection, rag_encoder = load_rag()
    except Exception as e:
        st.error(f"ChromaDB 로드 실패: {e}\n\n`python src/build_embeddings.py`를 먼저 실행해 주세요.")
        st.stop()

    st.success(f"ChromaDB 연결 완료: {rag_collection.count()}권의 도서 데이터 로드됨")

    def search_books(query, n_results=10):
        query_embedding = rag_encoder.encode([query]).tolist()
        results = rag_collection.query(
            query_embeddings=query_embedding,
            n_results=n_results,
            include=["documents", "metadatas", "distances"],
        )
        books = []
        if results["metadatas"] and results["metadatas"][0]:
            for meta, doc, dist in zip(
                results["metadatas"][0], results["documents"][0], results["distances"][0]
            ):
                books.append({**meta, "_distance": round(dist, 4)})
        return books

    def filter_books_by_price(min_price=0, max_price=999999, sort_by="판매지수", top_n=10):
        tmp = df.copy()
        for col_name in ["판매가", "판매지수"]:
            tmp[col_name] = pd.to_numeric(
                tmp[col_name].astype(str).str.replace(",", "", regex=False).str.strip(),
                errors="coerce",
            ).fillna(0)
        tmp = tmp[(tmp["판매가"] >= min_price) & (tmp["판매가"] <= max_price)]
        tmp = tmp.nlargest(top_n, "판매지수")
        results = []
        for _, r in tmp.iterrows():
            results.append({
                "순위": int(r["순위"]) if pd.notna(r["순위"]) else 0,
                "도서명": str(r["도서명"]),
                "저자": str(r["저자"]),
                "출판사": str(r["출판사"]),
                "판매가": f"{int(r['판매가']):,}원",
                "판매지수": f"{int(r['판매지수']):,}",
                "링크": str(r["링크"]),
            })
        return {"count": len(results), "books": results}

    def filter_books_by_sales_index(min_index=0, max_index=9999999, top_n=10):
        tmp = df.copy()
        for col_name in ["판매가", "판매지수"]:
            tmp[col_name] = pd.to_numeric(
                tmp[col_name].astype(str).str.replace(",", "", regex=False).str.strip(),
                errors="coerce",
            ).fillna(0)
        tmp = tmp[(tmp["판매지수"] >= min_index) & (tmp["판매지수"] <= max_index)]
        tmp = tmp.nlargest(top_n, "판매지수")
        results = []
        for _, r in tmp.iterrows():
            results.append({
                "순위": int(r["순위"]) if pd.notna(r["순위"]) else 0,
                "도서명": str(r["도서명"]),
                "저자": str(r["저자"]),
                "출판사": str(r["출판사"]),
                "판매가": f"{int(r['판매가']):,}원",
                "판매지수": f"{int(r['판매지수']):,}",
                "링크": str(r["링크"]),
            })
        return {"count": len(results), "books": results}

    tools = [
        {
            "type": "function",
            "function": {
                "name": "filter_books_by_price",
                "description": "판매가 범위로 도서를 필터링하고 판매지수 순으로 정렬하여 상위 결과를 반환합니다. 가격에 대한 질문 시 사용하세요.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "min_price": {
                            "type": "integer",
                            "description": "최소 판매가 (원). 기본값 0.",
                        },
                        "max_price": {
                            "type": "integer",
                            "description": "최대 판매가 (원). 기본값 999999.",
                        },
                        "top_n": {
                            "type": "integer",
                            "description": "반환할 도서 수. 기본값 10.",
                        },
                    },
                    "required": [],
                },
            },
        },
        {
            "type": "function",
            "function": {
                "name": "filter_books_by_sales_index",
                "description": "판매지수 범위로 도서를 필터링하고 상위 결과를 반환합니다. 판매지수에 대한 질문 시 사용하세요.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "min_index": {
                            "type": "integer",
                            "description": "최소 판매지수. 기본값 0.",
                        },
                        "max_index": {
                            "type": "integer",
                            "description": "최대 판매지수. 기본값 9999999.",
                        },
                        "top_n": {
                            "type": "integer",
                            "description": "반환할 도서 수. 기본값 10.",
                        },
                    },
                    "required": [],
                },
            },
        },
    ]

    def call_groq(messages, api_key, model, tools=None):
        payload = {
            "model": model,
            "messages": messages,
            "temperature": 0.7,
            "max_tokens": 2048,
        }
        if tools:
            payload["tools"] = tools
        payload_bytes = _json.dumps(payload, ensure_ascii=False).encode("utf-8")
        resp = _requests.post(
            "https://api.groq.com/openai/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json; charset=utf-8",
            },
            data=payload_bytes,
            timeout=30,
        )
        return resp.json()

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []

    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"], unsafe_allow_html=True)

    if user_input := st.chat_input("IT/모바일 도서 추천을 요청하거나 궁금한 점을 물어보세요!"):
        st.session_state.chat_history.append({"role": "user", "content": user_input})
        with st.chat_message("user"):
            st.markdown(user_input)

        retrieved = search_books(user_input, n_results=10)

        context_lines = []
        for b in retrieved:
            link = b.get("링크", "")
            title = b.get("도서명", "")
            link_markdown = f"[{title}]({link})" if link and link != "nan" else title
            context_lines.append(
                f"[{b.get('순위', '?')}위] {link_markdown} | 저자: {b.get('저자', '')} | "
                f"출판사: {b.get('출판사', '')} | 판매가: {b.get('판매가', '')}원 | "
                f"판매지수: {b.get('판매지수', '')} | 유사도 거리: {b.get('_distance', '?')}"
            )
        context_str = "\n".join(context_lines) if context_lines else "(검색 결과 없음)"

        system_prompt = f"""당신은 yes24 IT 모바일 베스트셀러 도서 추천 전문가입니다.
아래는 사용자 질문과 의미적으로 가장 유사한 도서 {len(retrieved)}권입니다 (ChromaDB 벡터 검색 결과).

{context_str}

규칙:
1. 당신은 IT, 모바일, 코딩, 프로그래밍, AI, 데이터, 디지털 관련 도서만 추천할 수 있습니다.
2. IT/모바일과 관련 없는 질문(유럽 여행, 요리, 운동, 소설 등)이 들어오면 반드시 아래 메시지로 답변하세요:
   "저는 IT 모바일 책 추천 챗봇이기 때문에 해당 주제 책 추천은 어렵습니다."
3. 가격이나 판매지수에 대한 질문이 들어오면 반드시 함수를 호출하세요. 함수 호출 결과를 바탕으로 답변하세요.
4. 위 검색 결과를 참고하여 사용자 질문에 가장 관련 있는 도서를 최대 5권 추천하세요.
5. 각 추천 도서에 대해 도서명(하이퍼링크), 저자, 출판사, 판매가, 추천 이유를 설명하세요.
6. 도서명 뒤에 [도서명](URL) 형태로 yes24 상세 페이지 링크를 포함하세요.
7. 추천할 만한 도서가 없으면 "현재 데이터에서 조건에 맞는 도서를 찾지 못했습니다. 다른 키워드로 다시 질문해 주세요."라고 답변하세요.
8. 한국어로 간결하게 답변하세요."""

        messages = [{"role": "system", "content": system_prompt}]
        recent = st.session_state.chat_history[-6:]
        messages += [{"role": m["role"], "content": m["content"]} for m in recent]

        import requests as _requests

        with st.chat_message("assistant"):
            try:
                result = call_groq(messages, groq_api_key, groq_model, tools=tools)

                if "error" in result:
                    assistant_msg = f"API 오류: {result['error'].get('message', result['error'])}"
                elif "choices" not in result or len(result["choices"]) == 0:
                    assistant_msg = "예상치 못한 응답 형식입니다. API Key와 모델 설정을 확인해 주세요."
                else:
                    choice = result["choices"][0]
                    msg = choice["message"]

                    if msg.get("tool_calls"):
                        tool_results = []
                        for tc in msg["tool_calls"]:
                            fn_name = tc["function"]["name"]
                            fn_args = tc["function"].get("arguments", "{}")
                            if isinstance(fn_args, str):
                                fn_args = _json.loads(fn_args)

                            if fn_name == "filter_books_by_price":
                                fn_result = filter_books_by_price(**fn_args)
                            elif fn_name == "filter_books_by_sales_index":
                                fn_result = filter_books_by_sales_index(**fn_args)
                            else:
                                fn_result = {"error": f"알 수 없는 함수: {fn_name}"}

                            tool_results.append({
                                "tool_call_id": tc["id"],
                                "role": "tool",
                                "content": _json.dumps(fn_result, ensure_ascii=False),
                            })

                        messages.append(msg)
                        messages.extend(tool_results)

                        final_result = call_groq(messages, groq_api_key, groq_model)
                        if "error" in final_result:
                            assistant_msg = f"API 오류: {final_result['error'].get('message', final_result['error'])}"
                        elif "choices" in final_result and len(final_result["choices"]) > 0:
                            assistant_msg = final_result["choices"][0]["message"]["content"]
                        else:
                            assistant_msg = "함수 호출 결과를 처리하는 중 오류가 발생했습니다."
                    else:
                        assistant_msg = msg.get("content", "응답을 생성하지 못했습니다.")

            except Exception as e:
                assistant_msg = f"오류가 발생했습니다: {e}"

            st.markdown(assistant_msg, unsafe_allow_html=True)

        st.session_state.chat_history.append({"role": "assistant", "content": assistant_msg})

    if st.session_state.chat_history:
        if st.button("대화 초기화"):
            st.session_state.chat_history = []
            st.rerun()
