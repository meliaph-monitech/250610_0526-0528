import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.font_manager as fm

# Korean font handling for Plotly hover/display
HANGUL_FONT = None
for f in fm.findSystemFonts(fontpaths=None, fontext='ttf'):
    if "NanumGothic" in f or "Malgun" in f:
        HANGUL_FONT = f
        break

st.set_page_config(layout="wide")
st.title("📊 5분 간격 생산 및 측정 데이터 분석 대시보드")

# --- Upload section ---
uploaded_file = st.file_uploader("📂 엑셀 파일을 업로드하세요 (.xlsx)", type=["xlsx"])
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    selected_sheets = st.multiselect("분석할 시트를 선택하세요:", all_sheets, default=all_sheets[:3])

    if selected_sheets:
        data_frames = []
        for sheet in selected_sheets:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = df.rename(columns={df.columns[0]: "Timestamp", df.columns[1]: "Quantity",
                                    df.columns[2]: "RR_RH_1", df.columns[3]: "RR_RH_2"})
            df["Timestamp"] = df["Timestamp"].astype(str)
            df["Sheet"] = sheet
            data_frames.append(df)

        df_all = pd.concat(data_frames, ignore_index=True)
        df_all.dropna(subset=["Timestamp"], inplace=True)

        # Force categorical time labels to prevent Plotly from misinterpreting
        df_all["Timestamp"] = df_all["Timestamp"].astype(str)
        unique_times_sorted = sorted(df_all["Timestamp"].unique())
        df_all["Timestamp"] = pd.Categorical(df_all["Timestamp"], categories=unique_times_sorted, ordered=True)

        st.subheader("📌 전처리된 통합 데이터")
        st.write(f"총 {len(df_all)}개의 데이터가 통합되었습니다.")
        st.dataframe(df_all.head(10))

        # --- Descriptive Statistics ---
        with st.expander("📈 기술 통계량 보기 (시트별 / 전체 통합)"):
            mode = st.radio("보기 모드 선택:", ["전체 통합", "시트별"], horizontal=True)
            if mode == "전체 통합":
                st.dataframe(df_all[["Quantity", "RR_RH_1", "RR_RH_2"]].describe())
            else:
                for sheet in selected_sheets:
                    st.markdown(f"**▶ 시트: {sheet}**")
                    st.dataframe(df_all[df_all["Sheet"] == sheet][["Quantity", "RR_RH_1", "RR_RH_2"]].describe())

        # --- Time-series Plot per Column ---
        st.subheader("📊 시간별 변수 시각화 (컬럼별)")
        for column in ["Quantity", "RR_RH_1", "RR_RH_2"]:
            st.markdown(f"**📌 {column}**")
            fig = px.bar(
                df_all, x="Timestamp", y=column, color="Sheet",
                labels={"Timestamp": "시간", column: column},
                title=f"{column} (시트별 구분)", height=350
            )
            fig.update_layout(
                xaxis_tickangle=90,
                xaxis_tickfont=dict(size=10),
                yaxis_title=column,
                margin=dict(l=40, r=20, t=50, b=120),
                legend_title="시트",
                font=dict(family="Nanum Gothic" if HANGUL_FONT else None)
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- Correlation Matrix (Pearson) ---
        st.subheader("🔗 상관관계 분석 (전체 통합)")
        numeric_cols = ["Quantity", "RR_RH_1", "RR_RH_2"]
        corr = df_all[numeric_cols].corr(method="pearson").round(3)

        fig = go.Figure(data=go.Heatmap(
            z=corr.values,
            x=corr.columns,
            y=corr.index,
            colorscale="RdBu",
            zmin=-1,
            zmax=1,
            text=corr.values,
            texttemplate="%{text}",
            hoverinfo="text"
        ))
        fig.update_layout(
            title="📌 Pearson 상관계수 히트맵",
            font=dict(family="Nanum Gothic" if HANGUL_FONT else None),
            height=400
        )
        st.plotly_chart(fig, use_container_width=True)

        # --- Cross Correlation ---
        st.subheader("⏱️ 시차 기반 상관관계 분석 (Cross Correlation)")
        ref_col = st.selectbox("기준 컬럼 선택:", numeric_cols)
        compare_col = st.selectbox("비교할 컬럼 선택:", [c for c in numeric_cols if c != ref_col])
        max_lag = st.slider("최대 시차 범위 (row shift)", 1, 100, 20)

        s1 = df_all[ref_col].dropna().reset_index(drop=True)
        s2 = df_all[compare_col].dropna().reset_index(drop=True)
        min_len = min(len(s1), len(s2))
        s1 = s1[:min_len]
        s2 = s2[:min_len]

        lags = list(range(-max_lag, max_lag + 1))
        xcorr = [s1.corr(s2.shift(lag)) for lag in lags]

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=lags, y=xcorr, mode="lines+markers", name="Cross Correlation"))
        fig.update_layout(
            title=f"📌 Cross Correlation: {ref_col} vs {compare_col}",
            xaxis_title="시차 (Lag)",
            yaxis_title="상관계수",
            font=dict(family="Nanum Gothic" if HANGUL_FONT else None)
        )
        st.plotly_chart(fig, use_container_width=True)

        # --- Sheet Comparison by Aggregates ---
        st.subheader("🧩 시트별 통계값 비교")
        stat_option = st.selectbox("비교할 통계 항목:", ["mean", "std", "min", "max"])
        sheet_stats = df_all.groupby("Sheet")[numeric_cols].agg(stat_option)

        fig = px.bar(
            sheet_stats.reset_index().melt(id_vars="Sheet"),
            x="Sheet", y="value", color="variable", barmode="group",
            title=f"{stat_option.upper()} 값 시트별 비교",
            labels={"value": "값", "variable": "컬럼"},
            height=400
        )
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Missing Value Heatmap (Table) ---
        st.subheader("🕳️ 결측값 개요 (Missing Value Overview)")
        missing_counts = df_all.groupby("Sheet")[numeric_cols].apply(lambda x: x.isna().sum()).reset_index()
        st.dataframe(missing_counts)
