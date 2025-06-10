import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

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
            df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
            df["Sheet"] = sheet
            data_frames.append(df)

        df_all = pd.concat(data_frames, ignore_index=True)
        df_all.sort_values(by="Timestamp", inplace=True)

        st.subheader("📌 시트 통합 및 전처리 완료")
        st.write(f"총 {len(df_all)}개의 데이터 포인트가 통합되었습니다.")
        st.dataframe(df_all.head(10))

        # --- Descriptive Stats ---
        with st.expander("📈 기술 통계량 보기 (per 시트 또는 전체)"):
            mode = st.radio("보기 모드 선택:", ["전체 통합", "시트별"], horizontal=True)
            if mode == "전체 통합":
                st.write(df_all.describe(include='all'))
            else:
                for sheet in selected_sheets:
                    st.markdown(f"**▶ 시트: {sheet}**")
                    st.dataframe(df_all[df_all["Sheet"] == sheet].describe(include='all'))

        # --- Visualization of Descriptive Stats ---
        st.subheader("📊 시간 기반 변수 시각화")
        for column in ["Quantity", "RR_RH_1", "RR_RH_2"]:
            st.markdown(f"**📌 {column}**")
            fig, ax = plt.subplots(figsize=(12, 3))
            for sheet in selected_sheets:
                sub = df_all[df_all["Sheet"] == sheet]
                ax.bar(sub["Timestamp"], sub[column], width=0.003, label=sheet)
            ax.set_ylabel(column)
            ax.set_xlabel("Timestamp")
            ax.legend()
            st.pyplot(fig)

        # --- Correlation Analysis ---
        st.subheader("🔗 상관관계 분석")
        numeric_cols = ["Quantity", "RR_RH_1", "RR_RH_2"]
        st.markdown("**▶ 전체 데이터 기준 상관계수 (Pearson)**")
        corr = df_all[numeric_cols].corr(method="pearson")
        fig, ax = plt.subplots()
        sns.heatmap(corr, annot=True, cmap="coolwarm", ax=ax)
        st.pyplot(fig)

        # --- Cross Correlation ---
        st.subheader("⏱️ 시차 기반 상관관계 분석 (Cross Correlation)")
        ref_col = st.selectbox("기준 컬럼 선택:", numeric_cols)
        compare_col = st.selectbox("비교할 컬럼 선택:", [col for col in numeric_cols if col != ref_col])
        max_lag = st.slider("최대 시차 범위 (단위: row)", 1, 100, 20)

        series1 = df_all[ref_col].dropna()
        series2 = df_all[compare_col].dropna()

        min_len = min(len(series1), len(series2))
        series1 = series1[:min_len]
        series2 = series2[:min_len]

        xcorr = [series1.corr(series2.shift(lag)) for lag in range(-max_lag, max_lag+1)]
        fig, ax = plt.subplots()
        ax.plot(range(-max_lag, max_lag+1), xcorr)
        ax.set_title(f"Cross Correlation: {ref_col} vs {compare_col}")
        ax.set_xlabel("Lag")
        ax.set_ylabel("Correlation")
        st.pyplot(fig)

        # --- Sheet Comparison ---
        st.subheader("🧩 시트 간 비교")
        compare_stat = st.selectbox("비교할 통계값:", ["mean", "std", "min", "max"])
        sheet_stats = df_all.groupby("Sheet")[numeric_cols].agg(compare_stat)

        fig, ax = plt.subplots(figsize=(8, 4))
        sheet_stats.plot(kind="bar", ax=ax)
        ax.set_title(f"시트별 {compare_stat} 값 비교")
        st.pyplot(fig)

        # --- Missing Value Heatmap ---
        st.subheader("🕳️ 결측값 개요")
        missing_counts = df_all.groupby("Sheet")[numeric_cols].apply(lambda x: x.isna().sum())
        fig, ax = plt.subplots()
        sns.heatmap(missing_counts, annot=True, cmap="Reds", fmt="d", ax=ax)
        ax.set_title("시트별 결측값 개수")
        st.pyplot(fig)
