import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(layout="wide")
st.title("📊 5분 간격 생산 및 측정 데이터 분석 대시보드")

# --- File upload ---
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

            # Convert h:mm time format to string (preserve as-is)
            df["Timestamp"] = df["Timestamp"].astype(str)

            df["Sheet"] = sheet
            data_frames.append(df)

        # Combine all sheets
        df_all = pd.concat(data_frames, ignore_index=True)
        df_all.dropna(subset=["Timestamp"], inplace=True)

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

        # --- Time-series Plots ---
        st.subheader("📊 컬럼별 시트 데이터 시각화 (시간 기준 막대 그래프)")
        for column in ["Quantity", "RR_RH_1", "RR_RH_2"]:
            st.markdown(f"**📌 {column}**")
            fig, ax = plt.subplots(figsize=(12, 3))
            for sheet in selected_sheets:
                sub = df_all[df_all["Sheet"] == sheet]
                ax.bar(sub["Timestamp"], sub[column], width=0.4, label=sheet)
            ax.set_ylabel(column)
            ax.set_xlabel("시간 (h:mm)")
            ax.legend()
            st.pyplot(fig)

        # --- Correlation Matrix ---
        st.subheader("🔗 상관관계 분석 (Pearson)")
        numeric_cols = ["Quantity", "RR_RH_1", "RR_RH_2"]
        corr = df_all[numeric_cols].corr(method="pearson")
        fig, ax = plt.subplots()
        sns.heatmap(corr, annot=True, cmap="coolwarm", ax=ax)
        st.pyplot(fig)

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

        lags = range(-max_lag, max_lag + 1)
        xcorr = [s1.corr(s2.shift(lag)) for lag in lags]

        fig, ax = plt.subplots()
        ax.plot(lags, xcorr)
        ax.set_title(f"Cross Correlation: {ref_col} vs {compare_col}")
        ax.set_xlabel("Lag")
        ax.set_ylabel("Correlation")
        st.pyplot(fig)

        # --- Sheet Comparison ---
        st.subheader("🧩 시트별 통계 비교")
        stat_option = st.selectbox("비교할 통계 항목:", ["mean", "std", "min", "max"])
        sheet_stats = df_all.groupby("Sheet")[numeric_cols].agg(stat_option)

        fig, ax = plt.subplots(figsize=(8, 4))
        sheet_stats.plot(kind="bar", ax=ax)
        ax.set_title(f"{stat_option.upper()} 값 시트별 비교")
        st.pyplot(fig)

        # --- Missing Data Heatmap ---
        st.subheader("🕳️ 결측값 개요 (Missing Value Heatmap)")
        missing_counts = df_all.groupby("Sheet")[numeric_cols].apply(lambda x: x.isna().sum())
        fig, ax = plt.subplots()
        sns.heatmap(missing_counts, annot=True, cmap="Reds", fmt="d", ax=ax)
        ax.set_title("시트별 결측값 개수")
        st.pyplot(fig)
