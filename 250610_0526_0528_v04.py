import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.font_manager as fm
from datetime import datetime, time

# Hangul font setup
HANGUL_FONT = None
for f in fm.findSystemFonts(fontpaths=None, fontext='ttf'):
    if "NanumGothic" in f or "Malgun" in f:
        HANGUL_FONT = f
        break

# --- Utility to handle Excel time-only values ---
def format_excel_time(t):
    if pd.isna(t):
        return np.nan
    if isinstance(t, (pd.Timestamp, datetime)):
        return t.strftime("%H:%M")
    elif isinstance(t, time):
        return t.strftime("%H:%M")
    elif isinstance(t, (int, float)):
        minutes = int(round(t * 24 * 60))
        return f"{minutes // 60:02}:{minutes % 60:02}"
    else:
        try:
            return pd.to_datetime(t).strftime("%H:%M")
        except:
            return str(t)

# --- Streamlit Layout ---
st.set_page_config(layout="wide")
st.title("📊 생산 데이터 통합 분석 대시보드")

# --- Sidebar Controls ---
st.sidebar.header("🧭 대시보드 설정")

uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])
if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    selected_sheets = st.sidebar.multiselect("시트 선택:", all_sheets, default=all_sheets[:3])

    if selected_sheets:
        data_frames = []
        for sheet in selected_sheets:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = df.rename(columns={df.columns[0]: "Timestamp", df.columns[1]: "Quantity",
                                    df.columns[2]: "RR_RH_1", df.columns[3]: "RR_RH_2"})
            df["Timestamp"] = df["Timestamp"].apply(format_excel_time)
            df["Sheet"] = sheet
            data_frames.append(df)

        df_all = pd.concat(data_frames, ignore_index=True)
        df_all.dropna(subset=["Timestamp"], inplace=True)
        df_all["Timestamp"] = pd.Categorical(df_all["Timestamp"],
                                             categories=sorted(df_all["Timestamp"].unique()),
                                             ordered=True)

        st.subheader("📌 전처리된 통합 데이터")
        st.dataframe(df_all.sample(10))

        # --- Descriptive Stats ---
        stat_mode = st.sidebar.radio("기술 통계 보기 방식:", ["전체 통합", "시트별"])
        with st.expander("📈 기술 통계량"):
            if stat_mode == "전체 통합":
                st.dataframe(df_all[["Quantity", "RR_RH_1", "RR_RH_2"]].describe())
            else:
                for sheet in selected_sheets:
                    st.markdown(f"**▶ 시트: {sheet}**")
                    st.dataframe(df_all[df_all["Sheet"] == sheet][["Quantity", "RR_RH_1", "RR_RH_2"]].describe())

        # --- Time-Series Plots ---
        st.subheader("📊 시간별 바 시각화 (컬럼별)")
        for col in ["Quantity", "RR_RH_1", "RR_RH_2"]:
            fig = px.bar(df_all, x="Timestamp", y=col, color="Sheet", height=350,
                         title=f"{col} (시트별 구분)", labels={"Timestamp": "시간", col: col})
            fig.update_layout(
                xaxis_tickangle=90,
                xaxis_tickfont=dict(size=10),
                margin=dict(l=40, r=20, t=50, b=120),
                font=dict(family="Nanum Gothic" if HANGUL_FONT else None)
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- Correlation Heatmap ---
        st.subheader("🔗 상관관계 분석")
        corr = df_all[["Quantity", "RR_RH_1", "RR_RH_2"]].corr().round(3)
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
        fig.update_layout(title="📌 Pearson 상관계수", font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Cross Correlation ---
        st.subheader("⏱️ Cross Correlation")
        numeric_cols = ["Quantity", "RR_RH_1", "RR_RH_2"]
        ref_col = st.sidebar.selectbox("기준 컬럼 선택:", numeric_cols)
        compare_col = st.sidebar.selectbox("비교할 컬럼 선택:", [c for c in numeric_cols if c != ref_col])
        max_lag = st.sidebar.slider("최대 시차 (lag)", 1, 100, 20)

        s1 = df_all[ref_col].dropna().reset_index(drop=True)
        s2 = df_all[compare_col].dropna().reset_index(drop=True)
        min_len = min(len(s1), len(s2))
        s1, s2 = s1[:min_len], s2[:min_len]

        lags = list(range(-max_lag, max_lag + 1))
        xcorr = [s1.corr(s2.shift(lag)) for lag in lags]

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=lags, y=xcorr, mode="lines+markers", name="Cross Corr"))
        fig.update_layout(title=f"Cross Correlation: {ref_col} vs {compare_col}",
                          xaxis_title="시차 (Lag)", yaxis_title="상관계수",
                          font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Scatter Plot of RR_RH-1 vs RR_RH-2 ---
        st.subheader("🧪 RR_RH-1 vs RR_RH-2 산점도")
        fig = px.scatter(df_all, x="RR_RH_1", y="RR_RH_2", color="Sheet", opacity=0.7,
                         title="RR_RH-1 vs RR_RH-2", labels={"RR_RH_1": "RR_RH-1", "RR_RH_2": "RR_RH-2"})
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Delta Plot (RR_RH-1 - RR_RH-2) ---
        st.subheader("📉 RR_RH-1 - RR_RH-2 차이")
        df_all["Delta"] = df_all["RR_RH_1"] - df_all["RR_RH_2"]
        fig = px.bar(df_all, x="Timestamp", y="Delta", color="Sheet", title="Delta: RR_RH-1 - RR_RH-2",
                     labels={"Timestamp": "시간", "Delta": "차이"})
        fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Rolling Mean and Rate of Change ---
        st.subheader("🔄 이동 평균 및 변화율")
        window = st.sidebar.slider("이동 평균 윈도우 (row 수)", 1, 20, 5)
        for col in ["RR_RH_1", "RR_RH_2"]:
            df_all[f"{col}_roll"] = df_all[col].rolling(window=window).mean()
            df_all[f"{col}_diff"] = df_all[col].diff()

            st.markdown(f"**{col} - 이동 평균**")
            fig = px.line(df_all, x="Timestamp", y=f"{col}_roll", color="Sheet", labels={"value": "값"})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

            st.markdown(f"**{col} - 변화율 (diff)**")
            fig = px.line(df_all, x="Timestamp", y=f"{col}_diff", color="Sheet", labels={"value": "값"})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Missing Value Patterns ---
        st.subheader("🕳️ 결측 패턴 시각화")
        nan_df = df_all[["Timestamp", "Sheet", "RR_RH_1", "RR_RH_2"]].copy()
        nan_df["RR_RH_1_missing"] = nan_df["RR_RH_1"].isna().astype(int)
        nan_df["RR_RH_2_missing"] = nan_df["RR_RH_2"].isna().astype(int)

        fig = px.bar(nan_df, x="Timestamp", y="RR_RH_1_missing", color="Sheet", title="RR_RH-1 결측 여부",
                     labels={"RR_RH_1_missing": "결측(1=결측)", "Timestamp": "시간"})
        fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        fig = px.bar(nan_df, x="Timestamp", y="RR_RH_2_missing", color="Sheet", title="RR_RH-2 결측 여부",
                     labels={"RR_RH_2_missing": "결측(1=결측)", "Timestamp": "시간"})
        fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)
