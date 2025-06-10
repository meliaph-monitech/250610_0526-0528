import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.font_manager as fm
from datetime import datetime, time

# Set Hangul font if available
HANGUL_FONT = None
for f in fm.findSystemFonts(fontpaths=None, fontext='ttf'):
    if "NanumGothic" in f or "Malgun" in f:
        HANGUL_FONT = f
        break

st.set_page_config(layout="wide")
st.title("📊 스마트 용접 신호 분석 대시보드")

# --- Excel time formatting fix ---
def format_excel_time(t):
    if pd.isna(t): return np.nan
    if isinstance(t, (pd.Timestamp, datetime)): return t.strftime("%H:%M")
    elif isinstance(t, time): return t.strftime("%H:%M")
    elif isinstance(t, (int, float)):
        minutes = int(round(t * 24 * 60))
        return f"{minutes // 60:02}:{minutes % 60:02}"
    else:
        try: return pd.to_datetime(t).strftime("%H:%M")
        except: return str(t)

# --- Sidebar Controls ---
st.sidebar.header("🧭 대시보드 설정")
uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    selected_sheets = st.sidebar.multiselect("시트 선택:", all_sheets, default=all_sheets[:3])

    if selected_sheets:
        dfs = []
        for sheet in selected_sheets:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = df.rename(columns={
                df.columns[0]: "Timestamp",
                df.columns[1]: "Quantity",
                df.columns[2]: "Sensor1",
                df.columns[3]: "Sensor2"
            })
            df["Timestamp"] = df["Timestamp"].apply(format_excel_time)
            df["Sheet"] = sheet
            df["Date"] = sheet[:4]
            df["SensorType"] = sheet.split("_")[-1]
            dfs.append(df)

        df_all = pd.concat(dfs, ignore_index=True)
        df_all.dropna(subset=["Timestamp"], inplace=True)
        df_all.fillna(0, inplace=True)

        # Derived columns
        df_all["Sensor1_per_unit"] = df_all["Sensor1"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Sensor2_per_unit"] = df_all["Sensor2"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Delta"] = df_all["Sensor1"] - df_all["Sensor2"]

        # Set timestamp as ordered categorical
        df_all["Timestamp"] = pd.Categorical(df_all["Timestamp"],
                                             categories=sorted(df_all["Timestamp"].unique()),
                                             ordered=True)

        # --- Overview Summary ---
        st.subheader("📌 데이터 요약")
        col1, col2 = st.columns(2)
        with col1:
            st.metric("전체 생산수량", int(df_all["Quantity"].sum()))
        with col2:
            st.metric("총 시트 개수", len(selected_sheets))

        st.dataframe(df_all[["Sheet", "Date", "SensorType", "Timestamp", "Quantity",
                             "Sensor1", "Sensor2", "Sensor1_per_unit", "Sensor2_per_unit", "Delta"]].head(10))

        # --- Time-series Plots ---
        st.subheader("⏱️ 시간별 생산수량 및 센서 평균값")
        for col in ["Quantity", "Sensor1", "Sensor2"]:
            fig = px.bar(df_all, x="Timestamp", y=col, color="Sheet", barmode="group",
                         title=f"{col} (시간 기준)", height=350)
            fig.update_layout(
                xaxis_tickangle=90,
                font=dict(family="Nanum Gothic" if HANGUL_FONT else None)
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- Sensor vs Quantity Correlation ---
        st.subheader("🔗 생산량과 센서 평균값의 상관관계")
        for sensor_col in ["Sensor1", "Sensor2"]:
            fig = px.scatter(df_all, x="Quantity", y=sensor_col, color="Sheet", trendline="ols",
                             title=f"{sensor_col} vs Quantity (산점도 + 추세선)")
            fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Normalized Signal (per unit) ---
        st.subheader("⚖️ 센서당 용접 단위당 평균 신호")
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            fig = px.line(df_all, x="Timestamp", y=col, color="Sheet", markers=True,
                          title=f"{col} (시간 순)")
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Delta Between Sensors ---
        st.subheader("📉 센서 차이: Sensor1 - Sensor2")
        fig = px.line(df_all, x="Timestamp", y="Delta", color="Sheet", markers=True,
                      title="Delta: Sensor1 - Sensor2")
        fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Rolling Mean & Diff ---
        st.subheader("🔄 이동 평균 및 변화량 분석")
        window = st.sidebar.slider("이동 윈도우 크기 (row)", 1, 20, 5)
        for col in ["Sensor1", "Sensor2"]:
            df_all[f"{col}_roll"] = df_all[col].rolling(window=window, min_periods=1).mean()
            df_all[f"{col}_diff"] = df_all[col].diff()

            fig1 = px.line(df_all, x="Timestamp", y=f"{col}_roll", color="Sheet", title=f"{col} 이동 평균",
                           labels={f"{col}_roll": f"{col} 이동 평균"})
            fig2 = px.line(df_all, x="Timestamp", y=f"{col}_diff", color="Sheet", title=f"{col} 변화량",
                           labels={f"{col}_diff": f"{col} 변화량"})

            fig1.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            fig2.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))

            st.plotly_chart(fig1, use_container_width=True)
            st.plotly_chart(fig2, use_container_width=True)

        # --- Sheet-level Summary ---
        st.subheader("📊 시트 요약 비교")
        sheet_summary = df_all.groupby("Sheet").agg({
            "Quantity": "sum",
            "Sensor1": "mean",
            "Sensor2": "mean",
            "Sensor1_per_unit": "mean",
            "Sensor2_per_unit": "mean",
            "Delta": "mean"
        }).reset_index()

        st.dataframe(sheet_summary.round(2))
        fig = px.bar(sheet_summary, x="Sheet", y="Quantity", title="시트별 총 생산량", color="Sheet")
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)
