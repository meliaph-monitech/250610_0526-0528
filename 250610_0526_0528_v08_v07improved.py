import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import matplotlib.font_manager as fm
from datetime import datetime, time

# Set Korean font if available
HANGUL_FONT = None
for f in fm.findSystemFonts(fontpaths=None, fontext='ttf'):
    if "NanumGothic" in f or "Malgun" in f:
        HANGUL_FONT = f
        break

st.set_page_config(layout="wide")
st.markdown("# 📊 스마트 용접 신호 분석 대시보드<br><span style='color:gray'>Smart Welding Signal Analysis Dashboard</span>", unsafe_allow_html=True)

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

# Sidebar
st.sidebar.header("🧭 대시보드 설정\n\nSensor Data Dashboard Settings")
uploaded_file = st.sidebar.file_uploader("엑셀 파일 업로드 (.xlsx)\n\nUpload Excel File", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    all_sheets = xls.sheet_names
    selected_sheets = st.sidebar.multiselect("시트 선택:\n\nSelect Sheets", all_sheets, default=all_sheets[:3])

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
            df["TimeKey"] = df["Sheet"] + "_" + df["Timestamp"]
            dfs.append(df)

        df_all = pd.concat(dfs, ignore_index=True)
        df_all.dropna(subset=["Timestamp"], inplace=True)
        df_all.fillna(0, inplace=True)

        # Feature engineering
        df_all["Sensor1_per_unit"] = df_all["Sensor1"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Sensor2_per_unit"] = df_all["Sensor2"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Delta"] = df_all["Sensor1"] - df_all["Sensor2"]

        # ──────────────────────────────────────────────
        # 📌 Data Summary
        st.markdown("## 📌 데이터 요약<br><span style='color:gray'>Data Summary</span>", unsafe_allow_html=True)
        st.metric("전체 생산수량", int(df_all["Quantity"].sum()))
        st.metric("총 시트 개수", len(selected_sheets))
        st.dataframe(df_all.head(10))

        # ──────────────────────────────────────────────
        # ⏱️ Sensor/Quantity Over Time
        st.markdown("## ⏱️ 시간별 센서 및 생산량<br><span style='color:gray'>Sensor/Quantity Over Time</span>", unsafe_allow_html=True)

        for col in ["Quantity", "Sensor1", "Sensor2"]:
            fig = px.bar(df_all, x="TimeKey", y=col, color="Sheet",
                         title=f"{col} (시간순)<br><span style='color:gray'>{col} Over Time</span>",
                         labels={"TimeKey": "시트+시간<br><span style='color:gray'>Sheet+Time</span>", col: col})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # ── Combined Dual-Y Axis Plot
        st.markdown("### 📊 통합 시계열 보기<br><span style='color:gray'>Combined Time Series of Quantity & Sensors</span>", unsafe_allow_html=True)
        fig = go.Figure()

        fig.add_bar(x=df_all["TimeKey"], y=df_all["Quantity"], name="Quantity", yaxis='y1', marker_color='rgba(100,149,237,0.6)')

        fig.add_trace(go.Scatter(x=df_all["TimeKey"], y=df_all["Sensor1"], name="Sensor1",
                                 yaxis='y2', mode='lines+markers', line=dict(color='firebrick')))
        fig.add_trace(go.Scatter(x=df_all["TimeKey"], y=df_all["Sensor2"], name="Sensor2",
                                 yaxis='y2', mode='lines+markers', line=dict(color='green')))

        fig.update_layout(
            title="생산량 및 센서값 통합 보기<br><span style='color:gray'>Quantity (bar) + Sensor1/2 (lines)</span>",
            xaxis=dict(title="시트+시간<br><span style='color:gray'>Sheet+Time</span>", tickangle=90),
            yaxis=dict(title="생산량<br><span style='color:gray'>Quantity</span>", side='left'),
            yaxis2=dict(title="센서 평균값<br><span style='color:gray'>Sensor Value</span>", overlaying='y', side='right'),
            legend=dict(x=1.01, y=1),
            font=dict(family="Nanum Gothic" if HANGUL_FONT else None)
        )
        st.plotly_chart(fig, use_container_width=True)

        # ──────────────────────────────────────────────
        # 🌐 Correlation Section
        st.markdown("## 📊 상관관계 분석<br><span style='color:gray'>Correlation Analysis</span>", unsafe_allow_html=True)
        corr_cols = ["Quantity", "Sensor1", "Sensor2", "Delta", "Sensor1_per_unit", "Sensor2_per_unit"]

        # Global Correlation
        st.markdown("#### 🌐 전체 상관계수<br><span style='color:gray'>Global Correlation Matrix</span>", unsafe_allow_html=True)
        global_corr = df_all[corr_cols].corr()
        fig = px.imshow(global_corr, text_auto=True, aspect="auto", color_continuous_scale="RdBu", zmin=-1, zmax=1)
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # Per-Sheet Correlations inside expander
        with st.expander("📂 시트별 상관계수 보기\n\nView Correlation Matrix per Sheet"):
            for sheet in selected_sheets:
                subset = df_all[df_all["Sheet"] == sheet]
                corr = subset[corr_cols].corr()
                st.markdown(f"**{sheet} 상관계수**<br><span style='color:gray'>{sheet} Correlation Matrix</span>", unsafe_allow_html=True)
                fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale="RdBu", zmin=-1, zmax=1)
                fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
                st.plotly_chart(fig, use_container_width=True)
