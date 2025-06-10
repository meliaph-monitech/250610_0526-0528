import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.font_manager as fm
from datetime import datetime, time

# Set Hangul font
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

        # Derived metrics
        df_all["Sensor1_per_unit"] = df_all["Sensor1"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Sensor2_per_unit"] = df_all["Sensor2"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Delta"] = df_all["Sensor1"] - df_all["Sensor2"]

        st.markdown("## 📌 데이터 요약<br><span style='color:gray'>Data Summary</span>", unsafe_allow_html=True)
        st.metric("전체 생산수량", int(df_all["Quantity"].sum()))
        st.metric("총 시트 개수", len(selected_sheets))
        st.dataframe(df_all.head(10))

        # --- Global Correlation Matrix ---
        st.markdown("## 🌐 전체 상관계수 분석<br><span style='color:gray'>Global Correlation Matrix</span>", unsafe_allow_html=True)
        corr_cols = ["Quantity", "Sensor1", "Sensor2", "Delta", "Sensor1_per_unit", "Sensor2_per_unit"]
        global_corr = df_all[corr_cols].corr()
        fig = px.imshow(global_corr, text_auto=True, aspect="auto", color_continuous_scale="RdBu", zmin=-1, zmax=1)
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Time-Series by TimeKey
        st.markdown("## ⏱️ 시간별 센서 및 생산량<br><span style='color:gray'>Sensor/Quantity Over Time</span>", unsafe_allow_html=True)
        for col in ["Quantity", "Sensor1", "Sensor2"]:
            fig = px.bar(df_all, x="TimeKey", y=col, color="Sheet",
                         title=f"{col} (시간순)<br><span style='color:gray'>{col} Over Time</span>",
                         labels={"TimeKey": "시트+시간<br><span style='color:gray'>Sheet+Time</span>", col: col})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Sensor per Unit vs Quantity
        st.markdown("## 📉 생산량 대비 단위당 센서 평균값<br><span style='color:gray'>Sensor Signal per Unit vs Quantity</span>", unsafe_allow_html=True)
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            fig = px.scatter(df_all, x="Quantity", y=col, color="Sheet", trendline="lowess",
                             title=f"{col} vs Quantity<br><span style='color:gray'>{col} vs Quantity</span>",
                             labels={"Quantity": "생산량<br><span style='color:gray'>Quantity</span>",
                                     col: "단위당 평균<br><span style='color:gray'>Per-Unit Average</span>"})
            fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Delta plots
        st.markdown("## ⚖️ 센서 차이 및 드리프트<br><span style='color:gray'>Sensor Delta & Drift</span>", unsafe_allow_html=True)
        fig = px.line(df_all, x="TimeKey", y="Delta", color="Sheet",
                      title="Sensor1 - Sensor2<br><span style='color:gray'>Delta Over Time</span>",
                      labels={"TimeKey": "시트+시간<br><span style='color:gray'>Sheet+Time</span>", "Delta": "센서 차이<br><span style='color:gray'>Delta</span>"})
        fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        fig = px.histogram(df_all, x="Delta", color="Sheet", nbins=50,
                           title="Sensor Delta 분포<br><span style='color:gray'>Distribution of Sensor Delta</span>")
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Rolling Mean
        st.markdown("## 🔄 단위당 신호의 이동 평균<br><span style='color:gray'>Rolling Mean of Signal per Weld</span>", unsafe_allow_html=True)
        window = st.sidebar.slider("이동 평균 윈도우 (row)\n\nRolling Window Size", 1, 20, 5)
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            df_all[f"{col}_roll"] = df_all[col].rolling(window=window, min_periods=1).mean()
            fig = px.line(df_all, x="TimeKey", y=f"{col}_roll", color="Sheet",
                          title=f"{col} 이동 평균<br><span style='color:gray'>Rolling Average</span>",
                          labels={"TimeKey": "시트+시간<br><span style='color:gray'>Sheet+Time</span>",
                                  f"{col}_roll": "이동 평균<br><span style='color:gray'>Rolling Average</span>"})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Time of Day Boxplot
        st.markdown("## 🕰️ 시간대별 센서 퍼 유닛 분포<br><span style='color:gray'>Signal per Weld by Time of Day</span>", unsafe_allow_html=True)
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            fig = px.box(df_all, x="Timestamp", y=col, color="SensorType",
                         title=f"{col} 시간대별 분포<br><span style='color:gray'>{col} by Time of Day</span>",
                         labels={"Timestamp": "시간<br><span style='color:gray'>Time</span>",
                                 col: "센서 퍼 유닛<br><span style='color:gray'>Signal per Weld</span>"})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Per-Sheet Correlations
        st.markdown("## 📊 시트별 상관계수 분석<br><span style='color:gray'>Per-Sheet Correlation Matrix</span>", unsafe_allow_html=True)
        for sheet in selected_sheets:
            subset = df_all[df_all["Sheet"] == sheet]
            corr = subset[corr_cols].corr()
            st.markdown(f"**{sheet} 상관계수**<br><span style='color:gray'>{sheet} Correlation Matrix</span>", unsafe_allow_html=True)
            fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale="RdBu", zmin=-1, zmax=1)
            fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Sensor Reliability Index (SRI)
        st.markdown("## 📏 센서 안정성 지수 (SRI)<br><span style='color:gray'>Sensor Reliability Index</span>", unsafe_allow_html=True)
        sheet_scores = df_all.groupby("Sheet").agg({
            "Sensor1_per_unit": "std",
            "Sensor2_per_unit": "std",
            "Delta": "mean"
        }).reset_index()
        sheet_scores["SRI"] = 1 - (
            sheet_scores["Sensor1_per_unit"] + sheet_scores["Sensor2_per_unit"] + sheet_scores["Delta"].abs()
        ) / 3
        fig = px.bar(sheet_scores.sort_values("SRI", ascending=False), x="Sheet", y="SRI", color="Sheet",
                     title="센서 안정성 지수 (높을수록 좋음)<br><span style='color:gray'>Higher = More Stable</span>",
                     labels={"SRI": "안정성 지수<br><span style='color:gray'>SRI</span>"})
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(sheet_scores.round(4))
