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
st.title("📊 스마트 용접 신호 분석 대시보드")

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
            df["TimeKey"] = df["Sheet"] + "_" + df["Timestamp"]
            dfs.append(df)

        df_all = pd.concat(dfs, ignore_index=True)
        df_all.dropna(subset=["Timestamp"], inplace=True)
        df_all.fillna(0, inplace=True)

        df_all["Sensor1_per_unit"] = df_all["Sensor1"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Sensor2_per_unit"] = df_all["Sensor2"] / df_all["Quantity"].replace(0, np.nan)
        df_all["Delta"] = df_all["Sensor1"] - df_all["Sensor2"]

        st.subheader("📌 데이터 요약")
        st.metric("전체 생산수량", int(df_all["Quantity"].sum()))
        st.metric("총 시트 개수", len(selected_sheets))

        # Show head of data
        st.dataframe(df_all[["Sheet", "Date", "SensorType", "Timestamp", "Quantity", "Sensor1", "Sensor2",
                             "Sensor1_per_unit", "Sensor2_per_unit", "Delta"]].head(10))

        # --- Time-Series by TimeKey (to avoid cross-day link)
        st.subheader("⏱️ 시간별 센서 및 생산량")
        for col in ["Quantity", "Sensor1", "Sensor2"]:
            fig = px.bar(df_all, x="TimeKey", y=col, color="Sheet", title=f"{col} (시간순)",
                         labels={"TimeKey": "시트+시간", col: col})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Quantity vs Sensor per unit
        st.subheader("📉 생산량 대비 단위당 센서 평균값")
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            fig = px.scatter(df_all, x="Quantity", y=col, color="Sheet", trendline="lowess",
                             title=f"{col} vs Quantity", labels={"Quantity": "생산량", col: "단위당 평균"})
            fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Delta plots
        st.subheader("⚖️ 센서 차이 및 드리프트")
        fig = px.line(df_all, x="TimeKey", y="Delta", color="Sheet", title="Sensor1 - Sensor2 (Delta)",
                      labels={"TimeKey": "시트+시간", "Delta": "차이"})
        fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        fig = px.histogram(df_all, x="Delta", color="Sheet", nbins=50,
                           title="Sensor Delta 분포 (Sensor1 - Sensor2)")
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        # --- Rolling mean per unit signal
        st.subheader("🔄 단위당 신호의 이동 평균")
        window = st.sidebar.slider("이동 평균 윈도우 (row)", 1, 20, 5)
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            df_all[f"{col}_roll"] = df_all[col].rolling(window=window, min_periods=1).mean()
            fig = px.line(df_all, x="TimeKey", y=f"{col}_roll", color="Sheet",
                          title=f"{col} 이동 평균", labels={f"{col}_roll": "이동 평균", "TimeKey": "시트+시간"})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Time-of-Day Boxplot
        st.subheader("🕰️ 시간대별 센서 퍼 유닛 분포 (모든 시트)")
        for col in ["Sensor1_per_unit", "Sensor2_per_unit"]:
            fig = px.box(df_all, x="Timestamp", y=col, color="SensorType",
                         title=f"{col} 시간대별 분포", labels={"Timestamp": "시간", col: "센서 퍼 유닛"})
            fig.update_layout(xaxis_tickangle=90, font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Sheet-level correlation matrix
        st.subheader("📊 시트별 상관계수 분석")
        grouped_corrs = []
        for sheet in selected_sheets:
            subset = df_all[df_all["Sheet"] == sheet]
            corr = subset[["Quantity", "Sensor1", "Sensor2", "Delta",
                           "Sensor1_per_unit", "Sensor2_per_unit"]].corr()
            grouped_corrs.append((sheet, corr))

        for sheet, corr in grouped_corrs:
            st.markdown(f"**{sheet} 상관계수**")
            fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale="RdBu", zmin=-1, zmax=1)
            fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
            st.plotly_chart(fig, use_container_width=True)

        # --- Reliability Index (SRI)
        st.subheader("📏 센서 안정성 지수 (SRI)")
        sheet_scores = df_all.groupby("Sheet").agg({
            "Sensor1_per_unit": "std",
            "Sensor2_per_unit": "std",
            "Delta": "mean"
        }).reset_index()

        # Simple scoring: lower std and smaller delta is better
        sheet_scores["SRI"] = 1 - (
            sheet_scores["Sensor1_per_unit"] + sheet_scores["Sensor2_per_unit"] + sheet_scores["Delta"].abs()
        ) / 3

        fig = px.bar(sheet_scores.sort_values("SRI", ascending=False),
                     x="Sheet", y="SRI", title="센서 안정성 지수 (높을수록 안정)",
                     labels={"SRI": "안정성 지수"}, color="Sheet")
        fig.update_layout(font=dict(family="Nanum Gothic" if HANGUL_FONT else None))
        st.plotly_chart(fig, use_container_width=True)

        st.dataframe(sheet_scores.round(4))
