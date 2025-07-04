import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import plotly.express as px
from io import BytesIO

st.set_page_config(layout="wide")

st.title("📊 Korean Partner Data Analysis Dashboard")

# File uploader
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=".xlsx")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    selected_sheets = st.sidebar.multiselect("Select sheet(s) to include", sheet_names)

    if selected_sheets:
        data_dict = {}
        all_times = []

        for sheet in selected_sheets:
            df = pd.read_excel(uploaded_file, sheet_name=sheet)
            df.columns = ['시간', '생산수량', 'RR RH-1', 'RR RH-2']
            df['시간'] = pd.to_datetime(df['시간'].astype(str))
            df = df.sort_values('시간').reset_index(drop=True)
            df = df[df['생산수량'] > 0]  # filter valid rows for RR RH values
            data_dict[sheet] = df
            all_times.extend(df['시간'].tolist())

        # Create unified timeline
        unified_time = pd.Series(sorted(pd.Series(all_times).unique()))

        # Align data to unified time index
        aligned_data = []
        for sheet, df in data_dict.items():
            df_aligned = pd.DataFrame({'시간': unified_time})
            df_merged = pd.merge(df_aligned, df, on='시간', how='left')
            df_merged['Sheet'] = sheet
            aligned_data.append(df_merged)

        combined_df = pd.concat(aligned_data, ignore_index=True)

        # --- Descriptive Statistics ---
        st.subheader("📌 Descriptive Statistics")
        for sheet in selected_sheets:
            st.markdown(f"#### Sheet: {sheet}")
            df = combined_df[combined_df['Sheet'] == sheet]
            st.dataframe(df[['생산수량', 'RR RH-1', 'RR RH-2']].describe())

        st.markdown("### Global Descriptive Statistics")
        st.dataframe(combined_df[['생산수량', 'RR RH-1', 'RR RH-2']].describe())

        # --- Pairwise Correlation Heatmap ---
        if st.button("Generate Pairwise Corr."):
            st.subheader("🔗 Correlation Heatmap")

            corr_mode = st.radio("View correlation heatmap per sheet or global", ["Per Sheet", "Global"], horizontal=True)

            if corr_mode == "Per Sheet":
                for sheet in selected_sheets:
                    st.markdown(f"#### Sheet: {sheet}")
                    df = combined_df[combined_df['Sheet'] == sheet]
                    fig, ax = plt.subplots()
                    sns.heatmap(df[['생산수량', 'RR RH-1', 'RR RH-2']].corr(), annot=True, cmap='coolwarm', ax=ax)
                    st.pyplot(fig)
            else:
                fig, ax = plt.subplots()
                sns.heatmap(combined_df[['생산수량', 'RR RH-1', 'RR RH-2']].corr(), annot=True, cmap='coolwarm', ax=ax)
                st.pyplot(fig)

        # --- Time-Series Plots ---
        st.subheader("📈 Time-Series Analysis")
        for col in ['RR RH-1', 'RR RH-2']:
            fig = px.line(
                combined_df.dropna(subset=[col]),
                x='시간', y=col, color='Sheet',
                title=f"{col} over Time (Aligned per Sheet)",
                labels={'시간': 'Time', col: col}
            )
            st.plotly_chart(fig, use_container_width=True)

        # --- Comparative Global Summary ---
        st.subheader("📊 Comparative Summary")
        summary = combined_df.groupby('Sheet')[['RR RH-1', 'RR RH-2']].mean().reset_index()
        st.dataframe(summary.set_index('Sheet'))

else:
    st.info("Please upload an Excel file to begin.")
