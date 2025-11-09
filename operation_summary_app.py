import streamlit as st
import pandas as pd
import re

def extract_operation_summary(file_content, file_name):
    lines = file_content.split('\n')
    well_name = ""
    rig_name = ""
    last_24_summary = ""
    next_24_forecast = ""
    for i, line in enumerate(lines):
        if "WELL NAME" in line.upper():
            cells = line.split('|')
            for j, cell in enumerate(cells):
                if "WELL NAME" in cell.upper() and j + 2 < len(cells):
                    well_name = cells[j + 2].strip()
                    break
        if "RIG NAME" in line.upper():
            cells = line.split('|')
            for j, cell in enumerate(cells):
                if "RIG NAME" in cell.upper() and j + 2 < len(cells):
                    rig_name = cells[j + 2].strip()
                    break
        if "LAST 24 SUMMARY" in line.upper():
            cells = line.split('|')
            for j, cell in enumerate(cells):
                if "LAST 24 SUMMARY" in cell.upper() and j + 2 < len(cells):
                    last_24_summary = cells[j + 2].strip()
                    last_24_summary = last_24_summary.replace(':-', '').strip()
                    break
        if "NEXT 24 FORECAST" in line.upper():
            cells = line.split('|')
            for j, cell in enumerate(cells):
                if "NEXT 24 FORECAST" in cell.upper() and j + 2 < len(cells):
                    next_24_forecast = cells[j + 2].strip()
                    next_24_forecast = next_24_forecast.replace(':', '').strip()
                    break
    return {
        'file_name': file_name,
        'well_name': well_name,
        'rig_name': rig_name,
        'last_24_summary': last_24_summary,
        'next_24_forecast': next_24_forecast
    }

def main():
    st.set_page_config(page_title="Drilling Reports Summary", layout="wide")
    st.title("🏗️ Drilling & Workover Reports Dashboard")
    st.markdown("### Operation Summary Extraction")

    sample_files = {
        "GANNA-27 DWOR#10 (TRUST-6) DATED 9-11-2025.xlsx": "Example content..."
    }
    results = []
    for file_name, file_content in sample_files.items():
        try:
            summary = extract_operation_summary(file_content, file_name)
            results.append(summary)
        except Exception as e:
            st.error(f"Error processing {file_name}: {str(e)}")

    if results:
        df = pd.DataFrame(results)
        df = df[['file_name', 'well_name', 'rig_name', 'last_24_summary', 'next_24_forecast']]
        st.subheader("📊 Operation Summary Overview")
        st.dataframe(df, use_container_width=True)
        st.subheader("📋 Detailed Operation Summaries")
        for result in results:
            with st.expander(f"{result['well_name']} - {result['rig_name']} ({result['file_name']})"):
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**LAST 24 HOURS SUMMARY**")
                    st.info(result['last_24_summary'])
                with col2:
                    st.markdown("**NEXT 24 HOURS FORECAST**")
                    st.success(result['next_24_forecast'])
    else:
        st.warning("No operation summaries found in the provided files.")

if __name__ == "__main__":
    main()