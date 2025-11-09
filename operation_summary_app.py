# drilling_reports_upload.py
import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

def extract_operation_summary_from_excel(uploaded_file):
    """
    Extract operation summary, well name, and rig name from uploaded Excel file
    """
    try:
        # Read the Excel file
        wb = load_workbook(filename=io.BytesIO(uploaded_file.read()), data_only=True)
        sheet = wb.active
        
        # Convert sheet to DataFrame for easier processing
        data = sheet.values
        columns = next(data)
        df = pd.DataFrame(data, columns=columns)
        
        # Initialize variables
        well_name = ""
        rig_name = ""
        last_24_summary = ""
        next_24_forecast = ""
        
        # Convert DataFrame to string representation for searching
        df_string = df.to_string()
        
        # Search for well name
        for row in sheet.iter_rows(values_only=True):
            row_str = str(row)
            if "WELL NAME" in str(row).upper():
                for cell in row:
                    if cell and "WELL NAME" not in str(cell).upper():
                        well_name = str(cell)
                        break
                # Also check neighboring cells
                for i, cell in enumerate(row):
                    if cell and "WELL NAME" in str(cell).upper() and i + 1 < len(row):
                        if row[i + 1] and "WELL NAME" not in str(row[i + 1]).upper():
                            well_name = str(row[i + 1])
                            break
        
        # Search for rig name
        for row in sheet.iter_rows(values_only=True):
            row_str = str(row)
            if "RIG NAME" in str(row).upper():
                for cell in row:
                    if cell and "RIG NAME" not in str(cell).upper():
                        rig_name = str(cell)
                        break
                # Also check neighboring cells
                for i, cell in enumerate(row):
                    if cell and "RIG NAME" in str(cell).upper() and i + 1 < len(row):
                        if row[i + 1] and "RIG NAME" not in str(row[i + 1]).upper():
                            rig_name = str(row[i + 1])
                            break
        
        # Search for LAST 24 SUMMARY
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "LAST 24 SUMMARY" in str(cell).upper():
                    # Get the summary from the next cell
                    if i + 1 < len(row) and row[i + 1]:
                        last_24_summary = str(row[i + 1])
                        break
                    # If not in next cell, try to find in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "LAST 24 SUMMARY" not in str(cell2).upper() and cell2:
                            last_24_summary = str(cell2)
                            break
        
        # Search for NEXT 24 FORECAST
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "NEXT 24 FORECAST" in str(cell).upper():
                    # Get the forecast from the next cell
                    if i + 1 < len(row) and row[i + 1]:
                        next_24_forecast = str(row[i + 1])
                        break
                    # If not in next cell, try to find in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "NEXT 24 FORECAST" not in str(cell2).upper() and cell2:
                            next_24_forecast = str(cell2)
                            break
        
        # Clean up the extracted data
        well_name = well_name.replace(':-', '').replace(':', '').strip() if well_name else "Not Found"
        rig_name = rig_name.replace(':-', '').replace(':', '').strip() if rig_name else "Not Found"
        last_24_summary = last_24_summary.replace(':-', '').replace(':', '').strip() if last_24_summary else "Not Found"
        next_24_forecast = next_24_forecast.replace(':-', '').replace(':', '').strip() if next_24_forecast else "Not Found"
        
        return {
            'file_name': uploaded_file.name,
            'well_name': well_name,
            'rig_name': rig_name,
            'last_24_summary': last_24_summary,
            'next_24_forecast': next_24_forecast
        }
        
    except Exception as e:
        st.error(f"Error processing file {uploaded_file.name}: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="Drilling Reports Analyzer", 
        layout="wide",
        page_icon="🏗️"
    )
    
    st.title("🏗️ Drilling Reports Analyzer")
    st.markdown("### Upload Excel files to extract operation summaries")
    
    # Sidebar with instructions
    st.sidebar.title("📋 Instructions")
    st.sidebar.markdown("""
    **How to Use:**
    1. Upload one or more Excel drilling report files
    2. The app will automatically extract:
       - Well Name
       - Rig Name  
       - Last 24 Hours Summary
       - Next 24 Hours Forecast
    3. View results in the summary table
    4. Click on each well for detailed view
    
    **File Requirements:**
    - Excel files (.xlsx)
    - Should contain "WELL NAME", "RIG NAME", 
      "LAST 24 SUMMARY", "NEXT 24 FORECAST" fields
    """)
    
    # File upload section
    st.subheader("📤 Upload Drilling Report Files")
    uploaded_files = st.file_uploader(
        "Choose Excel files",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Upload one or more drilling report Excel files"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
        
        # Process all uploaded files
        all_summaries = []
        
        with st.spinner("Processing files..."):
            for uploaded_file in uploaded_files:
                summary = extract_operation_summary_from_excel(uploaded_file)
                if summary:
                    all_summaries.append(summary)
        
        if all_summaries:
            # Create summary DataFrame
            df = pd.DataFrame(all_summaries)
            
            # Display summary table
            st.subheader("📊 Operation Summary Table")
            st.dataframe(
                df[['well_name', 'rig_name', 'last_24_summary', 'next_24_forecast']],
                use_container_width=True,
                height=400
            )
            
            # Display statistics
            st.subheader("📈 Summary Statistics")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Files", len(all_summaries))
            with col2:
                unique_wells = len(set([s['well_name'] for s in all_summaries]))
                st.metric("Unique Wells", unique_wells)
            with col3:
                unique_rigs = len(set([s['rig_name'] for s in all_summaries]))
                st.metric("Unique Rigs", unique_rigs)
            with col4:
                successful_extractions = len([s for s in all_summaries if s['last_24_summary'] != "Not Found"])
                st.metric("Successful Extractions", successful_extractions)
            
            # Detailed views for each file
            st.subheader("🔍 Detailed Operation Summaries")
            
            for i, summary in enumerate(all_summaries):
                with st.expander(f"📄 {summary['well_name']} - {summary['rig_name']} ({summary['file_name']})", expanded=False):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("**🕐 LAST 24 HOURS SUMMARY**")
                        if summary['last_24_summary'] != "Not Found":
                            st.info(summary['last_24_summary'])
                        else:
                            st.warning("Last 24 hours summary not found in file")
                    
                    with col2:
                        st.markdown("**🔮 NEXT 24 HOURS FORECAST**")
                        if summary['next_24_forecast'] != "Not Found":
                            st.success(summary['next_24_forecast'])
                        else:
                            st.warning("Next 24 hours forecast not found in file")
                    
                    # File info
                    st.markdown("**📋 File Information**")
                    info_col1, info_col2, info_col3 = st.columns(3)
                    with info_col1:
                        st.write(f"**Well Name:** {summary['well_name']}")
                    with info_col2:
                        st.write(f"**Rig Name:** {summary['rig_name']}")
                    with info_col3:
                        st.write(f"**File:** {summary['file_name']}")
                    
                    st.markdown("---")
            
            # Download button for the summary table
            st.subheader("💾 Download Results")
            csv = df.to_csv(index=False)
            st.download_button(
                label="Download Summary as CSV",
                data=csv,
                file_name="drilling_reports_summary.csv",
                mime="text/csv",
                help="Download the operation summaries as a CSV file"
            )
            
        else:
            st.error("❌ No valid operation summaries could be extracted from the uploaded files.")
    
    else:
        # Show sample structure when no files are uploaded
        st.info("👆 Please upload Excel drilling report files to get started")
        
        # Sample data structure
        st.subheader("📋 Expected File Structure")
        st.markdown("""
        Your Excel files should contain the following information in a table format:
        
        | WELL NAME | RIG NAME | OPERATION SUMMARY | LAST 24 SUMMARY | NEXT 24 FORECAST |
        |-----------|----------|-------------------|-----------------|------------------|
        | ABRAR-84  | EDC-11   | RUN 7" LINER      | [Summary text]  | [Forecast text]  |
        
        **Key fields the app looks for:**
        - `WELL NAME`
        - `RIG NAME` 
        - `LAST 24 SUMMARY` or `LAST 24 SUMMARY:-`
        - `NEXT 24 FORECAST` or `NEXT 24 FORECAST:`
        """)

if __name__ == "__main__":
    main()
