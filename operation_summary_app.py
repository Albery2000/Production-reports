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
        
        # Initialize variables
        well_name = ""
        rig_name = ""
        last_24_summary = ""
        next_24_forecast = ""
        
        # Search for well name
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "WELL NAME" in str(cell).upper():
                    # Get the well name from adjacent cells
                    if i + 1 < len(row) and row[i + 1]:
                        well_name = str(row[i + 1])
                        break
                    # Also check other cells in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "WELL NAME" not in str(cell2).upper() and cell2:
                            well_name = str(cell2)
                            break
                    break
        
        # Search for rig name
        for row in sheet.iter_rows(values_only=True):
            for i, cell in enumerate(row):
                if cell and "RIG NAME" in str(cell).upper():
                    # Get the rig name from adjacent cells
                    if i + 1 < len(row) and row[i + 1]:
                        rig_name = str(row[i + 1])
                        break
                    # Also check other cells in the row
                    for j, cell2 in enumerate(row):
                        if cell2 and "RIG NAME" not in str(cell2).upper() and cell2:
                            rig_name = str(cell2)
                            break
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

def create_operation_summary_display(last_24_summary, next_24_forecast):
    """
    Create a formatted operation summary for display
    """
    if last_24_summary == "Not Found" and next_24_forecast == "Not Found":
        return "❌ No operation summary found in file"
    
    summary_html = f"""
    <div style="padding: 10px; border-radius: 5px; background-color: #f0f8ff;">
        <div style="margin-bottom: 15px;">
            <h4 style="margin: 0; color: #1f77b4; font-size: 14px;">📅 LAST 24 HOURS:</h4>
            <p style="margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">{last_24_summary if last_24_summary != 'Not Found' else 'No data available'}</p>
        </div>
        <div>
            <h4 style="margin: 0; color: #2ca02c; font-size: 14px;">🔮 NEXT 24 HOURS:</h4>
            <p style="margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">{next_24_forecast if next_24_forecast != 'Not Found' else 'No data available'}</p>
        </div>
    </div>
    """
    return summary_html

def main():
    st.set_page_config(
        page_title="Drilling Reports Analyzer", 
        layout="wide",
        page_icon="🏗️"
    )
    
    st.title("🏗️ Drilling Operations Dashboard")
    st.markdown("### Upload Excel files to extract operation summaries")
    
    # Sidebar with instructions
    st.sidebar.title("📋 Instructions")
    st.sidebar.markdown("""
    **How to Use:**
    1. Upload Excel drilling report files
    2. View operation summaries in the table
    3. Click on rows for detailed information
    
    **The app extracts:**
    - 🔧 Rig & Well information
    - 📅 Last 24 hours activities
    - 🔮 Next 24 hours plans
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
        
        with st.spinner("🔍 Analyzing drilling reports..."):
            for uploaded_file in uploaded_files:
                summary = extract_operation_summary_from_excel(uploaded_file)
                if summary:
                    all_summaries.append(summary)
        
        if all_summaries:
            # Create the main summary table with two columns
            st.subheader("📊 Operations Summary")
            st.markdown("### Current Drilling Operations Overview")
            
            # Display statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📁 Total Files", len(all_summaries))
            with col2:
                unique_wells = len(set([s['well_name'] for s in all_summaries if s['well_name'] != "Not Found"]))
                st.metric("🛢️ Active Wells", unique_wells)
            with col3:
                unique_rigs = len(set([s['rig_name'] for s in all_summaries if s['rig_name'] != "Not Found"]))
                st.metric("🔧 Active Rigs", unique_rigs)
            
            # Create the main two-column display
            for i, summary in enumerate(all_summaries):
                # Create a container for each row
                with st.container():
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        # Rig and Well information
                        st.markdown(f"""
                        <div style="padding: 15px; background-color: #f8f9fa; border-radius: 10px; border-left: 4px solid #007bff;">
                            <h3 style="margin: 0 0 10px 0; color: #2c3e50;">{summary['well_name'] if summary['well_name'] != 'Not Found' else 'Unknown Well'}</h3>
                            <p style="margin: 0; color: #7f8c8d; font-size: 14px;">
                                <strong>Rig:</strong> {summary['rig_name'] if summary['rig_name'] != 'Not Found' else 'Unknown Rig'}
                            </p>
                            <p style="margin: 5px 0 0 0; color: #95a5a6; font-size: 12px;">
                                File: {summary['file_name']}
                            </p>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        # Operation summary
                        operation_display = create_operation_summary_display(
                            summary['last_24_summary'], 
                            summary['next_24_forecast']
                        )
                        st.markdown(operation_display, unsafe_allow_html=True)
                    
                    # Add some spacing between entries
                    st.markdown("<br>", unsafe_allow_html=True)
            
            # Detailed expandable sections
            st.subheader("🔍 Detailed Operation Views")
            st.markdown("Click on any operation below to see full details:")
            
            for i, summary in enumerate(all_summaries):
                with st.expander(f"🔧 {summary['well_name']} - {summary['rig_name']} | 📄 {summary['file_name']}", expanded=False):
                    
                    # Create two columns for detailed view
                    detail_col1, detail_col2 = st.columns(2)
                    
                    with detail_col1:
                        st.markdown("### 📋 Well & Rig Information")
                        st.info(f"""
                        **Well Name:** {summary['well_name'] if summary['well_name'] != 'Not Found' else '❌ Not found'}
                        \n**Rig Name:** {summary['rig_name'] if summary['rig_name'] != 'Not Found' else '❌ Not found'}
                        \n**Source File:** {summary['file_name']}
                        """)
                    
                    with detail_col2:
                        st.markdown("### 📊 Operation Status")
                        if summary['last_24_summary'] != "Not Found":
                            st.success("✅ Operations data successfully extracted")
                        else:
                            st.warning("⚠️ Limited operation data available")
                    
                    # Operation details in full width
                    st.markdown("### 🕐 Operation Details")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("#### 📅 Last 24 Hours")
                        if summary['last_24_summary'] != "Not Found":
                            st.info(summary['last_24_summary'])
                        else:
                            st.warning("No last 24 hours summary found")
                    
                    with col2:
                        st.markdown("#### 🔮 Next 24 Hours")
                        if summary['next_24_forecast'] != "Not Found":
                            st.success(summary['next_24_forecast'])
                        else:
                            st.warning("No next 24 hours forecast found")
                    
                    st.markdown("---")
            
            # Download section
            st.subheader("💾 Export Data")
            
            # Prepare data for download
            download_data = []
            for summary in all_summaries:
                download_data.append({
                    'Well Name': summary['well_name'],
                    'Rig Name': summary['rig_name'],
                    'Last 24 Hours Summary': summary['last_24_summary'],
                    'Next 24 Hours Forecast': summary['next_24_forecast'],
                    'Source File': summary['file_name']
                })
            
            download_df = pd.DataFrame(download_data)
            csv = download_df.to_csv(index=False)
            
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="📥 Download Summary as CSV",
                    data=csv,
                    file_name="drilling_operations_summary.csv",
                    mime="text/csv",
                    help="Download all operation summaries as a CSV file"
                )
            with col2:
                st.download_button(
                    label="📥 Download Summary as Excel",
                    data=download_df.to_csv(index=False),
                    file_name="drilling_operations_summary.xlsx",
                    mime="application/vnd.ms-excel",
                    help="Download all operation summaries as an Excel file"
                )
            
        else:
            st.error("❌ No valid operation summaries could be extracted from the uploaded files.")
            st.info("💡 Please make sure your Excel files contain the required fields: WELL NAME, RIG NAME, LAST 24 SUMMARY, and NEXT 24 FORECAST")
    
    else:
        # Show sample when no files uploaded
        st.info("👆 Please upload Excel drilling report files to get started")
        
        # Show sample output
        st.subheader("🎯 What You'll See")
        st.markdown("""
        After uploading files, you'll see a clean overview like this:
        """)
        
        # Sample preview
        sample_col1, sample_col2 = st.columns([1, 2])
        
        with sample_col1:
            st.markdown("""
            <div style="padding: 15px; background-color: #f8f9fa; border-radius: 10px; border-left: 4px solid #007bff;">
                <h3 style="margin: 0 0 10px 0; color: #2c3e50;">ABRAR-84</h3>
                <p style="margin: 0; color: #7f8c8d; font-size: 14px;">
                    <strong>Rig:</strong> EDC-11
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        with sample_col2:
            st.markdown("""
            <div style="padding: 10px; border-radius: 5px; background-color: #f0f8ff;">
                <div style="margin-bottom: 15px;">
                    <h4 style="margin: 0; color: #1f77b4; font-size: 14px;">📅 LAST 24 HOURS:</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">Running 7" liner operations, completed logging...</p>
                </div>
                <div>
                    <h4 style="margin: 0; color: #2ca02c; font-size: 14px;">🔮 NEXT 24 HOURS:</h4>
                    <p style="margin: 5px 0 0 0; font-size: 13px; line-height: 1.4;">Continue liner operations, prepare for cement job...</p>
                </div>
            </div>
            """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
