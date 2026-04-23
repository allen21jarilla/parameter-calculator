import streamlit as st
import pandas as pd
import io


import streamlit as st

# Custom CSS to hide Streamlit elements
hide_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    /* Hides the 'Hosted with Streamlit' button */
    .stAppDeployButton {display: none;}
    [data-testid="stStatusWidget"] {display: none;}
    /* Optional: Reduces top padding for a tighter fit in Looker */
    .block-container {padding-top: 2rem;}
    </style>
"""
st.markdown(hide_style, unsafe_allow_html=True)

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="EEI Corporation Parameter Tool", layout="wide")

st.title("EEI Corporation Parameter Tool")
st.markdown("Input project details, select activities, and calculate total duration, manhours, and equipment needs.")

# --- 1. PROJECT METADATA & SETTINGS ---
st.subheader("1. Project Information & Settings")
col_a, col_b, col_c, col_d = st.columns(4)
with col_a:
    project_name = st.text_input("Project Name", placeholder="e.g., Bridge Construction Phase 1")
with col_b:
    location = st.text_input("Location", placeholder="e.g., Manila City")
with col_c:
    prepared_by = st.text_input("Prepared By", placeholder="e.g., Juan Dela Cruz")
with col_d:
    # New Setting: Workday Hours
    workday_hours = st.number_input("Workday Hours", min_value=1.0, max_value=24.0, value=8.0, step=0.5, help="Affects the conversion of total hours into total days.")

st.divider()

# --- 2. CONNECT TO GOOGLE SHEETS ---
SHEET_ID = "13RdBUd8bolwBn3xHRelZTJWUIT-h9mQ5-ErkOSeWa9I"
TAB_NAME = "DPWH%20(Combined)"  # Web-safe format for spaces
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={TAB_NAME}"

@st.cache_data(ttl=300) 
def load_data(url):
    df = pd.read_csv(url)
    df.columns = df.columns.str.strip()
    return df

try:
    df = load_data(CSV_URL)
except Exception as e:
    st.error(f"Could not load data from the Google Sheet. Please ensure the link is public. Error: {e}")
    st.stop()

# --- CUSTOM COLUMN DISPLAY NAMES ---
# The left side is the Google Sheet name. The right side is what the user sees.
clean_names = {
    "Grouped_Activity": "Activity Description",
    "Output_Unit": "Unit",
    "Avg_Output_per_Hour": "Output/Hr",
    "Avg_Eqpt_Hrs_per_Unit": "Eqpt Factor",
    "Avg_Manhours_per_Unit": "Labor Factor",
    "Total_Duration (Days)": "Total Days",
    "Total_Manhours": "Total Manhours",
    "Total_Eqpt_Hours": "Total Eqpt Hrs",
    "Crew_Breakdown": "Crew Lineup",
    "Primary_Equipment_Required": "Equipment List"
}

# --- 3. USER SELECTION ---
st.subheader("2. Select Project Activities")

if 'Grouped_Activity' not in df.columns:
    st.error("Error: The column 'Grouped_Activity' was not found in your Google Sheet. Please check your column headers.")
    st.stop()

available_activities = df['Grouped_Activity'].dropna().unique().tolist()

selected_activities = st.multiselect(
    "Search and select the activities required for this project:",
    options=available_activities
)

if selected_activities:
    project_df = df[df['Grouped_Activity'].isin(selected_activities)].drop_duplicates(subset=['Grouped_Activity']).copy()
    
    # Add input columns for Quantity and Crews
    project_df.insert(0, 'Quantity', 0.0)
    project_df.insert(1, 'Number of Crews', 1) # Default to 1 crew per activity
    
    st.subheader("3. Input Quantities & Resources")
    st.write("Enter the required quantity and allocate the number of crews for each activity:")
    
    # --- 4. DYNAMIC INPUT GRID ---
    cols_to_fill = ['Avg_Output_per_Hour', 'Avg_Eqpt_Hrs_per_Unit', 'Avg_Manhours_per_Unit']
    for col in cols_to_fill:
        if col in project_df.columns:
            project_df[col] = pd.to_numeric(project_df[col], errors='coerce').fillna(0)

    # Added 'Number of Crews' to the editable grid
    edited_df = st.data_editor(
        project_df[['Quantity', 'Number of Crews', 'Grouped_Activity', 'Output_Unit', 'Avg_Output_per_Hour', 'Avg_Eqpt_Hrs_per_Unit', 'Avg_Manhours_per_Unit']],
        disabled=['Grouped_Activity', 'Output_Unit', 'Avg_Output_per_Hour', 'Avg_Eqpt_Hrs_per_Unit', 'Avg_Manhours_per_Unit'],
        hide_index=True,
        use_container_width=True,
        column_config=clean_names
    )
    
    # --- 5. CALCULATIONS ---
    calc_df = edited_df[edited_df['Quantity'] > 0].copy()
    
    if not calc_df.empty:
        st.divider()
        st.subheader("4. Project Summary & Requirements")
        
        final_df = pd.merge(
            calc_df, 
            df[['Grouped_Activity', 'Primary_Equipment_Required', 'Crew_Breakdown']].drop_duplicates(subset=['Grouped_Activity']), 
            on='Grouped_Activity', 
            how='left'
        )
        
        # New Math Logic: Duration is divided by the Number of Crews
        final_df['Total_Duration (Hours)'] = final_df.apply(
            lambda row: (row['Quantity'] / row['Avg_Output_per_Hour']) / max(1, row['Number of Crews']) if row['Avg_Output_per_Hour'] > 0 else 0, axis=1
        )
        # Days calculated using the custom workday_hours setting
        final_df['Total_Duration (Days)'] = final_df['Total_Duration (Hours)'] / workday_hours
        
        # Manhours and Equipment Hours based on total output requirement
        final_df['Total_Manhours'] = final_df['Quantity'] * final_df['Avg_Manhours_per_Unit']
        final_df['Total_Eqpt_Hours'] = final_df['Quantity'] * final_df['Avg_Eqpt_Hrs_per_Unit']
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Project Duration (Days)", f"{final_df['Total_Duration (Days)'].sum():,.2f}")
        col2.metric("Total Manhours", f"{final_df['Total_Manhours'].sum():,.2f}")
        col3.metric("Total Equipment Hours", f"{final_df['Total_Eqpt_Hours'].sum():,.2f}")
        
        st.write("### Detailed Breakdown")
        st.dataframe(
            final_df[['Grouped_Activity', 'Quantity', 'Output_Unit', 'Number of Crews', 'Total_Duration (Days)', 'Total_Manhours', 'Total_Eqpt_Hours']], 
            hide_index=True, 
            use_container_width=True
            column_config=clean_names
        )
        
        st.write("### Resource Lineup Required")
        colA, colB = st.columns(2)
        with colA:
            st.write("**Crew Breakdown per Activity:**")
            st.dataframe(final_df[['Grouped_Activity', 'Crew_Breakdown']].fillna("None Listed"), hide_index=True, use_container_width=True, column_config=clean_names)
        with colB:
            st.write("**Equipment Required per Activity:**")
            st.dataframe(final_df[['Grouped_Activity', 'Primary_Equipment_Required']].fillna("None Listed"), hide_index=True, use_container_width=True, column_config=clean_names)
    

        # --- 6. EXPORT TO EXCEL ---
        st.divider()
        st.subheader("5. Export Project Plan")
        
        export_df = final_df[['Grouped_Activity', 'Quantity', 'Output_Unit', 'Number of Crews', 'Total_Duration (Days)', 'Total_Manhours', 'Total_Eqpt_Hours', 'Crew_Breakdown', 'Primary_Equipment_Required']]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, index=False, sheet_name='Project Estimate', startrow=6)
            
            workbook = writer.book
            worksheet = writer.sheets['Project Estimate']
            
            header_format = workbook.add_format({'bold': True, 'font_size': 14})
            bold_format = workbook.add_format({'bold': True})
            
            worksheet.write('A1', 'EEI Corporation Parameter Tool - Estimate', header_format)
            worksheet.write('A2', f"Project Name: {project_name if project_name else 'N/A'}", bold_format)
            worksheet.write('A3', f"Location: {location if location else 'N/A'}", bold_format)
            worksheet.write('A4', f"Prepared By: {prepared_by if prepared_by else 'N/A'}", bold_format)
            worksheet.write('A5', f"Workday Hours: {workday_hours} hours/day", bold_format)
            
            for i, col in enumerate(export_df.columns):
                column_len = max(export_df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, column_len)

        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Download Final Project Estimate (Excel)",
            data=excel_data,
            file_name=f"EEI_Estimate_{project_name.replace(' ', '_') if project_name else 'Project'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
