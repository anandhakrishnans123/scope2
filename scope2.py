import streamlit as st
import pandas as pd
import io

# Function to process the Excel file
def process_excel(file):
    # Load the Excel file
    excel_data = pd.ExcelFile(file)
    
    # Define the specified sheets
    specified_sheets = ['SSLL', 'FZE - Office', 'DWC', 'AL ROSTAMANI', 'M&M Global', "ALIA MOH'D TRADING", 'AL SAYEGH', 'TB07', 'GLIF']
    
    # Initialize an empty DataFrame for storing the merged data
    merged_data = pd.DataFrame()
    
    # Loop through the specified sheet names and merge them
    for sheet_name in specified_sheets:
        if sheet_name in excel_data.sheet_names:
            df = pd.read_excel(file, sheet_name=sheet_name)
            merged_data = pd.concat([merged_data, df], ignore_index=True)
    
    # Define the path for the template workbook
    template_workbook_path = r'Electricity-Sample.xlsx'
    
    # Define column mapping
    column_mapping = {
        'Country': 'Country',
        'Facility': 'Office/Factory/Site/\nLocation(Optional)',
        'Energy Consumption': 'Units Consumed (in kWh)',
        'Res_Date': 'Start Date (DD/MM/YYYY Format)'
    }
    
    # Load the template workbook and get the specified sheet
    template_df = pd.read_excel(template_workbook_path, sheet_name=None)
    template_sheet_name = 'Electricity'
    template_data = template_df[template_sheet_name]
    
    # Preserve the first row (header) of the template
    preserved_header = template_data.iloc[:0, :]
    
    # Create a DataFrame with the template columns
    matched_data = pd.DataFrame(columns=template_data.columns)
    
    # Map and copy data based on column_mapping
    for template_col, client_col in column_mapping.items():
        if client_col in merged_data.columns:
            matched_data[template_col] = merged_data[client_col]
        else:
            st.write(f"Column '{client_col}' not found in merged_data")
    
    # Combine header and matched data
    final_data = pd.concat([preserved_header, matched_data], ignore_index=True)
    final_data['CF Standard'] = "IMO"
    final_data['Energy Unit'] = "kWh"
    final_data['Activity Unit'] = "kWh"
    final_data['Energy Type'] = "India"
    final_data['Gas'] = "CO2"
    final_data['Activity'] = 0
    final_data['Res_Date'] = pd.to_datetime(final_data['Res_Date']).dt.date
    
    # Split data into different DataFrames based on 'Facility'
    final_data_SSL = final_data[final_data['Facility'] == 'Shreyas Shipping and Logistics Limited']
    final_data_FZE = final_data[final_data['Facility'] == 'TW Logistics FZE']
    final_data_DWC = final_data[final_data['Facility'].isin(['DWC', 'AL ROSTAMANI', 'M&M Global', 'ALIA MOHD TRADING', 'AL SAYEGH', 'TB07', 'Global Logistics Investments FZE'])]
    
    # Save data to buffers
    buffer_SSL = io.BytesIO()
    buffer_FZE = io.BytesIO()
    buffer_DWC = io.BytesIO()
    
    with pd.ExcelWriter(buffer_SSL, mode='xlsx') as writer:
        final_data_SSL.to_excel(writer, sheet_name='SSL', index=False)
    
    with pd.ExcelWriter(buffer_FZE, mode='xlsx') as writer:
        final_data_FZE.to_excel(writer, sheet_name='FZE', index=False)
    
    with pd.ExcelWriter(buffer_DWC, mode='xlsx') as writer:
        final_data_DWC.to_excel(writer, sheet_name='DWC', index=False)
    
    buffer_SSL.seek(0)
    buffer_FZE.seek(0)
    buffer_DWC.seek(0)
    
    return buffer_SSL, buffer_FZE, buffer_DWC

# Streamlit UI
st.title('Excel Data Processing App')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    buffer_SSL, buffer_FZE, buffer_DWC = process_excel(uploaded_file)
    
    st.download_button(
        label="Download SSL Data",
        data=buffer_SSL,
        file_name="SSL_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.download_button(
        label="Download FZE Data",
        data=buffer_FZE,
        file_name="FZE_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.download_button(
        label="Download DWC Data",
        data=buffer_DWC,
        file_name="DWC_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
