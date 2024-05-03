import streamlit as st
import pandas as pd
import time
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import date
 
st.set_page_config(page_title="Topsides Plant Maintenance Data Analysis")
 
st.markdown(
    """
    <h1 style='text-align: center; font-size: 36px; color: #2F80ED;'>
        Topsides Plant Maintenance Data Analysis
    </h1>
    
    <p style='text-align: center; font-size: 18px;'>
        Upload an Excel file and use the dropdown menus to filter and analyze the data.
    </p>
    """,
    unsafe_allow_html=True
)
 
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
 
if uploaded_file is not None:
    progress_text = "Uploading file..."
    my_bar = st.progress(0, text=progress_text)
 
    for percent_complete in range(100):
        time.sleep(0.01)
        my_bar.progress(percent_complete + 1, text=progress_text)
 
    excel_data = BytesIO(uploaded_file.getvalue())
    workbook = load_workbook(excel_data)
    sheet = workbook["Data Base"]
 
    sheet.delete_cols(1, 1)
    sheet.delete_cols(sheet.max_column - 2, 3)
    sheet.delete_rows(1, 4)
 
    sheet.cell(row=1, column=sheet.max_column + 1, value="Today's Date")
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=sheet.max_column, value=date.today())
 
    table_name = "MainTable"
    table_range = f"A1:{chr(ord('A') + sheet.max_column - 1)}{sheet.max_row}"
    table = Table(displayName=table_name, ref=table_range)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style
    sheet.add_table(table)
 
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
 
    df = pd.read_excel(output)
 
    if df.columns.nlevels == 1:
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        df.reset_index(drop=True, inplace=True)
 
        if "SECE STATUS" in df.columns:
            df["SECE STATUS"].fillna("Non-SCE", inplace=True)
 
        st.write("Cleaned Table:")
        st.write(df.head(10))
 
        column_options = df.columns.tolist()
        main_column = st.radio("Select the main column", column_options)
 
        filter_columns = [col for col in column_options if col != main_column]
        selected_columns = st.multiselect("Select columns to filter", filter_columns)
 
        if selected_columns:
            filtered_df = df[[main_column] + selected_columns]
 
            filter_values = {}
            for column in selected_columns:
                unique_values = filtered_df[column].unique()
                filter_values[column] = st.multiselect(f"Select values to filter '{column}'", unique_values)
