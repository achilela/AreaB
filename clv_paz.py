import streamlit as st
import pandas as pd
import time
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import date

# Set page title
st.set_page_config(page_title="Topsides Plant Maintenance Data Analysis")

# Add a title and description
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

# File upload in the sidebar
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Show progress bar during file upload
    progress_text = "Uploading file..."
    my_bar = st.progress(0, text=progress_text)

    for percent_complete in range(100):
        time.sleep(0.01)
        my_bar.progress(percent_complete + 1, text=progress_text)

    # Rest of the code...

    # Save the modified workbook to a BytesIO object
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Read the modified Excel file into a pandas DataFrame
    df = pd.read_excel(output)

    # Check if the Excel file is already in table form
    if df.columns.nlevels == 1:
        # Remove rows with missing data
        df.dropna(how='all', inplace=True)

        # Remove columns with missing data
        df.dropna(axis=1, how='all', inplace=True)

        # Reset index
        df.reset_index(drop=True, inplace=True)

        # Replace NaN values in the "SECE STATUS" column with "Non-SCE"
        if "SECE STATUS" in df.columns:
            df["SECE STATUS"].fillna("Non-SCE", inplace=True)

        # Display the first 10 rows of the cleaned table
        st.write("Cleaned Table:")
        st.write(df.head(10))

        # Get unique values for dropdown menus
        column_options = df.columns.tolist()

        # Create radio button to select the main column
        main_column = st.radio("Select the main column", column_options)

        # Create multiselect dropdowns for filtering other columns
        filter_columns = [col for col in column_options if col != main_column]
        selected_columns = st.multiselect("Select columns to filter", filter_columns)

        if selected_columns:
            # Filter the DataFrame based on selected columns
            filtered_df = df[[main_column] + selected_columns]

            # Create a dictionary to store the filter values for each selected column
            filter_values = {}

            # Create multiselect dropdowns for filtering each selected column
            for column in selected_columns:
                unique_values = filtered_df[column].unique()
                filter_values[column] = st.multiselect(f"Select values to filter '{column}'", unique_values)

            # Filter the DataFrame based on the filter values
            for column, values in filter_values.items():
                if values:
                    filtered_df = filtered_df[filtered_df[column].isin(values)]

            # Group the data by the main column and selected columns
            grouped_data = filtered_df.groupby([main_column] + selected_columns).size().reset_index(name='Count')

            # Pivot the grouped data to create a table with the main column as rows and selected columns as columns
            pivot_table = grouped_data.pivot_table(index=main_column, columns=selected_columns, values='Count', fill_value=0)

            # Add a "Grand Total" column to the pivot table
            pivot_table["Grand Total"] = pivot_table.sum(axis=1)

            # Add a "Total" row to the pivot table
            pivot_table.loc["Total"] = pivot_table.sum()

            # Display the pivot table
            st.write("Filtered Table:")
            st.write(pivot_table)

    else:
        st.write("The uploaded Excel file is not in table form.")
