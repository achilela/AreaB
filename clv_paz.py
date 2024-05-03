import streamlit as st
import pandas as pd
import time
from io import BytesIO
import xlrd
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
    workbook = xlrd.open_workbook(file_contents=excel_data.read())
    sheet = workbook.sheet_by_name("Data Base")

    data = []
    for row in range(4, sheet.nrows):
        row_data = []
        for col in range(1, sheet.ncols - 2):
            value = sheet.cell_value(row, col)
            row_data.append(value)
        row_data.append(date.today())
        data.append(row_data)

    columns = [sheet.cell_value(0, col) for col in range(1, sheet.ncols - 2)]
    columns.append("Today's Date")

    df = pd.DataFrame(data, columns=columns)

    if df.columns.nlevels == 1:
        df.dropna(how='all', inplace=True)
        df.dropna(axis=1, how='all', inplace=True)
        df.reset_index(drop=True, inplace=True)

        if "SECE STATUS" in df.columns:
            df["SECE STATUS"].fillna("Non-SCE", inplace=True)

        st.write("Cleaned Table:")
        st.write(df.head(10))

        column_options = df.columns.tolist()
        main_column = st.selectbox("Select the main column", column_options)

        filter_columns = [col for col in column_options if col != main_column]
        selected_columns = st.multiselect("Select columns to filter", filter_columns)

        if selected_columns:
            filtered_df = df[[main_column] + selected_columns]

            filter_values = {}
            for column in selected_columns:
                unique_values = filtered_df[column].unique()
                filter_values[column] = st.multiselect(f"Select values to filter '{column}'", unique_values)

            for column, values in filter_values.items():
                if values:
                    filtered_df = filtered_df[filtered_df[column].isin(values)]

            if "SECE STATUS" in selected_columns:
                sce_status = st.selectbox("Filter SECE STATUS", ["All", "SCE", "Non-SCE"])
                if sce_status == "SCE":
                    filtered_df = filtered_df[filtered_df["SECE STATUS"] == "SCE"]
                elif sce_status == "Non-SCE":
                    filtered_df = filtered_df[filtered_df["SECE STATUS"] == "Non-SCE"]

            grouped_data = filtered_df.groupby([main_column] + selected_columns).size().reset_index(name='Count')
            pivot_table = grouped_data.pivot_table(index=main_column, columns=selected_columns, values='Count', fill_value=0)
            pivot_table["Grand Total"] = pivot_table.sum(axis=1)
            pivot_table.loc["Total"] = pivot_table.sum()

            st.write("Filtered Table:")
            st.write(pivot_table)

            # Display summary statistics for numeric columns
            numeric_columns = pivot_table.select_dtypes(include=['number']).columns
            if len(numeric_columns) > 0:
                st.write("Summary Statistics for Numeric Columns:")
                st.write(pivot_table[numeric_columns].describe())

            # Display unique values and counts for categorical columns
            categorical_columns = pivot_table.select_dtypes(include=['object']).columns
            if len(categorical_columns) > 0:
                st.write("Unique Values and Counts for Categorical Columns:")
                for column in categorical_columns:
                    st.write(f"{column}:")
                    st.write(pivot_table[column].value_counts())

    else:
        st.write("The uploaded Excel file is not in table form.")
