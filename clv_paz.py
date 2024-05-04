import streamlit as st
import pandas as pd
import time
import numpy as np
from datetime import datetime, timedelta

# Set page title
st.set_page_config(page_title="Area B - Methods Engineering Data Analysis")

# Add a title and description
st.markdown(
    """
    <h1 style='text-align: center; font-size: 36px; color: #2F80ED;'>
        Area B - Methods Engineering Data Analysis
    </h1>

    <p style='text-align: center; font-size: 18px;'>
        Upload the excel file raw data from SAP reports.
    </p>
    """,
    unsafe_allow_html=True
)

# Add a sidebar
sidebar = st.sidebar

# Add a logo and description to the sidebar
logo_path_computer = "/home/atalibamiguel/dataCamp PythonExcel/CLOV.png"
logo_path_github = "https://raw.githubusercontent.com/achilela/AreaB/main/CLOV.png"

sidebar.markdown(
    """
    <style>
    .logo-description {
        style='text-align: center; 
        color: #2F80ED;
        font-size: 20px;
       
        margin-top: 10px;
        font-family: Arial, sans-serif;
        line-height: 1.5;
    }
    </style>
  

    <div class="logo-description">
        Area B - CLOV & PAZFLOR
    </div>
    """,
    unsafe_allow_html=True
)

try:
    sidebar.image(logo_path_computer, width=250)
except FileNotFoundError:
    sidebar.image(logo_path_github, width=150)



# File upload in the sidebar
uploaded_file = sidebar.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Show progress bar during file upload
    progress_text = "Uploading file..."
    my_bar = st.progress(0, text=progress_text)

    for percent_complete in range(100):
        time.sleep(0.01)
        my_bar.progress(percent_complete + 1, text=progress_text)

    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(uploaded_file)

    # Add a new column 'Today's Date' with today's date
    today_date = pd.Timestamp.now().date()
    df['Today\'s Date'] = today_date.strftime('%m/%d/%Y')

    # Format columns
    if "Order" in df.columns:
        df["Order"] = df["Order"].astype(int)
    if "Last Insp/" in df.columns:
        df["Last Insp/"] = pd.to_datetime(df["Last Insp/"]).dt.strftime('%m/%d/%Y')
    if "Next Insp/" in df.columns:
        df["Next Insp/"] = pd.to_datetime(df["Next Insp/"]).dt.strftime('%m/%d/%Y')
    if "Due Date" in df.columns:
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors='coerce').dt.strftime('%m/%d/%Y')
    if "Compl Date" in df.columns:
        df["Compl Date"] = pd.to_datetime(df["Compl Date"]).dt.strftime('%m/%d/%Y')
    if "Year" in df.columns:
        df["Year"] = df["Year"].astype(str).str[:4]

    # Calculate delay
    if "Delay" in df.columns and "Due Date" in df.columns:
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors='coerce')
        df["Delay"] = np.where(today_date - df["Due Date"].dt.date > pd.Timedelta(days=1095), "> 3 Yrs",
                               np.where(today_date - df["Due Date"].dt.date > pd.Timedelta(days=730), "2 Yrs < x <3 Yrs",
                                      np.where(today_date - df["Due Date"].dt.date > pd.Timedelta(days=365), "1 Yrs < x <2 Yrs",
                                               np.where(today_date - df["Due Date"].dt.date > pd.Timedelta(days=182), "6 Months < x <1 Yrs",
                                                        "< 6 Months"))))

    # Calculate backlog
    if "Backlog" not in df.columns:
        df["Backlog"] = np.nan
        df["Backlog Date"] = pd.NaT
        df["Backlog Days"] = np.nan

    if "Due Date" in df.columns and "Order Status" in df.columns:
        # Convert "Due Date" to datetime dtype, handling invalid values
        df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce")
        today_date = pd.Timestamp(today_date)

        # Calculate the sum of "Due Date" and 28 days
        due_date_plus_28 = df["Due Date"] + pd.Timedelta(days=28)

        # Compare the sum with today's date
        is_backlog = due_date_plus_28.dt.date == today_date

        # Assign "Yes" or "No" to the "Backlog" column based on the condition
        df["Backlog"] = np.where((df["Order Status"].isin(["WIP", "HOLD", "WREL"])) & is_backlog, "Yes", "No")

        # Assign the backlog date when the equipment enters into backlog
        df.loc[(df["Order Status"].isin(["WIP", "HOLD", "WREL"])) & is_backlog, "Backlog Date"] = due_date_plus_28.dt.strftime('%m/%d/%Y')

        # Calculate the cumulative days the equipment has been in backlog
        df["Backlog Days"] = (today_date - pd.to_datetime(df["Backlog Date"], errors="coerce")).dt.days
        df["Backlog Days"] = df["Backlog Days"].fillna(0).astype(int)

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
        st.write("Completed Pre-Processed Table")
        st.write(df.head(5))

        # Get unique values for dropdown menus
        column_options = df.columns.tolist()

        # Create dropdown menu in the sidebar
        selected_columns = sidebar.multiselect("Select Columns", column_options)

        if selected_columns:
            # Filter the DataFrame based on selected columns
            selected_columns_df = df[selected_columns]

            # Create a dictionary to store the filter values for each selected column
            filter_values = {}

            # Create multiselect dropdowns for filtering each selected column in the sidebar
            for column in selected_columns:
                unique_values = selected_columns_df[column].unique()
                filter_values[column] = sidebar.multiselect(f"Select values to filter '{column}'", unique_values)

            # Filter the DataFrame based on the filter values
            filtered_df = selected_columns_df.copy()
            for column, values in filter_values.items():
                if values:
                    filtered_df = filtered_df[filtered_df[column].isin(values)]

            # Display the filtered DataFrame as a comprehensive table
            st.write("Filtered Table:")
            filtered_table = filtered_df.pivot_table(index=selected_columns[0], columns=selected_columns[1:], aggfunc='size', fill_value=0)

            # Convert filtered_table to a DataFrame
            filtered_table = pd.DataFrame(filtered_table)

            filtered_table["Grand Total"] = filtered_table.sum(axis=1)

            # Calculate the total for each column
            column_totals = filtered_table.sum()

            # Add the total row to the DataFrame
            filtered_table.loc["Total"] = column_totals

            filtered_table.loc["Grand Total"] = filtered_table.sum(axis=1)

            # Style the filtered table
            styled_table = filtered_table.style.format('{:,.0f}') \
                .set_properties(**{'text-align': 'center'}) \
                .set_table_styles([
                    {'selector': 'th', 'props': [('background-color', '#2F80ED'), ('color', 'white')]},
                    {'selector': 'td', 'props': [('border', '1px solid #ddd')]},
                    {'selector': 'tr:nth-child(even)', 'props': [('background-color', '#f2f2f2')]},
                    {'selector': 'tr:hover', 'props': [('background-color', '#e6e6e6')]}
                ])

            st.write(styled_table)

    else:
        st.write("The uploaded Excel file is not in table form.")
