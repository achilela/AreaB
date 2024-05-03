import streamlit as st
import pandas as pd
import time

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
    df['Today\'s Date'] = pd.Timestamp.now().date()

    # Format columns
    if "Order" in df.columns:
        df["Order"] = df["Order"].astype(int)
    if "Last Insp/" in df.columns:
        df["Last Insp/"] = pd.to_datetime(df["Last Insp/"]).dt.date
    if "Next Insp/" in df.columns:
        df["Next Insp/"] = pd.to_datetime(df["Next Insp/"]).dt.date
    if "Due Date" in df.columns:
        df["Due Date"] = pd.to_datetime(df["Due Date"]).dt.date
    if "Compl Date" in df.columns:
        df["Compl Date"] = pd.to_datetime(df["Compl Date"]).dt.date
    if "Year" in df.columns:
        df["Year"] = df["Year"].astype(str).str[:4]

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
        st.write(df.head(10))

        # Get unique values for dropdown menus
        column_options = df.columns.tolist()

        # Create dropdown menus
        selected_columns = st.multiselect("Select Columns", column_options)

        if selected_columns:
            # Filter the DataFrame based on selected columns
            selected_columns_df = df[selected_columns]

            # Display the selected columns DataFrame
            st.write(selected_columns_df)

            # Create a dictionary to store the filter values for each selected column
            filter_values = {}

            # Create multiselect dropdowns for filtering each selected column
            for column in selected_columns:
                unique_values = selected_columns_df[column].unique()
                filter_values[column] = st.multiselect(f"Select values to filter '{column}'", unique_values)

            # Filter the DataFrame based on the filter values
            filtered_df = selected_columns_df.copy()
            for column, values in filter_values.items():
                if values:
                    filtered_df = filtered_df[filtered_df[column].isin(values)]

            # Display the filtered DataFrame as a comprehensive table
            st.write("Filtered Table:")
            filtered_table = filtered_df.pivot_table(index=selected_columns[0], columns=selected_columns[1:], aggfunc='size', fill_value=0)
            filtered_table["Grand Total"] = filtered_table.sum(axis=1)
            filtered_table.loc["Total"] = filtered_table.sum()
            filtered_table.loc["Grand Total"] = filtered_table.sum(axis=1)
            st.write(filtered_table)

    else:
        st.write("The uploaded Excel file is not in table form.")
