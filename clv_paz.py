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
        Upload the SAP extracted raw data excel file for either CLV or PAZ .
    </p>
    """,
    unsafe_allow_html=True
)

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Show progress bar during file upload
    progress_text = "Uploading file..."
    my_bar = st.progress(0, text=progress_text)

    for percent_complete in range(100):
        time.sleep(0.01)
        my_bar.progress(percent_complete + 1, text=progress_text)

    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(uploaded_file)

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

        # Create dropdown menus
        selected_columns = st.multiselect("Select Columns", column_options)

        if selected_columns:
            # Filter the DataFrame based on selected columns
            selected_columns_df = df[selected_columns]

            # Display the selected columns DataFrame
            st.write(selected_columns_df)

            # Create a dictionary to store the filter values for each selected column
            filter_values = {}

            # Create dropdown menus for filtering each selected column
            for column in selected_columns:
                unique_values = selected_columns_df[column].unique()
                filter_value = st.selectbox(f"Select a value to filter '{column}'", [""] + list(unique_values))
                if filter_value:
                    filter_values[column] = filter_value

            # Filter the DataFrame based on the filter values
            filtered_df = selected_columns_df.copy()
            for column, value in filter_values.items():
                filtered_df = filtered_df[filtered_df[column] == value]

            # Display the filtered DataFrame as summary descriptive statistics
            st.write("Filtered Table Summary:")
            st.write(filtered_df.describe())

            # Filter columns based on data type
            numeric_columns = filtered_df.select_dtypes(include=['number']).columns.tolist()
            categorical_columns = filtered_df.select_dtypes(include=['object']).columns.tolist()

            # Display summary statistics for numeric columns
            if numeric_columns:
                st.write("Summary Statistics for Numeric Columns:")
                st.write(filtered_df[numeric_columns].describe())

            # Display unique values and counts for categorical columns
            if categorical_columns:
                st.write("Unique Values and Counts for Categorical Columns:")
                for column in categorical_columns:
                    st.write(f"Column: {column}")
                    st.write(filtered_df[column].value_counts())

    else:
        st.write("The uploaded Excel file is not in table form.")
