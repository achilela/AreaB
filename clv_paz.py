import streamlit as st
import pandas as pd
import pygwalker as pyg
from datetime import date
import openpyxl

def preprocess_data(df):
    # Delete the first and last 2 columns
    df = df.iloc[:, 2:-2]
    
    # Add a "Today Date" column with the TODAY() formula
    df['Today Date'] = date.today()
    
    return df

def filter_data(df, selected_columns, filter_values):
    filtered_df = df.copy()
    
    for column, values in filter_values.items():
        if values:
            filtered_df = filtered_df[filtered_df[column].isin(values)]
    
    return filtered_df

def display_filtered_data(filtered_df):
    st.write("Filtered Data:")
    st.write(filtered_df)
    
    # Display the filtered data in a professional table format
    st.write(filtered_df.style.format({"Today Date": lambda x: f'=TODAY()'}))
    
    # Use pygwalker for data visualization
    #fig = pyg.plot(filtered_df)
    #st.pyplot(fig)

def main():
    st.title("Excel Data Analysis App")
    
    # Upload the Excel file using Streamlit sidebar
    uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(uploaded_file)
        
        # Pre-process the data
        df = preprocess_data(df)
        
        st.write("Cleaned Data:")
        st.write(df)
        
        # Use multiselect to filter columns and contents
        selected_columns = st.multiselect("Select columns to filter", df.columns)
        
        filter_values = {}
        for column in selected_columns:
            unique_values = df[column].unique()
            filter_values[column] = st.multiselect(f"Select values for {column}", unique_values)
        
        filtered_df = filter_data(df, selected_columns, filter_values)
        
        display_filtered_data(filtered_df)

if __name__ == "__main__":
    main()
