import streamlit as st
import pandas as pd
 
# Function to clean the data by removing the specified columns and rows
def clean_data(file):
    # Load the Excel or CSV file
    df = pd.read_csv(file)
    
    # Removing specified columns (Assuming we are using CSV for now)
    # Adapted to remove appropriate columns by index or name (adjust accordingly)
    columns_to_remove = ['A', 'F', 'G', 'H', 'S', 'W', 'X', 'Y', 'Z']
    df.drop(columns=columns_to_remove, inplace=True, errors='ignore')
    
    # Removing the first 4 rows
    df = df.iloc[4:].reset_index(drop=True)
    
    return df
 
# Function to group the data and create the backlog table
def create_backlog_table(df):
    # Grouping data by 'Class', 'Delay', 'SECE STATUS', 'Backlog'
    # Adapt based on the column names from the CSV/Excel file
    grouped_data = df.groupby(['Class', 'Delay', 'SECE STATUS', 'Backlog']).size().unstack(fill_value=0)
    
    return grouped_data
 
# Function to display the cleaned data and backlog analysis
def analyze_backlog(file):
    df = clean_data(file)
    
    # Display cleaned data
    st.subheader("Cleaned Data")
    st.write(df)
    
    # Creating backlog table
    backlog_table = create_backlog_table(df)
    
    # Display backlog analysis table
    st.subheader("Backlog Table Analysis")
    st.write(backlog_table)
 
# Streamlit App layout
def main():
    st.title("Backlog Analysis WebApp")
    
    # File upload option for the Excel/CSV file
    uploaded_file = st.file_uploader("Upload your Excel/CSV file", type=["xlsm", "csv"])
    
    # Button to perform analysis
    if st.button("Analyze Backlog"):
        if uploaded_file is not None:
            analyze_backlog(uploaded_file)
        else:
            st.error("Please upload a file to analyze.")
    
if __name__ == "__main__":
    main()
 
