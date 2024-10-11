# Methods Engineering

The Methods Engineering App is designed to analyze and filter Inspection Plant Maintenance data from SAP harvested into excel format. It provides a user-friendly interface to upload the so-called file, perform data cleansing, and allow users to select columns and apply filters for data analysis. 

## Features

- File Upload: Users can upload an Excel file containing topsides plant maintenance data.
- Data Cleansing: The App performs data cleansing tasks.
- Column Selection: Users can select specific columns from the uploaded Excel file for analysis.
- Data Filtering: Users can apply filters to the selected columns using dropdown menus to narrow down the data based on specific values.
- Summary Statistics: The webapp displays summary descriptive statistics for the filtered data in a condensed table format.
- Numeric Column Analysis: It provides summary statistics for numeric columns in the filtered data.
- Categorical Column Analysis: It shows unique values and their counts for categorical columns in the filtered data.

## Usage

1. Run the Streamlit webapp by executing the Python script.
2. Upload an Excel file containing topsides plant maintenance data using the file uploader.
3. The App will perform data cleansing tasks automatically.
4. Select the desired columns for analysis using the multiselect dropdown menus.
5. Apply filters to the selected columns using the corresponding dropdown menus.
6. The webapp will display the filtered data as summary descriptive statistics.
7. Explore the summary statistics for numeric columns and unique values and counts for categorical columns.

## Dependencies

- Python 3.x
- streamlit
- pandas
- pygwalker

## Data Cleansing Steps

1. Remove rows with missing data.
2. Remove columns with missing data.
3. Reset the index of the DataFrame.
4. Replace NaN values in the "SECE STATUS" column with "Non-SCE".

## Customization

- Styling: The webapp's title and description can be customized by modifying the HTML code in the `st.markdown()` function.
- File Types: The accepted file types for upload can be changed in the `st.file_uploader()` function.
- Column Selection: The column options for selection can be modified based on the specific requirements of the analysis.
- Filtering: The filtering mechanism can be extended to support additional data types or custom filtering logic.

## Limitations

- The webapp assumes that the uploaded Excel file is in a table format with column headers.
- The webapp may not handle large datasets efficiently and may require optimization for better performance.

## Future Work

- Upload raw SAP data without table format.
- Perform data cleansing techniques to transform the raw data into table format and defining the column naming headers.
- Attempt to leverage the .groupby() method to replicate the Pivot Tables functionalities.
- Implement graph functionalities for each filtered/sorted feature.
- Improve functionalities to handle large datasets efficiently with some optimization for better performance.


## License

This project is licensed under the [MIT License](LICENSE).
