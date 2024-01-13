# Data Transformation in Pandas/Python

This repository contains a Python script that automates the process of transforming and pivoting data within an Excel file. The script is designed to work with Excel files that contain transactional data for merchants, actions, and metrics.

## Technology Stack and Techniques

This project is implemented using Python and utilizes several programming techniques and libraries to manipulate and transform data efficiently.

### Key Python Features

-   **Object-Oriented Programming (OOP)**: The script is structured using classes and objects, encapsulating the data transformation logic within a `DataTransformer` class to enhance modularity and reusability.
-   **Pandas Library**: A powerful data manipulation tool that provides data structures and operations for manipulating numerical tables and time series.

### Data Manipulation Techniques

-   **Data Cleaning**: Standardizing the action names to ensure consistency across the dataset.
-   **Pivoting**: Transforming data from long to wide format, summarizing the actions and metrics for each merchant customer ID.
-   **Aggregation**: Computing summary statistics for each group of data, particularly counting occurrences of actions for each merchant.
-   **Column Renaming**: Modifying column names to merge similar actions into unified categories for clearer analysis.
-   **Dataframe Apply**: Using the apply method to process data row-wise, enabling complex logic that is dependent on multiple columns.

## Features

-   Cleans data by standardizing/renaming action names.
-   Pivots data into a summary table, highlighting the count of each action by merchant and metric.
-   Generates a 'Stats' string based on predefined rules and adds it to a new Excel sheet.
-   Separates records into different sheets based on the highest priority action.
-   Outputs a final Excel file with four sheets: 'Pivot Table', 'Revoke', 'Softblock', and 'Warn'.

## Usage

To use this script, ensure that you have an Excel file in the following format:

| Merchant Customer ID | Action | Metrics |
| -------------------- | ------ | ------- |
| 123456789            | Soft1  | CR      |
| ...                  | ...    | ...     |

The script will:

1. Replace 'Soft1' and 'Soft2' actions with 'Softblock', and 'Warn1' and 'Warn2' with 'Warn'.
2. Pivot the data to create a summary table.
3. Generate a 'Stats' string for each merchant and categorize it into the appropriate action sheet.
4. Save the output in a new Excel file with multiple sheets.

### Running the Script

Make sure to have `pandas` installed in your Python environment:

```bash
pip install pandas
```

Run the script with the input Excel file path, and it will generate the new Excel file with the transformed data:

```python
transformer = DataTransformer('path/to/your/input.xlsx', 'start_date', 'end_date')
transformer.clean_data()
```

```python
pivot_table = transformer.pivot_data()
stats_sheets = transformer.create_stats_sheets(pivot_table)
transformer.save_to_excel(pivot_table, stats_sheets, 'path/to/your/output.xlsx')
```

Replace `'path/to/your/input.xlsx'` with the path to your input Excel file and `'path/to/your/output.xlsx'` with the desired path for the output file.
