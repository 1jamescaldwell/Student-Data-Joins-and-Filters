# Graduated students
# James Caldwell UVA IRA
# 8/19/24

# This script performs a task similar to a VLOOKUP in Excel. It merges data from two spreadsheets, adds a new column for graduation term information, and creates a new Excel file with the updated data. 

import pandas as pd

# File paths
file_graduated_raw = r"...\Graduated students raw.xlsx"
file_graduated = r"...\Graduated students.xlsx"
file_graduated_new = r"...\Graduated students 2.xlsx"

# Load the Excel files into DataFrames
df_graduated_raw = pd.read_excel(file_graduated_raw)
df_graduated = pd.read_excel(file_graduated)

# Merge the two DataFrames based on 'Student System ID'
df_merged = pd.merge(df_graduated, df_graduated_raw[['Student System ID', 'Completion Term Desc','Degree Level Desc']], 
                     on='Student System ID', how='left')

# Rename the new column to avoid potential conflicts
df_merged.rename(columns={'Completion Term Desc': 'New Completion Term Desc'}, inplace=True)

# Format for final NPSAS submission
# Add the "Graduated" column based on the condition
# If a student graduated with a bachlors in 2023-2024, we want a yes. Otherwise, put a no.
df_merged['Graduated Bachelors 2023-2024 Year'] = df_merged.apply(
    lambda row: 'Yes' if row['Degree Level Desc'] == 1 and row['New Completion Term Desc'] in ['2024 Spring', '2023 Fall'] else 'No',
    axis=1
)

# Save the updated DataFrame back to the Graduated Students file
df_merged.to_excel(file_graduated_new, index=False)

print("File updated successfully.")
