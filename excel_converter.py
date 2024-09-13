import pandas as pd

# Define the data for the Excel file
data = {
    'file_column': [
        'file1.txt',
        'file2.txt',
        'file3.txt'
    ]
}

# Create a DataFrame
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
excel_file_path = 'sample_files.xlsx'
df.to_excel(excel_file_path, index=False)

print(f"Sample Excel file created: {excel_file_path}")
