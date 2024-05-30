import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta

# Load the raw data
raw_data_path = 'F:\excel_auto\excel_automation\Generated files\Raw.xlsx'  # replace with your actual file path
raw_df = pd.read_excel(raw_data_path)

# Load the format file
format_file_path = r'F:\excel_auto\excel_automation\format file\format.xlsx'  # replace with your actual format file path
wb = load_workbook(format_file_path)
ws = wb.active

# Extract unique names and dates from the raw data
names = raw_df['Name'].unique()
dates = pd.to_datetime(raw_df['Date'].unique())

# Helper function to find the anchor cell of merged ranges
def get_anchor_cell(ws, cell):
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell

# Ensure dates are sorted and continuous
min_date = min(dates)
max_date = max(dates)
all_dates = pd.date_range(start=min_date, end=max_date)

# Calculate days based on dates
days = [date.strftime("%A") for date in all_dates]

# Fill in the names starting from cell C7
for row_index, name in enumerate(names, start=7):
    cell = ws.cell(row=row_index, column=3)
    if cell.value is None:
        cell.value = name

# Fill in the days starting from cell 4D
for col_index, day in enumerate(days, start=0):  # Start from column D
    cell = ws.cell(row=4, column=col_index * 2 + 4)  # Adjust column index
    if cell.value is None:
        cell.value = day

# Fill in the dates starting from cell 5D
for col_index, date in enumerate(all_dates, start=0):  # Start from column D
    cell = ws.cell(row=5, column=col_index * 2 + 4)  # Adjust column index
    if cell.value is None:
        cell.value = date.strftime("%d-%b-%y")

# Loop through each name and populate the worksheet
for name_index, name in enumerate(names, start=0):
    # Filter data for the current name
    name_df = raw_df[raw_df['Name'] == name]
    
    for date_index, date in enumerate(all_dates, start=0):
        # Filter data for the current date
        date_str = date.strftime("%d-%b-%y")
        date_df = name_df[name_df['Date'] == date_str]
        
        # Calculate the row index for SH and AH
        row_index = name_index + 7
        col_index = date_index * 2 + 4
        
        # Scheduled Hours (SH) cell
        sh_cell = ws.cell(row=row_index, column=col_index)
        if sh_cell.value is None:
            sh_cell.value = date_df.iloc[0]['Scheduled Hours'] if not date_df.empty else 0
        
        # Actual Hours (AH) cell
        ah_cell = ws.cell(row=row_index, column=col_index + 1)
        if ah_cell.value is None:
            ah_cell.value = date_df.iloc[0]['Actual Hours'] if not date_df.empty else 0

# Save the workbook to a new file
formatted_data_path = 'F:\excel_auto\excel_automation\Generated files\dashboard.xlsx'  # replace with your desired file path
wb.save(formatted_data_path)
