import pandas as pd

# Load data from the provided files
scheduled_vsactual_path = 'excel_automation/Input files/scheduled_vs_actual_GT.csv'
timeoffrequest_path = 'excel_automation/Input files/timeoffrequest_report_GT.csv'
output_file = 'excel_automation/Generated files/GT_Staff_Working_Hours.xlsx'

# Read the CSV files into DataFrames
scheduled_df = pd.read_csv(scheduled_vsactual_path)
timeoff_df = pd.read_csv(timeoffrequest_path)

# Ensure date columns are in datetime format
scheduled_df['Date'] = pd.to_datetime(scheduled_df['Date'], format='%b %d, %Y')
timeoff_df['Start Date'] = pd.to_datetime(timeoff_df['Start Date'], format='%b %d, %Y')
timeoff_df['End Date'] = pd.to_datetime(timeoff_df['End Date'], format='%b %d, %Y')

# Create a DataFrame to store the final output
output_df = scheduled_df.copy()

# Initialize the Time Off Hours column
output_df['Time Off Hours'] = 0

# Convert the 'Time Off Amount' from string to float, removing the dollar sign and commas
timeoff_df['Time Off Amount'] = timeoff_df['Time Off Amount'].replace('[\$,]', '', regex=True).astype(float)

# Add Time Off Requests to the data
for index, row in timeoff_df.iterrows():
    mask = (output_df['Date'] >= row['Start Date']) & (output_df['Date'] <= row['End Date']) & (output_df['Name'] == row['Name'])
    output_df.loc[mask, 'Time Off Hours'] += row['Time Off Amount']

# Calculate the total hours (Scheduled Hours - Time Off Hours + Actual Hours)
output_df['Total Hours'] = output_df['Scheduled Hours'] - output_df['Time Off Hours'] + output_df['Actual Hours']

# Reshape the DataFrame to match the desired output format
pivot_table = output_df.pivot_table(index='Name', columns='Date', values=['Scheduled Hours', 'Actual Hours', 'Time Off Hours', 'Total Hours'], aggfunc='sum')

# Flatten the multi-index columns
pivot_table.columns = [f'{col[1]}_{col[0]}' for col in pivot_table.columns]
pivot_table.reset_index(inplace=True)

# Create a new Excel writer object
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Write the pivot table to the Excel file
pivot_table.to_excel(writer, sheet_name='Sheet1', index=False, startrow=3)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Define the format for merged cells
merge_format = workbook.add_format({
    'bold': 1,
    'align': 'center',
    'valign': 'vcenter'
})

# Add day and date rows
dates = scheduled_df['Date'].unique()
days = pd.to_datetime(dates).strftime('%a')

# Create headers for the first two rows
header_days = []
header_dates = []

for date in dates:
    header_days.extend([pd.to_datetime(date).strftime('%a')] * 4)
    header_dates.extend([pd.to_datetime(date).strftime('%d-%b')] * 4)

# Write the headers to the worksheet
worksheet.write_row('B1', header_days, merge_format)
worksheet.write_row('B2', header_dates, merge_format)

# Write the SH, AH, TOH, and TH headers
sh_ah_toh_th_headers = []
for _ in dates:
    sh_ah_toh_th_headers.extend(['SH', 'AH', 'TOH', 'TH'])

worksheet.write_row('B3', sh_ah_toh_th_headers, merge_format)

# Merge the day and date cells appropriately
for i in range(len(dates)):
    col_index = 4 * i + 1
    worksheet.merge_range(0, col_index, 0, col_index + 3, pd.to_datetime(dates[i]).strftime('%a'), merge_format)
    worksheet.merge_range(1, col_index, 1, col_index + 3, pd.to_datetime(dates[i]).strftime('%d-%b'), merge_format)

# Set column widths
worksheet.set_column('A:A', 15)
worksheet.set_column('B:Z', 12)

# Save and close the writer
writer.close()
