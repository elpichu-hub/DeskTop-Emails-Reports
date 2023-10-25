from openpyxl import load_workbook, Workbook
import pandas as pd

# Load the .xlsm file
xlsm_filename = 'Guardians 2023 Attendance Tracker.xlsm'
wb = load_workbook(xlsm_filename, data_only=True)

# Create a new .xlsx workbook
xlsx_wb = Workbook()

# Get the active sheet (default created) and then remove it after creating at least one sheet
default_sheet = xlsx_wb.active

# Copy data from .xlsm to .xlsx, excluding columns A to C, after column AH,
# removing rows 1 to 3, and every row after row 17
for sheet in wb:
    xlsx_sheet = xlsx_wb.create_sheet(title=sheet.title)
    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        # Exclude rows 1 to 3 and every row after row 17
        if 1 <= row_idx <= 3 or row_idx > 17:
            continue
        
        # Exclude columns A to C by starting from the 4th column (D)
        # and exclude everything after column AH (35th column)
        filtered_row = row[3:36]
        xlsx_sheet.append(filtered_row)

# Now remove the default sheet
xlsx_wb.remove(default_sheet)

# Save the xlsx_wb workbook to a file
xlsx_filename = 'Guardians 2023 Attendance Tracker.xlsx'
xlsx_wb.save(xlsx_filename)

# Load the .xlsx file
wb = load_workbook(xlsx_filename)

# sheets containing the month in col A
sheets_with_month_col_a = []

# Search for the value 'MAR' in column A of every sheet
value_to_search = 'MAR'
for sheet in wb:
    for row_idx, row in enumerate(sheet.iter_rows(min_col=1, max_col=1, values_only=True), start=1):
        if row[0] == value_to_search:
            sheets_with_month_col_a.append((sheet.title, row_idx))
            break

print(sheets_with_month_col_a)

# Specify the Excel file and the range of cells to read
cell_range = 'A1:AF13'

# Specify the columns
columns = ['MONTH', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']



def check_value_on_feb_3(df):
    # Check if the value in the cell corresponding to February 3rd is not NaN
    try:
        value = df.loc[df['2023'] == 'FEB', '3']
    except Exception as e:
        print(e)
    return not pd.isna(value.values[0])



# Read the data from every sheet in the workbook and print the DataFrame
with pd.ExcelFile(xlsx_filename) as xls:
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        print(f"Data from {sheet_name}:")
        # Example usage:
        # df is the DataFrame you provided
        value_exists = check_value_on_feb_3(df)
        print(value_exists)




