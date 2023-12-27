import pandas as pd
import openpyxl

# Path to the Excel file
file_path = 'Guardians Coaching Tracker 2023.xlsx'

# Specify the sheet name you want to access
sheet_name = 'Mayelín F.'

workbook = openpyxl.load_workbook('Guardians Coaching Tracker 2023.xlsx')

# Get the sheet names
current_sheet = workbook[sheet_name]
print(current_sheet)

# Load the data from the specified sheet, assuming no header row
data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Access column "AX" by its index
column_index_ax = 49  # Assuming 'AX' is the 50th column
column_ax = data.iloc[:, column_index_ax]

# Define the indices for each week
first_week_indices = [50, 51]
second_week_indices = [53, 54]
third_week_indices = [56, 57]
fourth_week_indices = [59, 60]
fifth_week_indices = [62, 63]



# Find the row where 'Oct' is located
oct_rows = column_ax[column_ax.str.contains('Jun', na=False)]

# Assuming you want to print the first occurrence of 'Oct'
if not oct_rows.empty:
    first_oct_index = oct_rows.index[0]

    # Select columns from 'AX' to 'BO' for the specific row by their indices
    ax_to_bo_indices = range(49, 66)  # Adjust these indices based on your file
    selected_columns = data.iloc[first_oct_index, ax_to_bo_indices]

    # Print the selected range of columns
    print(f"Row containing 'Oct' (Columns 'AX' to 'BO'):\n{selected_columns}")
else:
    print("No row containing 'Oct' was found.")
