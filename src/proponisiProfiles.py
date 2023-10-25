def createProponisiProfilesStringAndExcel(supervisor, location, file_path, team=''):
    import pandas as pd
    import os

    # Load the .xlsx file into a DataFrame
    df = pd.read_excel(file_path, engine='openpyxl')

    # Identifying rehires based on patterns
    patterns = r"\(R\)|\(R \)|\( r \)|\( r\)"
    df['Rehires'] = df.apply(lambda row: 'yes' if any(row.str.contains(patterns, case=False, na=False, regex=True)) else 'no', axis=1)

    # Creating the FullName column
    df['FullName'] = df['First Name'] + ' ' + df['Last Name']

    # Removing the (R) patterns, symbols, and extra spaces from the FullName column
    symbols = r'[-.]'  # regex pattern to match - and .
    df['FullName'] = (df['FullName'].str.replace(patterns, '', case=False, regex=True)
                    .str.replace(symbols, '', regex=True)
                    .str.replace(r'\s+', ' ', regex=True)
                    .str.strip())

    # Update the values in the existing Email column based on the pattern
    df['Email '] = df['FullName'].str.replace(r'\s+', '.', regex=True).str.lower() + '@conduent.com'

    # Remove unwanted characters from first name and last name for email creation
    def clean_for_email(name):
        return name.str.replace(r"[-,',]", '', regex=True)

    split_names = df['FullName'].str.split(' ', n=1, expand=True)
    split_names[0] = clean_for_email(split_names[0])
    split_names[1] = clean_for_email(split_names[1])

    # Update the values in the existing Email column based on the pattern
    df['Email '] = split_names[0].str.lower() + '.' + split_names[1].str.replace(' ', '').str.lower() + '@conduent.com'

    # Ensure the Start Date column is in datetime format
    df['Start Date'] = pd.to_datetime(df['Start Date'])

    # Format the Start Date column
    df['Start Date'] = df['Start Date'].dt.strftime('%m/%d/%Y')

    # Create a new dataframe to write to the Excel file, without Rehires and FullName columns
    df_to_write = df.drop(columns=['Rehires', 'FullName'])

    # Write the modified dataframe back to the .xlsx file, overwriting the original data
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df_to_write.to_excel(writer, index=False, sheet_name='Sheet1')  # specify the sheet name to overwrite

    # Add constant values and blank spaces to the dataframe
    df['blank1'] = ''
    df['blank2'] = ''
    df['blank3'] = ''
    df['blank4'] = ''
    df['supervisor'] = supervisor
    df['CSR'] = 'CSR'
    df['Team'] = team
    df['Call Center'] = 'Call Center'
    df['Location'] = location

    # Specifying columns in the desired order
    ordered_columns = ['Win #', 'blank1', 'blank2', 'blank3', 'blank4', 'FullName', 'Email ', 'supervisor', 'CSR', 'Team', 'Call Center', 'Location', 'Start Date']

    # Creating a subset of the dataframe based on the desired columns
    df_subset = df[ordered_columns]

    # Save the subset dataframe to a .txt file with CSV format and no spaces
    df_subset.to_csv('bulk_proponisi_profiles.txt', sep=',', index=False, header=False)
    os.startfile(file_path)
    os.startfile('bulk_proponisi_profiles.txt')