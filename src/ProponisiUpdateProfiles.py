import pandas as pd
import os

def clean_for_email(name):
    """
    Function to remove unwanted characters from name for email creation.
    """
    return name.str.replace(r"[-,',]", '', regex=True)

def update_proponisi_profiles_nesting_to_team(location, file_path, file_path_leads):
    
    # Load the .xlsx file into a DataFrame
    df = pd.read_excel(file_path, engine='openpyxl')

    # Process 'NAME' column to get it in 'First Name Last Name' format
    name_split = df['NAME'].str.split(',', expand=True)
    name_split = name_split.apply(lambda x: x.str.strip())  # Trim white spaces
    df['NAME'] = name_split[1] + ' ' + name_split[0]

    # Remove unwanted characters from names
    df['NAME'] = df['NAME'].str.replace(r'[-.*]', ' ', regex=True)

    # Process 'NAME' column for email creation
    split_names = df['NAME'].str.split(' ', n=1, expand=True)
    split_names = split_names.apply(clean_for_email)
    df['Email'] = split_names[0].str.lower() + '.' + split_names[1].str.replace(' ', '').str.lower() + '@conduent.com'

    # Replace multiple spaces with a single space
    df['Transferred Team'] = df['Transferred Team'].str.replace(r'\s+', ' ', regex=True)

    # Strip leading and trailing spaces
    df['Transferred Team'] = df['Transferred Team'].str.strip()

    # Read the leads per subteam file
    df_leads = pd.read_excel(file_path_leads, engine='openpyxl')

    # Merge the two DataFrames on the 'Transferred Team' column
    merge_df = pd.merge(df, df_leads, on='Transferred Team', how='inner')

    merge_df['location'] = location
    merge_df['TEAM'] = merge_df['Transferred Team'].str.replace(r'[\d\s]+$','', regex=True)
    merge_df[['blank1', 'blank2', 'blank3', 'blank4']] = ''  # Assign empty strings
    merge_df['CSR'] = 'CSR'
    merge_df['Call Center'] = 'Call Center'

    # Specifying columns in the desired order
    ordered_columns = ['FEPS ID', 'blank1', 'blank2', 'blank3', 'blank4', 'NAME', 'Email', 'WIN ID', 'CSR', 'TEAM', 'Call Center', 'location', 'Hire Date']
    df_for_txt = merge_df[ordered_columns]

    # Output the processed data to a text file
    output_file_path = 'updateProfiles.txt'
    df_for_txt.to_csv(output_file_path, sep=',', index=False, header=False)

    # Open the output file using the default application
    os.startfile(output_file_path)

