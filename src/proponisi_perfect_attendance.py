
from fuzzywuzzy import fuzz
import pandas as pd
import re
import os

def is_same_name(name1, name2, threshold=80):
    similarity_ratio = fuzz.token_set_ratio(name1, name2)
    if similarity_ratio > threshold:
        return True
    else:
        return False
    
file_path1 = 'Proponisi_report_lazaro_gonzalez_4.24.2023.csv'
file_path2 = 'perfect_attendance.csv'

def mark_attendance(file_path1, file_path2):
    try:
        df1 = pd.read_csv(file_path1, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            df1 = pd.read_csv(file_path1, encoding='latin1')
        except UnicodeDecodeError:
            df1 = pd.read_csv(file_path1, encoding='iso-8859-1')

    try:
        df2 = pd.read_csv(file_path2, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            df2 = pd.read_csv(file_path2, encoding='latin1')
        except UnicodeDecodeError:
            df2 = pd.read_csv(file_path2, encoding='iso-8859-1')

    df2['FullName'] = df2['FullName'].apply(lambda x: " ".join(str(x).split()))
    df2['FullName'] = df2['FullName'].apply(lambda x: re.sub(r'\s-\s', '-', x))
    df2['FullName'] = df2['FullName'].apply(lambda x: re.sub(r'[^\w\s-]', '', x))
    df2['FullName'] = df2['FullName'].apply(lambda x: re.sub(r'^\d+ - ', '', x))
    df2['FullName'] = df2['FullName'].apply(lambda x: re.sub(r'^\d+-', '', x))

    # Iterate over the 'FullName' column of df1
    for name1 in df1['FullName']:
        # Compare this name with all names in df2
        for name2 in df2['FullName']:
            if is_same_name(name1, name2):
                df1.loc[df1['FullName'] == name1, 'Perfect Attendance'] = 'Y'
                # Break the inner loop as we've found a match
                break

    dir_path = os.path.dirname(file_path1)
    base_name = os.path.basename(file_path1)
    new_file_path = os.path.join(dir_path, 'perfect_attendance_' + base_name)

    df1.to_csv(new_file_path, index=False)
    os.startfile(new_file_path)


