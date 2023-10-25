import os
import pandas as pd

file_path = r"C:\MyStuff\Projects\Developing\Tests\StatsOrginizer\src\data\adherence2.csv"

# Read the CSV file into a pandas dataframe
df = pd.read_csv(file_path)

# Specify the column names you want to keep
UserId = 'UserId'
StatusDateTime = 'StatusDateTime'
statusTime = 'StatusTime'
StatusKey = 'StatusKey'
EndDateTime = 'EndDateTime'
StateDuration = 'StateDuration'
LastName = 'LastName'
FirstName = 'FirstName'
JobTitle = 'JobTitle'

# Keep only selected columns in the dataframe
df_custom = df[[UserId, statusTime, StatusDateTime, StatusKey, EndDateTime, StateDuration, LastName, FirstName, JobTitle]].copy()

# Convert StatusDateTime to datetime
df_custom[StatusDateTime] = pd.to_datetime(df_custom[StatusDateTime])

# Sort by UserId and StatusDateTime
df_custom.sort_values([UserId, StatusDateTime], inplace=True)

# Group by UserId
grouped = df_custom.groupby(UserId)

for name, group in grouped:
    first_code = None
    first_code_time = None
    last_code = None
    last_code_time = None
    last_code_before_gonehome = None
    scheduled_breaks_duration = pd.Timedelta(0)
    lunch_duration = pd.Timedelta(0)
    unscheduled_breaks_duration = pd.Timedelta(0)
    first_code_after_gonehome_set = False
    agent_name = f'{group[FirstName].values[0]} {group[LastName].values[0]}'

    for index, row in group.iterrows():
        if row[StatusKey] != 'gone home':
            if first_code is None:
                first_code = row[StatusKey]
                first_code_time = row[StatusDateTime]
            elif not first_code_after_gonehome_set and row[StatusDateTime] != first_code_time:
                first_code_after_gonehome_set = True
                
        # Calculate break durations
        if row[StatusKey] == 'scheduled break':
            scheduled_breaks_duration += pd.Timedelta(seconds=row[StateDuration])
        elif row[StatusKey] == 'at lunch':
            lunch_duration += pd.Timedelta(seconds=row[StateDuration])
        elif row[StatusKey] == 'unscheduled break':
            unscheduled_breaks_duration += pd.Timedelta(seconds=row[StateDuration])

    for index in reversed(group.index):
        row = group.loc[index]
        if last_code is None:
            last_code = row[StatusKey]
            last_code_time = row[StatusDateTime]
        if row[StatusKey] == 'gone home':
            previous_records = group.loc[group.index[group.index < index]]
            if not previous_records.empty:
                last_code_before_gonehome = previous_records.iloc[-1][StatusKey]
            break
            
        # Calculate break durations
        if row[StatusKey] == 'scheduled break':
            scheduled_breaks_duration += pd.Timedelta(seconds=row[StateDuration])
        elif row[StatusKey] == 'at lunch':
            lunch_duration += pd.Timedelta(seconds=row[StateDuration])
        elif row[StatusKey] == 'unscheduled break':
            unscheduled_breaks_duration += pd.Timedelta(seconds=row[StateDuration])

    
    print(f"UserId: {group[FirstName].values[0]}")
    print(f"First code of the day: {first_code} at {first_code_time}")
    
    print(f"Duration of scheduled breaks: {scheduled_breaks_duration}")
    print(f"Duration of lunch: {lunch_duration}")
    print(f"Duration of unscheduled breaks: {unscheduled_breaks_duration}")

    print(f"Last code of the day: {last_code} at {last_code_time}")
    print(f"Last code before last 'gone home': {last_code_before_gonehome}")
    print('---')

