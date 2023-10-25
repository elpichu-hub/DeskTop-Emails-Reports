# file_path1 = "User Productivity Summary stats two call center 5.15.2023 teams.csv"
def run_agent_time_report(file_path1):
    import pandas as pd
    import os
    
    def seconds_to_time_format(seconds):
        hours, remainder = divmod(seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    
    try:
        # read the CSV file into a pandas dataframe
        df = pd.read_csv(file_path1)   

        # specify the column names you want to keep
        user_id = 'i3user'
        queue = 'queue'
        avg_hold = 'totTHoldACD'
        avg_acw = 'totTACW'
        tot_talk = 'totTTalkACD'
        full_name = 'DisplayUserName'
        team = 'JobTitle'  # new column

        # keep only selected columns in the dataframe
        df_custom = df[[user_id, queue, avg_hold, avg_acw, tot_talk, full_name, team]]

        # # filter out rows where queue contains 'WebChat'
        # df_filtered = df_custom[~df_custom[queue].str.contains('WebChat', case=False)]
        df_filtered = df_custom

        # group by user_id and calculate total time per agent across all queues
        df_total_time_per_agent = df_filtered.groupby(user_id).agg({
            avg_hold: 'sum',
            avg_acw: 'sum',
            tot_talk: 'sum',
            team: 'first'
        }).reset_index()

        df_total_time_per_agent['Total Time In Phone'] = df_total_time_per_agent[avg_hold] + df_total_time_per_agent[avg_acw] + df_total_time_per_agent[tot_talk]
        
        # group the data by the 'queue' and 'i3user' columns, keeping the first occurrence of 'DisplayUserName'
        df_grouped = df_filtered.groupby([queue, user_id]).agg({
            avg_hold: 'sum',
            avg_acw: 'sum',
            tot_talk: 'sum',
            full_name: 'first',
            team: 'first'  # This line was added to include the 'team' column in the aggregation
        }).reset_index()
        
        # Merge total time per agent into df_grouped
        df_grouped = df_grouped.merge(df_total_time_per_agent[[user_id, 'Total Time In Phone']], on=user_id, how='left')

        # Convert the time columns to a time format (HH:MM:SS)
        df_grouped['Total Time'] = (df_grouped[avg_hold] + df_grouped[avg_acw] + df_grouped[tot_talk]).apply(seconds_to_time_format)
        df_grouped['Total Time In Phone'] = df_grouped['Total Time In Phone'].apply(seconds_to_time_format)

        # convert 'totTHoldACD' and 'totTACW' columns to time format
        df_grouped['totTHoldACD'] = df_grouped['totTHoldACD'].apply(seconds_to_time_format)
        df_grouped['totTACW'] = df_grouped['totTACW'].apply(seconds_to_time_format)
        df_grouped['totTTalkACD'] = df_grouped['totTTalkACD'].apply(seconds_to_time_format)
        

        # Export the data to a CSV file
        file_name = "Total_Time_Per_Agent_In_Queues.csv"
        df_grouped.to_csv(file_name, index=False, encoding='utf-8-sig')

        os.startfile(file_name)

        print(f"Report exported to: {file_name}")
        return f"Total Time Per Agent In Queues Report exported to: {file_name}"

    except Exception as e:
        print(e)
        return e
