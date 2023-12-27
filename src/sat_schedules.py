import pandas as pd
from itertools import cycle
from collections import defaultdict
import os

# Define the agent names
agents_bilingual = [
    "Andrea Arango",
    "Jenyse Arroyo",
    "Jusmanny Anzalota",
    "Aniuska Rodriguez",
    "Guadalupe Ocampo",
    "Mayelin Ferrales",
]

agents_non_bilingual = [
    "Fiona Parker",
    "Hope Smith",
    "Kadine Nesbeth",
    "Michael Jacobson",
    "Kenelm Doleyres",
    "Keshawna Morris",
    "Sean Scott"
]

# Generate all Saturdays in the year 2024
saturdays = pd.date_range(start='2024-01-06', end='2024-12-28', freq='W-SAT')

# Cycle through the agents indefinitely
bilingual_cycle = cycle(agents_bilingual)
non_bilingual_cycle = cycle(agents_non_bilingual)

# Create a schedule dictionary
schedule = defaultdict(list)

# Function to get the next weekday for an agent offset
def get_next_weekday(date, add_days):
    next_day = date + pd.DateOffset(days=add_days)
    while next_day.weekday() > 4 or next_day.weekday() < 1:  # Skip weekends, 0 is Monday, 6 is Sunday
        next_day += pd.DateOffset(days=1)
    return next_day


# Define a dictionary to store the previous offset day for each agent
previous_offset_days = {}

# Assign agents to each Saturday and determine their offset day
for saturday in saturdays:
    # Assign bilingual agents
    for _ in range(2):
        agent = next(bilingual_cycle)
        schedule[saturday].append(agent)
        # Get the next weekday for offset
        previous_offset_day = previous_offset_days.get(agent, 2)  # Start with a 2-day offset which should be Tuesday
        offset_date = get_next_weekday(saturday, previous_offset_day)
        while offset_date in schedule and any(agent in entry for entry in schedule[offset_date]):
            # If the offset date is already in the schedule and the agent is already assigned an offset on that day, increment the offset day by 1
            previous_offset_day = (previous_offset_day + 1) % 7
            offset_date = get_next_weekday(saturday, previous_offset_day)
        schedule[offset_date].append(f'{agent} (Offset)')
        # Update the previous offset day for the agent
        previous_offset_days[agent] = (previous_offset_day + 1) % 7  # Increment the offset day by 1, wrapping around to 0 if it exceeds 6

    # Assign non-bilingual agents
    for _ in range(2):
        agent = next(non_bilingual_cycle)
        schedule[saturday].append(agent)
        # Get the next weekday for offset
        previous_offset_day = previous_offset_days.get(agent, 2)  # Start with a 2-day offset which should be Tuesday
        offset_date = get_next_weekday(saturday, previous_offset_day)
        while offset_date in schedule and any(agent in entry for entry in schedule[offset_date]):
            # If the offset date is already in the schedule and the agent is already assigned an offset on that day, increment the offset day by 1
            previous_offset_day = (previous_offset_day + 1) % 7
            offset_date = get_next_weekday(saturday, previous_offset_day)
        schedule[offset_date].append(f'{agent} (Offset)')
        # Update the previous offset day for the agent
        previous_offset_days[agent] = (previous_offset_day + 1) % 7  # Increment the offset day by 1, wrapping around to 0 if it exceeds 6

# Convert the schedule to a DataFrame
df_schedule = pd.DataFrame([(date, agent) for date, agents in schedule.items() for agent in agents], columns=['Date', 'Agent'])
df_schedule['Day'] = df_schedule['Date'].dt.day_name()

# Sort the DataFrame by date
df_schedule.sort_values('Date', inplace=True)

# Print the first few entries to verify
print(df_schedule.head(10))

os.startfile('agent_schedule_2024_with_offsets.csv')

# Optionally, save to a CSV file
df_schedule.to_csv('agent_schedule_2024_with_offsets.csv', index=False)

