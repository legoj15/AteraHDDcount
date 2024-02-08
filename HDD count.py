import pandas as pd
import os

# Get a list of all .xlsx files in the current directory
xlsx_files = [f for f in os.listdir('.') if os.path.isfile(f) and f.endswith('.xlsx')]

for xlsx_file in xlsx_files:
    # Load the spreadsheet
    xl = pd.ExcelFile(xlsx_file)

    # Load a sheet into a DataFrame by its name
    df = xl.parse('sheet1')

    # Replace 'Unspecified' with 'HDD' in 'Disk Type' column
    df['Disk Type'] = df['Disk Type'].replace('Unspecified', 'HDD')

    # Get the unique agent names
    unique_agents = df['Agent Name'].unique()

    # List to store the agent names that fit the criteria
    agents_with_hdd = []

    # Iterate over the unique agent names
    for agent in unique_agents:
        # Get the rows for this agent
        agent_rows = df[df['Agent Name'] == agent]
        
        # If 'SSD' is not in the 'Disk Type' column for this agent
        if 'SSD' not in agent_rows['Disk Type'].values:
            # If 'HDD' is in the 'Disk Type' column for this agent
            if 'HDD' in agent_rows['Disk Type'].values:
                agents_with_hdd.append(agent)

    # Sort the list of agent names
    agents_with_hdd.sort()

    # Write the results to a .txt file
    with open(xlsx_file.replace('.xlsx', '.txt'), 'w') as f:
        f.write(f"The number of unique agents with 'HDD' and without 'SSD' is: {len(agents_with_hdd)}\n")
        f.write("The agent names that fit the criteria are:\n")
        for agent in agents_with_hdd:
            f.write(f"{agent}\n")
