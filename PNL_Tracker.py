import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Create a Tkinter root window (this will be hidden)
root = Tk()
root.withdraw()  # Hide the root window

# Prompt user to select an Excel file
file_path = askopenfilename(title='Select the Excel file', filetypes=[('Excel Files', '*.xlsx;*.xls')])

# Check if a file was selected
if not file_path:
    print("No file selected. Exiting.")
    exit()

# Load the Excel file
df = pd.read_excel(file_path)

# Multiply the relevant columns (index 8, 9, 12 in this example) to get PNL
df['Current PNL'] = df.iloc[:, 8] * df.iloc[:, 9] * df.iloc[:, 12]

# Ensure the 'Date' column is in datetime format (assuming it is in the first column)
df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0])

# Group by 'User' column (assuming there's a "User" column with user IDs)
grouped = df.groupby('User')

# Initialize a dictionary to store each user's cumulative PNL
user_cumulative_pnl = {}

# Calculate cumulative PNL for each user and store in dictionary
for user, user_df in grouped:
    user_df = user_df.sort_values(by=df.columns[0])  # Sort by date
    user_df['User Cumulative PNL'] = user_df['Current PNL'].cumsum()  # Calculate cumulative PNL directly
    df.loc[user_df.index, str(user) + 'Cumulative PNL'] = user_df['User Cumulative PNL']
    user_cumulative_pnl[user] = user_df[['Event Date', 'User Cumulative PNL']]

# Plot cumulative PNL for each user
plt.figure(figsize=(12, 6))

# Plot for each user
for user, pnl_df in user_cumulative_pnl.items():
    plt.plot(pnl_df['Event Date'], pnl_df['User Cumulative PNL'], marker='o', linestyle='-', markersize=3, label=user)

# Plot the overall cumulative PNL
df['Cumulative PNL'] = df['Current PNL'][::-1].cumsum()[::-1]
plt.plot(df.iloc[:, 0], df['Cumulative PNL'], marker='o', linestyle='-', color='b', markersize=3, label='Team Cumulative PNL')

# Customize plot
plt.title('Cumulative PNL for Each User vs Date')
plt.xlabel('Date')
plt.ylabel('Cumulative PNL')
plt.xticks(rotation=45)
plt.grid()

plt.gca().get_yaxis().set_major_formatter(ticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))

plt.legend()
plt.tight_layout()

# Save the plot
plt.savefig('Cumulative_PNL_per_User_vs_Date.png')
plt.show()

# Write the updated DataFrame back to the Excel file
df.to_excel(file_path, index=False)

print("PNL for Each User Calculated, Cumulative PNL Added, and Graph Created", file_path)

