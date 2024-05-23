import os
import hashlib
import pandas as pd
from openpyxl import Workbook


# Function to generate a unique ID for a user
def generate_user_id(user_id):
    # Use a hash function to generate a unique ID from the user ID
    return hashlib.sha1(user_id.encode()).hexdigest()[:10]


# File paths
input_excel_file = r'C:\Temp\Test Folder\File Input\test.xlsx'
output_folder = r'C:\Temp\Test Folder\File Output'

# Read the spreadsheet into a pandas DataFrame
df = pd.read_excel(input_excel_file)

# Filter rows where AuditData mentions portfolio@searchflow.co.uk
df_filtered = df[df['AuditData'].str.contains('Portfolio@Searchflow.co.uk', na=False)]

# Group by UserID and Operation, and count occurrences
summary_df = df_filtered.groupby(['UserId', 'Operation']).size().reset_index(name='Count')

# Group rows by UserID
grouped_df = df_filtered.groupby('UserId')

# Create a new Excel workbook
wb = Workbook()

# Create a summary sheet
summary_sheet = wb.active
summary_sheet.title = 'Summary'
summary_sheet.append(['UserID', 'Operation', 'Count'])

# Iterate through each group (user)
for user_id, group in grouped_df:
    # Generate a unique ID for the user
    user_id_hash = generate_user_id(user_id)

    # Write the UserID to the summary sheet
    summary_sheet.append([user_id])

    # Dictionary to store operation counts for the current user
    operation_counts = {}

    # Count occurrences of each operation for the current user
    for index, row in group.iterrows():
        operation = row['Operation']
        if operation in operation_counts:
            operation_counts[operation] += 1
        else:
            operation_counts[operation] = 1

    # Write operation counts to the summary sheet
    for operation, count in operation_counts.items():
        summary_sheet.append(['', operation, count])

    # Create a new worksheet with the generated ID
    ws = wb.create_sheet(title=user_id)

    # Write the rows for the current user to the worksheet
    ws.append(df.columns.tolist())
    for index, row in group.iterrows():
        ws.append(row.tolist())

# Save the workbook to the output folder
output_file = os.path.join(output_folder, 'summary.xlsx')
wb.save(output_file)
