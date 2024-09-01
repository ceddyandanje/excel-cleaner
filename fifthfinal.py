import os
import pandas as pd
from tqdm import tqdm
import tkinter as tk
from tkinter import filedialog
import signal

# Handle user interrupt (Ctrl+C)
def signal_handler(sig, frame):
    print('You pressed Ctrl+C! Exiting...')
    exit(0)

signal.signal(signal.SIGINT, signal_handler)

# Function to open file explorer for selecting the file
def get_file_path(prompt_message):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title=prompt_message, filetypes=[("Excel files", "*.xlsx")])
    return file_path

# Function to save file
def get_save_location(prompt_message):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    save_path = filedialog.asksaveasfilename(title=prompt_message, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    return save_path

# Function to log progress
def log_progress(message):
    with open("process_log.txt", "a") as log_file:
        log_file.write(message + "\n")
    print(message)

# Prompt user for file location
print("Please select the Excel file to clean.")
file_path = get_file_path("Select the Excel file to clean")

# Read the Excel file
log_progress("Reading the Excel file...")
df = pd.read_excel(file_path, sheet_name=0, dtype={'Phone 1': str, 'Phone 2': str, 'Phone 3': str})
log_progress("File read successfully.")

# Step 1: Duplicate the first worksheet and retain all rows with any Attendance Status
log_progress("Step 1: Duplicating the worksheet and retaining rows with any Attendance Status...")
df_clean = df.copy()

# Retain only necessary columns
columns_to_keep = ['Business Name', 'Phone 1', 'Phone 2', 'Phone 3', 'Email 1', 'Email 2', 'Email 3', 'Attendance Status']
df_clean = df_clean[columns_to_keep]

# Remove rows without 'Attendance Status'
log_progress("Filtering rows without Attendance Status...")
df_clean = df_clean.dropna(subset=['Attendance Status'])

log_progress("Step 1 completed: Retained specific columns and filtered rows based on the presence of Attendance Status.")

# Step 2: Create a second cleaned sheet with only phone numbers
log_progress("Step 2: Creating a second sheet with only phone numbers...")
df_phones = df_clean[['Business Name', 'Phone 1', 'Phone 2', 'Phone 3']]

# Remove rows without any phone number
log_progress("Filtering rows without any phone number...")
df_phones = df_phones.dropna(subset=['Phone 1', 'Phone 2', 'Phone 3'], how='all')

log_progress("Step 2 completed: Retained rows with phone numbers.")

# Step 3: Create a third cleaned sheet with company names and emails
log_progress("Step 3: Creating a third sheet with company names and emails...")
df_emails = df_clean[['Business Name', 'Email 1', 'Email 2', 'Email 3']]

# Remove rows without any email
log_progress("Filtering rows without any email...")
df_emails = df_emails.dropna(subset=['Email 1', 'Email 2', 'Email 3'], how='all')

log_progress("Step 3 completed: Retained rows with emails.")

# Step 4: Create a fourth sheet with only specific attendance statuses
log_progress("Step 4: Creating a fourth sheet with specific attendance statuses...")
top_statuses = ['24boy-top', '24boy-interested', 'top 10']

# Filter based on attendance status, case-insensitive
df_top_cleaned = df_clean[df_clean['Attendance Status'].str.lower().isin(top_statuses)]

log_progress("Step 4 completed: Retained rows with specified Attendance Statuses.")

# Prompt user for save location
print("Please select where to save the cleaned Excel file.")
save_path = get_save_location("Select save location for cleaned Excel file")

# Save the cleaned data to a new Excel file with progress tracking
log_progress("Saving the cleaned data to a new Excel file...")
with pd.ExcelWriter(save_path) as writer:
    tqdm.pandas(desc="Saving Original Data")
    df.progress_apply(lambda x: None, axis=1)  # Progress bar for writing original sheet
    df.to_excel(writer, sheet_name='Original Data', index=False)
    
    tqdm.pandas(desc="Saving Cleaned Data")
    df_clean.progress_apply(lambda x: None, axis=1)  # Progress bar for writing first cleaned sheet
    df_clean.to_excel(writer, sheet_name='Cleaned Data', index=False)
    
    tqdm.pandas(desc="Saving Phone Numbers")
    df_phones.progress_apply(lambda x: None, axis=1)  # Progress bar for writing second cleaned sheet
    df_phones.to_excel(writer, sheet_name='Phone Numbers', index=False)
    
    tqdm.pandas(desc="Saving Emails")
    df_emails.progress_apply(lambda x: None, axis=1)  # Progress bar for writing third cleaned sheet
    df_emails.to_excel(writer, sheet_name='Emails', index=False)

    tqdm.pandas(desc="Saving Top Cleaned Data")
    df_top_cleaned.progress_apply(lambda x: None, axis=1)  # Progress bar for writing fourth cleaned sheet
    df_top_cleaned.to_excel(writer, sheet_name='Top Cleaned', index=False)

log_progress("Final Step: Data saved to " + save_path)

# Sorting the sheets alphabetically
log_progress("Sorting the data alphabetically...")
df_clean.sort_values(by='Business Name', inplace=True)
df_phones.sort_values(by='Business Name', inplace=True)
df_emails.sort_values(by='Business Name', inplace=True)
df_top_cleaned.sort_values(by='Business Name', inplace=True)

log_progress("Data cleaning completed and saved successfully.")
