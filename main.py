import pandas as pd
from datetime import datetime, timedelta

# Redirecting console output to "output.txt"
import sys

sys.stdout = open("output.txt", "w")

# Loading the Excel file into a DataFrame
file_path = "C:\Sem long\Internshala\Bluejay delivery\Assignment_Timecard.xlsx"
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print(f"File not found: {file_path}")
    exit(1)
except Exception as e:
    print(f"An error occurred while reading the file: {str(e)}")
    exit(1)

# Ensuring that the required columns are present in the DataFrame
required_columns = ["Employee Name", "Position ID", "Time", "Time Out"]
if not all(col in df.columns for col in required_columns):
    print("Required columns are missing in the Excel file.")
    exit(1)

# Converting 'Time' and 'Time Out' columns to datetime objects
df["Time"] = pd.to_datetime(df["Time"])
df["Time Out"] = pd.to_datetime(df["Time Out"])

# Sorting the DataFrame by employee name and time in
df.sort_values(["Employee Name", "Time"], inplace=True)

# Initializing variables to track consecutive days and time between shifts
consecutive_days = 0
prev_employee = None
prev_time_out = None

# Iterating through the DataFrame
for index, row in df.iterrows():
    employee = row["Employee Name"]
    time_in = row["Time"]
    time_out = row["Time Out"]

    # Checking if it's the same employee
    if employee == prev_employee:
        # Checking for consecutive days
        if time_in - prev_time_out <= timedelta(days=1):
            consecutive_days += 1
        else:
            consecutive_days = 0

        # Checking for time between shifts
        time_between_shifts = (time_in - prev_time_out).total_seconds() / 3600  # hours
        if 1 < time_between_shifts < 10:
            print(
                f"Employee: {employee}, Position: {row['Position ID']} - Less than 10 hours between shifts"
            )

        # Checking for long shifts
        shift_duration = (time_out - time_in).total_seconds() / 3600  # hours
        if shift_duration > 14:
            print(
                f"Employee: {employee}, Position: {row['Position ID']} - Worked more than 14 hours in a single shift"
            )

    else:
        consecutive_days = 0

    # Checking for 7 consecutive days
    if consecutive_days == 6:
        print(
            f"Employee: {employee}, Position: {row['Position ID']} - Worked 7 consecutive days"
        )

    prev_employee = employee
    prev_time_out = time_out
