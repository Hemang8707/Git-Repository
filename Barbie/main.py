import pandas as pd
from datetime import timedelta

def analyze_employee_data(file_path):
    # Load the Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Drop rows with missing values in 'Time' or 'Timecard Hours (as Time)'
    df.dropna(subset=['Time', 'Timecard Hours (as Time)'], inplace=True)

    # Convert 'Time' column to datetime format
    df['Time'] = pd.to_datetime(df['Time'])

    # Preprocess 'Timecard Hours (as Time)' column to handle different formats
    def convert_to_timedelta(value):
        try:
            return pd.to_timedelta(value)
        except ValueError:
            if ':' in str(value):
                # If the value is in 'hh:mm:ss' format
                time_components = str(value).split(':')
                hours = int(time_components[0])
                minutes = int(time_components[1]) if len(time_components) > 1 else 0
                seconds = int(time_components[2]) if len(time_components) > 2 else 0
                return pd.Timedelta(hours=hours, minutes=minutes, seconds=seconds)
            else:
                # Assume the value is already in numeric hours
                return pd.Timedelta(hours=float(value))

    df['Timecard Hours (as Time)'] = df['Timecard Hours (as Time)'].apply(convert_to_timedelta)

    # Sort the DataFrame by 'Employee Name' and 'Time' for consecutive day analysis
    df.sort_values(['Employee Name', 'Time'], inplace=True)

    # Group by 'Employee Name'
    grouped_data = df.groupby('Employee Name')

    # Function to check consecutive days worked
    def consecutive_days_worked(group):
        return group['Time'].diff().dt.days == 1

    # Function to check time between shifts and single-shift hours
    def analyze_shifts(group):
        less_than_10_hours = group['Time'].diff() < pd.Timedelta(hours=10)
        more_than_14_hours = group['Timecard Hours (as Time)'] > pd.Timedelta(hours=14)
        return group[less_than_10_hours], group[more_than_14_hours]

    # Iterate through each group (Employee)
    for name, group in grouped_data:
        consecutive_days = consecutive_days_worked(group)
        less_than_10_hours, more_than_14_hours = analyze_shifts(group)

        # Check and print employees who meet the conditions
        if consecutive_days.all():
            print(f"{name} has worked for 7 consecutive days. Position: {group['Position ID'].iloc[0]}")

        if not less_than_10_hours.empty:
            print(f"{name} has less than 10 hours between shifts but greater than 1 hour. Position: {less_than_10_hours['Position ID'].iloc[0]}")

        if not more_than_14_hours.empty:
            print(f"{name} has worked for more than 14 hours in a single shift. Position: {more_than_14_hours['Position ID'].iloc[0]}")

# Example usage:
file_path = "Assignment_Timecard.xlsx"
analyze_employee_data(file_path)
