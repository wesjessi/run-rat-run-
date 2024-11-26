import openpyxl
import pandas as pd
import glob
import os
from datetime import datetime

# Constants
ROW_SPLIT = 719  # Adjusted to split active/inactive cycles correctly
WHEEL_CIRCUMFERENCE = 1.081

# Function to calculate running bouts
def calculate_running_bouts(wheel_turns):
    bouts = []
    for i in range(len(wheel_turns)):
        if i == 0:  # Special case for the first row
            if wheel_turns[i] >= 3:
                bouts.append(1)
            else:
                bouts.append(0)
        else:  # Normal logic for subsequent rows
            if wheel_turns[i - 1] < 3 and wheel_turns[i] >= 3:
                bouts.append(1)
            else:
                bouts.append(0)
    return bouts

# Function to calculate metrics
def calculate_metrics(data, sheet, date, segment, debug_data):
    metrics = {}
    # Only count values >= 3 for wheel turns
    valid_wheel_turns = data['Activity'][data['Activity'] >= 3]

    metrics['Total_Bouts'] = data['Running_Bout'].sum()
    metrics['Minutes_Running'] = len(valid_wheel_turns)  # Same as (data['Activity'] >= 3).sum()
    metrics['Total_Wheel_Turns'] = valid_wheel_turns.sum()  # Sum only valid wheel turns
    metrics['Distance_m'] = metrics['Total_Wheel_Turns'] * WHEEL_CIRCUMFERENCE

    if metrics['Total_Bouts'] > 0:
        metrics['Avg_Distance_per_Bout'] = metrics['Distance_m'] / metrics['Total_Bouts']
    else:
        metrics['Avg_Distance_per_Bout'] = 0

    if metrics['Total_Bouts'] > 0:
        metrics['Avg_Bout_Length'] = metrics['Minutes_Running'] / metrics['Total_Bouts']
    else:
        metrics['Avg_Bout_Length'] = 0

    metrics['Speed'] = metrics['Distance_m'] / metrics['Minutes_Running'] if metrics['Minutes_Running'] > 0 else 0

    # Collect debug information
    debug_data.append({
        'Rat': sheet,
        'Date': date,
        'Segment': segment,
        'Wheel Turns Sum (>=3)': valid_wheel_turns.sum(),
        'Minutes Running (>= 3)': len(valid_wheel_turns),
        'Distance (meters)': metrics['Distance_m'],
        'Total Bouts': data['Running_Bout'].sum()
    })

    return metrics
    
# Save Hourly Data Function
def save_hourly_data(output_dir, hourly_data, filename, phase_label):
    """
    Saves hourly data to an Excel file with separate sheets for each hour.
    """
    with pd.ExcelWriter(os.path.join(output_dir, filename)) as writer:
        for hour, data in hourly_data.items():
            if data:  # Only process if there is data for this hour
                rows = {'Rat': []}  # Initialize a dictionary for rows
                for rat, dates in data.items():
                    rows['Rat'].append(rat)  # Add the rat ID
                    for day_index, (date, metrics) in enumerate(dates.items(), start=1):  # Iterate over days
                        for metric, value in metrics.items():
                            col_name = f"{metric} Day {day_index}"  # Create a column name
                            if col_name not in rows:
                                rows[col_name] = []
                            rows[col_name].append(value)
                # Pad columns to the same length
                max_len = max(len(col) for col in rows.values())
                for col in rows:
                    while len(rows[col]) < max_len:
                        rows[col].append(None)
                # Create and save the DataFrame
                hour_df = pd.DataFrame(rows)
                hour_df.to_excel(writer, sheet_name=f"{phase_label} Hour {hour + 1}", index=False)
                
# Helper Functions
def save_data_to_excel(output_dir, data_dict, filename):
    """
    Save the given data dictionary to an Excel file with multiple sheets.
    """
    with pd.ExcelWriter(os.path.join(output_dir, filename)) as writer:
        for metric, rat_data in data_dict.items():
            if rat_data:
                pd.DataFrame(rat_data).T.to_excel(writer, sheet_name=metric)

def save_hourly_data(output_dir, hourly_data, filename, phase_label):
    """
    Save hourly data to an Excel file with separate sheets for each hour.
    """
    with pd.ExcelWriter(os.path.join(output_dir, filename)) as writer:
        has_data = False
        for hour, data in hourly_data.items():
            if data:
                has_data = True
                rows = {'Rat': []}
                for rat, dates in data.items():
                    rows['Rat'].append(rat)
                    for day_index, (date, metrics) in enumerate(dates.items(), start=1):
                        for metric, value in metrics.items():
                            col_name = f"{metric} Day {day_index}"
                            if col_name not in rows:
                                rows[col_name] = []
                            rows[col_name].append(value)
                max_len = max(len(col) for col in rows.values())
                for col in rows:
                    while len(rows[col]) < max_len:
                        rows[col].append(None)
                pd.DataFrame(rows).to_excel(writer, sheet_name=f"{phase_label} Hour {hour + 1}", index=False)

        if not has_data:
            pd.DataFrame({"Message": ["No data available"]}).to_excel(writer, sheet_name="No Data")

                
# Main processing function
def main_process(input_dir, output_dir):
    # Get list of Excel files
    excel_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in directory: {input_dir}")

    # Sort files by date, stripping the prefix and parsing the date
    excel_files = sorted(
        excel_files,
        key=lambda x: datetime.strptime(
            os.path.basename(x).replace('MT14 running data ', '').split('.')[0],
            '%m-%d-%y'
        )
    )

    # Initialize data dictionaries
    active_data = {metric: {} for metric in ['Total_Bouts', 'Minutes_Running', 'Total_Wheel_Turns', 'Distance_m', 'Avg_Distance_per_Bout', 'Avg_Bout_Length', 'Speed']}
    inactive_data = {metric: {} for metric in ['Total_Bouts', 'Minutes_Running', 'Total_Wheel_Turns', 'Distance_m', 'Avg_Distance_per_Bout', 'Avg_Bout_Length', 'Speed']}

    # Separate hourly data dictionaries for active and inactive phases
    hourly_data_active = {hour: {} for hour in range(12)}  # Active hours: 0 to 11
    hourly_data_inactive = {hour: {} for hour in range(12)}  # Inactive hours: 0 to 11

    # Initialize debug data for logging and troubleshooting
    debug_data = []

    # Process each file and collect metrics
    for file in excel_files:
        xl = pd.ExcelFile(file)
        rat_sheets = xl.sheet_names
        date = os.path.basename(file).split('.')[0]  # Use file name as date
        for sheet in rat_sheets:
            try:
                df_rat = xl.parse(sheet, usecols="A")  # Only read column A
                # Check if 'Activity' column exists and is valid
                if 'Activity' in df_rat.columns and not df_rat['Activity'].isnull().all():
                    df_rat['Running_Bout'] = calculate_running_bouts(df_rat['Activity'])

                    # Determine if active or inactive cycle is first
                    first_cycle = 'Active' if df_rat['Activity'].iloc[:ROW_SPLIT].sum() > df_rat['Activity'].iloc[ROW_SPLIT:].sum() else 'Inactive'

                    # Split data into active and inactive cycles
                    if first_cycle == 'Active':
                        df_active = df_rat.iloc[:ROW_SPLIT]
                        df_inactive = df_rat.iloc[ROW_SPLIT:]
                    else:
                        df_inactive = df_rat.iloc[:ROW_SPLIT]
                        df_active = df_rat.iloc[ROW_SPLIT:]

                    # Process hourly data (existing logic remains)
                    for hour in range(24):
                        phase = "Active" if hour < 12 else "Inactive"
                        start_row = hour * 60
                        end_row = start_row + 60
                        hourly_df = df_rat.iloc[start_row:end_row]

                        if len(hourly_df) < 60:  # Handle missing data
                            padding = pd.DataFrame(0, index=range(60 - len(hourly_df)), columns=hourly_df.columns)
                            hourly_df = pd.concat([hourly_df, padding], ignore_index=True)

                        metrics = calculate_metrics(hourly_df, sheet, date, f"Hour {hour + 1}", debug_data)
                        if phase == "Active":
                            if sheet not in hourly_data_active[hour % 12]:
                                hourly_data_active[hour % 12][sheet] = {}
                            hourly_data_active[hour % 12][sheet][date] = metrics
                        else:
                            if sheet not in hourly_data_inactive[hour % 12]:
                                hourly_data_inactive[hour % 12][sheet] = {}
                            hourly_data_inactive[hour % 12][sheet][date] = metrics

                   # Calculate metrics for active and inactive cycles (existing logic remains)
                    active_metrics = calculate_metrics(df_active, sheet, date, "Active", debug_data)
                    inactive_metrics = calculate_metrics(df_inactive, sheet, date, "Inactive", debug_data)

                   # Add data to dictionaries (existing logic remains)
                    for metric, value in active_metrics.items():
                        if sheet not in active_data[metric]:
                            active_data[metric][sheet] = {}
                        active_data[metric][sheet][date] = value
                    for metric, value in inactive_metrics.items():
                        if sheet not in inactive_data[metric]:
                            inactive_data[metric][sheet] = {}
                        inactive_data[metric][sheet][date] = value

                else:
                    # Log issue with missing or invalid 'Activity' column
                    debug_data.append({'File': file, 'Sheet': sheet, 'Error': 'Missing or invalid Activity column'})

            except Exception as e:
                debug_data.append({'File': file, 'Sheet': sheet, 'Error': str(e)})

            # Save Active and Inactive Data
            save_data_to_excel(output_dir, active_data, 'Active_Data.xlsx')
        save_data_to_excel(output_dir, inactive_data, 'Inactive_Data.xlsx')

    # Save Hourly Data
    save_hourly_data(output_dir, hourly_data_active, 'Active_Hourly_Data.xlsx', "Active")
    save_hourly_data(output_dir, hourly_data_inactive, 'Inactive_Hourly_Data.xlsx', "Inactive")

    # Save Debug Data
    debug_df = pd.DataFrame(debug_data)
    debug_df.to_excel(os.path.join(output_dir, 'Debug_Output.xlsx'), index=False)
