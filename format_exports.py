import pandas as pd
import numpy as np
from datetime import datetime
import re
import requests
from io import BytesIO
import os
import shutil

# URL of the Excel file
EXCEL_URL = "https://balanca.economia.gov.br/balanca/semanal/Setores_Produtos.xlsx"

# Create directories if they don't exist
os.makedirs('datasets/historicals', exist_ok=True)
os.makedirs('datasets/downloads', exist_ok=True)
os.makedirs('datasets/current', exist_ok=True)
os.makedirs('datasets/historicals/backups', exist_ok=True)

def extract_date_info(text):
    """Extract month, year from header text."""
    # Example: "Mar/2025: 13 dias úteis"
    match = re.search(r'(\w+)/(\d{4}):', text)
    if match:
        month_str, year = match.groups()
        # Convert Portuguese month abbreviation to number
        month_map = {
            'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4, 'Mai': 5, 'Jun': 6,
            'Jul': 7, 'Ago': 8, 'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12
        }
        return int(year), month_map.get(month_str, 0)
    return None, None

def calculate_weekly_values(current_row, historical_data, year, month, working_days):
    """
    Calculate the weekly values based on whether it's a new month or mid-month update.
    Returns a dictionary with the calculated values.
    """
    # Find all historical data for this product in the current month/year
    product_history = historical_data[
        (historical_data['Year'] == year) &
        (historical_data['Month'] == month) &
        (historical_data['Product'] == current_row['Description'])
    ]
    
    if len(product_history) > 0:
        # Mid-month update: Calculate the difference from sum of all previous entries
        previous_total_value = product_history['ValueFOB'].sum()
        previous_total_volume = product_history['Volume'].sum()
        last_entry = product_history.iloc[-1]
        previous_working_days = last_entry['Running_Work_Days']
        working_days_delta = working_days - previous_working_days
        
        # Calculate the differences from current total minus sum of all previous weeks
        value_fob = float(current_row['Value_FOB_Total']) - previous_total_value
        volume = float(current_row['Volume_Total']) - previous_total_volume
        
        print(f"Mid-month update for {current_row['Description']}:")
        print(f"Current total: {current_row['Value_FOB_Total']}")
        print(f"Previous weeks total: {previous_total_value}")
        print(f"Previous working days: {previous_working_days}")
        print(f"Current working days: {working_days}")
        print(f"Delta days: {working_days_delta}")
        print(f"Value FOB delta: {value_fob}")
        print(f"Volume delta: {volume}")
    else:
        # New month: Use the current values directly
        working_days_delta = working_days
        value_fob = float(current_row['Value_FOB_Total'])
        volume = float(current_row['Volume_Total'])
        
        print(f"New month data for {current_row['Description']}:")
        print(f"Working days: {working_days}")
        print(f"Value FOB: {value_fob}")
        print(f"Volume: {volume}")
    
    # Calculate daily averages based on the period's values and working days
    value_fob_daily = value_fob / working_days_delta if working_days_delta > 0 else 0
    volume_daily = volume / working_days_delta if working_days_delta > 0 else 0
    
    print(f"Daily averages:")
    print(f"Value FOB daily: {value_fob_daily}")
    print(f"Volume daily: {volume_daily}\n")
    
    return {
        'ValueFOB': value_fob,
        'Volume': volume,
        'WorkingDaysDelta': working_days_delta,
        'ValueFOB_Daily': value_fob_daily,
        'Volume_Daily': volume_daily
    }

def check_if_week_exists(historical_data, year, month, week):
    """Check if data for this specific week already exists in historical data."""
    exists = len(historical_data[
        (historical_data['Year'] == year) &
        (historical_data['Month'] == month) &
        (historical_data['Week_Number_In_Month'] == week)
    ]) > 0
    return exists

# Create backup of historical files
print("Creating backup of historical files...")
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
exports_backup = f'datasets/historicals/backups/Brazil_Secex_Weekly_Exports_{timestamp}.csv'
imports_backup = f'datasets/historicals/backups/Brazil_Secex_Weekly_Imports_{timestamp}.csv'

shutil.copy2('datasets/historicals/Brazil_Secex_Weekly_Exports.csv', exports_backup)
shutil.copy2('datasets/historicals/Brazil_Secex_Weekly_Imports.csv', imports_backup)
print(f"Backups created:\n{exports_backup}\n{imports_backup}")

# Read the historical files
print("\nReading historical data...")
historical_exports = pd.read_csv('datasets/historicals/Brazil_Secex_Weekly_Exports.csv')
historical_imports = pd.read_csv('datasets/historicals/Brazil_Secex_Weekly_Imports.csv')

# Download and save the Excel file
print("Downloading Excel file from source...")
response = requests.get(EXCEL_URL)
if response.status_code == 200:
    # Save the Excel file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f'datasets/downloads/Setores_Produtos_{timestamp}.xlsx'
    with open(excel_filename, 'wb') as f:
        f.write(response.content)
    print(f"Excel file saved as: {excel_filename}")
    
    # Read the Excel file from memory for both tabs
    excel_data = BytesIO(response.content)
    
    # First read the header information to get working days and week
    header_info = pd.read_excel(excel_data, sheet_name='EXP', header=None, nrows=5)
    
    # Extract working days and date info from the header
    working_days_info = header_info.iloc[3, 0]
    RUNNING_WORK_DAYS = int(re.search(r':\s*(\d+)\s*dias', working_days_info).group(1))
    YEAR, MONTH = extract_date_info(working_days_info)
    
    # Extract week number from the header
    week_info = header_info.iloc[3, 0]
    WEEK_NUMBER = int(re.search(r'(\d+)ª\s*Semana', week_info).group(1))
    
    print(f"\nExtracted information:")
    print(f"Year: {YEAR}")
    print(f"Month: {MONTH}")
    print(f"Week Number: {WEEK_NUMBER}")
    print(f"Running Work Days: {RUNNING_WORK_DAYS}")
    
    # Check if this week's data already exists
    exports_exist = check_if_week_exists(historical_exports, YEAR, MONTH, WEEK_NUMBER)
    imports_exist = check_if_week_exists(historical_imports, YEAR, MONTH, WEEK_NUMBER)
    
    if exports_exist or imports_exist:
        print("\nWARNING: Data for this week already exists in historical files!")
        print("Exports exist:", exports_exist)
        print("Imports exist:", imports_exist)
        print("Skipping processing to avoid duplicate entries.")
        print("If you need to update this week's data, please remove the existing entries first.")
        exit(0)
    
    # Reset file pointer for next read
    excel_data.seek(0)
    
    # Now read both tabs
    current_exports = pd.read_excel(excel_data, sheet_name='EXP', header=None)
    current_imports = pd.read_excel(excel_data, sheet_name='IMP', header=None)
    
    # Process exports data
    current_exports = current_exports.iloc[8:].copy()  # Skip the first 8 rows
    current_exports.columns = ['Description', 'Value_FOB_Total', 'Value_FOB_Total_2024', 'Value_FOB_Daily', 'Value_FOB_Daily_2024', 
                             'Volume_Total', 'Volume_Total_2024', 'Volume_Daily', 'Volume_Daily_2024', 
                             'Price_2025', 'Price_2024', 'Price_Var', 'Volume_Var', 'Price_Var_Pct']
    current_exports = current_exports.reset_index(drop=True)
    
    # Process imports data
    current_imports = current_imports.iloc[8:].copy()  # Skip the first 8 rows
    current_imports.columns = ['Description', 'Value_FOB_Total', 'Value_FOB_Total_2024', 'Value_FOB_Daily', 'Value_FOB_Daily_2024', 
                             'Volume_Total', 'Volume_Total_2024', 'Volume_Daily', 'Volume_Daily_2024', 
                             'Price_2025', 'Price_2024', 'Price_Var', 'Volume_Var', 'Price_Var_Pct']
    current_imports = current_imports.reset_index(drop=True)
    
    # Apply the name replacements for exports
    current_exports['Description'] = current_exports['Description'].str.replace('Milho não moído, exceto milho doce', 'Corn', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Soja', 'Soybeans', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Algodão em bruto', 'Cotton', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Carne bovina fresca, refrigerada ou congelada', 'Beef', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Carne suína fresca, refrigerada ou congelada', 'Pork', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Carnes de aves e suas miudezas comestíveis, frescas, refrigeradas ou congeladas', 'Poultry', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Açúcares e melaços', 'Sugar', regex=False)
    current_exports['Description'] = current_exports['Description'].str.replace('Couro', 'Beef_Skin', regex=False)
    
    # Apply the name replacements for imports
    current_imports['Description'] = current_imports['Description'].str.replace('Trigo e centeio, não moídos', 'Wheat', regex=False)
    current_imports['Description'] = current_imports['Description'].str.replace('Adubos ou fertilizantes químicos (exceto fertilizantes brutos)', 'Fertilizers', regex=False)
    current_imports['Description'] = current_imports['Description'].str.replace('Inseticidas, rodenticidas, fungicidas, herbicidas, reguladores de crescimento para plantas, desinfetantes e semelhantes', 'Crop_Chemicals', regex=False)
    
    # Create new DataFrames with the same structure as historical
    formatted_exports = pd.DataFrame(columns=historical_exports.columns)
    formatted_imports = pd.DataFrame(columns=historical_imports.columns)
    
    print("\nProcessing export products...")
    # Process exports data
    for _, row in current_exports.iterrows():
        if pd.notna(row['Description']) and row['Description'] in ['Corn', 'Soybeans', 'Cotton', 'Beef', 'Pork', 'Poultry', 'Sugar', 'Beef_Skin']:
            # Calculate the weekly values using exports historical
            weekly_values = calculate_weekly_values(row, historical_exports, YEAR, MONTH, RUNNING_WORK_DAYS)
            
            new_row = {
                'Year': YEAR,
                'Month': MONTH,
                'Week_Number_In_Month': WEEK_NUMBER,
                'Date': datetime.now().strftime('%m/%d/%Y'),
                'Running_Work_Days': RUNNING_WORK_DAYS,
                'Product': row['Description'],
                **weekly_values  # Unpack the calculated values
            }
            formatted_exports = pd.concat([formatted_exports, pd.DataFrame([new_row])], ignore_index=True)
    
    print("\nProcessing import products...")
    # Process imports data
    for _, row in current_imports.iterrows():
        if pd.notna(row['Description']) and row['Description'] in ['Wheat', 'Fertilizers', 'Crop_Chemicals']:
            # Calculate the weekly values using imports historical
            weekly_values = calculate_weekly_values(row, historical_imports, YEAR, MONTH, RUNNING_WORK_DAYS)
            
            new_row = {
                'Year': YEAR,
                'Month': MONTH,
                'Week_Number_In_Month': WEEK_NUMBER,
                'Date': datetime.now().strftime('%m/%d/%Y'),
                'Running_Work_Days': RUNNING_WORK_DAYS,
                'Product': row['Description'],
                **weekly_values  # Unpack the calculated values
            }
            formatted_imports = pd.concat([formatted_imports, pd.DataFrame([new_row])], ignore_index=True)
    
    # Save the current formatted data
    current_exports_filename = f'datasets/current/current_exports_{datetime.now().strftime("%Y%m%d")}.csv'
    current_imports_filename = f'datasets/current/current_imports_{datetime.now().strftime("%Y%m%d")}.csv'
    
    formatted_exports.to_csv(current_exports_filename, index=False)
    formatted_imports.to_csv(current_imports_filename, index=False)
    
    print(f"\nCurrent formatted data saved to:")
    print(f"Exports: {current_exports_filename}")
    print(f"Imports: {current_imports_filename}")
    
    # Update historical files with new data
    updated_exports = pd.concat([historical_exports, formatted_exports], ignore_index=True)
    updated_imports = pd.concat([historical_imports, formatted_imports], ignore_index=True)
    
    # Sort the updated data by date
    updated_exports = updated_exports.sort_values(by=['Year', 'Month', 'Week_Number_In_Month'])
    updated_imports = updated_imports.sort_values(by=['Year', 'Month', 'Week_Number_In_Month'])
    
    # Save the updated data back to the original historical files
    updated_exports.to_csv('datasets/historicals/Brazil_Secex_Weekly_Exports.csv', index=False)
    updated_imports.to_csv('datasets/historicals/Brazil_Secex_Weekly_Imports.csv', index=False)
    
    print("\nHistorical files updated with new data")
    
    # Print the formatted data for verification
    print("\nFormatted exports data:")
    print(formatted_exports)
    print("\nFormatted imports data:")
    print(formatted_imports)
else:
    print(f"Failed to download file. Status code: {response.status_code}") 