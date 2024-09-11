#Inventory Analysis Package v1.1.0 (IAP)
#Ehsun Yazdani, Crinetics Pharmaceuticals
#v1.1.0 update includes scientist's initials extractions and better date formatting

import os
import pandas as pd
from datetime import datetime
import re

# Function to extract Dewar, Rack, and Box from sheet names
def extract_dewar_rack_box(sheet_name):
    parts = [part for part in sheet_name.split(' ') if part]  # Remove empty parts due to spaces
    
    if len(parts) == 2:  # Case: "BB8 Rack1Box9" - no spaces between Rack and Box
        dewar = parts[0]
        rack = parts[1][:5]  # Extract "RackX"
        box = parts[1][5:]   # Extract "BoxY"
    elif len(parts) == 3:  # Case: "BB8 Rack2 Box1" - spaces between Rack and Box
        dewar = parts[0]
        rack = parts[1]
        box = parts[2]
    elif len(parts) == 5:  # Case: "BB8 Rack 2 Box 4" - spaces between "Rack", number, and "Box", number
        dewar = parts[0]
        rack = parts[1] + parts[2]  # Combine "Rack" and "number"
        box = parts[3] + parts[4]   # Combine "Box" and "number"
    else:
        dewar, rack, box = 'Unknown', 'Unknown', 'Unknown'
    
    return dewar, rack, box

# Function to convert the 9x9 grid to a flat index (1 to 81)
def grid_to_index(row, col):
    return row * 9 + col + 1

# Function to handle initials followed by a date in YYYYMMDD
def extract_initials_and_date(text):
    match = re.search(r'\b([A-Z]{2})\s(\d{8})$', text)
    if match:
        initials = match.group(1)  # Extract initials
        date_str = match.group(2)  # Extract date as string (YYYYMMDD)
        formatted_date = datetime.strptime(date_str, '%Y%m%d').strftime('%Y-%m-%d')  # Format date as YYYY-MM-DD
        return initials, formatted_date
    return None, None

# Combined logic to handle date formats including "1DEC2022" and initials followed by a date
def extract_and_format_dates_v6(text):
    initials, formatted_date = extract_initials_and_date(text)
    if initials and formatted_date:
        return formatted_date
    
    date_patterns = [
        r'(\d{1,2}/\d{1,2}/\d{4})',  # MM/DD/YYYY or M/D/YYYY
        r'(\d{1,2}/\d{1,2}/\d{2})',  # MM/DD/YY or M/D/YY
        r'(\d{1,2}-\d{1,2}-\d{4})',  # MM-DD-YYYY
        r'(\d{1,2}-\d{1,2}-\d{2})',  # MM-DD-YY
        r'(\d{1,2}/\d{4})',          # MM/YYYY
        r'(\d{1,2}/\d{2})',          # MM/YY
        r'\((\d{1,2}/\d{1,2}/\d{2,4})\)', # Dates in parentheses, e.g., (5/23/24)
        r'(\d{2}[a-zA-Z]{3}\d{4})',   # DDMMMYYYY, e.g., 1DEC2022
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        if matches:
            match = matches[0]
            try:
                if len(match.split('/')) == 3 or len(match.split('-')) == 3:
                    if len(match.split('/')[-1]) == 4 or len(match.split('-')[-1]) == 4:  # MM/DD/YYYY or MM-DD-YYYY
                        return datetime.strptime(match, '%m/%d/%Y').strftime('%Y-%m-%d') if '/' in match else datetime.strptime(match, '%m-%d-%Y').strftime('%Y-%m-%d')
                    else:  # MM/DD/YY or MM-DD-YY
                        return datetime.strptime(match, '%m/%d/%y').strftime('%Y-%m-%d') if '/' in match else datetime.strptime(match, '%m-%d-%y').strftime('%Y-%m-%d')
                elif len(match.split('/')) == 2:
                    if len(match.split('/')[-1]) == 4:  # MM/YYYY
                        return datetime.strptime(match, '%m/%Y').strftime('%Y-%m')
                    else:  # MM/YY
                        return datetime.strptime(match, '%m/%y').strftime('%Y-%m')
                elif len(match) == 9:  # Handle DDMMMYYYY
                    return datetime.strptime(match, '%d%b%Y').strftime('%Y-%m-%d')
            except ValueError:
                continue
    return None

# Function to extract initials, handling stars like "JN*"
def extract_final_two_letter_initials_v3(text):
    initials, _ = extract_initials_and_date(text)
    if initials:
        return initials

    match = re.search(r'\b([A-Z]{2})\*?\b\s*$', text)  # Handle initials with or without a trailing star
    return match.group(1) if match else None

# Load the Excel file
current_dir = os.path.dirname(os.path.abspath(__file__))
file_name = 'CRNX LN2 Inventory R2D2.xlsx'  # Replace with the correct file path
file_path = os.path.join(current_dir, file_name)
excel_file = pd.ExcelFile(file_path)

# Extract sheet names and load each sheet for inspection
sheet_names = excel_file.sheet_names
sheet_data = {sheet: pd.read_excel(excel_file, sheet_name=sheet) for sheet in sheet_names}

# Initialize an empty list to store the data
data = []

# Loop through each relevant sheet and extract data
for sheet_name in sheet_names:
    if "Rack" in sheet_name and "Box" in sheet_name:  # Exclude summary sheet
        dewar, rack, box = extract_dewar_rack_box(sheet_name)
        sheet = sheet_data[sheet_name].iloc[1:, 1:10]  # Skip header, focus on 9x9 grid
        
        # Iterate through the 9x9 grid to capture tile data and its flat index
        for row in range(9):
            for col in range(9):
                value = sheet.iloc[row, col]
                index = grid_to_index(row, col)
                if pd.notna(value):  # Only include non-empty cells
                    # Extract and format dates with all formats handled correctly
                    formatted_dates = extract_and_format_dates_v6(str(value))
                    # Extract two-letter initials, handling stars and special cases like "PA 20240805"
                    initials = extract_final_two_letter_initials_v3(str(value))
                    data.append([dewar, rack, box, index, value, formatted_dates, initials])

# Convert the collected data into a DataFrame
output_df = pd.DataFrame(data, columns=["Dewar", "Rack", "Box", "Tile Index", "Contents", "Formatted Date", "Initials"])

# Save to a new Excel file
output_file = os.path.join(current_dir, "IAPv1.1.0_GONKY_Output.xlsx")
output_df.to_excel(output_file, index=False)
