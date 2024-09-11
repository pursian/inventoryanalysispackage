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

# Function to extract and format dates from a given string
def extract_and_format_dates(text):
    # Define several date patterns that might appear in the tiles
    date_patterns = [
        r'(\d{1,2}/\d{1,2}/\d{4})',   # M/D/YYYY or MM/DD/YYYY
        r'(\d{1,2}/\d{1,2}/\d{2})',   # M/D/YY or MM/DD/YY
        r'(\d{1,2}-\d{1,2}-\d{2,4})', # M-D-YY, MM-DD-YY or MM-DD-YYYY
        r'\((\d{1,2}/\d{1,2}/\d{2,4})\)', # Dates in parentheses, e.g., (5/23/24)
        r'(\d{2}[a-zA-Z]{3}\d{4})',   # DDMMMYYYY, e.g., 17May2023
    ]
    
    # List to hold formatted dates
    for pattern in date_patterns:
        matches = re.findall(pattern, text)
        if matches:
            # Process the first matched date
            match = matches[0]
            try:
                if '/' in match or '-' in match:
                    if len(match.split('/')[-1]) == 4 or len(match.split('-')[-1]) == 4:
                        return datetime.strptime(match, '%m/%d/%Y').strftime('%Y-%m-%d') if '/' in match else datetime.strptime(match, '%m-%d-%Y').strftime('%Y-%m-%d')
                    else:
                        return datetime.strptime(match, '%m/%d/%y').strftime('%Y-%m-%d') if '/' in match else datetime.strptime(match, '%m-%d-%y').strftime('%Y-%m-%d')
                elif len(match) == 9:  # Handle DDMMMYYYY
                    return datetime.strptime(match, '%d%b%Y').strftime('%Y-%m-%d')
            except ValueError:
                continue
    return None

# Load the Excel file
current_dir = os.path.dirname(os.path.abspath(__file__))
file_name = 'CRNX LN2 Inventory C3PO.xlsx'  # Replace with the correct file path
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
                    # Extract and format dates
                    formatted_dates = extract_and_format_dates(str(value))
                    if formatted_dates:  # If a date is found, append it to the data
                        data.append([dewar, rack, box, index, value, formatted_dates])
                    else:
                        data.append([dewar, rack, box, index, value, None])

# Convert the collected data into a DataFrame
output_df = pd.DataFrame(data, columns=["Dewar", "Rack", "Box", "Tile Index", "Contents", "Formatted Dates"])

# Save to a new Excel file
output_file = os.path.join(current_dir, "IAPv1.0.2_C3PO_Output.xlsx")
output_df.to_excel(output_file, index=False)
