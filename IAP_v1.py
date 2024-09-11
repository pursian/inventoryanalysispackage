import os
import pandas as pd

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

# Load the Excel file
current_dir = os.path.dirname(os.path.abspath(__file__))
file_name = 'CRNX LN2 Inventory XXXX.xlsx'  # Replace with the correct file path
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
                    data.append([dewar, rack, box, index, value])

# Convert the collected data into a DataFrame
output_df = pd.DataFrame(data, columns=["Dewar", "Rack", "Box", "Tile Index", "Contents"])

# Save to a new Excel file

output_file = os.path.join(current_dir, "XXXX_Output.xlsx")
output_df.to_excel(output_file, index=False)
