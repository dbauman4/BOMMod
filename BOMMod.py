import pandas as pd
import json
from tkinter import filedialog
from tkinter import Tk

def select_excel_file():
    Tk().withdraw()  # Avoid creating the main window
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xls")])
    return file_path

def apply_transformation(config_file, output_csv):
    excel_file = select_excel_file()
    if not excel_file:
        print("No file selected. Exiting.")
        return
    
    # Read Excel file
    df = pd.read_excel(excel_file, engine='xlrd', skiprows=1)
    
    # Keep only QTY and CATALOG columns
    # df = df[['CATALOG', 'QTY']]
    
    # Reorder columns
    df = df[['CATALOG', 'QTY']]
    
    

    # Read config file
    with open(config_file, 'r') as config_json:
        config_data = json.load(config_json)
    
    # Apply corrections from config
    df['CATALOG'] = df['CATALOG'].replace(config_data['corrections'])
    
    # Remove specified catalog numbers
    df = df[~df['CATALOG'].isin(config_data['remove'])]
    
    print(df)

    # Save to CSV
    df.to_csv(output_csv, index=False)

apply_transformation('config.json', 'output.csv')