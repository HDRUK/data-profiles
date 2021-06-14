import os
import json
import pathlib
import pandas as pd
from pandas_profiling import ProfileReport

DATA_FOLDER = "sample-data"

#######################################################################
# # WhiteRabbit Export to JSON
WR_SCAN_EXCEL_FILE = "sample-data/reports/whiterabbit/ScanReport.xlsx"
wr_scan_report = pd.read_excel(WR_SCAN_EXCEL_FILE, None)

# Create missing folders if needed
pathlib.Path(os.path.join(DATA_FOLDER, "reports", "whiterabbit")).mkdir(parents=True, exist_ok=True) 

for sheet_name in wr_scan_report.keys():
    print("Reading WhiteRabbit Sheet:", sheet_name)
    sheet = pd.read_excel(WR_SCAN_EXCEL_FILE, sheet_name=sheet_name)
    folder = os.path.dirname(WR_SCAN_EXCEL_FILE)
    json_filename = os.path.splitext(sheet_name)[0] + ".json"
    print("Writing WhiteRabbite Profile:", json_filename)
    with open(os.path.join(folder,json_filename), 'w') as json_file:
        json.dump(sheet.to_dict(), json_file)

#######################################################################

#######################################################################
# Pandas-Profiling export to JSON
SAMPLE_DATA_FILES = [f for f in os.listdir(DATA_FOLDER) if os.path.isfile(os.path.join(DATA_FOLDER,f)) and f not in ['.DS_Store', 'sample-data.zip']]

# Create missing folders if needed
pathlib.Path(os.path.join(DATA_FOLDER, "reports", "pandas-profiling")).mkdir(parents=True, exist_ok=True) 

for file in SAMPLE_DATA_FILES:
    print("Reading file:", file)
    df = pd.read_csv(os.path.join(DATA_FOLDER,file))
    print("Generating Profile:", file)
    profile = ProfileReport(df, title=file)
    profile_filename = os.path.splitext(file)[0]
    
    print("Writing profile report:", os.path.join(DATA_FOLDER, "reports", "pandas-profiling", profile_filename))
    # Write HTML report
    profile_filename_html = profile_filename + ".html"
    profile.to_file(os.path.join(DATA_FOLDER, "reports", "pandas-profiling",profile_filename_html))
    # JSON Profile
    profile_filename_json = profile_filename + ".json"
    with open(os.path.join(DATA_FOLDER, "reports", "pandas-profiling", profile_filename_json), 'w') as profile_json:
        profile_json.write(profile.to_json())

#######################################################################