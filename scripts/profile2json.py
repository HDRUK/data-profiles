import sys
import argparse
import json
import os
from collections import OrderedDict

import pandas as pd
import numpy as np
import xlrd

class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        else:
            return super(NpEncoder, self).default(obj)

def parse_wr_overview_sheet(filename, sheet_name='Overview', DATA=None):
    DATA = DATA if DATA else {}
    try:
        sheet = pd.read_excel(filename, sheet_name='Overview')
    except xlrd.biffh.XLRDError:
        sheet = parse_wr_overview_sheet(filename, 'Field Overview', DATA)
    
    sheet = pd.DataFrame(sheet)


    for i, row in sheet.iterrows():
        if not pd.isnull(row['Table']):
            TABLE = row['Table']
            DATA.setdefault(TABLE, OrderedDict([
                ('name', TABLE),
            ]))
            DATA[TABLE][row['Field']] = OrderedDict({
                'field': row.get('Field', ""),
                'type': row.get('Type', ""),
                'decription': row.get('Description', ""),
                'length': int(row.get('Max length', -1)),
                'rows': int(row.get('N rows', -1)),
                'rows_checked': int(row.get('N rows checked', -1)),
                'emptiness': row.get('Fraction empty', -1),
                'uniqeness': row.get('Fraction unique', -1),
                'statistics': {
                    'mean': row.get('Average', -1),
                    'standard_deviation': row.get('Standard Deviation', -1),
                    'min': row.get('Min', -1),
                    '25_percentile': row.get('25%', -1),
                    'median': row.get('Median', -1),
                    '75_percentile': row.get('75%', -1),
                    'max': row.get('Max', -1),
                }
                

            })
    
    return DATA

def parse_wr_table_sheet(filename, table_name):
    DATA = OrderedDict()
    # BUG: WhiteRabbit Profile truncates sheet name :O
    sheets = pd.read_excel(filename, None)
    for sheet_name in sheets.keys():
        if table_name.startswith(sheet_name):
            table_name = sheet_name
            break
    
    sheet = pd.read_excel(filename, sheet_name=table_name)
    sheet = pd.DataFrame(sheet)

    # loop through columns two at a time
    column_length = len(sheet.columns)
    for i in range(1,column_length+1,2):
        # extract first two columns
        value_freq = sheet.iloc[:, :i+1]
        # extract first field
        values = value_freq.iloc[:, i-1]
        freqs = value_freq.iloc[:, i]

        not_nan_values = values.dropna()
        not_nan_freqs = freqs.dropna()
        if len(not_nan_values) or len(not_nan_freqs):
            DATA.setdefault(values.name, OrderedDict())
            
            for index, value in values.items():
                value = None if pd.isnull(value) else value
                DATA[values.name][value] = None if pd.isnull(freqs[index]) else freqs[index]
                if str(value).startswith('List truncated'):
                    break
    return DATA

def merge_frequency_table(overview_table, frequency_table):
    for field, value in frequency_table.items():
        overview_table[field]['frequencies'] = value
    return overview_table

def parse_wr_report(filename):
    data  = parse_wr_overview_sheet(filename)
    tables = list(data.keys())

    for table in tables:
        freq_table = parse_wr_table_sheet(filename, table)
        data[table] = merge_frequency_table(data[table], freq_table)
    return data

def read_json(filename):
    with open(filename, 'r') as file:
        return json.load(file)

def write_json(data, filename=None, indent=2):
   with (open(filename, 'w') if filename else sys.stdout) as file:
        json.dump(data, file, indent=indent, cls=NpEncoder)

def main(args):
    PID = args['pid']
    FILE = args['FILE']
    OUTPUT = args['out']
    if not PID and FILE.startswith('profiles/'):
        _, PID, filename = FILE.split('/')
        OUTPUT = os.path.join("profiles", PID)
    
    OUT_FILENAME = PID + ".white_rabbit.profile.json"
    if not OUTPUT:
        OUTPUT = os.path.join("./", OUT_FILENAME)
    else:
        OUTPUT = os.path.join(OUTPUT, OUT_FILENAME)
    
    data = parse_wr_report(args['FILE'])
    data = {
        'pid': PID,
        'dataClasses': data
    }
    write_json(data, OUTPUT)



if __name__ == '__main__':
    parser = argparse.ArgumentParser(prog="profile2json", 
                                     description='Process data profile reports into a json format')
    parser.add_argument('FILE', type=str, help='Data profile file')
    parser.add_argument('--pid', type=str, help="Dataset PID")
    parser.add_argument('--format', type=str, help="Profile format")
    parser.add_argument('--out', type=str, help="Output path")

    args = parser.parse_args()
    print(vars(args))

    sys.exit(main(vars(args)))
