#!/usr/bin/python3

import sys
import os
import yaml
import pandas
import xlsxwriter
import copy

def flatten(fileDict):
    flattened_data = []
    for key, value in fileDict.items():
        for k1, v1 in value['end_ports'].items():
            if k1 != 'circuits': # remove references to bundled circuits in CID pathway
                row = copy.deepcopy(v1)
                row['recordId'] = key
                row['end_port_name'] = k1

                # cleaning out data
                if value.get('description') != None:
                    row['description'] = value['description']
                if row.get('circuits') != None:
                    row.pop('circuits')

                flattened_data.append(row)
    return flattened_data

def getDataFrame(collection):
    return pandas.DataFrame(collection)

def main():
    # opens first filename in sys.argv as yaml and saves output to second filename in sys.argv as an Excel table

    with open(sys.argv[1]) as yaml_file:
        # convert yaml file to pandas dataframe
        pythonized_yaml = yaml.full_load(yaml_file)
        flattened_data = flatten(pythonized_yaml)
        df = getDataFrame(flattened_data)

        # setup Excel workbook object to export pandas dataframe
        writer = pandas.ExcelWriter(sys.argv[2], engine='xlsxwriter')
        df.to_excel(writer, sheet_name='circuits', index=False) # startrow=1, header=False, 

        # define shorthand
        wb = writer.book
        ws = writer.sheets['circuits']
        rc = df.shape[0]-1 # subtract one (1) from row count because xlsxwriter's index starts at zero (0)
        cc = df.shape[1]-1 # subtract one (1) from col count because xlswwriter's index starts at zero (0)

#        h1 = wb.add_format({
#            'bold': True,
#            'text_wrap': True,
#            'valign': 'top',
#            'fg_color': '#D7E4BC',
#            'border': 1})

#        for col, value in enumerate(df.columns.values):
#            ws.write(0, col, value, h1)

        # convert all cells to text format
        c1 = wb.add_format({'num_format': '@'})
        ws.set_column(0, cc, None, c1)

        # convert entire dataframe to a large table
        ws.add_table(0, 0, rc, cc, {'header_row': True})

        # restore column headers in table

        writer.save()

if __name__ == "__main__":
    main()
