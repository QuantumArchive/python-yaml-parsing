#!/usr/bin/python3

import sys
import os
import yaml
import pandas
import xlsxwriter
import copy

def flatten(fileDict):
    flattened_data = []

    for recordId, record in fileDict.items():
        for end_port, end_port_record in record['end_ports'].items():
            # remove references to bundled circuits in CID pathway
            if end_port == 'circuits':
                continue

            # construct next row
            row = copy.deepcopy(end_port_record)
            row['recordId'] = recordId
            row['end_port_name'] = end_port

            # protect high level CID description from being overwritten by a hop description
            if record.get('description') != None:
                row['record_description'] = record.get('description')

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
        headers = ['recordId', 'end_port_name', 'record_description', 'owner', 'circuit_id', 'circuit_label', 'serial', 'serial_num', 'serial_number',
                   'serial_id', 'site', 'cage', 'device', 'interface', 'channel', 'patch_panel', 'panel', 'customer_panel', 'ports', 'port', 'order_id',
                   'service_order', 'order_number', 'order_no', 'work_order', ]
        writer = pandas.ExcelWriter(sys.argv[2], engine='xlsxwriter')
        df.to_excel(writer, sheet_name='circuits', index=False, columns=headers)

        # define shorthand
        wb = writer.book
        ws = writer.sheets['circuits']
        rc = df.shape[0]
        cc = len(headers) - 1 # subtract one (1) from col count because xlswwriter's index starts at zero (0)

        # convert all cells to text format
        c1 = wb.add_format({'num_format': '@'})
        ws.set_column(0, cc, None, c1)

        # convert entire dataframe to a large table
        headers = [{'header': i} for i in headers]
        ws.add_table(0, 0, rc, cc, {'name': 'CIDs', 'columns': headers})

        writer.save()

if __name__ == "__main__":
    main()
