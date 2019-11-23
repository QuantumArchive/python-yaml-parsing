#!/usr/bin/python3

import yaml
import os
import copy
import pandas

def flatten(fileDict):
    flattened_data = []

    for recordId, record in fileDict.items():
        for end_port, end_port_record in record['end_ports'].items():

            if end_port == 'circuits':
                continue

            row = copy.deepcopy(end_port_record)
            row['recordId'] = recordId
            row['end_port_name'] = end_port

            # Cleaning out data
            if record.get('description') != None:
                row['record_description'] = record['description']

            flattened_data.append(row)
    return flattened_data

def getDataFrame(collection):
    return pandas.DataFrame(collection)

def main():
    # going to have to supply with the name of the yaml you care about
    fileName = 'Untitled2'
    cwd = os.getcwd()
    with open("%s/assets/%s.yaml" % (cwd, fileName)) as file:
        pythonized_yaml = yaml.full_load(file)
        flattened_data = flatten(pythonized_yaml)
        data_frame = getDataFrame(flattened_data)
        data_frame.to_csv("%s/assets/%s.csv" % (cwd, fileName))

if __name__ == "__main__":
    main()
