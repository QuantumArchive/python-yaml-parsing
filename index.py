#!/usr/bin/python3

import yaml
import os
import copy
import pandas

def flatten(fileDict):
    flattened_data = []
    for key, value in fileDict.items():
        for k1, v1 in value['end_ports'].items():
            row = copy.deepcopy(v1)
            row['recordId'] = key
            row['description'] = value['description']
            row['end_port_name'] = k1
            flattened_data.append(row)
    return flattened_data

def getDataFrame(collection):
    return pandas.DataFrame(collection)

def main():
    # going to have to supply with the name of the yaml you care about
    
    cwd = os.getcwd()
    with open("%s/assets/Untitled.yaml" % cwd) as file:
        pythonized_yaml = yaml.full_load(file)
        flattened_data = flatten(pythonized_yaml)
        data_frame = getDataFrame(flattened_data)
        data_frame.to_csv("%s/assets/Untitled.csv" % cwd)

if __name__ == "__main__":
    main()
