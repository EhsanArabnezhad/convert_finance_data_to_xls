import os
import json
import sys
import argparse
import pathlib
from class_module import ProcessData

AIDAExport = False


def is_file(file_name):

    if not os.path.isfile(file_name) or not pathlib.Path(file_name).suffix == '.xlsx':
        raise ValueError("You must provide a valid excel filename as parameter")


if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    # parser.add_argument("-id", "--id", help="ID of the company")
    parser.add_argument("-od", "--od", help="Output directory")
    parser.add_argument("-ot", "--ot", help="Output type (currency or not)", action='store_true')
    known_args, file_to_open = parser.parse_known_args()

    if len(file_to_open) != 1:  
        print('You need to provide name of the excel file like "xxx.xlsx" (mandatory) '
              'location to save json (-od)')
        sys.exit(1)

    try:
        fileName = file_to_open[0]
        is_file(fileName)
    except ValueError as e:
        print("You must provide a valid excel filename as parameter")
        sys.exit(1)

    with open('questionario_indicatori_crisi_aziendale.json') as json_file:
        json_data = json.load(json_file)

    dir_to_save = known_args.__dict__["od"]
    aida = known_args.__dict__["ot"]
    
    # print(known_args.__dict__["od"])
    # print(known_args.__dict__["ot"])
    if not dir_to_save:
        dir_to_save = 'json_dir'
        print('Results will be saved in', dir_to_save)
    if aida:
        AIDAExport = True
    else:
        AIDAExport = False

    company_class = ProcessData(fileName, json_data, dir_to_save, AIDAExport)
    msg = company_class.convert_data()
    print(msg)
