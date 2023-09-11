import os
import time
import pandas
import json
from pathlib import Path
from art import tprint


def excel_to_json(file_path='gov.xlsx'):
    if file_path == "":
        file_path = "gov.xlsx"
    if Path(file_path).is_file() and Path(file_path).suffix == '.xlsx':
        json_type = json.loads(pandas.read_excel(file_path).to_json())
        file_name = Path(file_path).stem
        print(f"[+] Original file: {Path(file_path).stem}")
        print("[+] Processing...")
        with open(f'{file_name}.json', 'w') as file:
            json.dump(json_type, file, indent=2)
        return "convert successfully"
    else:
        return "There is no file at the given path"


def main():
    os.system("color 40")
    time.sleep(0.2)
    tprint('EXCEL<<TO<<JSON', font='cybermedum')
    path_to_file = input("path to excel file: \n")
    print(excel_to_json(path_to_file))
    s = input()


if __name__ == '__main__':
    main()
