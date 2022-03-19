import sys
import os
from pyxlsb import open_workbook

def create_csv_directory(directory_name: str) -> None:
    # Check whether the specified path exists or not
    isExist = os.path.exists(directory_name)
    if not isExist:
        # Create a new directory because it does not exist 
        os.makedirs(directory_name)

def convert(xlsb_path: str, csv_path: str) -> None:
    basepath = xlsb_path
    newbasepath = csv_path
    file_name = 1
    for entry in os.listdir(basepath):
        if os.path.isfile(os.path.join(basepath, entry)):
            with open_workbook(os.path.join(basepath, entry)) as wb:
                for sheetname in wb.sheets:
                    csv_content = ''
                    with wb.get_sheet(sheetname) as sheet:
                        for row in sheet.rows():
                            values = [str(r.v) for r in row if r.v is not None]
                            csv_line = ','.join(values)
                            csv_content += csv_line + '\n'
                    csv_file_name = str(file_name) + '.csv'
                    file_name += 1
                    with open(os.path.join(newbasepath, csv_file_name), 'w', encoding='utf-8') as f:
                        f.write(csv_content)

if __name__ == '__main__':
    no_args = len(sys.argv)
    if (no_args != 2):
        raise Exception('Place all .xlsb files in one directory\n' + 'Usage: ConvertToCSV.py /path/to/xlsbs')
    else:
        xlsb_path = sys.argv[1]
        csv_path = './csvs'
        create_csv_directory(csv_path)
        convert(xlsb_path, csv_path)
