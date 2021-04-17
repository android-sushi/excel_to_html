import openpyxl as xl
import pathlib


def read_excel_accidents():
    dir_path = pathlib.Path('Excelヒヤリハット/2021/04/202104_ヒヤリハット.xlsx')
    messages = []

    wb = xl.load_workbook(dir_path)
    ws = wb['ヒヤリハット']
    accidents = [row[1].value for row in ws.iter_rows(min_row=2)]
    
    return accidents


if __name__ == '__main__':
    accidents = read_excel_accidents()
    print(accidents)