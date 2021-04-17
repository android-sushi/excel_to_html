import openpyxl as xl
import pathlib


def read_excel_messages():
    dir_path = pathlib.Path('Excel連絡事項/2021/04/202104_連絡事項.xlsx')
    messages = []

    wb = xl.load_workbook(dir_path)
    ws = wb['連絡事項']
    messages = [row[0].value for row in ws.iter_rows(min_row=2)]
    
    return messages


if __name__ == '__main__':
    messages = read_excel_messages()
    print(messages)