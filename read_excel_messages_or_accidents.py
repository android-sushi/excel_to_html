import openpyxl as xl
import pathlib


def read_excel_messages_or_accidents(i):
    if i == 0:
        file_name = '連絡事項'
    else:
        file_name = 'ヒヤリハット'

    dir_path = pathlib.Path('Excel連絡事項・ヒヤリハット/2021/04/202104_'+file_name+'.xlsx')
    wb = xl.load_workbook(dir_path)
    ws = wb[file_name]
    return_list = [row[i].value for row in ws.iter_rows(min_row=2)]

    return return_list


if __name__ == '__main__':
    for i in range(2):
        return_list = read_excel_messages_or_accidents(i)
        print(return_list)