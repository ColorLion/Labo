import xlrd
import os
import openpyxl

def open_xls(excel, filename, xlsx):
    # xls 파일을 열기 위해 사용
    sheets = excel.nsheets
    fn = filename.split('.')[0]

    for sheet in range(sheets):
        sheet_name = excel.sheet_by_index(sheet)
        sn = excel.sheet_names()


def main():
    # Extract_MainHo.xlsx 첫행
    first_row = ['항목명', 'data_type', 'max_length']

    # 작업 위치 지정
    folder = os.getcwd()
    notice = "현재 작업 위치 : "
    print(notice + folder)

    # 출력물 위치 지정
    folder = os.getcwd()
    dir_name = 'output_metadata'
    if os.path.isdir(dir_name) == False:
        os.mkdir(dir_name)
    mov = folder + '\\' + dir_name + '\\'

    name = mov + 'Extract_MainHo.xlsx'

    filenames = os.listdir(folder)
    for filename in filenames:
        if len(filename.split('.')) == 2:
            if filename.split('.')[1] == 'xls':
                print(filename)
                first_row.append(filename.split('.')[0])
                excel = xlrd.open_workbook(filename, formatting_info=True)
                xlsx = open_xls(excel, filename, xlsx)

main()