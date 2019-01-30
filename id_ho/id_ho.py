import xlrd
import os.path
import openpyxl

def extract_data(file, mho_dir):
    target = xlrd.open_workbook(file, formatting_info=True)

    xlsx1 = openpyxl.Workbook()
    mho_output = xlsx1.active

    ws = target.sheet_by_index(0)
    for i in range(ws.nrows):
        row = ws.row_values(i)
        # 주호판정
        if "주호" in row[17]:
            # 이순, 통, 호 자리수 변경
            # if str(row[8]) in r'\'':
            if type(row[8]) == float:
                if len(str(row[8]).split('.')[0]) == 1:             # 이순
                    row[8] = "00" + str(row[8]).split('.')[0]
                elif len(str(row[8]).split('.')[0]) == 2:
                    row[8] = "0" + str(row[8]).split('.')[0]
            else:
                row[8] = "10" + str(row[8]).replace("\'", "")
            if len(str(row[11]).split('.')[0]) == 1:
                row[11] = "0" + str(row[11]).split('.')[0]
            if len(str(row[13]).split('.')[0]) == 1:
                row[13] = "0" + str(row[13]).split('.')[0]
            mho_output.append(row)
    save_file_mho = mho_dir + file.split('.')[0] + "_mho.xlsx"
    xlsx1.save(save_file_mho)

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")
    mho_dir = work_dir + '\\' + "output_주호분리" + '\\'

    # output 저장 폴더 확인
    if os.path.isdir(mho_dir) == 0:
        os.mkdir(mho_dir)

    files = os.listdir(work_dir)

    for file in files:
        if len(file.split('.')) == 2:
            if file.split('.')[1] == 'xls':
                print(file)
                extract_data(file, mho_dir)
main()