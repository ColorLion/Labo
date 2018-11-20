import xlrd
import os.path
import openpyxl

# 데이터 추출 함수
def extract_data(file, nob_dir, sla_dir):
    target = xlrd.open_workbook(file, formatting_info=True)

    xlsx1 = openpyxl.Workbook()
    xlsx2 = openpyxl.Workbook()
    nob_output = xlsx1.active
    sla_output = xlsx2.active

    ws = target.sheet_by_index(0)

    nob_output.append(ws.row_values(0))
    sla_output.append(ws.row_values(0))

    # 계급 판별 후 저장 파일 결정
    for i in range(1, ws.nrows):
        row = ws.row_values(i)
        if "奴" in row[18] or "婢" in row[18]:
            sla_output.append(row)
        else:
            nob_output.append(row)

    # 저장
    save_file_nob = nob_dir + file.split('.')[0] + "_nob.xlsx"
    save_file_sla = sla_dir + file.split('.')[0] + "_sla.xlsx"
    xlsx1.save(save_file_nob)
    xlsx2.save(save_file_sla)

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")
    nob_dir = work_dir + '\\' + "output_양반" + '\\'
    sla_dir = work_dir + '\\' + "output_노비" + '\\'

    # output 저장 폴더 확인
    if os.path.isdir(nob_dir) == 0:
        os.mkdir(nob_dir)
    if os.path.isdir(sla_dir) == 0:
        os.mkdir(sla_dir)

    files = os.listdir(work_dir)

    for file in files:
        if len(file.split('.')) == 2:
            if file.split('.')[1] == 'xls':
                print(file)
                extract_data(file, nob_dir, sla_dir)

main()
