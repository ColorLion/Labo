import os.path
import openpyxl

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")

    step2_dir = work_dir + '\\' + "step2" + '\\'
    save_file_step2 = step2_dir + "remove_dup.xlsx"

    # mkdir
    if os.path.isdir(step2_dir) == 0:
        os.mkdir(step2_dir)

    # input
    xlsx1 = openpyxl.load_workbook(filename='mho_id.xlsx', data_only=True)
    ws1 = xlsx1.active

    # output
    xlsx2 = openpyxl.Workbook()
    ws2 = xlsx2.active

    values = []
    for i in range(2, ws1.max_row + 1):
        if ws1.cell(row=i, column=3).value in values:
            pass  # if already in list do nothing
        else:
            values.append(ws1.cell(row=i, column=3).value)

    for value in values:
        ws2.append([value])

    xlsx2.save(save_file_step2)
main()