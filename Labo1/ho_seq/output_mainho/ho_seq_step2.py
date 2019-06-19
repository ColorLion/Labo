import os.path
import openpyxl

def main():
    dup_data = 0
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
    ws1 = xlsx1['Sheet']

    # output
    xlsx2 = openpyxl.Workbook()
    ws2 = xlsx2.active

    data = []
    for i in ws1.rows:
        tmp = [i[2].value] + [i[3].value] + [i[4].value] + [i[5].value] \
                    + [i[6].value] + [i[7].value] + [i[8].value] + [i[9].value] \
                    + [i[10].value] + [i[11].value] + [i[12].value] + [i[13].value] \
                    + [i[14].value] + [i[15].value] + [i[16].value]
        if tmp not in data:
            data.append(tmp)
        else:
            dup_data += 1


    for i in range(len(data)):
        ws2.append(data[i])

    xlsx2.save(save_file_step2)
    print(dup_data)
main()