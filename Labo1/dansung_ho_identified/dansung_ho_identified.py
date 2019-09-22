import os.path
import openpyxl

def extract_data(i, save_file):
    # 추출 데이터
        # 성/명(한자, 한글), 출생년도, 간지, 년도, 이순, 통, 호 면명, 리명
    a = [(str(i[5].value) + "-" + str(i[0].value))]       # ID
    a.append(i[5].value)            # 년도
    a.append(i[7].value)            # 면명
    a.append(i[11].value)           # 리명
    a.append(i[9].value)            # 이순
    a.append(i[12].value)           # 통
    a.append(i[14].value)           # 호
    a.append(i[15].value)           # 주호(한자)
    a.append(i[16].value)           # 주호(한글)
    a.append(i[21].value)           # 성(한자)
    a.append(i[22].value)           # 명(한자)
    a.append(i[23].value)           # 성(한글)
    a.append(i[24].value)           # 명(한글)
    a.append(i[18].value)           # 호내위상
    if type(i[25].value) == int:    # 출생년도
        a.append(i[5].value - i[25].value)
    else:
        a.append(i[25].value)
    a.append(i[26].value)           # 간지(한자)
    a.append(i[27].value)           # 간지(한글)

    save_file.append(a)


def extract_xlsx(file, main_hoid_output):
    # target file open
    xlsx = openpyxl.load_workbook(filename=file, data_only=True)
    sheet = xlsx['Sheet']

    for i in sheet.rows:
        # 첫 번째 줄 처리
        if i[0].value == "ID":
            continue
        else:
            extract_data(i, main_hoid_output)

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")
    main_id_dir = work_dir + '\\' + "output_ho" + '\\'

    save_file_ho = main_id_dir + "ho_identified.xlsx"

    # report나올 것
    xlsx1 = openpyxl.Workbook()
    ho_first = ["id", "연도", "면명", "리명", "이순", "통", "호", "주호(한자)", "주호(한글)", "성(한자)", \
               "명(한자)", "성(한글)", "명(한글)", "호내위상", "출생년도", "간지(한자)", "간지(한글)"]
    main_hoid_output = xlsx1.active
    main_hoid_output.append(ho_first)

    # output 저장 폴더 확인
    if os.path.isdir(main_id_dir) == 0:
        os.mkdir(main_id_dir)

    files = os.listdir(work_dir)

    for file in files:
        if len(file.split('.')) == 2:
            if file.split('.')[1] == 'xlsx':
                print(file)
                extract_xlsx(file, main_hoid_output)

    xlsx1.save(save_file_ho)

main()