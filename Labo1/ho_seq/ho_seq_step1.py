import os.path
import openpyxl

main_ho = []

def extract_data(i, save_file):
    # 추출 데이터
        # 성/명(한자, 한글), 출생년도, 간지, 년도, 이순, 통, 호 면명, 리명
    a = [i[4].value]                # 년도
    a.append(i[6].value)            # 면명
    a.append(i[10].value)           # 리명
    a.append(i[8].value)            # 이순
    a.append(i[11].value)           # 통
    a.append(i[13].value)           # 호
    a.append(i[14].value)           # 주호(한자)
    a.append(i[15].value)           # 주호(한글)
    a.append(i[20].value)           # 성(한자)
    a.append(i[21].value)           # 명(한자)
    a.append(i[22].value)           # 성(한글)
    a.append(i[23].value)           # 명(한글)
    a.append(i[17].value)           # 호내위상
    if type(i[24].value) == int:    # 출생년도
        a.append(i[4].value - i[24].value)
    else:
        a.append(i[24].value)
    a.append(i[25].value)           # 간지(한자)
    a.append(i[26].value)           # 간지(한글)

    save_file.append(a)
    #return a

def gather_static(file, i, save_file1):
    print("gather static")


def extract_xlsx(file, main_hoid_output, static_output):
    global main_ho
    # target file open
    xlsx = openpyxl.load_workbook(filename=file, data_only=True)
    sheet = xlsx['Sheet']

    # static_first = ["file_name", "ALL", "평민 이상", "노비", "공백", "x포함"]
    static_count = [file, 0, 0, 0, 0, 0]

    for i in sheet.rows:
        # 첫 번째 줄 처리
        if i[0].value == "原本":
            continue
        else:
            extract_data(i, main_hoid_output)

        static_count[1] += 1
        job = i[18].value
        if job == None:
            static_count[4] += 1
        elif '奴' in job or '婢' in job:
            static_count[3] += 1
        elif 'x' in job:
            static_count[5] += 1
        else:
            static_count[2] += 1

    static_output.append(static_count)

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")
    main_id_dir = work_dir + '\\' + "output_ho" + '\\'
    static_id_dir = work_dir + '\\' + "output_static" + '\\'

    save_file_ho = main_id_dir + "ho_id.xlsx"
    save_file_static = static_id_dir + "statics.xlsx"

    # report나올 것
    xlsx1 = openpyxl.Workbook()
    ho_first = ["연도", "면명", "리명", "이순", "통", "호", "주호(한자)", "주호(한글)", "성(한자)", \
               "명(한자)", "성(한글)", "명(한글)", "호내위상", "출생년도", "간지(한자)", "간지(한글)"]
    main_hoid_output = xlsx1.active
    main_hoid_output.append(ho_first)

    # 통계 자료 출력용
    xlsx2 = openpyxl.Workbook()
    static_output1 = xlsx2.active
    static_output1.title = "기본 통계자료"
    static_first = ["file_name", "ALL", "평민 이상", "노비", "공백", "x포함"]
    static_output1.append(static_first)

    # output 저장 폴더 확인
    if os.path.isdir(main_id_dir) == 0:
        os.mkdir(main_id_dir)
    if os.path.isdir(static_id_dir) == 0:
        os.mkdir(static_id_dir)

    files = os.listdir(work_dir)

    for file in files:
        if len(file.split('.')) == 2:
            if file.split('.')[1] == 'xlsx':
                print(file)
                extract_xlsx(file, main_hoid_output, static_output1)

    xlsx1.save(save_file_ho)
    xlsx2.save(save_file_static)

main()