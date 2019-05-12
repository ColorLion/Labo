import os.path
import xlrd
import openpyxl


def make_report(file):
    # 단성호적 파일 open
    target = xlrd.open_workbook(file, formatting_info=True)
    ws = target.sheet_by_index(0)
    # report에 사용할 변수 선언
    ri = []
    tong = []
    ho = []
    report_rows = []
    ho_all = 0
    ho_nob = 0
    ho_slv = 0
    ho_etc = 0

    # 연도 채우기 - year
    year = file.split('-')[0]
    for i in range(ws.nrows):
        row = ws.row_values(i)
        # 리의 수 채우기 - ri
        if row[8] not in ri:
            ri.append(row[8])
        # 통의 수 채우기 - tong
        if row[11] not in tong:
            tong.append(row[11])
        # 호의 수 채우기 - ho
        if row[13] not in ho:
            ho.append(row[13])
            '''
            # 디버깅용 코드(ho가 빈칸인 경우)
            if row[13] == "":
                print(i)
            '''
        # 총 주호의 수 채우기 - ho_all
        row = ws.row_values(i)
        if "주호" in row[17]:
            ho_all += 1
            # 주호 중에서 양반인 경우 채우기 - ho_nob
                # 주호가 성이 있는 경우
            if row[22] != "":
                ho_nob += 1
            # 주호 중 노비인 경우 채우기 - ho_slv
                # 주호가 주성명이 있는 경우 or 직역에 노, 비 한자가 있는 경우
            elif row[37] != "":
                ho_slv += 1
            # 주호 중 위 두가지 조건에 해당하지 않는 항목 채우기 - ho_etc
            else:
                ho_etc += 1

    # 전체 인원수 채우기 - people_all
    people_all = ws.nrows - 1

    # 배열 채우기
    report_rows.append(year)
    ri_data = str(len(ri) - 1) + "("
    for i in range(1, len(ri)):
        if type(ri[i]) == float:
            ri[i] = int(ri[i])
        if ri[i] == '':
            ri[i] = "Null"
        ri_data = ri_data + str(ri[i]) + ", "
    ri_data = ri_data + ")"
    report_rows.append(ri_data)

    tong_data = str(len(tong) - 1) + "("
    for i in range(1, len(tong)):
        if type(tong[i]) == float:
            tong[i] = int(tong[i])
        if tong[i] == '':
            tong[i] = "Null"
        tong_data = tong_data + str(tong[i]) + ", "
    tong_data = tong_data + ")"
    report_rows.append(tong_data)
    ho_data = str(len(ho) - 1) + "("
    for i in range(1, len(ho)):
        if type(ho[i]) == float:
            ho[i] = int(ho[i])
        if ho[i] == '':
            ho[i] = "Null"
        ho_data = ho_data + str(ho[i]) + ", "
    ho_data = ho_data + ")"
    report_rows.append(ho_data)
    report_rows.append(people_all)
    report_rows.append(ho_all)
    report_rows.append(ho_nob)
    report_rows.append(ho_slv)
    report_rows.append(ho_etc)

    print("리 ", ri)
    print("통 ", tong)
    print("호 ", ho)
    print("report ", report_rows)

    return report_rows

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")
    report_dir = work_dir + '\\' + "output_report" + '\\'
    location = input("report 대상 이름: ")

    # output 저장 폴더 확인
    if os.path.isdir(report_dir) == 0:
        os.mkdir(report_dir)

    files = os.listdir(work_dir)

    # report sheet 첫번째 행
    first_row = ['연도', '리(구성)', '통(구성)', '호(구성)', 'ALL', '주호', '주호_양반', '주호_노비', '주호_그외']
    print(first_row)

    # report 저장용 xlsx 파일
    xlsx = openpyxl.Workbook()

    # 연도별 파일의 rows 수 저장
    report_1 = xlsx.active
    report_1.title = 'Statistics'
    report_1.append(first_row)

    for file in files:
        if len(file.split('.')) == 2:
            if file.split('.')[1] == 'xls':
                print(file)
                report_rows = make_report(file)
                report_1.append(report_rows)

    # report 파일 저장
    xlsx.save(report_dir + '\\' + location + "_report.xlsx")

main()