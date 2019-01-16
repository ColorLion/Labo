import os.path
import openpyxl

def extract_data(file, id_dir):
    #save file open
    xlsx1 = openpyxl.Workbook()
    hoid_output = xlsx1.active

    #target file open
    #하나 배웠다, data_only안하면 cell에 써져있는 그대로 가져오고
    #data only하면 출력되는 모습 그대로 가져오는 듯 하다
    xlsx = openpyxl.load_workbook(filename=file, data_only=True)
    sheet = xlsx['Sheet']
    for i in sheet.rows:
        #a = [i[4].value]        #연도
        #a.append(i[8].value)    #이순
        #a.append(i[11].value)   #통(번호)
        #a.append(i[13].value)   #호(번호)
        a = [str(i[4].value) + str(i[8].value) \
             + str(i[11].value) + str(i[13].value)]   #hoid
        #a.append(i[16].value)   #주호
        #a.append(i[18].value)   #호내위상
        a.append(i[19].value)   #직역
        a.append(i[22].value)   #성
        a.append(i[23].value)   #명
        if type(i[24].value) == int:
            a.append(i[4].value - i[24].value)   #출생년도
        else:
            a.append(i[24].value)
        a.append(i[26].value)   #간지
        a.append(i[42].value)   #주성명, 노비인 경우에 사용
        a.append(i[46].value)   #부명
        a.append(i[50].value)   #모명
        a.append(i[55].value)   #조명
        a.append(i[59].value)   #증조명
        a.append(i[63].value)   #외조명
        hoid_output.append(a)

    save_file_hoid = id_dir + file.split('.')[0] + "_hoid.xlsx"
    xlsx1.save(save_file_hoid)

def main():
    # 작업 위치 및 저장 위치 정의
    work_dir = os.getcwd()
    print(work_dir + "현재 작업 위치")
    id_dir = work_dir + '\\' + "output_식별데이터" + '\\'

    # output 저장 폴더 확인
    if os.path.isdir(id_dir) == 0:
        os.mkdir(id_dir)

    files = os.listdir(work_dir)

    for file in files:
        if len(file.split('.')) == 2:
            if file.split('.')[1] == 'xlsx':
                print(file)
                extract_data(file, id_dir)
main()