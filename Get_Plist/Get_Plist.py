import xlrd
import os.path
import sys
import openpyxl

# 데이터 추출 함수
def extract_plist(wb, xlsx, ax, fn, mov):
	# showing parameter
	'''
		wb = 추출할 대상 xls 파일
		xlsx = 추출한 데이터를 저장할 xlsx파일
		ax = xlsx의 active를 담은 변수
		fn = xlsx의 파일 명
		mov = xlsx를 저장할 위치
	'''
	
	# check index
	ws = wb.sheet_by_index(0)
	a = ['region_id', '성', '명', '출생년도', '간지']
	ax.append(a)
	
	# data row
	for i in range(ws.nrows):
		ws.row_values(i)
		row = ws.row_values(i)
		
		if row[22] == "":
			row[22] = '.'
		if row[23] == "":
			row[23] = '.'
		if row[26] == "":
			row[26] = '.'
		if type(row[24]) == str:
			row[24] = '.'
		else:
			row[24] = str(row[4] - row[24]).split('.')[0]

		if type(row[4]) == float:
			a = [str(row[1]), row[22], row[23], row[24], row[26]]
			ax.append(a)

	save = mov + '\\' + fn
	xlsx.save(save)
	

# 파일 이름을 만들기 위한 함수
def make_filename(xl):
	xlsx_name = xl.split('.')[0] + "_plist.xlsx"
	return xlsx_name

def main():
	folder = os.getcwd()
	notice = "현재 작업 위치 : "
	print(notice+folder)
	dir_name = 'output_plist'

	if os.path.isdir(dir_name) == False:
		os.mkdir(dir_name)
		mov = folder + '\\' + dir_name + '\\'
	else:
		mov = folder + '\\' + dir_name + '\\'

	filenames = os.listdir(folder)

	for filename in filenames:
		if len(filename.split('.')) == 2:
			if filename.split('.')[1] == 'xls':
				xl = filename
				print(xl)
				wb = xlrd.open_workbook(xl, formatting_info=True)

				xlsx = openpyxl.Workbook()
				ax = xlsx.active
				fn = make_filename(xl)
				
				extract_plist(wb, xlsx, ax, fn, mov)

main()
