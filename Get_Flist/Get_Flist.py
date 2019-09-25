import xlrd
import os.path
import sys
import openpyxl

# global variable
main_ho = []

# 데이터 추출 함수
def search_row(wb, xlsx, ax, fn):
	global main_ho
	ws = wb.sheet_by_index(0)

	i = 1
	for i in range(ws.nrows):
		row = ws.row_values(i)
		if row[17] == "주호":
			extract_juho(row, ax)
		else:
			number_stats(row[17], ax)

	# save output file
	xlsx.save(fn)

# 주호에 대한 정보 추출
def extract_juho(row, ax):
	global main_ho
	main_ho = []

	# 호 아이디 뒷자리
	if type(row[11]) == float and type(row[13]) == float:
		sub_id = '{0:0>2}{1:0>2}'.format(int(row[11]), int(row[13]))
	elif type(row[11]) == float and type(row[13]) != float:
		sub_id = '{0:0>2}{1:0>2}'.format(int(row[11]), row[13])
	elif type(row[11]) != float and type(row[13]) == float:
		sub_id = '{0:0>2}{1:0>2}'.format(row[11], int(row[13]))
	else:
		sub_id = '{0:0>2}{1:0>2}'.format(row[11], row[13])
	# ho_id
	ho_id = str(int(row[4]))+str(row[2])+str(sub_id)

	# 계급 판별
	# 0 = 빈칸, 1 = 양반, 2 = 노비
	if "奴" in row[18] or "婢" in row[18]:
		stratum = 2
	elif row[18] == "":
		stratum = 0
	else:
		stratum = 1

	# 호 id
	main_ho.append(ho_id)
	# 성
	main_ho.append(row[22])
	# 명
	main_ho.append(row[23])
	# 주호의 region_id
	main_ho.append(str(row[1]))
	# 계급
	main_ho.append(stratum)
	# 부명
	main_ho.append(row[46])
	# 부명_한자
	main_ho.append(row[45])
	# 조명
	main_ho.append(row[55])
	# 조명_한자
	main_ho.append(row[54])
	# 주성명
	main_ho.append(row[42])
	# 주성명_한자
	main_ho.append(row[41])

	# sub_info
	main_ho.append(0)
	main_ho.append(0)
	main_ho.append(0)
	main_ho.append(0)

	ax.append(main_ho)

# 호의 구성원에 대한 정보
def number_stats(row, ax):
	row_pointer = ax.max_row
	edit_1 = 'L' + str(row_pointer)
	value_1 = ax[edit_1].value
	edit_2 = 'M' + str(row_pointer)
	value_2 = ax[edit_2].value
	edit_3 = 'O' + str(row_pointer)
	value_3 = ax[edit_3].value
	edit_4 = 'N' + str(row_pointer)
	value_4 = ax[edit_4].value
	# 에러 처리) 주호가 없는 가족 구성원이 존재 할 시
	if type(value_1) == str and type(value_2) == str and type(value_3) == str and type(value_4) == str:
		return
	# 처
	if row == "처" or row == "첩" or row == "후첩":
		ax[edit_1] = value_1 + 1
	# 자녀
	if row == "자" or row == "녀":
		ax[edit_2] = value_2 + 1
	# 노비
	if row == "노비":
		ax[edit_3] = value_3 + 1
	# 가족
	else:
		ax[edit_4] = value_4 + 1	

def make_filename(xl, mov):
	filename = mov + xl.split('.')[0] + "_flist.xlsx"
	return filename

def main():
	# 첫 번째 행
	first_row = ['family_id', '성', '명', 'region_id', '계급', '부명', '부명_한자', '조명', '조명_한자', '주성명', '주성명_한자','처', '자녀', '가족', '노비']
	# 작업 위치 지정
	folder = os.getcwd()
	notice = "현재 작업 위치 : "
	print(notice+folder)
	dir_name = 'output_flist'

	# 출력물 위치 지정
	if os.path.isdir(dir_name) == False:
		os.mkdir(dir_name)
		mov = folder + '\\' + dir_name + '\\'
	else:
		mov = folder + '\\' + dir_name + '\\'

	filenames = os.listdir(folder)
	for filename in filenames:
		if len(filename.split('.')) == 2:
			if filename.split('.')[1] == 'xls':
				# open target file
				xl = filename
				print(xl)
				wb = xlrd.open_workbook(xl, formatting_info=True)
				fn = make_filename(xl, mov)

				# open output file
				xlsx = openpyxl.Workbook()
				ax = xlsx.active
				ax.append(first_row)
				ax.column_dimensions['A'].width = 12
				ax.column_dimensions['D'].width = 12
				ax.column_dimensions['K'].width = 12
				
				search_row(wb, xlsx, ax, fn)

main()