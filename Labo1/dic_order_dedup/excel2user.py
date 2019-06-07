import openpyxl

f = open('sql.txt', 'w', encoding='UTF8')

xlsx = openpyxl.load_workbook(filename='ID.xlsx')

sheet = xlsx['ID']
sql = 'test'
for i in sheet.rows:
    sql = "INSERT INTO `global_id` (`global_ID`, `first_name`, `last_name`, `birth`, `gan`)  VALUES (`" \
          + str(i[0].value) + "`,`" + i[1].value + "`,`" + i[2].value + "`,`" \
          + i[3].value + "`,`" + i[4].value + "`)"
    f.write(sql)

xlsx.close
f.close()