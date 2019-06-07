def main():
    fileA = open('고려인명.txt', 'r', encoding='UTF-8')
    fileB = open('고려지명.txt', 'r', encoding='UTF-8')
    fileC = open('고려관직.txt', 'r', encoding='UTF-8')

    saveA = open('dic.txt', 'w', encoding='UTF-8')
    saveB = open('nng.txt', 'w', encoding='UTF-8')
    saveC = open('nnp.txt', 'w', encoding='UTF-8')


    # 고려인명.txt 처리
    for line in fileA.readlines():
        input_line = line.replace('\n', "")
        output_lineA = input_line + "\tNNP\n"
        output_lineC = input_line + "/NNP\t<인간>\n"
        saveA.write(output_lineA)
        saveC.write(output_lineC)

    # 고려지명.txt 처리
    for line in fileB.readlines():
        input_line = line.replace('\n', "")
        output_lineA = input_line + "\tNNP\n"
        output_lineC = input_line + "/NNP\t<지역>\n"
        saveA.write(output_lineA)
        saveC.write(output_lineC)

    # 고려관직.txt 처리
    for line in fileC.readlines():
        input_line = line.replace('\n', "")
        output_lineA = input_line + "\tNNG\n"
        output_lineB = input_line + "/NNG\t<관직>\n"
        saveA.write(output_lineA)
        saveB.write(output_lineB)


    fileA.close()
    fileB.close()
    fileC.close()
    saveA.close()
    saveB.close()
    saveC.close()

main()