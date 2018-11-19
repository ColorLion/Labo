import codecs
import os
import re

# Tag: Regular Expression a
TAG_RE = re.compile(r'<[^>]+>')

'''
xml_file = codecs.open('kr_000r.xml', 'r', 'utf-8')
u = xml_file.read()

TAG_RE = re.compile(r'<[^>]+>')

a = TAG_RE.sub('', u)

txt = codecs.open('test.txt', 'w', encoding='utf8')

print(a)
print(a, file=txt)

txt.close()
'''
def convert_xml2txt(xml_file, txt_name):
    # print("convert_xml2txt")
    # open files
    xml_pointer = codecs.open(xml_file, 'r', 'utf-8')
    xml = xml_pointer.read()
    txt = codecs.open(txt_name, 'w', encoding='utf8')

    remove_tag = TAG_RE.sub('', xml)
    print(remove_tag, file=txt)
    txt.close()

def get_name(xml_file):
    txt_name = xml_file.split('.')[0] + ".txt"
    return txt_name
    
def main():
    # variable
    xml_cnt = 0
    pass_cnt = 0
    # print job directory
    job_dir = os.getcwd()
    notice = "작업 위치 : "
    print(notice + job_dir)

    # Get File name in job directory
    all_files = os.listdir(job_dir)
    for xml_file in all_files:
            if len(xml_file.split('.')) == 2:
                if xml_file.split('.')[1] == 'xml':
                    xml_cnt += 1
                    txt_name = get_name(xml_file)
                    convert_xml2txt(xml_file, txt_name)
                else:
                    pass_cnt += 1
                    print("pass file: " + xml_file)

    print("===============================")
    print("xml files: " + str(xml_cnt) + " / pass files: " + str(pass_cnt))
    input("press enter plz")

main()