import fnmatch
import os
import zipfile

import mammoth
import openpyxl
from bs4 import BeautifulSoup
from pdf2docx import Converter
import re
import copy

# ZIP 파일 압축 풀기
extract_path = "/Users/gimsanha/개발/PythonWorkspace"

zipfile.ZipFile('/Users/gimsanha/개발/PythonWorkspace/sample.zip').extractall(extract_path)

zipfile_list = os.walk(extract_path + "/sample")


# os.rmdir(extract_path +"/__macosx")


# 파일 구분

xlsx_pattern = "*.xlsx"

pdf_pattern = "*.pdf"
docx_pattern = "*.docx"
html_pattern = "*.html"
standard_html_pattern = "standard.html"
standard_excel_pattern = "standard.xlsx"
result = []

template_re = re.compile('([1-9][0-9]*)\?')
template_path = dict()
max_index = 0


def dfs_children(obj, p):
    global max_index
    for i, c in enumerate(obj.children):
        if c.name == 'p' and template_re.match(c.string):
            print(template_re.findall(c.text)[0])
            index = int(template_re.findall(c.text)[0])
            template_path[index] = copy.deepcopy(p)
            max_index = max(max_index, index)
            print(template_path)
        p.append(i)
        dfs_children(c, p)
        p.pop()

def navigate(obj, p):
    for i in p:
        obj = obj.children[i]
    return obj.text


for root, dirs, files in zipfile_list:
    for name in files:
        # docs 파일이라면 html 로 변경
        if fnmatch.fnmatch(name, docx_pattern):
            docx_path = os.path.join(root, name)
            with open(docx_path, "rb") as docx_file:
                output = mammoth.convert_to_html(docx_file)
            with open(docx_path.replace("docx", "html"), "w") as html_file:
                html_file.write(output.value)

        # pdf 파일 -> docs 파일 -> html 파일로 변경
        elif fnmatch.fnmatch(name, pdf_pattern):
            pdf_path = os.path.join(root, name)
            docx_path = pdf_path.replace("pdf", "docx")
            cv = Converter(pdf_path)
            cv.convert(docx_path)
            cv.close()

            # docs 파일 -> html 파일
            with open(docx_path, "rb") as docx_file:
                output = mammoth.convert_to_html(docx_file)
            with open(docx_path.replace("docx", "html"), "w") as html_file:
                html_file.write(output.value)


        # excel 기준 파일 1? 2? 3? 4? 5? 6? 7? 데이터 위치 가져오기
        elif fnmatch.fnmatch(name, standard_html_pattern):
            html_name = os.path.join(root, name)

            html_name_html = open(html_name, "r")
            soup = BeautifulSoup(html_name_html, "html.parser", from_encoding='utf=8')

            dfs_children(soup, [])
        else:
            print("바꾸지 못하는 파일이 존재합니다")

# 날짜 추출
for root, dirs, files in zipfile_list:
    for name in files:
        if fnmatch.fnmatch(name, html_pattern):
            html_name = os.path.join(root, name)

            html_name_html = open(html_name, "r")
            soup = BeautifulSoup(html_name_html, "html.parser", from_encoding='utf=8')

            for i in range(1, max_index + 1):
                content = navigate(soup, template_path[i])
                print(content)
                # i-th content of html

# html -> excel 파일 변경

for root, dirs, files in zipfile_list:
    for name in files:
        if fnmatch.fnmatch(name, standard_excel_pattern):
            excel_path = os.path.join(root, name)
            excelfile = openpyxl.load_workbook(excel_path)
            sheet = excelfile.active
            sheet.append(content)
            excelfile.save("standard_fin.xlsx")
            print("마무리 되었습니다!")
