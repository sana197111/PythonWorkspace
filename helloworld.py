from re import S
import mammoth
import zipfile
import os, sys
from pdf2docx import Converter
import fnmatch 
import openpyxl
from bs4 import BeautifulSoup
import re

# ZIP 파일 압축 풀기

extract_path = "/Users/gimsanha/개발/PythonWorkspace"
zipfile.ZipFile('/Users/gimsanha/개발/PythonWorkspace/sample.zip').extractall(extract_path)
zipfile_list = os.walk(extract_path+"/sample")

# os.rmdir(extract_path +"/__macosx")

# 파일 구분
xlsx_pattern = "*.xlsx"
pdf_pattern = "*.pdf"
docx_pattern = "*.docx"
html_pattern = "*.html"
standard_html_pattern = "standard.html"
standard_excel_pattern = "standard.xlsx"
result = []

for root, dirs, files in zipfile_list:
    for name in files:
            # docs 파일이라면 html 로 변경
            if fnmatch.fnmatch(name, docx_pattern):
                docx_path = os.path.join(root, name)
                with open(docx_path, "rb") as docx_file:
                    output = mammoth.convert_to_html(docx_file)
                with open(docx_path.replace("docx","html"), "w") as html_file:
                    html_file.write(output.value)

            # pdf 파일 -> docs 파일 -> html 파일로 변경
            elif fnmatch.fnmatch(name, pdf_pattern):
                pdf_path = os.path.join(root, name)
                docx_path = pdf_path.replace("pdf","docx")
                cv = Converter(pdf_path)
                cv.convert(docx_path)
                cv.close()

                # docs 파일 -> html 파일
                with open(docx_path, "rb") as docx_file:
                    output = mammoth.convert_to_html(docx_file)
                with open(docx_path.replace("docx","html"), "w") as html_file:
                    html_file.write(output.value)


            #excel 기준 파일 1? 2? 3? 4? 5? 6? 7? 데이터 위치 가져오기        
            elif fnmatch.fnmatch(name, standard_html_pattern):
                html_name = os.path.join(root,name)

                html_name_html = open(html_name, "r")
                soup = BeautifulSoup(html_name_html, "html.parser", from_encoding='utf=8')

            else:
                print("바꾸지 못하는 파일이 존재합니다")



        # re.match("<*>", objs_text)
        # p = re.compile(r'^<\d+$>')
        # mc = p.findall(objs_text)
        # print(mc)

# sibling 판단 후, 몇번째 자식인지 결정하는 함수
def obj_next(objs):
    if objs.find_next_siblings == None :
        return
    else:
        objs_text = objs.find_next_siblings

def obj_prev(objs):
    if objs.find_previous_siblings == None :
        return
    else:
        objs.find_previous_siblings


# selector 뽑아내는 함수
def obj_path(objs):
    if objs.parent == None :
        return

    objs = objs.parent
    obj_path(objs)

    # print(objs)

# 기준 파일에 필요한 노드의 selector 데이터 저장
standard_list = ["1?","2?","3?","4?","5?","6?","7?","8?","9?","10?","11?","12?","13?","14?","15?","16?","17?","18?","19?","20?"]

for obj in soup.find_all('p') : 
    for i in standard_list:
        if obj.text == i :
            print(obj)
            print(obj.parent)

print(soup.find('p').name)

            # obj.findParent
            # print(obj_next(obj.parent))
            # print(obj_prev(obj.parent))
            # obj_next(obj)
            # print(obj.parent.parent.find_previous_siblings)
            # print(obj.parent.parent.find_next_siblings)
            # print(obj.find_parent)
            # obj_path(obj)
            # obj_next(obj.parent.parent)

# print(soup.find_all("table"))

print(soup.select_one('table:nth-child(5) > tr:nth-child(2) > td:nth-child(1)'))
# body > table:nth-child(5) > tbody > tr:nth-child(2) > td:nth-child(1) > p > strong

# 날짜 추출
for root, dirs, files in zipfile_list:
    for name in files:
        if fnmatch.fnmatch(name, html_pattern):
            html_name = os.path.join(root,name)

            html_name_html = open(html_name, "r")
            soup = BeautifulSoup(html_name_html, "html.parser", from_encoding='utf=8')
            for obj in soup.find_all('p') : 
                if obj.text == "\d\d\d\d. \d\d. \d\d":
                    day = obj.text
                    print(day)


# html -> excel 파일 변경

for root, dirs, files in zipfile_list:
    for name in files:
        if fnmatch.fnmatch(name, standard_excel_pattern):
            excel_path = os.path.join(root, name)
            excelfile = openpyxl.load_workbook(excel_path)
            sheet = excelfile.active
            sheet.append([day])
            excelfile.save("standard_fin.xlsx")