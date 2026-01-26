import os
from docx import Document

filename = ''
folder_name = 'CODE'
folder_name_len = len(folder_name) + 3
sample_file = 'template.docx'
type_file = []
prodl_name = ""
with open("..\\settings.txt", "r",  encoding='utf-8', errors='ignore') as f:
    tmp = f.read().split("\n")
    type_file = tmp[0].replace(" ", "").split(",")
    filename = "..\\" + str(tmp[1]) + " 12 01 (Текст программы)"
    prodl_name = tmp[2]

def check(file_name):
    if file_name.split(".")[-1] in type_file:
        return True
    return False

def read_file(filename):
    with open(filename, "r",  encoding='utf-8', errors='ignore') as f:
        return f.read()

folder = []
for i in os.walk("..\\" + folder_name):
    folder.append(i)

paths = []
for address, dirs, files in folder:
    for file in files:
        if check(file):
            paths.append(address + '\\' + file)

total_code_file = []
print(paths)
for i, path in enumerate(paths):
    if i == 0:
        catalog = path.split('\\')[1]
        # total_code_file.append('Каталог ' + catalog + '\n')
    if path[folder_name_len:].split('\\')[1] != catalog:
        catalog = path[folder_name_len:].split('\\')[1]
        total_code_file.append('Каталог ' + catalog + '\n')

    total_code_file.append('Файл ' + path[folder_name_len:] + '\n')
    total_code_file.append(read_file(path))
    total_code_file += '\n'

doc = Document(sample_file)
# for style in doc.styles:
# #     if style.type == 1 and not style.builtin:  # paragraph styles, кастомные
#       print(style.name)
for p in doc.paragraphs:
    if '<КОД ПРОГРАММЫ>' in p.text:
        p.text = ''
        for line in total_code_file:
            if line.rstrip() > '' and line.split()[0] == 'Каталог':
                p.insert_paragraph_before(line.rstrip(), 'Heading 2')
            elif line.rstrip() > '' and line.split()[0] == 'Файл':
                p.insert_paragraph_before(line.rstrip(), 'Heading 3')
            else:
                p.insert_paragraph_before(line.rstrip(), 'КОД')
    elif "<НАЗВАНИЕ>" in p.text.upper():
        p.text = ''
        p.insert_paragraph_before(prodl_name, "Заг искл огл")
    elif "<НОМЕР>" in p.text.upper():
        p.text = ''
        p.insert_paragraph_before(filename[3:-18], "Заг искл огл")
    elif "<НОМЕР2>" in p.text.upper():
        p.text = ''
        p.insert_paragraph_before(filename[3:-18] + "-ЛУ", "Заг искл огл")


    elif "<НОМЕР1>" in p.text.upper():
        p.text = ''
        p.insert_paragraph_before("             " + filename[3:-18] + "-ЛУ", "еспд-дец-1")

    for section in doc.sections:
        header = section.header
        for paragraph in header.paragraphs:
            if "643.ХХХХ.ХХХХХ-01 12 01" in paragraph.text:
                paragraph.text = ''
                paragraph.insert_paragraph_before(filename[3:-18], "Title")

from datetime import datetime
year = str(datetime.now().year)


def replace_year(container):
    for p in container.paragraphs:
        # print(p.text)
        if "<data>" in p.text:
            p.text = p.text.replace("<data>", year)

    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_year(cell)

def replace_num_dc(container):
    for p in container.paragraphs:
        # print(p.text)
        if "<НОМЕР_ВХ>" in p.text.upper():
            p.text = p.text.replace("<НОМЕР_ВХ>", filename[3:-18])

    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_num_dc(cell)

for section in doc.sections:
    replace_year(section.footer)
    replace_year(section.first_page_footer)
    replace_num_dc(section.header)




doc.save(filename + '.docx')



import win32com.client as win32
from pathlib import Path

path = Path(filename + '.docx').resolve()

word = win32.Dispatch("Word.Application")
word.Visible = False

doc = word.Documents.Open(str(path))
doc.TablesOfContents(1).Update()   # обновить содержание
doc.Save()
doc.Close()
word.Quit()
