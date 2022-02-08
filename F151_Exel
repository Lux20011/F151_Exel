import xlrd
import os, fnmatch
import shutil
import pathlib


input("Press Enter to begin...")

if not os.path.isdir("Ф151_Exel"):
    os.mkdir("Ф151_Exel")

old_path = ''
new_path = 'Ф151_Exel'

listOfFiles = os.listdir('.')
pattern = "*.XLS"
for entry in listOfFiles:

    if fnmatch.fnmatch(entry, pattern):
        print(entry)
        workbook = xlrd.open_workbook(entry)
        sheet = workbook.sheet_by_name('1')
        name_budget = sheet.cell_value(6,1)
        data_file = sheet.cell_value(2,6)
        new_filename = os.rename(entry, data_file+"_"+(name_budget + ".xls"))
        print(new_filename)
        #  Read data
        print(data_file)
        print(name_budget)
input("Press Enter to move...")
for new_filename in listOfFiles:
    file_ext = pathlib.Path(new_filename).suffix
    if file_ext not in ('.exe', '.py','',' ','.XML','.doc','.DOC','.RTF','.rtf','.docx','.txt','.spec','.html','.pdf'):
        print(file_ext)
        shutil.move(old_path + new_filename, new_path)

