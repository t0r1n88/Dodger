from docxcompose.composer import Composer
from docx import Document as Document_compose
import os
def combine_all_docx(filename_master, files_list):
    number_of_sections = len(files_list)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document_compose(files_list[i])
        composer.append(doc_temp)
    open('1.txt', 'w')
    composer.save("ALL_SERTIFICATES.docx")

files = []
# Получаем список всех файлов с расширением .docx в текущем каталоге.
for filedocx in os.listdir():
    if filedocx.endswith(".docx"):
        files.append(filedocx)





filename_master = files[0]

combine_all_docx(filename_master, files)
