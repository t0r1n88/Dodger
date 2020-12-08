from docx import Document
from docxcompose.composer import  Composer
from docx import Document as Document_compose
import os


files = []
# Получаем список всех файлов с расширением .docx в текущем каталоге.
for filedocx in os.listdir():
    if filedocx.endswith(".docx"):
        files.append(filedocx)



def combine_all_docx(filename_master,files_list):
    number_of_sections=len(files_list)
    master = Document_compose(filename_master)
    composer = Composer(master)
    for i in range(0, number_of_sections):
        doc_temp = Document_compose(files_list[i])
        composer.append(doc_temp)
    composer.save("combined_file.docx")

filename_master = 'all.docx'
combine_all_docx(filename_master,files)