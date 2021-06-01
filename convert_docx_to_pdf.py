# import tkinter as tk
# import tkinter.ttk as ttk
# from  tkinter.filedialog import askopenfile
# from tkinter.messagebox  import showinfo
from docx2pdf import convert
#
"""
usage: docx2pdf [-h] [--keep-active] [--version] input [output]

Example Usage:

Convert single docx file in-place from myfile.docx to myfile.pdf:
    docx2pdf myfile.docx

Batch convert docx folder in-place. Output PDFs will go in the same folder:
    docx2pdf myfolder/

Convert single docx file with explicit output filepath:
    docx2pdf input.docx output.docx

Convert single docx file and output to a different explicit folder:
    docx2pdf input.docx output_dir/

Batch convert docx folder. Output PDFs will go to a different explicit folder:
    docx2pdf input_dir/ output_dir/

positional arguments:
  input          input file or folder. batch converts entire folder or convert
                 single file
  output         output file or folder

optional arguments:
  -h, --help     show this help message and exit
  --keep-active  prevent closing word after conversion
  --version      display version and exit
"""
# win = tk.Tk()
# win.title("Word To PDF Converter")
# def openfile():
#     file = askopenfile(filetypes = [('Word Files','*.docx')])
#     print(file)
#     convert(file.name)
#     showinfo("Done","File Successfully Converted")
#
# label = tk.Label(win,text='Choose File: ')
# label.grid(row=0,column=0,padx=5,pady=5)
#
# button = ttk.Button(win,text='Select',width=30,command=openfile)
# button.grid(row=0,column=1,padx=5,pady=5)
#
# win.mainloop()

convert('resources/')