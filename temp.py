from docxtpl import DocxTemplate

doc = DocxTemplate('resources/test_template.docx')
tbl_contents = ['Cilian', 'Booth', 'Lindy']

context = {'tbl_contents':tbl_contents}
doc.render(context)
doc.save('Test.docx')