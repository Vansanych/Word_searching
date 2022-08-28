import docx

doc = docx.Document(r'C:\Users\Иван\Desktop\Иван\python\pythonProject\python-docx\01-ТК 001-ППР001-08-2020 Геодез.docx')

tables = doc.tables

print(len(tables))
