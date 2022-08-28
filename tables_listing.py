import docx

doc = docx.Document(r'C:\Users\Иван\Desktop\Иван\python\pythonProject\python-docx\01-ТК 001-ППР001-08-2020 Геодез.docx')

tables = doc.tables

print("tables amount -", len(tables))

for table in range(len(tables)):
    print("\ntable #", table)
    for i in range(len(tables[table].rows)):
        print("row #", i)
    print('end')
