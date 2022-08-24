import docx

doc = docx.Document(r'C:\Users\Иван\Desktop\Иван\python\pythonProject\python-docx\01-ТК 001-ППР001-08-2020 Геодез.docx')

print(len(doc.tables))
print(len(doc.tables[1].rows))
print(len(doc.tables[1].rows[0].cells))

target_cell = doc.tables[1].rows[2].cells[0]
target_text = target_cell.paragraphs[0].runs[1].text

target_cell.paragraphs[0].runs[1].text = str('Какой-то текст 4')
print(target_cell.paragraphs[0].runs[1].text)


doc.save(r'C:\Users\Иван\Desktop\Иван\python\pythonProject\python-docx\01-ТК 001-ППР001-08-2020 Геодез1.docx')

tables_text = []
for row in range(len(doc.tables[1].rows)):
    for cell in range(len(doc.tables[1].rows[row].cells)):
        tables_text.append(doc.tables[1].rows[row].cells[cell].text)


tables_word = []
for row in range(len(doc.tables[1].rows)):
    for cell in range(len(doc.tables[1].rows[row].cells)):
        for word in range(len(doc.tables[1].rows[row].cells[cell].text.split())):
            if doc.tables[1].rows[row].cells[cell].text.split()[word] not in tables_word:
                tables_word.append(doc.tables[1].rows[row].cells[cell].text.split()[word])

tables_letter = ''
for word in tables_word:
    tables_letter += word

letter_list = list(tables_letter)


def searching(letters):

    def search(i):
        if i < len(list(letters)):
            if letters[i] in letter_list:
                search(i+1)
            else:
                print('нет такого сочетания букв')
        else:
            print('В тексте найдено слово (' + letters + ')')

    search(0)


searching('Пунга')
