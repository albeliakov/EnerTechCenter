from openpyxl import Workbook



def writeListInXlsx(tableName, rows):
    book = Workbook()
    sheet = book.active
    for row in rows:
        # print(row)
        sheet.append(row)

    book.save('{}.xlsx'.format(tableName))