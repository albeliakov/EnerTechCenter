from openpyxl import load_workbook
import xlwt

# Функция получения данных из эксель файла в виде словаря
# {кол-во_КА : ср_мощн, (позиции)}. Добавить функцию метки позиций символом заданным символом
def loadResFromXlsx(wbName, sheetName, cellsInd, onLabel = True):
    wb = load_workbook(wbName)
    sheet = wb.get_sheet_by_name(sheetName)

    feederNameInd = wbName.find('выхДанные')
    label = str(input("Введите метку для позиций "
                      "{}: ".format(wbName[feederNameInd:-5]))) if onLabel else ''

    dictReturn = {}
    for cellObj in sheet[cellsInd[0]:cellsInd[1]]:
        posWithLabel = tuple(label+pos.strip()+' '
                             for pos in str(cellObj[2].value).split(',')
                             if pos != 'None')
        dictReturn[int(cellObj[1].value)] = (float(cellObj[0].value),
                                             posWithLabel)
    return dictReturn


# сортировка по мат ожиданию (по возрастанию)
def sortByME(outData):
    return sorted(outData, key=lambda data: data[0])

# Псстрочная запись в эксель-файл
def write2Excel(book, sheetName, cols, colNames, writeRows):
    # именование столбцов
    rowNames = sheetName.row(0)
    for index, col in enumerate(cols):
        rowNames.write(index, colNames[index])

    # заполнение значений
    for ind in range(len(writeRows)):
        row = sheetName.row(ind+1)
        for index, col in enumerate(cols):
            value = writeRows[ind][index]
            row.write(index, value)

    # Save the result
    book.save("C:/EnerTechCenter/Результаты/объедин_БелоречкаГагарский.xls")

# полученные результаты по сети 1
gridResults_1 = loadResFromXlsx('C:/EnerTechCenter/Результаты/выхДанные_Белоречка.xlsx',
                      'Лист1', ('A6', 'C13'))
# полученные результаты по сети 2
gridResults_2 = loadResFromXlsx('C:/EnerTechCenter/Результаты/выхДанные_Гагарский.xlsx',
                      'Лист1', ('A6', 'C14'))

# gridResults_1 = {0: (500, ()), 1: (450, ('К_2',)), 2: (370, ('К_3', 'К_4')), 3: (300, ('К_2', 'К_4','К_5'))}
# gridResults_2 = {0: (450, ()), 1: (450, ('П_3',)), 2: (320, ('П_3', 'П_5')), 3: (250, ('П_3', 'П_6','П_7'))}

numbKA = int(input("Введите количесвто доступных КА: "))
maxKA_1 = max(gridResults_1.keys())
maxKA_2 = max(gridResults_2.keys())

# Поиск комбинаций с наименьщими потерями
dictConsolidRes = {}
for totalAmountKA in range(numbKA+1):
    minPower = (gridResults_1[0][0] + gridResults_2[0][0]) * 2
    for amounKA_1 in range(totalAmountKA+1):
        amounKA_2 = totalAmountKA -amounKA_1

        if  amounKA_1 <= maxKA_1 and amounKA_2 <= maxKA_2:
            sumPowers = round(gridResults_1[amounKA_1][0] + gridResults_2[amounKA_2][0], 1)
            if sumPowers < minPower:
                minPower = sumPowers
                tplConsoPos = gridResults_1[amounKA_1][1] + gridResults_2[amounKA_2][1]
                dictConsolidRes[totalAmountKA] = (sumPowers, tplConsoPos)

for key, val in dictConsolidRes.items():
     print(key, val)


# Рассчет для крайних (максимальных) значений КА
# maxKAGrid_1 = {0: (500, ()), 1: (450, ()), 2: (370, ()), 3: (300, ())}
# maxKAGrid_2 = {0: (450, ()), 1: (450, ()), 2: (320, ()), 3: (250, ())}
maxKAGrid_1 = loadResFromXlsx('C:/EnerTechCenter/Результаты/выхДанные_Белоречка.xlsx',
                      'Лист1', ('A2', 'C5'), onLabel=False)
maxKAGrid_2 = loadResFromXlsx('C:/EnerTechCenter/Результаты/выхДанные_Гагарский.xlsx',
                      'Лист1', ('A2', 'C5'), onLabel=False)

dictMaxKARes = {}
for amKA_1 in maxKAGrid_1.keys():
    for amKA_2 in maxKAGrid_2.keys():
        sumMaxKA = amKA_1 + amKA_2
        sumMaxPower = round(maxKAGrid_1[amKA_1][0] + maxKAGrid_2[amKA_2][0], 1)
        consolidPos = maxKAGrid_1[amKA_1][1] + maxKAGrid_2[amKA_2][1]
        dictMaxKARes[sumMaxKA] = (sumMaxPower, consolidPos)

for key, val in dictMaxKARes.items():
    print(key, val)

print('----------------------------------------------------\n'
      '----------------------------------------------------\n')

book = xlwt.Workbook(encoding="utf-8")
# Add a sheet to the workbook
sheet1 = book.add_sheet('outData')

cols = ['A', 'B', 'C']
colNames = ['Среднее значение отключаемой мощности',
            'Количество устанавливаемых КА',
            'Места для установки']

consolidDicts = dictConsolidRes.copy()
consolidDicts.update(dictMaxKARes) # объединение словарей
for key, val in consolidDicts.items():
    print(key, val[0], val[1])

arrayOutData = [(powAndPos[0], amountKA, powAndPos[1])
                for amountKA, powAndPos
                in consolidDicts.items()]

write2Excel(book, sheet1, cols, colNames, sortByME(arrayOutData))