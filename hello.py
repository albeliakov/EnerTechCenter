from openpyxl import load_workbook

def loadDataTP(wbName, sheetName, cellsInd):
    wb = load_workbook(wbName)
    sheet = wb.get_sheet_by_name(sheetName)

    dictPowerLine = {}
    for cellObj in sheet[cellsInd[0]:cellsInd[1]]:
        dictPowerLine[int(cellObj[0].value)] = (str(cellObj[1].value),
                                                int(cellObj[2].value),
                                                float(cellObj[3].value))
    return dictPowerLine

wbName = './inMatrix.xlsx'
sheetConjMatrix = 'conjMatrix'
cellConjMatrix = ('C3', 'EW153')
sheetDataTP = 'TP'
cellDataTP = ('B3', 'E153')
# Загрузка словаря из эксель с данными о ТП (позиция: №, ном.мощн, ст.загрузки)
dictDataTP = loadDataTP(wbName='./inMatrix.xlsx', sheetName=TP, cellsInd=('B3', 'E153'))

print("HELLO")