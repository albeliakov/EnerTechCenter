import pandas as pd
from queue import deque


#---------------------- Запись графа линий в файл txt -----------------------------------------------------------

def writeFileDict(filename, int_key, int_list):
    with open(filename, 'a') as out:
        out.write(
            '{}:{}\n'.format(int_key,
                      ' '.join(str(int_list[i]) for i in range(len(int_list)))
                             ))

def readFileDict(filename):
    readDict = {}
    with open(filename) as inp:
        for i in inp.readlines():
            if i != None:
                key, val = i.strip().split(':')
                readDict[int(key)] = [int(i) for i in val.split()]
    return readDict



# ------------------------------ Формирование графа линий в виде словаря ------------------------

# Заполнение словаря графа {int_key : [int_list]}
# int_key - номер вершины, int_list - список соседних вершин
def getGraphDict(saveFileDict):
    fileWithDict = saveFileDict + '.txt'
    dictGrPosits = dict()
    numbNodes = int(input("Количество вершин: "))

    print("\nВНИМАНИЕ: входные данные для каждой вершины")
    print("          должны вводиться через пробел\n")

    for node in range(1, numbNodes+1):
        while True:
            inList = input('Вершина {}: '.format(node)).split()
            try:
                for i in range(len(inList)):
                    inList[i] = int(inList[i])
                break
            except ValueError:
                print("ВНИМАНИЕ: Введено неверное знаечение, повторите ввод")
        dictGrPosits[node] = tuple(inList)
        writeFileDict(fileWithDict, node, dictGrPosits[node])
    return dictGrPosits


#-----------------------------ФОрмирование матрицы зависимостей------------------------------------------

# graph = {1: (2, 3),
#          2: (4, 5),
#          3: (6,7,8),
#          4: (9, 10),
#          5: (),
#          6: (),
#          7: (),
#          8: (),
#          9: (),
#          10: ()}

def bfs(graph, startNode):
    visited, queue = [], deque([startNode])
    while queue:
        vertex = queue.pop()
        if vertex not in visited:
            visited.append(vertex)
            queue.extendleft(set(graph[vertex]) - set(visited))
    return visited

def graphDict2Matr(dictGraph):
    graphSize = len(dictGraph)
    matrDepend = [[0]*graphSize for i in range(graphSize)]
    # Формирование матрицы зависимостей вершин графа
    for node in dictGraph.keys():
        lstDependNodes = bfs(dictGraph, node)
        #print(node, lstDependNodes)
        for dependNode in lstDependNodes:
            matrDepend[node - 1][dependNode - 1] = 1
    return matrDepend

def getMatrNeigh(dictGraph):
    graphSize = len(dictGraph)
    matrNeigh = [[0]*graphSize for i in range(graphSize)]
    for node in dictGraph.keys():
        for daughNode in dictGraph[node]:
            matrNeigh[node-1][daughNode-1] = 1
    return matrNeigh

def writeGrMatrInXlsx(tableName, matrName):
    df_ = pd.DataFrame(matrName)
    df_.to_excel(tableName + '.xlsx')

#------------------ main -----------------------------

def main():
    feederName = input("Введите название фидера (имя папки): ")
    myGraphDict = getGraphDict(feederName+'/graphDict') 
    writeGrMatrInXlsx(feederName+'/dependMatr', graphDict2Matr(myGraphDict)) # матрица зависимостей
    writeGrMatrInXlsx(feederName+'/matrNeigh', getMatrNeigh(myGraphDict))   # матрица соседей

if __name__== "__main__":
    main()