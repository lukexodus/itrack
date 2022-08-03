import threading
import time
import os
import typing
from typing import List, Dict, Mapping


def getExcelFiles() -> List[str]:
    excelFiles: List[str] = []
    for file in os.listdir('.'):
        if file.endswith('.xlsx'):
            excelFiles.append(file)
    return excelFiles


def getData(ws) -> Dict[int, list]:
    rows = list(ws.rows)
    rowsDict = {}
    for i, row in enumerate(rows):
        rowsDict.setdefault(i + 1, [])
        for cell in row:
            rowsDict[i + 1].append(cell.value)

    columns = list(ws.columns)
    columnsDict = {}
    for i, column in enumerate(columns):
        columnsDict.setdefault(i + 1, [])
        for cell in column:
            columnsDict[i + 1].append(cell.value)
    return rowsDict, columnsDict


def removeFromList(targetList, value, duration):
    time.sleep(duration)
    targetList.remove(value)
