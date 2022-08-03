import datetime
import logging
import os
import sys
import time
import traceback
from contextlib import suppress
import threading
from typing import List, Dict, Mapping, Tuple, Union, Optional, Any
import random
import re
import collections

import cv2
import numpy as np
from pyzbar.pyzbar import decode
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils.cell import get_column_letter, column_index_from_string
import pyinputplus
import pyinputplus as pyip

from util import getExcelFiles, getData

defaultExcelFilename = "GARSHS_INVENTORY.xlsx"

logging.basicConfig(level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
# logging.disable(logging.INFO)

while True:
    excelFiles = getExcelFiles()

    noInitialFile = False

    if len(excelFiles) == 0:
        noInitialFile = True
        print('No xlsx files found in the directory.')
        print(
            f'Please enter a filename for the excel file to be created. Press enter if you would like the default name [{defaultExcelFilename}]')
        excelFilename = pyinputplus.inputStr(prompt='> ', blank=True)
        if excelFilename == '':
            excelFilename = defaultExcelFilename
        if not excelFilename.endswith('.xlsx'):
            excelFilename += '.xlsx'
        wb = openpyxl.Workbook()
        wb.save(excelFilename)
        break
    elif len(excelFiles) == 1:
        excelFilename = excelFiles[0]
        print(f'Found {excelFilename}')
        print('Exit the program (Ctrl+C) if this is not the excel file you want.')
        break
    elif len(excelFiles) == 2 and \
            (excelFiles[0].startswith('~$') or excelFiles[1].startswith('~$')):
        print('!!! The excel file is open in another window !!!')
        print('Please close the excel file then press enter to continue.')
        continueProgram = input('> ')
        continue
    else:
        exelFileIsOpen = False
        for file in os.listdir('.'):
            if exelFileIsOpen:
                break
            if file.startswith('~$') and file.endswith('.xlsx'):
                print('!!! An excel file is open in another window !!!')
                print('Please close the excel file/s then press enter to continue.')
                continueProgram = input('> ')
                for filename in os.listdir('.'):
                    if filename.startswith('~$') and file.endswith('.xlsx'):
                        exelFileIsOpen = True
                        break
        if exelFileIsOpen:
            continue
        print('Multiple xlsx files found in the directory.')
        excelFilename: str = pyinputplus.inputMenu(getExcelFiles(), numbered=True)
        break

# Detects if the excel file is not initialized (data is not yet inputted)
# and if not, write the column headers, and then prompts the user to fill it up
while True:
    wb = openpyxl.load_workbook(excelFilename)
    ws = wb.active
    rowsDict, columnsDict = getData(ws)
    if not rowsDict and not columnsDict:
        ws.freeze_panes = 'B3'

        if ws['A1'].value != '':
            ws['A1'].value = 'GARSHS INVENTORY'
            ws['A1'].font = Font(size=20)
        if ws['A2'].value != '':
            ws['A2'].value = 'Equipment'
            ws['A2'].font = Font(size=14)
        if ws['B2'].value != '':
            ws['B2'].value = 'Working'
            ws['B2'].font = Font(size=14)
        if ws['C2'].value != '':
            ws['C2'].value = 'Available'
            ws['C2'].font = Font(size=14)
        if ws['D2'].value != '':
            ws['D2'].value = 'Place'
            ws['D2'].font = Font(size=14)
        if ws['E2'].value != '':
            ws['E2'].value = 'Person In Charge'
            ws['E2'].font = Font(size=14)

    try:
        wb.save(excelFilename)
        wb = openpyxl.load_workbook(excelFilename)
        ws = wb.active
        rowsDict, columnsDict = getData(ws)
    except PermissionError:
        print("\n!!!  The excel file is open on another window  !!!")
        print("Close the excel file to let the program manipulate it. Press enter if done closing the window.")
        continueProgram = input('> ')
        continue

    wb.save(excelFilename)
    break

# Puts the data to a dictionary


try:
    print("\nOpening the webcam... please wait\n")
    cap = cv2.VideoCapture(0)
    cap.set(3, 640)
    cap.set(4, 480)

    print("Webcam is running.     You can now show your QR code to the webcam.")
    print("--------------------------ATTENDANCE_LOG---------------------------")

    while True:
        success, img = cap.read()
        for qrcode in decode(img):
            name = qrcode.data.decode("utf-8")
            print(name)
            # time.sleep(0.5)

            pts = np.array([qrcode.polygon], np.int32)
            pts = pts.reshape((-1, 1, 2))
            cv2.polylines(img, [pts], True, (255, 0, 255), 5)
            pts2 = qrcode.rect
            cv2.putText(img, name, (pts2[0], pts2[1]), cv2.FONT_HERSHEY_COMPLEX, 0.9, (255, 0, 255), 2)

        cv2.imshow("Please show your QR code to the webcam.", img)
        cv2.waitKey(1)
except (KeyboardInterrupt, UserWarning):
    print("Program ended.")













