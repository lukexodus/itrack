#! python3

import os, sys

if len(sys.argv) != 2:
    print('Usage: batMaker <FOLDER_PATH>')
    sys.exit()
else:
    folderPath = sys.argv[1]
# print('Enter folder path:')
# folderPath = input()
# os.chdir(folderPath)
# folderPath = "C:\pyFiles"

pythonExeIsInside = False
for folderName, subfolders, filenames in os.walk(folderPath):
    for filename in filenames:
        if filename == 'python.exe':
            pythonExeIsInside = True
            pythonExeAbsPath = os.path.join(folderName, 'python.exe')
            break


pyFiles = []
for folderName, subfolders, filenames in os.walk(folderPath):
    for filename in filenames:
        if filename.endswith(".py") or filename.endswith(".pyw"):
            pyFiles.append(os.path.join(folderName, filename))

makeBatFor = []
for file in pyFiles:
    text = open(file)
    textString = text.readlines()
    if not textString:
        continue
    if textString[0] == "#! python3" or textString[0] == "#! python3\n":
        makeBatFor.append(file)
    text.close()


def extensionLen(fileName):
    if fileName.endswith(".py"):
        return -3
    elif fileName.endswith(".pyw"):
        return -4


notHaveBat = []
for i in range(len(makeBatFor)):
    if not os.path.exists(makeBatFor[i][: extensionLen(makeBatFor[i])] + ".bat"):
        notHaveBat.append(makeBatFor[i])

for file in notHaveBat:
    newBatName = os.path.basename(file)[: extensionLen(file)] + ".bat"
    newBatPath = os.path.join(os.path.dirname(file), newBatName)
    newBat = open(newBatPath, "w")
    if pythonExeIsInside:
        newBat.write(f"@{pythonExeAbsPath} {file} %*")
    else:
        newBat.write(f"@py.exe {file} %*")
    newBat.close()
