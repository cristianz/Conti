import openpyxl
from openpyxl import load_workbook
import win32com.client as win32
import TasisExplorer
import os
from datetime import datetime
from distutils.dir_util import copy_tree

today = datetime.today().strftime('%d-%m-%Y')
print(today)

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.mkdir(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def excelExecute():

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    path_wb = (str(excel.ActiveWorkbook.FullName))
    path = path_wb
    print(path)
    folderName = str(excel.ActiveWorkbook.Name)
    folderName = folderName.split(".xlsx")[0]
    print(folderName)
    wb = load_workbook(path_wb)
    path = str(path.split(excel.ActiveWorkbook.Name)[0] + folderName + "\\")
    createFolder(path)
    print('Root path in which we put results: ' + path)
    firstSheet = wb.active
    resultCol = firstSheet['P']
    firstCol = firstSheet['A']

    for x in range(len(firstCol)):
        if (firstCol[x].value == None) or (x == 0) or (resultCol[x].value == None):
            continue
        else:
            testID_DIR = str(firstCol[x].value)
            reportDir = str(resultCol[x].value)
            aux = reportDir
            aux = aux.split(",")[0]
            aux = aux.strip('=HYPERLINK(')
            aux = aux.strip('"')
            print(aux)
    finalPath = (path + testID_DIR)
    createFolder(finalPath)
    copy_tree(aux, finalPath)

    excel.Visible = True
