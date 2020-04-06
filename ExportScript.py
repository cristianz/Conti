from openpyxl import load_workbook
import win32com.client as win32
import TestCommander
import os
import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import datetime
from distutils.dir_util import copy_tree

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = True
today = str(datetime.today().strftime('%d-%m-%Y'))

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.mkdir(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)

class TASISCopyResults(QtWidgets.QMainWindow, TestCommander.Ui_Exporter):

    def __init__(self, parent = None):
        super().__init__(parent)
        self.setupUi(self)
        self.show()
        self.cancelButton.clicked.connect(self.functionCancel)
        self.refreshButton.clicked.connect(self.functionRefresh)
        self.copyButton.clicked.connect(self.functionCopy)
        self.excelFile.clear()

    def functionRefresh(self):
        self.excelFile.clear()
        self.excelFile.addItem(excel.ActiveWorkbook.FullName)
        self.progressBar.setValue(0)

    def functionCopy(self):
        excel.Visible = False
        path_wb = (str(excel.ActiveWorkbook.FullName))
        path = path_wb
        print(path)
        folderName = str(excel.ActiveWorkbook.Name)
        folderName = folderName.split(".xlsx")[0]
        folderName = folderName + "_" + today
        print(folderName)
        wb = load_workbook(path_wb)
        path = str(path.split(excel.ActiveWorkbook.Name)[0] + folderName + "\\")
        createFolder(path)
        print('Root path in which we put results: ' + path)
        firstSheet = wb.active
        resultCol = firstSheet['P']
        firstCol = firstSheet['A']
        print(len(resultCol))
        self.progressBar.setMaximum(len(resultCol)-1)
        for x in range(len(firstCol)):
            if (firstCol[x].value == None) or (x == 0) or (resultCol[x].value == None):
                self.progressBar.setValue(x)
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
                self.progressBar.setValue(x)

        excel.Visible = True

    def functionCancel(self):
        self.close()

def main():
    app = QtWidgets.QApplication(sys.argv)
    main = TASISCopyResults()
    app.exec_()

if __name__ == '__main__':
    main()