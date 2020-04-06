from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Exporter(object):
    def setupUi(self, Exporter):
        Exporter.setObjectName("Exporter")
        Exporter.resize(484, 249)
        Exporter.setFixedSize(484,249)
        self.centralwidget = QtWidgets.QWidget(Exporter)
        self.centralwidget.setObjectName("centralwidget")

        self.copyButton = QtWidgets.QPushButton(self.centralwidget)
        self.copyButton.setGeometry(QtCore.QRect(40, 130, 111, 41))
        self.copyButton.setObjectName("copyButton")

        self.refreshButton = QtWidgets.QPushButton(self.centralwidget)
        self.refreshButton.setGeometry(QtCore.QRect(190, 130, 111, 41))
        self.refreshButton.setObjectName("refreshButton")

        self.cancelButton = QtWidgets.QPushButton(self.centralwidget)
        self.cancelButton.setGeometry(QtCore.QRect(340, 130, 111, 41))
        self.cancelButton.setObjectName("cancelButton")

        self.excelFile = QtWidgets.QComboBox(self.centralwidget)
        self.excelFile.setGeometry(QtCore.QRect(40, 70, 411, 31))
        self.excelFile.setObjectName("excelFile")

        self.Label = QtWidgets.QLabel(self.centralwidget)
        self.Label.setGeometry(QtCore.QRect(40, 30, 401, 31))

        font = QtGui.QFont()
        font.setPointSize(12)
        self.Label.setFont(font)
        self.Label.setObjectName("Label")
        Exporter.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(Exporter)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 484, 21))
        self.menubar.setObjectName("menubar")
        Exporter.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(Exporter)
        self.statusbar.setObjectName("statusbar")
        Exporter.setStatusBar(self.statusbar)

        self.retranslateUi(Exporter)
        QtCore.QMetaObject.connectSlotsByName(Exporter)

    def retranslateUi(self, Exporter):
        _translate = QtCore.QCoreApplication.translate
        Exporter.setWindowTitle(_translate("Exporter", "Tasis Exporter"))
        self.copyButton.setText(_translate("Exporter", "Copy"))
        self.refreshButton.setText(_translate("Exporter", "Refresh"))
        self.cancelButton.setText(_translate("Exporter", "Cancel"))
        self.Label.setText(_translate("Exporter", "Select TASIS Excel related file:"))



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Exporter = QtWidgets.QMainWindow()
    ui = Ui_Exporter()
    ui.setupUi(Exporter)
    Exporter.show()
    sys.exit(app.exec_())
