import sys
from PySide2.QtCore import QUrl, QTime , QFileInfo
from PySide2.QtWidgets import (QLabel,QListWidgetItem, QApplication, QFileDialog, QDoubleSpinBox, QTableWidgetItem, QLineEdit, QGroupBox, QVBoxLayout, QWidget, QComboBox, QHBoxLayout, QSizePolicy, QMainWindow)
from PySide2.QtGui import QPixmap
from PySide2.QtCore import Qt
from tika import parser
import re
from ui_cvSearchDesign_v4 import Ui_MainWindow

class MainWindow(QMainWindow):

    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.pBAjoutCV.clicked.connect(self.ajoutCV)

        self.ui.pBAddKeyWord.clicked.connect(self.ajoutKeyWord)

        self.ui.pBResearch.clicked.connect(self.rechKeyWord)

        self.ui.pBNewRequest.clicked.connect(self.newRequest)

        self.updateTable()

    def updateTable(self):
        self.ui.tableWidget.setColumnCount(1)
        self.ui.tableWidget.setColumnWidth(0, 300)
        self.ui.tableWidget.setHorizontalHeaderLabels(['Liste des CVs'])

        self.ui.tWKeyWord.setColumnCount(1)
        self.ui.tWKeyWord.setColumnWidth(0, 200)
        self.ui.tWKeyWord.setHorizontalHeaderLabels(['Mot Clef'])

        self.ui.tWResults.setColumnCount(2)
        self.ui.tWResults.setHorizontalHeaderLabels(['Nom CV', 'Nombre KeyWord'])
        self.ui.tWResults.setColumnWidth(0, 200)
        self.ui.tWResults.setColumnWidth(1, 150)

    def ajoutCV (self):
        newFiles = QFileDialog.getOpenFileNames(self, "Choix CVs", "/home", "CV (*pptx *pdf *docx)")
        n = len(newFiles[0])
        for f in range(0,n):
            fInfo = QFileInfo(newFiles[0][f])
            fShortName = fInfo.baseName()
            addCV = QListWidgetItem(fShortName)
            addCV.setToolTip(newFiles[0][f])
            ajoutCV = QTableWidgetItem(fShortName)
            ajoutCV.setToolTip(newFiles[0][f])
            if self.ui.tableWidget.rowCount() == 0 :
                self.ui.tableWidget.insertRow(0)
                self.ui.tableWidget.setItem(0,0,QTableWidgetItem(ajoutCV))
            else :
                lgTable = self.ui.tableWidget.rowCount()
                self.ui.tableWidget.insertRow(lgTable)
                self.ui.tableWidget.setItem(lgTable, 0, QTableWidgetItem(ajoutCV))

    def ajoutKeyWord(self):
        keyW = self.ui.lEKeyWord.text()
        if self.ui.tWKeyWord.rowCount()== 0 :
            self.ui.tWKeyWord.insertRow(0)
            self.ui.tWKeyWord.setItem(0, 0, QTableWidgetItem(keyW))
        else :
            lgTable = self.ui.tWKeyWord.rowCount()
            self.ui.tWKeyWord.insertRow(lgTable)
            self.ui.tWKeyWord.setItem(lgTable, 0, QTableWidgetItem(keyW))
        self.ui.lEKeyWord.clear()

    def listKeyWord(self):
        listKW = []
        n = self.ui.tWKeyWord.rowCount()
        for i in range (0, n):
            keyW = self.ui.tWKeyWord.item(i, 0).text()
            keyW = keyW.lower()
            # keyW = " " + keyW + [^a-z]
            listKW.append(keyW)
        return (listKW)

    def rechKeyWord (self):
        self.ui.tWResults.setRowCount(0)
        kWToTest = self.listKeyWord()
        klen = len(kWToTest)
        n = self.ui.tableWidget.rowCount()
        for cv in range (0, n):
            nomCV = self.ui.tableWidget.item(cv, 0).text()
            currentTable = self.ui.tableWidget.item(cv, 0).toolTip()
            cvATester = QUrl.fromLocalFile(currentTable)
            cvATestPath = QUrl.toLocalFile(cvATester)
            text = parser.from_file(cvATestPath)
            cvTransfText = text['content']
            cvAcompare = cvTransfText.lower()
            cptKw = 0
            for k in range (0, klen) :
                kWSearch=kWToTest[k]
                # (r'{}\[^a-z]'.format(kWSearch))
                if re.search(r'\W{}\W'.format(kWSearch), cvAcompare) != None :
                # if kWSearch in cvAcompare:
                    cptKw += 1
            if cptKw !=0 :
                if self.ui.tWResults.rowCount() == 0:
                    self.ui.tWResults.insertRow(0)
                    self.ui.tWResults.setItem(0, 0, QTableWidgetItem(nomCV))
                    self.ui.tWResults.setItem(0,1, QTableWidgetItem(str(cptKw)))
                else:
                    lgTable = self.ui.tWResults.rowCount()
                    self.ui.tWResults.insertRow(lgTable)
                    self.ui.tWResults.setItem(lgTable, 0, QTableWidgetItem(nomCV))
                    self.ui.tWResults.setItem(lgTable, 1, QTableWidgetItem(str(cptKw)))
        self.ui.tWResults.sortItems(1,order = Qt.DescendingOrder)

    def newRequest(self):
        self.ui.tableWidget.clear()
        self.ui.tableWidget.setRowCount(0)
        self.ui.tWResults.clear()
        self.ui.tWResults.setRowCount(0)
        self.ui.tWKeyWord.clear()
        self.ui.tWKeyWord.setRowCount(0)
        self.updateTable()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    cvSearch = MainWindow()
    cvSearch.show()
    cvSearch.resize(800,600)

    sys.exit(app.exec_())