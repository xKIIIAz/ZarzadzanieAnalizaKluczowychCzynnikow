# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\basic.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import os
from pathlib import Path
import funkcjonalnosci as f
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
from matplotlib.figure import Figure


class Ui_MainWindow(QtWidgets.QMainWindow):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1075, 758)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.przyciskOdswiez = QtWidgets.QPushButton(self.centralwidget)
        self.przyciskOdswiez.setGeometry(QtCore.QRect(580, 690, 88, 27))
        self.przyciskOdswiez.setObjectName("przyciskOdswiez")
        self.wykresGlowny = QtWidgets.QFrame(self.centralwidget)
        self.wykresGlowny.setGeometry(QtCore.QRect(-1, -1, 631, 671))
        self.wykresGlowny.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.wykresGlowny.setFrameShadow(QtWidgets.QFrame.Raised)
        self.wykresGlowny.setObjectName("wykresGlowny")
        self.wykresPrawyGora = QtWidgets.QFrame(self.centralwidget)
        self.wykresPrawyGora.setGeometry(QtCore.QRect(630, 0, 451, 331))
        self.wykresPrawyGora.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.wykresPrawyGora.setFrameShadow(QtWidgets.QFrame.Raised)
        self.wykresPrawyGora.setObjectName("wykresPrawyGora")
        self.wykresPrawyDol = QtWidgets.QFrame(self.centralwidget)
        self.wykresPrawyDol.setGeometry(QtCore.QRect(630, 330, 451, 341))
        self.wykresPrawyDol.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.wykresPrawyDol.setFrameShadow(QtWidgets.QFrame.Raised)
        self.wykresPrawyDol.setObjectName("wykresPrawyDol")
        self.przyciskEdytuj = QtWidgets.QPushButton(self.centralwidget)
        self.przyciskEdytuj.setGeometry(QtCore.QRect(420, 690, 151, 27))
        self.przyciskEdytuj.setObjectName("przyciskEdytuj")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1075, 21))
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(77, 100, 202))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(77, 100, 202))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Inactive, QtGui.QPalette.Base, brush)
        brush = QtGui.QBrush(QtGui.QColor(120, 120, 120))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Text, brush)
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Disabled, QtGui.QPalette.Base, brush)
        self.menubar.setPalette(palette)
        self.menubar.setAutoFillBackground(False)
        self.menubar.setObjectName("menubar")
        self.menuDane = QtWidgets.QMenu(self.menubar)
        self.menuDane.setObjectName("menuDane")
        self.menuEksportuj = QtWidgets.QMenu(self.menuDane)
        self.menuEksportuj.setObjectName("menuEksportuj")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionImport = QtWidgets.QAction(MainWindow)        
        self.actionImport.setObjectName("actionImport")
        self.actionWykresy = QtWidgets.QAction(MainWindow)
        self.actionWykresy.setObjectName("actionWykresy")
        self.actionArkusz = QtWidgets.QAction(MainWindow)
        self.actionArkusz.setObjectName("actionArkusz")
        self.actionArkusz.triggered.connect(MainWindow.zapisz_aktualny_arkusz)
        self.actionImportuj = QtWidgets.QAction(MainWindow)
        self.actionImportuj.setObjectName("actionImportuj")
        self.actionImportuj.triggered.connect(MainWindow.otworz_wybrany_excel)
        self.menuEksportuj.addAction(self.actionWykresy)
        self.menuEksportuj.addAction(self.actionArkusz)
        self.menuDane.addAction(self.menuEksportuj.menuAction())
        self.menuDane.addAction(self.actionImportuj)
        self.menubar.addAction(self.menuDane.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def view_starting_popup(self):  #dodac mozliwosc ze arkusz juz istnieje
        start_dialog = QtWidgets.QDialog()
        label = QtWidgets.QLabel()
        label.setAlignment(QtCore.Qt.AlignLeft)
        label.setText("Witaj!\nProgram wygenerował dla Ciebie arkusz Excel\nw którym możesz dodać informacje potrzebne\ndo przeprowadzenia analizy.")
        label.move(50,50)
        popup_button = QtWidgets.QPushButton()
        popup_button.setText("OK")
        popup_button.clicked.connect(self.zamknij_okienko)
        popup_layout = QtWidgets.QVBoxLayout()
        popup_layout.addWidget(label)
        popup_layout.addWidget(popup_button)
        start_dialog.setLayout(popup_layout)
        start_dialog.setWindowTitle("Informacja")
        start_dialog.show()
        start_dialog.exec_()

    def zamknij_okienko(self):
        QtCore.QCoreApplication.instance().quit()

    def open_file(self, file = "nazwa_firmy.xlsx"):
        if sys.platform == 'linux2':
            subprocess.call(["xdg-open", file])
        else:
            os.startfile(file)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Analiza Kluczowych Czynników Sukcesu"))
        self.przyciskOdswiez.setWhatsThis(_translate("MainWindow", "<html><head/><body><p>Naciśnij kiedy zamkniesz uzupełniony arkusz.</p></body></html>"))
        self.przyciskOdswiez.setText(_translate("MainWindow", "Odśwież"))
        self.przyciskEdytuj.setText(_translate("MainWindow", "Edytuj Arkusz"))
        self.menuDane.setTitle(_translate("MainWindow", "Dane"))
        self.menuEksportuj.setTitle(_translate("MainWindow", "Eksportuj"))
        self.actionImport.setText(_translate("MainWindow", "Importuj"))
        self.actionWykresy.setText(_translate("MainWindow", "Wykresy"))
        self.actionArkusz.setText(_translate("MainWindow", "Arkusz"))
        self.actionImportuj.setText(_translate("MainWindow", "Importuj"))
        self.actionImportuj.setWhatsThis(_translate("MainWindow", "Kliknij jeżeli w folderze programu znajduje się arkusz o nazwie \"Dane.xlsx\""))

class MainWindow(Ui_MainWindow):
    
    workbook = f.Workbook()
    flaga_istnieje_plik = False
    flaga_stworzono_przykladowy = False
    filename = ""
    def __init__(self):
        super(MainWindow, self).__init__()
        Ui_MainWindow.setupUi(self, self)
        self.dodaj_wykresy_do_frameow()
        filepath = "Dane.xlsx"
        data_file = Path("Dane.xlsx")
        print(data_file)
        if data_file.is_file():
            self.flaga_istnieje_plik = True
            #dzialanie jak jest plik
            pass
        else:
            #tworzenie pliku
            x = f.daj_przykladowe_dane()
            f.exporter(x[0],x[1],x[2],x[3],x[4],x[5])
            self.flaga_stworzono_przykladowy = True
            self.flaga_istnieje_plik = True

    #rysowanie_wykresow
    def dodaj_wykresy_do_frameow(self):
        layout = QtWidgets.QVBoxLayout()
        sc = MplCanvas(self, width=30, height=10, dpi=100)
        sc.axes.plot([0,1,2,3,4], [10,1,20,3,40])
        layout.addWidget(sc)
        self.wykresGlowny.setLayout(layout)

    def odswiez(self):
        pass

    def otworz_wybrany_excel(self):
        self.filename = self.openExcelFileNameDialog()
        self.workbook = f.load_workbook(self.filename)
        self.odswiez()

    def zapisz_aktualny_arkusz(self):
        savename = self.saveArkuszDialog()
        self.workbook.save(savename)

    def openExcelFileNameDialog(self):
        options = QtWidgets.QFileDialog.Options()
        #options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(self,"Wybierz dokument excel z danymi.", "Dane","Excel Files (*.xlsx)", options=options)
        if fileName:
           return fileName
        
    def saveArkuszDialog(self):
        options = QtWidgets.QFileDialog.Options()
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self,"Zapisz arkusz jako...","dane_firm","Excel Files (*.xlsx)", options=options)
        if fileName:
            return fileName

class MplCanvas(FigureCanvasQTAgg):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot()
        super(MplCanvas, self).__init__(fig)

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    if window.flaga_stworzono_przykladowy:
        window.view_starting_popup()
    else:
        pass
    app.exec_()