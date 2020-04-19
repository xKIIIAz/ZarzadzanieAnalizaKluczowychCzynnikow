# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\basic.ui'
#
# Created by: PyQt5 UI code generator 5.14.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtWidgets, QtCore, QtGui

import sys, os


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1075, 758)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(-1, -1, 1081, 731))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1075, 21))
        self.menubar.setAutoFillBackground(False)
        self.menubar.setObjectName("menubar")
        self.menuDane = QtWidgets.QMenu(self.menubar)
        self.menuDane.setObjectName("menuDane")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionImport = QtWidgets.QAction(MainWindow)
        self.actionImport.setObjectName("actionImport")
        self.actionEksportuj = QtWidgets.QAction(MainWindow)
        self.actionEksportuj.setObjectName("actionEksportuj")
        self.menuDane.addAction(self.actionImport)
        self.menuDane.addAction(self.actionEksportuj)
        self.menubar.addAction(self.menuDane.menuAction())

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        # self.open_file()
        
    def view_starting_popup(self, istnieje = "False"):  #dodac mozliwosc ze arkusz juz istnieje
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

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Analiza Kluczowych Czynnikow Sukcesu"))
        self.menuDane.setTitle(_translate("MainWindow", "Dane"))
        self.actionImport.setText(_translate("MainWindow", "Importuj"))
        self.actionEksportuj.setText(_translate("MainWindow", "Eksportuj..."))

    def open_file(self, file = "nazwa_firmy.xlsx"):
        if sys.platform == 'linux2':
            subprocess.call(["xdg-open", file])
        else:
            os.startfile(file)

if __name__ == "__main__":
    # os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QtWidgets.QApplication(sys.argv)

    window = QtWidgets.QMainWindow()
    ui_window = Ui_MainWindow()
    ui_window.setupUi(window)
    window.show()
    ui_window.view_starting_popup()
    app.exec_()
    

