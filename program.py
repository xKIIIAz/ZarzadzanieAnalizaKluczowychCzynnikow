import showexcel
from plotdrawer import Plotter
from gui import Ui_MainWindow as Ui
from gui import QtWidgets
from gui import os
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QObjectCleanupHandler
from pathlib import Path, PureWindowsPath
import funkcjonalnosci as f
import matplotlib
matplotlib.use('Qt5Agg')
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import numpy as np

'''
solution to not being able to save excel file while the program is running:
https://stackoverflow.com/questions/31416842/openpyxl-does-not-close-excel-workbook-in-read-only-mode
'''


class MainWindow(Ui):

    '''
    here is a logic of a program
    todo:
    - handler for exception if someone clicks X when opening or saving a file
    - handler for editing a workbook
    - add function "addListeners" // not that important, but would add more readability

    '''
    
    flaga_istnieje_plik = False
    flaga_stworzono_przykladowy = False
    filepath = ""
    dane_firm = []

    def __init__(self):
        super(MainWindow, self).__init__()
        Ui.setupUi(self, self)
        self.filepath = "Dane.xlsx"
        data_file = Path("Dane.xlsx")
        if data_file.is_file():
            self.flaga_istnieje_plik = True
            self.dane_firm = f.full_import(self.filepath)
            del data_file
        else:
            #tworzenie pliku
            x = f.daj_przykladowe_dane()
            f.exporter(x[0],x[1],x[2],x[3],x[4],x[5])
            self.flaga_stworzono_przykladowy = True
            self.flaga_istnieje_plik = True
        #for x in self.dane_firm:
        #    print(x)
        self.odswiez()
        #self.workbook.close()
 
    
    #rysowanie_wykresow
    def dodaj_wykres_glowny(self):
        layout = self.wykresGlowny.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.wykresGlowny.layout())
        layout = QtWidgets.QVBoxLayout()
        self.scGlowny = MplCanvas(self)
        for dane in self.dane_firm:
            self.scGlowny.axes.plot(dane[1],dane[4],label = dane[0], linewidth = 4)
        #sc.axes.text(1,3.6,"Wyniki firm względem czynników", horizontalalignment = 'center', verticalalignment = 'center', fontsize=12)
        self.scGlowny.axes.set_title("Wykres wyników firm względem czynników.")
        self.scGlowny.axes.grid(axis = 'y',color = 'gray', linestyle='-', linewidth = 1)
        self.scGlowny.axes.legend()
        self.scGlowny.correct_ticks()
        layout.addWidget(self.scGlowny)
        self.wykresGlowny.setLayout(layout)

    def dodaj_wykres_prawy_gora(self):
        Plotter.draw_and_print(self.dane_firm)
        layout = self.wykresPrawyGora.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.wykresPrawyGora.layout())
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel(self)
        pixmap = QPixmap('plt/bar_chart.png')
        pixmap = pixmap.scaled(451,331, Qt.KeepAspectRatio, Qt.FastTransformation)
        label.setPixmap(pixmap)
        layout.addWidget(label)
        self.wykresPrawyGora.setLayout(layout)

    def dodaj_wykres_prawy_dolny(self):
        layout = self.wykresPrawyDol.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.wykresPrawyDol.layout())
        layout = QtWidgets.QVBoxLayout()
        self.scPie = MplCanvas(self)
        wyniki = []
        labele = []
        for dane in self.dane_firm:
            wyniki.append(sum(dane[4]))
            labele.append(dane[0])
        explode = []
        dex = wyniki.index(max(wyniki))
        for i in range(0,len(wyniki)):
            if i == dex:
                explode.append(0.05)
            else:
                explode.append(0.01)
        wyniki = np.array(wyniki)
        labele = np.array(labele)
        self.scPie.axes.pie(x = wyniki,labels = labele, explode = explode, shadow = False, autopct='%1.1f%%')
        self.scPie.axes.set_title("Udział w wynikach każdej z firm")
        layout.addWidget(self.scPie)
        self.wykresPrawyDol.setLayout(layout)
            
    def czytaj_dane(self):
        self.dane_firm = [x for x in f.full_import(self.filepath)]

    def odswiez(self):
        self.czytaj_dane()
        self.dodaj_wykres_glowny()
        self.dodaj_wykres_prawy_gora()
        self.dodaj_wykres_prawy_dolny()
        #self.dodaj_wykres_glowny()
        #self.dodaj_wykres_prawy_gora()

    def edytuj_excel(self):
        '''
        opens excel file under the specified filename
        or if the file is already opened, sets focus to that
        excel window.
        '''
        p = str(PureWindowsPath(self.filepath))
        command = "start " + "\"title\" " + "\"" + p + "\""
        if not showexcel.show_excel_window(os.path.basename(self.filepath)):
            os.system(command)

    def otworz_wybrany_excel(self):
        new_filepath = self.openExcelFileNameDialog()
        if new_filepath == None or new_filepath == "":
            pass
        else:
            self.filepath = new_filepath
            self.workbook = f.load_workbook(self.filepath)
        self.odswiez()

    def openExcelFileNameDialog(self):
        options = QtWidgets.QFileDialog.Options()
        #options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, odd = QtWidgets.QFileDialog.getOpenFileName(self,"Wybierz dokument excel z danymi.", "Dane","Excel Files (*.xlsx)", options=options)
        if fileName != None:
            return fileName
        else:
            return None 
    
    def eksportuj_wykres_glowny(self):
        filepath = self.saveWykresDialog("wykres_wyniki_czynniki")
        if filepath == None or filepath == "":
            return
        else:
            Plotter.export_plot(self.dane_firm, filepath)

    def eksportuj_wykres_kolowy(self):
        filepath = self.saveWykresDialog("wykres_kolowy")
        if filepath == None or filepath == "":
            return
        else:
            Plotter.export_pie_chart(self.dane_firm, filepath)

    def eksportuj_wykres_slupkowy(self):
        filepath = self.saveWykresDialog("wykres_slupkowy")
        if filepath == None or filepath == "":
            return
        else:
            Plotter.export_bar_chart(self.dane_firm,filepath)

    def zapisz_aktualny_arkusz(self):
        '''
        deprecated, do not use...
        '''
        savename = self.saveArkuszDialog()
        self.workbook.save(savename)
        
    def saveWykresDialog(self, nazwa_pliku):
        options = QtWidgets.QFileDialog.Options()
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self,"Zapisz wykres jako...",nazwa_pliku,"PNG files (*.png);; PDF files (*.pdf)", options=options)
        return fileName

class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super(MplCanvas, self).__init__(self.fig)
    def correct_ticks(self):
        for xtick in self.axes.get_xticklabels():
            xtick.set_rotation(20)

    def export(self, filepath):
        self.fig.savefig(filepath)
    

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = MainWindow()
    window.show()
    if window.flaga_stworzono_przykladowy:
        window.view_starting_popup()
    else:
        pass
    app.exec_()