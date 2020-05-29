
from plotdrawer import Plotter
from gui import Ui_MainWindow as Ui
from gui import QtWidgets
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt
from PyQt5.QtCore import QObjectCleanupHandler
from pathlib import Path
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
    - handlers for buttons underneath the plots
    - dwo additional plots - I'm doing it now
    - add an excel file to a workbook
    - add function "addListeners"

    '''
    
    flaga_istnieje_plik = False
    flaga_stworzono_przykladowy = False
    filepath = ""
    dane_firm = []
    ranking_firm = []

    def __init__(self):
        super(MainWindow, self).__init__()
        Ui.setupUi(self, self)
        self.filepath = "Dane.xlsx"
        data_file = Path("Dane.xlsx")
        print(data_file)
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
        for x in self.dane_firm:
            print(x)
        self.odswiez()
        #self.workbook.close()
 
    
    #rysowanie_wykresow
    def dodaj_wykres_glowny(self):
        layout = self.wykresGlowny.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.wykresGlowny.layout())
        layout = QtWidgets.QVBoxLayout()
        sc = MplCanvas(self)
        for dane in self.dane_firm:
            sc.axes.plot(dane[1],dane[4],label = dane[0], linewidth = 4)
        #sc.axes.text(1,3.6,"Wyniki firm względem czynników", horizontalalignment = 'center', verticalalignment = 'center', fontsize=12)
        sc.axes.set_title("Wykres wyników firm względem czynników.")
        sc.axes.legend()
        layout.addWidget(sc)
        self.wykresGlowny.setLayout(layout)
    '''
    def dodaj_wykres_prawy_gora(self):
        layout1 = QtWidgets.QHBoxLayout()
        barplot = MplCanvas(self)
        pos = np.arange(len(self.ranking_firm))
        width = 0.35
        barplot.axes.bar(pos, self.ranking_firm[1], align = 'center', width = width)
        layout1.addWidget(barplot)
        self.wykresPrawyGora.setLayout(layout1)
    '''
    def dodaj_wykres_prawy_gora(self):
        Plotter.draw_and_print(self.dane_firm)
        layout = self.wykresPrawyGora.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.wykresPrawyGora.layout())
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel(self)
        pixmap = QPixmap('fig.png')
        pixmap = pixmap.scaled(451,331, Qt.KeepAspectRatio, Qt.FastTransformation)
        label.setPixmap(pixmap)
        layout.addWidget(label)
        self.wykresPrawyGora.setLayout(layout)

    def dodaj_wykres_prawy_dolny(self):
        layout = self.wykresPrawyDol.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.wykresPrawyDol.layout())
        layout = QtWidgets.QVBoxLayout()
        sc = MplCanvas(self)
        wyniki = []
        labele = []
        for dane in self.dane_firm:
            wyniki.append(sum(dane[4]))
            labele.append(dane[0])
        wyniki = np.array(wyniki)
        labele = np.array(labele)
        print(type(wyniki))
        print(type(labele))
        print(wyniki)
        print(labele)
        sc.axes.pie(x = wyniki,labels = labele)
        #sc.axes.text(1,3.6,"Wyniki firm względem czynników", horizontalalignment = 'center', verticalalignment = 'center', fontsize=12)
        sc.axes.set_title("Udział w wynikach każdej z firm")
        layout.addWidget(sc)
        self.wykresPrawyDol.setLayout(layout)

    def wylicz_ranking(self):
        '''
        funkcja wylicza sume wynikow kazdej z firm i umieszcza na liscie ranking_firm w postaci par
        tzn.:
        [ [nazwa_firmy, laczny_wynik] , [nazwa_firmy, laczny_wynik], ..., [nazwa_firmy, laczny_wynik] ]
        zmiana:
        Umieszcza je w postaci:
        [ [nazwa_firmy, nazwa_firmy1..., nazwa_firmyN] , [laczny_wynik, laczny_wynik1, ..., laczny_wynikN] ]
        '''
        self.ranking_firm = []
        for x in self.dane_firm:
            firma = []
            firma.append(x[0])
            self.ranking_firm.append(firma)
        for x in self.dane_firm:
            firma = []
            firma.append(sum(filter(None,x[4])))
            self.ranking_firm.append(firma)
        print(self.ranking_firm)
            
    def czytaj_dane(self):
        self.dane_firm = [x for x in f.full_import(self.filepath)]

    def odswiez(self):
        self.czytaj_dane()
        self.wylicz_ranking()
        self.dodaj_wykres_glowny()
        self.dodaj_wykres_prawy_gora()
        self.dodaj_wykres_prawy_dolny()
        #self.dodaj_wykres_glowny()
        #self.dodaj_wykres_prawy_gora()

    def otworz_wybrany_excel(self):
        self.filename = self.openExcelFileNameDialog()
        if filename != None:
            self.workbook = f.load_workbook(self.filename)
        self.odswiez()

    def zapisz_aktualny_arkusz(self):
        savename = self.saveArkuszDialog()
        self.workbook.save(savename)

    def openExcelFileNameDialog(self):
        options = QtWidgets.QFileDialog.Options()
        #options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, odd = QtWidgets.QFileDialog.getOpenFileName(self,"Wybierz dokument excel z danymi.", "Dane","Excel Files (*.xlsx)", options=options)
        if fileName != None:
            print("none")
            return fileName
        else:
            return None  
        
    def saveArkuszDialog(self):
        options = QtWidgets.QFileDialog.Options()
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self,"Zapisz arkusz jako...","dane_firm","Excel Files (*.xlsx)", options=options)
        if fileName:
            return fileName
            
class Canvas(FigureCanvas):
    def __init__(self, parent = None, width = 5, height = 4, dpi=100):
        fig = Figure(figsize=(width,height),dpi=dpi)
        self.axes = fig.add_subplot(111)
        FigureCanvas.__init__(self,fig)
        self.setParent(parent)
    def plot(self, vals, labels,):
        ax = self.figure.add_subplot(111)
        ax.pie(vals, labels=labels)
        return ax

class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)
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