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
from numpy import array

'''
solution to not being able to save excel file while the program is running:
https://stackoverflow.com/questions/31416842/openpyxl-does-not-close-excel-workbook-in-read-only-mode
'''


class MainWindow(Ui):

    '''
    here is a logic of a program
    '''
    flag_file_exists = False
    sample_xl_created_flag = False
    filepath = ""
    input_data = []

    def __init__(self):
        super(MainWindow, self).__init__()
        Ui.setupUi(self, self)
        self.filepath = "Dane.xlsx"
        data_file = Path("Dane.xlsx")
        if data_file.is_file():
            self.flag_file_exists = True
            self.input_data = f.full_import(self.filepath)
            del data_file
        else:
            #tworzenie pliku
            x = f.daj_przykladowe_dane()
            f.exporter(x[0],x[1],x[2],x[3],x[4],x[5])
            self.sample_xl_created_flag = True
            self.flag_file_exists = True
        #for x in self.dane_firm:
        #    print(x)
        self.refresh()
        #self.workbook.close()
 
    
    #rysowanie_wykresow
    def add_main_plot(self):
        layout = self.mainPlot.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.mainPlot.layout())
        layout = QtWidgets.QVBoxLayout()
        self.scGlowny = MplCanvas(self)
        for dane in self.input_data:
            self.scGlowny.axes.plot(dane[1],dane[4],label = dane[0], linewidth = 4)
        #sc.axes.text(1,3.6,"Wyniki firm względem czynników", horizontalalignment = 'center', verticalalignment = 'center', fontsize=12)
        self.scGlowny.axes.set_title("Wykres wyników firm względem czynników.")
        self.scGlowny.axes.grid(axis = 'y',color = 'gray', linestyle='-', linewidth = 1)
        self.scGlowny.axes.legend()
        self.scGlowny.correct_ticks()
        layout.addWidget(self.scGlowny)
        self.mainPlot.setLayout(layout)

    def add_bar_chart(self):
        Plotter.draw_and_print(self.input_data)
        layout = self.barChart.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.barChart.layout())
        layout = QtWidgets.QVBoxLayout()
        label = QtWidgets.QLabel(self)
        pixmap = QPixmap('plt/bar_chart.png')
        pixmap = pixmap.scaled(451,331, Qt.KeepAspectRatio, Qt.FastTransformation)
        label.setPixmap(pixmap)
        layout.addWidget(label)
        self.barChart.setLayout(layout)

    def add_pie_chart(self):
        layout = self.pieChart.layout()
        if layout != None:
            QObjectCleanupHandler().add(self.pieChart.layout())
        layout = QtWidgets.QVBoxLayout()
        self.scPie = MplCanvas(self)
        wyniki = []
        labele = []
        for dane in self.input_data:
            wyniki.append(sum(dane[4]))
            labele.append(dane[0])
        explode = []
        dex = wyniki.index(max(wyniki))
        for i in range(0,len(wyniki)):
            if i == dex:
                explode.append(0.05)
            else:
                explode.append(0.01)
        wyniki = array(wyniki)
        labele = array(labele)
        self.scPie.axes.pie(x = wyniki,labels = labele, explode = explode, shadow = False, autopct='%1.1f%%')
        self.scPie.axes.set_title("Udział w wynikach każdej z firm")
        layout.addWidget(self.scPie)
        self.pieChart.setLayout(layout)
            
    def read_data(self):
        self.input_data = [x for x in f.full_import(self.filepath)]

    def refresh(self):
        self.read_data()
        self.add_main_plot()
        self.add_bar_chart()
        self.add_pie_chart()
        #self.dodaj_wykres_glowny()
        #self.dodaj_wykres_prawy_gora()

    def edit_xl(self):
        '''
        opens excel file under the specified filename
        or if the file is already opened, sets focus to that
        excel window.
        '''
        p = str(PureWindowsPath(self.filepath))
        command = "start " + "\"title\" " + "\"" + p + "\""
        if not showexcel.show_excel_window(os.path.basename(self.filepath)):
            os.system(command)

    def open_given_xl(self):
        new_filepath = self.openExcelFileNameDialog()
        if new_filepath == None or new_filepath == "":
            pass
        else:
            self.filepath = new_filepath
            self.workbook = f.load_workbook(self.filepath)
        self.refresh()

    def openExcelFileNameDialog(self):
        options = QtWidgets.QFileDialog.Options()
        #options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, odd = QtWidgets.QFileDialog.getOpenFileName(self,"Wybierz dokument excel z danymi.", "Dane","Excel Files (*.xlsx)", options=options)
        if fileName != None:
            return fileName
        else:
            return None 
    
    def export_main_plot(self):
        filepath = self.savePlotDialog("wykres_wyniki_czynniki")
        if filepath == None or filepath == "":
            return
        else:
            Plotter.export_plot(self.input_data, filepath)

    def export_pie_chart(self):
        filepath = self.savePlotDialog("wykres_kolowy")
        if filepath == None or filepath == "":
            return
        else:
            Plotter.export_pie_chart(self.input_data, filepath)

    def export_bar_chart(self):
        filepath = self.savePlotDialog("wykres_slupkowy")
        if filepath == None or filepath == "":
            return
        else:
            Plotter.export_bar_chart(self.input_data,filepath)

    def savePlotDialog(self, filename):
        options = QtWidgets.QFileDialog.Options()
        fileName, _ = QtWidgets.QFileDialog.getSaveFileName(self,"Zapisz wykres jako...", filename,"PNG files (*.png);; PDF files (*.pdf)", options=options)
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
    if window.sample_xl_created_flag:
        window.view_starting_popup()
    else:
        pass
    app.exec_()