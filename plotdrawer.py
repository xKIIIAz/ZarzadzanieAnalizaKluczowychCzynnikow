from matplotlib import pyplot as plt
from numpy import array

class Plotter:
    @staticmethod
    def licz_dane_do_wykresu_slupkowego(dane_firm):
        nazwy_firm = []
        sumy_wynikow_firm = []
        paczka = []
        for x in dane_firm:
            tupla = []
            tupla.append(x[0])
            tupla.append(sum(filter(None,x[4])))
            paczka.append(tupla)
        paczka.sort(key= lambda x: x[1], reverse = True)
        for x in paczka:
            nazwy_firm.append(x[0])
            sumy_wynikow_firm.append(x[1])
        return nazwy_firm, sumy_wynikow_firm
    
    @staticmethod
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            stheight = "{:.1f}".format(height)
            diff = len(stheight.replace(".","")) * 0.005
            plt.text(rect.get_x() + rect.get_width()/2 - diff, height*0.9,
                stheight,
                ha='center', va='bottom')

    @staticmethod
    def draw_and_print(dane_firm):
        a,b = Plotter.licz_dane_do_wykresu_slupkowego(dane_firm)
        Plotter.plot_bar_chart(a,b)

    @staticmethod
    def export_bar_chart(dane_firm, filepath):
        a,b = Plotter.licz_dane_do_wykresu_slupkowego(dane_firm)
        Plotter.save_bar_chart(a,b,filepath)

    @staticmethod
    def plot_bar_chart(nazwy_firm, sumy_wynikow_firm):
        index = sumy_wynikow_firm.index(max(sumy_wynikow_firm))
        bars = plt.bar(nazwy_firm,sumy_wynikow_firm, width=0.6,color='gray')
        bars[index].set_color('magenta')
        Plotter.autolabel(bars)
        plt.title("Podsumowanie wyników względem firm")
        plt.savefig("plt/bar_chart.png")
        plt.clf()

    @staticmethod
    def save_bar_chart(nazwy_firm, sumy_wynikow_firm, filepath):
        index = sumy_wynikow_firm.index(max(sumy_wynikow_firm))
        bars = plt.bar(nazwy_firm,sumy_wynikow_firm, width=0.6,color='gray')
        bars[index].set_color('magenta')
        Plotter.autolabel(bars)
        plt.title("Podsumowanie wyników względem firm")
        plt.savefig(filepath, dpi = 300)
        plt.clf()
    
    @staticmethod
    def export_pie_chart(dane_firm, filepath):
        wyniki = []
        labele = []
        for dane in dane_firm:
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
        pie = plt.pie(x = wyniki, labels = labele, explode = explode, shadow = False, autopct='%1.1f%%')
        plt.title("Udział w wynikach każdej z firm")
        plt.savefig(filepath, dpi = 300)
        plt.clf()
    
    @staticmethod
    def export_plot(dane_firm, filepath):
        for dane in dane_firm:
            plt.plot(dane[1],dane[4],label = dane[0], linewidth = 4)
        #sc.axes.text(1,3.6,"Wyniki firm względem czynników", horizontalalignment = 'center', verticalalignment = 'center', fontsize=12)
        plt.title("Wykres wyników firm względem czynników.")
        plt.legend()    
        plt.grid(axis = 'y',color = 'gray', linestyle='-', linewidth = 1)
        _ , xtic = plt.xticks()
        for tic in xtic:
            tic.set_rotation(20)
        plt.subplots_adjust(bottom=0.2)
        plt.savefig(filepath, dpi = 300, pad_inches = 50)
        plt.clf()