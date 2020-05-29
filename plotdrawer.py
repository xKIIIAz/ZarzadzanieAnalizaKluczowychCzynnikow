from matplotlib import pyplot as plt

class Plotter:
    @staticmethod
    def licz_dane_do_wykresu_slupkowego(dane_firm):
        nazwy_firm = []
        sumy_wynikow_firm = []
        for x in dane_firm:
            nazwy_firm.append(x[0])
            sumy_wynikow_firm.append(sum(filter(None,x[4])))
        return nazwy_firm, sumy_wynikow_firm
    
    @staticmethod
    def draw_and_print(dane_firm):
        a,b = Plotter.licz_dane_do_wykresu_slupkowego(dane_firm)
        Plotter.plot_bar_chart(a,b)


    @staticmethod
    def plot_bar_chart(nazwy_firm, sumy_wynikow_firm):
        index = sumy_wynikow_firm.index(max(sumy_wynikow_firm))
        bars = plt.bar(nazwy_firm,sumy_wynikow_firm, width=0.6,color='gray')
        bars[index].set_color('magenta')
        plt.title("Podsumowanie wyników względem firm")
        plt.savefig("fig.png")
        plt.clf()
    