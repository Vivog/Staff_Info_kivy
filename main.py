import pathlib
import random 
from kivymd.app import App
from kivy.uix.gridlayout import GridLayout
from pathlib import Path
import pandas as pd
from btn_commands import Sort

class Conteiner(GridLayout):
    def aver_oklad(self):
        if self.ids['aver_oklad'].text == "":
            df = pd.read_excel(Path(pathlib.Path.cwd(), 'NDL6Staff.xlsx'))
            aver = int(df['Oklad'].mean())
            aver = round(aver)
            strAverOklad = f"{aver}"
            all_sum = 0
            for i in range(0, len(df.Oklad)):
                all_sum += df.Oklad[i] + int(df.Oklad[i] * (df.Bonus[i] / 100))
            averZar = all_sum / len(df.Oklad)
            averZar = round(averZar)
            strAverZar = f'{averZar}'
            strAverL = f'{strAverOklad}\n{strAverZar}'
            self.ids['text_aver'].text = strAverL
            self.ids['aver_oklad'].text = "!"
        elif self.ids['aver_oklad'].text == '!':
            strAverL = '    \n    '
            self.ids['text_aver'].text = strAverL
            self.ids['aver_oklad'].text = ""

    def tabel(self, up = True):
        self.ids['text_A'].text = Sort.sortByKey(Sort(), 'FIO', 'Tabel', up, 'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByKey(Sort(), 'FIO', 'Tabel', up, 'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByKey(Sort(), 'FIO', 'Tabel', up, 'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByKey(Sort(), 'FIO', 'Tabel', up, 'NDL6Staff')[3]
    def up(self):
        self.tabel()
    def down(self):
        self.tabel(False)


class LabInfoApp(App):
    def build(self):
        return Conteiner()








LabInfoApp().run()