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

    def oklad(self, up = True):
        self.ids['text_A'].text = Sort.sortByValue(Sort(), 'FIO', 'Oklad', up,int,'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByValue(Sort(), 'FIO', 'Oklad', up,int,'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByValue(Sort(), 'FIO', 'Oklad', up,int,'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByValue(Sort(), 'FIO', 'Oklad', up,int,'NDL6Staff')[3]

    def prof(self, up=True):
        self.ids['text_A'].text = Sort.sortByValue(Sort(), 'FIO', 'Prof', up, str, 'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByValue(Sort(), 'FIO', 'Prof', up, str, 'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByValue(Sort(), 'FIO', 'Prof', up, str, 'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByValue(Sort(), 'FIO', 'Prof', up, str, 'NDL6Staff')[3]

    def workplace(self, up = True):
        self.ids['text_A'].text = Sort.sortByKey(Sort(), 'FIO', 'WorkPlace', up, 'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByKey(Sort(), 'FIO', 'WorkPlace', up, 'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByKey(Sort(), 'FIO', 'WorkPlace', up, 'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByKey(Sort(), 'FIO', 'WorkPlace', up, 'NDL6Staff')[3]

    def bonus(self, up = True):
        self.ids['text_A'].text = Sort.sortByValue(Sort(), 'FIO', 'Bonus', up,int,'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByValue(Sort(), 'FIO', 'Bonus', up,int,'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByValue(Sort(), 'FIO', 'Bonus', up,int,'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByValue(Sort(), 'FIO', 'Bonus', up,int,'NDL6Staff')[3]

    def phone(self, up = True):
        self.ids['text_A'].text = Sort.sortByKey(Sort(), 'FIO', 'Phone', up, 'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByKey(Sort(), 'FIO', 'Phone', up, 'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByKey(Sort(), 'FIO', 'Phone', up, 'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByKey(Sort(), 'FIO', 'Phone', up, 'NDL6Staff')[3]

    def adress(self, up = True):
        self.ids['text_A'].text = Sort.sortByKey(Sort(), 'FIO', 'Adress', up, 'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByKey(Sort(), 'FIO', 'Adress', up, 'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByKey(Sort(), 'FIO', 'Adress', up, 'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByKey(Sort(), 'FIO', 'Adress', up, 'NDL6Staff')[3]

    def birthday(self, up = True):
        self.ids['text_A'].text = Sort.sortByDate(Sort(), 'FIO', 'Birthday', up, 'NDL6Staff')[0]
        self.ids['text_C'].text = Sort.sortByDate(Sort(), 'FIO', 'Birthday', up, 'NDL6Staff')[2]
        self.ids['text_B'].text = Sort.sortByDate(Sort(), 'FIO', 'Birthday', up, 'NDL6Staff')[1]
        self.ids['text_D'].text = Sort.sortByDate(Sort(), 'FIO', 'Birthday', up, 'NDL6Staff')[3]





    # def up(self):
    #     if '56001' in self.ids['text_C'].text:
    #         self.tabel()
    #     elif '56001' not in self.ids['text_C'].text:
    #         self.oklad()
    # def down(self):
    #     if '56001' in self.ids['text_C'].text:
    #         self.tabel(False)
    #     elif '56001' not in self.ids['text_C'].text:
    #         self.oklad(False)


class LabInfoApp(App):
    def build(self):
        return Conteiner()








LabInfoApp().run()




