from datetime import datetime
import pathlib
from pathlib import Path
import pandas as pd

class Sort():
    def make_dict(self, column_x='', column_y='', sheet_name = 'NDL6Staff'):
        df = pd.read_excel(Path(pathlib.Path.cwd(),'NDL6Staff.xlsx'), sheet_name=sheet_name)
        self.dict_FT = {}
        for i in range(0, len(df.FIO)):
            self.dict_FT[df[column_x][i]] = df[column_y][i]
        return self.dict_FT

    def sortByKey(self, column_x='', column_y='', up=True, sheet_name = 'NDL6Staff'):
        listTextA, listTextB, listTextC, listTextD = '', '', '',''
        self.dictSort = self.make_dict(column_x, column_y, sheet_name)
        if up:
            counter = 0
            self.listSort = sorted(self.dictSort.keys(), key=str)
            for i in self.listSort:
                if counter < 18:
                    if str(self.dictSort[i]) == '0':  # для исключения отображения 0 надбавки
                        continue
                    else:
                        if column_y == 'WorkPlace':
                            listTextA +='{0:<20}{1}\n'.format(i, 0)
                            listTextC +='{0:>5}\n'.format(str(self.dictSort[i]))
                            counter += 1
                        else:
                            listTextA +='{0:<20}\n'.format(i)
                            listTextC +='{0:>5}\n'.format(str(self.dictSort[i]))
                            counter += 1
                else:
                    if str(self.dictSort[i]) == '0':
                        continue
                    else:
                        if column_y == 'WorkPlace':
                            listTextB +='{0:<20}{1}\n'.format(i, 0)
                            listTextD +='{0:>5}\n'.format(str(self.dictSort[i]))
                        else:
                            listTextB +='{0:<20}\n'.format(i)
                            listTextD +='{0:>5}\n'.format(str(self.dictSort[i]))
            return (listTextA, listTextB, listTextC, listTextD)
        else:
            counter = 0
            self.listSort = sorted(self.dictSort.keys(), key=str, reverse=True)
            for i in self.listSort:
                if counter < 18:
                    if str(self.dictSort[i]) == '0':  # для исключения отображения 0 надбавки
                        continue
                    else:
                        if column_y == 'WorkPlace':
                            listTextA +='{0:<20}{1}\n'.format(i, 0)
                            listTextC +='{0:>5}\n'.format(str(self.dictSort[i]))
                            counter += 1
                        else:
                            listTextA +='{0:<20}\n'.format(i)
                            listTextC +='{0:>5}\n'.format(str(self.dictSort[i]))
                            counter += 1
                else:
                    if str(self.dictSort[i]) == '0':  # для исключения отображения 0 надбавки
                        continue
                    else:
                        if column_y == 'WorkPlace':
                            listTextB +='{0:<20}{1}\n'.format(i, 0)
                            listTextD +='{0:>5}\n'.format(str(self.dictSort[i]))
                        else:
                            listTextB +='{0:<20}\n'.format(i)
                            listTextD +='{0:>5}\n'.format(str(self.dictSort[i]))
            return (listTextA, listTextB, listTextC, listTextD)

    def sortByValue(self, column_x='', column_y='', up=True, key=str, sheet_name = 'NDL6Staff'):
        listTextA, listTextB = '', ''
        counter = 0
        self.dictSort = self.make_dict(column_x, column_y, sheet_name)
        if up:
            self.listSort = sorted(self.dictSort.values(), key=key)
            for i in self.listSort:
                for k in self.dictSort.keys():
                    if self.dictSort[k] == i:
                        if counter < 20:
                            if i == 0:
                                continue
                            else:
                                listTextA += k + '\t' + str(self.dictSort.pop(k)) + '\n'
                                counter += 1
                                break
                        else:
                            if i == 0:
                                continue
                            else:
                                listTextB += k + '\t' + str(self.dictSort.pop(k)) + '\n'
                                break
            return (listTextA, listTextB)
        else:
            self.listSort = sorted(self.dictSort.values(), key=key, reverse=True)
            for i in self.listSort:
                for k in self.dictSort.keys():
                    if self.dictSort[k] == i:
                        if counter < 20:
                            if i == 0:
                                continue
                            else:
                                listTextA += k + '\t' + str(self.dictSort.pop(k)) + '\n'
                                counter += 1
                                break
                        else:
                            if i == 0:
                                continue
                            else:
                                listTextB += k + '\t' + str(self.dictSort.pop(k)) + '\n'
                                break
            return (listTextA, listTextB)

    def sortByDate(self, column_x ='', column_y ='', up=True, sheet_name = 'NDL6Staff'):
        listTextA, listTextB = '', ''
        counter = 0
        self.dictSort = self.make_dict(column_x, column_y, sheet_name)
        if up:
            self.listSort = sorted(self.dictSort.values(), key= lambda d: datetime.strptime(d, '%d.%m.%Y'))
            for i in self.listSort:
                for k in self.dictSort.keys():
                    date = datetime.strptime(i, '%d.%m.%Y')
                    today = date.today()
                    age = today.year - date.year - ((today.month, today.day) < (date.month, date.day))
                    age = str(age)
                    if self.dictSort[k] == i:
                        if counter < 20:
                            listTextA += k + '\t' + str(self.dictSort.pop(k)) + '    ' + age + '\n'
                            counter += 1
                            break
                        else:
                            listTextB += k + '\t' + str(self.dictSort.pop(k)) + '    ' + age + '\n'
                            break
            return (listTextA, listTextB)
        else:
            self.listSort = sorted(self.dictSort.values(), key=lambda  d: datetime.strptime(d, '%d.%m.%Y'), reverse=True)
            for i in self.listSort:
                for k in self.dictSort.keys():
                    date = datetime.strptime(i, '%d.%m.%Y')
                    today = date.today()
                    age = today.year - date.year - ((today.month, today.day) < (date.month, date.day))
                    age = str(age)
                    if self.dictSort[k] == i:
                        if counter < 20:
                            listTextA += k + '\t' + str(self.dictSort.pop(k)) + '    ' + age + '\n'
                            counter += 1
                            break
                        else:
                            listTextB += k + '\t' + str(self.dictSort.pop(k)) + '    ' + age + '\n'
                            break
            return (listTextA, listTextB)

    def sort_BY_age(self, age_min=0, age_max=100, up=True, sheet_name = 'NDL6Staff'):
        listTextA= ''
        self.dictSort = self.make_dict('FIO', 'Birthday', sheet_name)
        if up:
            self.listSort = sorted(self.dictSort.values(), key=lambda d: datetime.strptime(d, '%d.%m.%Y'))
            for i in self.listSort:
                for k in self.dictSort.keys():
                    date = datetime.strptime(i, '%d.%m.%Y')
                    today = date.today()
                    age = today.year - date.year - ((today.month, today.day) < (date.month, date.day))
                    if self.dictSort[k] == i and age >= age_min and age <= age_max:
                        listTextA += k + '\t' + str(self.dictSort.pop(k)) + '    ' + str(age) + '\n'
                        break
            return (listTextA)
        else:
            self.listSort = sorted(self.dictSort.values(), key=lambda d: datetime.strptime(d, '%d.%m.%Y'), reverse=True)
            for i in self.listSort:
                for k in self.dictSort.keys():
                    date = datetime.strptime(i, '%d.%m.%Y')
                    today = date.today()
                    age = today.year - date.year - ((today.month, today.day) < (date.month, date.day))
                    if self.dictSort[k] == i and age >= age_min and age <= age_max:
                        listTextA += k + '\t' + str(self.dictSort.pop(k)) + '    ' + str(age) + '\n'
                        break
            return (listTextA)

    def sort_BY_oklad(self, oklad_min=0, oklad_max=0, up=True, sheet_name = 'NDL6Staff'):
        listTextA= ''
        self.dictSort = self.make_dict('FIO', 'Oklad', sheet_name)
        for i in self.dictSort:
            self.dictSort[i] = int(self.dictSort[i])
        if up:
            self.listSort = sorted(self.dictSort.values(), key=int)
            for i in self.listSort:
                for k in self.dictSort.keys():
                    if self.dictSort[k] == i and int(i) >= oklad_min and int(i) <= oklad_max:
                        listTextA += k + '\t' + str(self.dictSort.pop(k)) + '\n'
                        break
            return (listTextA)
        else:
            self.listSort = sorted(self.dictSort.values(), key=int, reverse=True)
            for i in self.listSort:
                for k in self.dictSort.keys():
                    if self.dictSort[k] == i and int(i) >= oklad_min and int(i) <= oklad_max:
                        listTextA += k + '\t' + str(self.dictSort.pop(k)) + '\n'
                        break
            return (listTextA)


