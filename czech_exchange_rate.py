from tkinter import *
from tkinter import filedialog, messagebox
import os
from datetime import datetime
import numpy as np
import requests
import urllib.request
import xmltodict
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import zipfile
import warnings
import yadisk

warnings.simplefilter("ignore")

def plus0(n):
    if len(n) == 1:
        n = '0'+n
    return n

def file_from_dir(directory):
    """Находит абсолютно все файлы из директории
    и сохраняет их в список paths"""
    paths = []
    files = os.listdir(directory)
    for file in files:
        if not file.endswith('.py'):
            paths.append(f'{directory}\{file}')
    return paths
                        
class CurrencyRate:
    respons = requests.get(r'https://www.cbr-xml-daily.ru/daily_json.js')
    data = respons.json()['Valute']['CZK']
    data_fr = []
    maxi = 0
    mini = 0

    def __init__(self, app):
        self.app = app
        #Текущий курс
        self.current = Button(app, text = 'Получить текущий курс\nЧешской кроны',
                                command = self.CurrentCourse)
        self.current.place(x = 115, y = 15, width = 180, height = 40)
        #Курс по дате
        self.daterate1 = Label(app, text = 'Узнать курс валюты на:')
        self.daterate1.place(x = 50, y = 65, width = 130, height = 15)
        self.daterate_day = Entry(app)
        self.daterate_day.place(x = 185, y = 65, width = 20, height = 15)
        self.daterate_day_1 = Label(app, text = 'д.')
        self.daterate_day_1.place(x = 208, y = 65, width = 10, height = 15)
        self.daterate_month = Entry(app)
        self.daterate_month.place(x = 220, y = 65, width = 20, height = 15)
        self.daterate_month_1 = Label(app, text = 'м.')
        self.daterate_month_1.place(x = 243, y = 65, width = 10, height = 15)
        self.daterate_year = Entry(app)
        self.daterate_year.place(x = 255, y = 65, width = 35, height = 15)
        self.daterate_year_1 = Label(app, text = 'г.')
        self.daterate_year_1.place(x = 293, y = 65, width = 10, height = 15)
        self.daterate_btn = Button(app, text = 'OK',
                                   command = self.DateRate)
        self.daterate_btn.place(x = 305, y = 60, width = 25, height = 25)
        #Надпись
        self.Downld = Label(app,text = 'Получить данные:')
        self.Downld.place(x = 50, y = 95, width = 100, height = 25)  
        #Источники данных
        self.data_from_csv = Button(app, text = 'Загрузить из файла',
                                    command = self.DataFromFile)
        self.data_from_csv.place(x = 130, y = 125, width = 200, height = 25)
        self.data_from_cbr = Button(app, text = 'Загрузить с сайта ЦБ РФ',
                                    command = self.DataFromCBR)
        self.data_from_cbr.place(x = 130, y = 155, width = 200, height = 25)
        #Надпись
        self.DataST1 = Label(app, text = f'Статус данных:')
        self.DataST1.place(x = 50, y = 180, width = 100, height = 25)
        self.DataST2 = Label(self.app, text = "Недоступно", fg = 'Red')
        self.DataST2.place(x = 145, y = 180, width = 65, height = 25)
        #Анализ
        self.analyze = Button(app, text = 'Проанализировать валюту',
                              command = self.RateAnalyze)
        self.analyze.place(x = 130, y = 220, width = 200, height = 25)
        #Вывод
        self.txt = Text(app, fg = "green")
        self.txt.place(x = 200, y = 250, width = 250, height = 140)
            
   
    def CurrentCourse(self):
        """Получает текущий курс Чешской кроны"""
        valute = self.data['Value']

        out_box = Message(text = round(valute/10, 2))
        out_box['bg'] = 'grey'
        out_box.place(x = 300, y = 15, width = 45, height = 40)

    def DateRate(self):
        """Получает курс Чешской кроны по дате"""
        if (1 <= len(self.daterate_day.get()) <= 2 and\
            1 <= len(self.daterate_month.get()) <= 2 and\
            len(self.daterate_year.get()) == 4) and\
            (self.daterate_day.get().isdigit() and\
            self.daterate_month.get().isdigit() and\
            self.daterate_year.get().isdigit()) and\
            (0 < int(self.daterate_day.get()) <= 31 and\
            0 < int(self.daterate_month.get()) <= 12 and\
            2010 <= int(self.daterate_year.get()) <= 2025):
            
            self.daterate_mes1 = Label(self.app,
                                       text = f'На {plus0(self.daterate_day.get())}.'
                                       f'{plus0(self.daterate_month.get())}.'
                                       f'{self.daterate_year.get()} '
                                       'курс Чешской кроны был:')
            self.daterate_mes1.place(x = 50, y = 83, width = 220, height = 15)
            
            dt = [plus0(self.daterate_day.get()), plus0(self.daterate_month.get()),
                            self.daterate_year.get()]
            date = '/'.join(dt)
            url = f"https://www.cbr.ru/scripts/XML_daily.asp?date_req={date}"
            response = urllib.request.urlopen(url).read()
            root = xmltodict.parse(response)

            for val in root['ValCurs']['Valute']:
                if val['@ID'] == self.data['ID']:
                    rate = round(float(val['VunitRate'].replace(',', '.')), 2)
                    self.daterate_mes2 = Message(text = rate)
                    self.daterate_mes2['bg'] = 'grey'
                    self.daterate_mes2.place(x = 275, y = 83, width = 35, height = 15)
        else:
            messagebox.showerror('Ошибка', 'Формат введённых данных неверен!')
   
    def DataFromCBR(self):
        """Загружает архив курса Чешской кроны за 5 лет с сайта ЦБ"""
        valuteID = self.data['ID']
        data_today = str(datetime.now().date()).split('-')[::-1]

        b = '.'.join(data_today)
        a = '01.01.2021'
        
        DDa, MMa, GGa = a.split('.')
        DDb, MMb, GGb = b.split('.')
        url = (f'https://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98956?'
               f'Posted=True&so=1&mode=1&VAL_NM_RQ={valuteID}&From={a}&To={b}&'
                f'FromDate={MMa}%2F{DDa}%2F{GGa}&ToDate={MMb}%2F{DDb}%2F{GGb}')
        urllib.request.urlretrieve(url, "CZK_data.csv")
        data = pd.read_excel("CZK_data.csv")
        if os.path.isfile("CZK_data.csv"):
            self.DataST2.config(text = 'Доступно', fg = 'Green')
            data['curs'] /= 10
            CurrencyRate.data_fr = pd.DataFrame(data, columns = ['data', 'curs'])
            CurrencyRate.data_fr.rename(columns = {'data' : 'Дата', 'curs' : 'Курс'},
                                        inplace = True)
            CurrencyRate.data_fr = CurrencyRate.data_fr.sort_values('Дата')
            CurrencyRate.data_fr.reset_index(drop=True, inplace=True)

    def DataFromFile(self):
        """Загружает данные курса Чешской кроны за 5 лет из файла с компьютера"""
        file_path = ''
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"),
                                                          ("Excel files", "*.xls"),
                                                          ("Excel files", "*.xlsx")])
        if file_path != '':
            data = pd.read_excel(file_path)
            data['curs'] /= 10
            CurrencyRate.data_fr = pd.DataFrame(data, columns = ['data', 'curs'])
            CurrencyRate.data_fr.rename(columns = {'data' : 'Дата', 'curs' : 'Курс'},
                                        inplace = True)
            CurrencyRate.data_fr = CurrencyRate.data_fr.sort_values('Дата')
            CurrencyRate.data_fr.reset_index(drop=True, inplace=True)
            dt1 = str(CurrencyRate.data_fr.Дата.iloc[0]).split(' ')[0].split('-')[0]
            dt2 = str(CurrencyRate.data_fr.Дата.iloc[-1]).split(' ')[0].split('-')[0]
            if abs(int(dt1) - int(dt2)) >= 4:
                self.DataST2.config(text = 'Доступно', fg = 'Green')
            else:
                messagebox.showerror('Ошибка', 'В файле недостаточно данных')

    def MaxMinSearch(self):
        """Вычисляет максимальное и минимальное значения курса Чешской кроны за 5 лет"""
        self.mini = round(float(CurrencyRate.data_fr['Курс'].min()),2)
        self.maxi = round(float(CurrencyRate.data_fr['Курс'].max()),2)
        mm_data = {'Минимальный курс' : [self.mini], 'Максимальный курс' : [self.maxi]}
        mm_fr = pd.DataFrame(mm_data)
        mm_fr.to_excel('Результаты.xlsx', index = False)
        self.txt.insert(END, f'Минимальный курс: {self.mini}\n'
                        f'Максимальный курс: {self.maxi}')

    def AvgSearch(self):
        """Вычисляет среднее значение курса Чешской кроны за каждый год"""
        yr = int(str(CurrencyRate.data_fr.Дата.iloc[0]).split(' ')[0].split('-')[0])
        goda = [yr]
        vals = []
        summ = 0
        count = 0
        for index, row in CurrencyRate.data_fr.iterrows():
            if str(yr) in str(row['Дата']):
                summ += round(float(row['Курс']),2)
                count += 1
            else:
                avg = round(summ/count, 2)
                yr += 1
                summ = 0
                count = 0
                goda.append(yr)
                vals.append(avg)
        #Для последнего значения
        if count > 0:
            avg = round(summ / count, 2)
            vals.append(avg)
                
        avg_fr = pd.read_excel('Результаты.xlsx')
        num_rows = avg_fr.shape[0] + 5
        new_fr = pd.DataFrame({'Год' : goda, 'Среднее' : vals})
        res_fr = pd.concat([avg_fr, new_fr], axis = 1)
        res_fr.to_excel('Результаты.xlsx', index = False)

        for i in range(len(goda)):
            self.txt.insert(END, f'\nСредний курс в {goda[i]} году: {vals[i]}')
        self.txt.insert(END, f'\n>>> Файл "Результаты.xlsx"')
        

    def fiveyearsgraph(self):
        """Строит график курса Чешской кроны за 5 лет"""
        max_date = ''
        min_date = ''
        for index, row in CurrencyRate.data_fr.iterrows():
            if round(row['Курс'], 2) == self.maxi:
                max_date = row['Дата']
            if round(row['Курс'], 2) == self.mini:
                min_date = row['Дата']

        years = mdates.YearLocator()
        months = mdates.MonthLocator()
        x = CurrencyRate.data_fr['Дата']
        y = CurrencyRate.data_fr['Курс']
        figure = plt.figure(figsize=[13, 5])
        ax = figure.add_axes([0.1, 0.15, 0.8, 0.6])
        ax.plot(x, y, 'g')
        ax.set_yticks(np.arange(round(y.min(),2), round(y.max()+0.05,2), 0.5)) 
        ax.xaxis.set_major_locator(years)
        ax.xaxis.set_minor_locator(months)
        plt.scatter(max_date, self.maxi, s = 50, color='red', label= 'Максимум')
        plt.text(max_date, self.maxi,'Максимум')
        plt.scatter(min_date, self.mini, s = 50, color='blue', label= 'Минимум')
        plt.text(min_date, self.mini,'Минимум')
        ax.set_xlabel("Дата")
        ax.set_ylabel("Курс")
        ax.set_title("Курс Чешской кроны")
        plt.legend(['Чешская крона'], loc = 2)
        plt.show()
        figure.savefig('ЧК за 5 лет.jpg')
        self.txt.insert(END, f'\n>>> Файл "ЧК за 5 лет.jpg"')
        
    def Diagram(self):
        """Строит диаграмму средних значений курса Чешской кроны по годам"""
        avg_data = pd.read_excel('Результаты.xlsx')
        figure = plt.figure(figsize = [6, 8])
        ax = figure.add_axes([0.1, 0.1, 0.8, 0.9])
        ax.bar(avg_data['Год'], avg_data['Среднее'], color = 'orange')
        ax.set_yticks(avg_data['Среднее'])
        ax.set_xlabel("Год")
        ax.set_ylabel("Средний курс")
        plt.legend(['Чешская крона'], loc = 2)
        plt.show()
        figure.savefig('ЧК Средн.jpg')
        self.txt.insert(END, f'\n>>> Файл "ЧК Средн.jpg"')
        
    def pregr_years(self):
        """Запрашивает начальную и конечнуб даты для графика за промежуток времени"""
        fstpart = Label(self.app, text = "Построить график с")
        fstpart.place(x = 10, y = 395, width = 120, height = 15)
        #c
        self.fr_day = Entry(self.app)
        self.fr_day.place(x = 132, y = 396, width = 20, height = 15)
        fr_day_1 = Label(self.app, text = 'д.')
        fr_day_1.place(x = 155, y = 395, width = 10, height = 15)
        self.fr_month = Entry(self.app)
        self.fr_month.place(x = 167, y = 396, width = 20, height = 15)
        fr_month_1 = Label(self.app, text = 'м.')
        fr_month_1.place(x = 190, y = 395, width = 10, height = 15)
        self.fr_year = Entry(self.app)
        self.fr_year.place(x = 202, y = 396, width = 35, height = 15)
        fr_year_1 = Label(self.app, text = 'г.')
        fr_year_1.place(x = 240, y = 395, width = 10, height = 15)
        #no
        topart = Label(self.app, text = "по")
        topart.place(x = 252, y = 395, width = 15, height = 15)
        self.to_day = Entry(self.app)
        self.to_day.place(x = 270, y = 396, width = 20, height = 15)
        to_day_1 = Label(self.app, text = 'д.')
        to_day_1.place(x = 293, y = 395, width = 10, height = 15)
        self.to_month = Entry(self.app)
        self.to_month.place(x = 305, y = 396, width = 20, height = 15)
        to_month_1 = Label(self.app, text = 'м.')
        to_month_1.place(x = 327, y = 395, width = 10, height = 15)
        self.to_year = Entry(self.app)
        self.to_year.place(x = 340, y = 396, width = 35, height = 15)
        to_year_1 = Label(self.app, text = 'г.')
        to_year_1.place(x = 377, y = 395, width = 10, height = 15)
        yr_btn = Button(self.app, text = 'OK',
                                   command = self.gr_years)
        yr_btn.place(x = 195, y = 415, width = 70, height = 20)

    def gr_years(self):
        """Строит график курса Чешской кроны за выбранный промежуток времени"""
        day_st = plus0(self.fr_day.get())
        month_st = plus0(self.fr_month.get())
        year_st = self.fr_year.get()
        st = [year_st, month_st, day_st]
        start_date = '-'.join(st)
        day_fn = plus0(self.to_day.get())
        month_fn = plus0(self.to_month.get())
        year_fn = self.to_year.get()
        fn = [year_fn, month_fn, day_fn]
        finish_date = '-'.join(fn)

        ind_start_date = -3
        ind_finish_date = -3
        for index, row in CurrencyRate.data_fr.iterrows():
            if start_date in str(row['Дата']):
                ind_start_date = index
            elif finish_date in str(row['Дата']):
                ind_finish_date = index
            else:
                continue
            
        if ind_start_date == -3:
            messagebox.showerror('Ошибка', 'Начальная дата не найдена в базе данных')
        elif ind_finish_date == -3:
            messagebox.showerror('Ошибка', 'Конечная дата не найдена в базе данных')
        else:
            x = CurrencyRate.data_fr['Дата'].iloc[ind_start_date:ind_finish_date]
            y = CurrencyRate.data_fr['Курс'].iloc[ind_start_date:ind_finish_date]
            years = mdates.YearLocator()
            months = mdates.MonthLocator()
            figure = plt.figure(figsize=[12, 5])
            ax = figure.add_axes([0.1, 0.15, 0.8, 0.6])
            ax.plot(x, y, 'r')
            ax.set_yticks(np.arange(round(y.min(),2), round(y.max()+0.05,2), 0.25)) 
            ax.xaxis.set_major_locator(years)
            ax.xaxis.set_minor_locator(months)
            ax.set_xlabel("Дата")
            ax.set_ylabel("Курс")
            ax.set_title("Курс Чешской кроны")
            plt.legend(['Чешская крона'], loc = 0)
            plt.show()
            figure.savefig(f'ЧК с {start_date} по {finish_date}.jpg')
            self.txt.insert(END, f'\n>>> Файл "ЧК с {start_date} по {finish_date}.jpg"')
            
    def RateAnalyze(self):
        """Выполняет анализ курса и предлагает построить графики и сохранить файлы"""
        if self.DataST2['text'] == 'Доступно':
            self.MaxMinSearch()
            self.AvgSearch()
            postgr = Label(self.app, text = "Составить график:")
            postgr.place(x = 50, y = 250, width = 100, height = 25)
            gr_five = Button(self.app, text = 'Построить график за 5 лет',
                             command = self.fiveyearsgraph)
            gr_five.place(x = 10, y = 280, width = 185, height = 25)
            gr_diaga = Button(self.app, text = 'Построить диаграмму среднего\nкурса',
                             command = self.Diagram)
            gr_diaga.place(x = 10, y = 310, width = 185, height = 40)
            gr_goda = Button(self.app, text = 'Построить график за время',
                             command = self.pregr_years)
            gr_goda.place(x = 10, y = 355, width = 185, height = 25)
            saveLB = Label(self.app, text = "Сохранить файлы:")
            saveLB.place(x = 180, y = 437, width = 100, height = 25)
            saveLC = Button(self.app, text = 'Сохранить локально',
                            command = self.saveLocal)
            saveLC.place(x = 75, y = 465, width = 150, height = 25)
            saveCS = Button(self.app, text = 'Сохранить в облако',
                            command = self.saveCloud)
            saveCS.place(x = 235, y = 465, width = 150, height = 25)
        else:
            messagebox.showerror('Ошибка', 'Вы не добавили данные')
                
    def saveLocal(self):
        """Формирует zip-файл из файлов, полученных в ходе выполнения программы,
            и сохраняет в указанное место"""
        file_path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
        with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file in file_from_dir(os.getcwd()):
                zipf.write(file)
        messagebox.showinfo('Сохранение', 'Файлы успешно сохранены')        

    def saveCloud(self):
        """Сохраняет файлы, полученнык в ходе выполнения программы, в облачное хранилище"""
        y = yadisk.YaDisk(token = '')
        if y.check_token():
        #Проверка наличия директории
            if y.is_dir('/РГЗ') == False:
                y.mkdir('РГЗ')
            #Сохраняем все файлы в облако
            for file in file_from_dir(os.getcwd()):
                y.upload(file, f'/РГЗ/{os.path.basename(file)}',
                         overwrite = True)
            messagebox.showinfo('Сохранение', 'Файлы успешно сохранены')
            self.txt.insert(END, f'\n>>> https://disk.yandex.ru/d/qgEeY5kj56TdWQ"')
        else:
            messagebox.showerror('Ошибка', 'Проблемы с токеном')
            
window = Tk()
window.geometry("460x500")
window.title("Чешская крона")
CurrencyRate(window)
window.mainloop()
