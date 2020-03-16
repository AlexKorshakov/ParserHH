import csv
import os
import requests
import re
import win32com.client as com_client
from datetime import datetime
from typing import Union, List
from bs4 import BeautifulSoup as bs

global start_def


class ExcelApp(object):

    @classmethod
    def app_open(self):
        # открываем Excel в скрытом режиме, отключаем обновление экрана и сообщения системы
        excel = com_client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        return print('Книга excel открыта')

    @classmethod
    def app_close(self):
        # включаем обновление экрана и сообщения системы
        excel = com_client.Dispatch("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = True
        excel.ScreenUpdating = True
        # выходим из Excel
        excel.Quit()
        return print('Книга excel закрыт')

    @classmethod
    def file_create(self, full_path):
        excel = com_client.Dispatch("Excel.Application")
        wbook = excel.Workbooks.Add()
        wbook.Worksheets.Add()
        wbook.Worksheets.Add()
        wbook.SaveAs(full_path)
        return print('Книга создана в full_path')


def HH_parse(base_url, headers):
    # парсим сайт
    start_def: datetime = datetime.now()
    jobs: list = []
    urls: list = []
    urls.append(base_url)  # список url
    session = requests.Session()  # создаём сессию
    request = session.get(base_url, headers=headers)
    if request.status_code == 200:  # если сервер ответил то

        soup = bs(request.content, 'lxml')
        try:
            pagination = soup.find_all('a', attrs={'data-qa': 'pager-page'})
            count = int(pagination[-1].text)
            for i in range(count):
                if i >= 1:
                    url = str(base_url + '&page=' + str(i))
                    if url not in urls:
                        urls.append(url)
        except:
            pass
    for url in urls:
        request = session.get(url, headers=headers)
        soup = bs(request.content, 'lxml')
        divs = soup.find_all('div', attrs={'data-qa': 'vacancy-serp__vacancy'})
        iRow = 0
        for div in divs:
            try:
                title = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-title'}).text
                href = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-title'})['href']
                company = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-employer'}).text
                text1 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'}).text
                try:
                    text2 = div.find('div', attrs={'data-qa': 'vacancy-vacancy-serp__vacancy_snippet_requirement'}).text
                    content = str(text1) + ' ' + str(text2)
                except:
                    content = str(text1)
                iRow: Union[int, iRow] = iRow + 1
                jobs.append({
                    'rowNom': iRow,
                    'title': title,
                    'href': href,
                    'company': company,
                    'content': content
                })
            except:
                pass

        print('Всего:' + str(len(jobs)) + ' ' + 'Время выполнения lxml: ' + str(datetime.now() - start_def))
    else:
        print('Error or Done ' + str(request.status_code))
    return jobs


def deep_pars(jobs, headers):
    # парсим найденные страницы
    global start
    deepcont: list = []
    try:
        start = datetime.now()
        for job in jobs:
            session = requests.Session()
            request = session.get(job['href'], headers=headers)
            if request.status_code == 200:
                soup = bs(request.content, 'lxml')
                try:
                    pagination = soup.find_all('div', attrs={'data-qa': 'vacancy-description'})
                    pag_text = clear_string(str(pagination))
                    deepcont.append({'deepcontent': pag_text})
                except:
                    print('Err pagination')
                    pass
    except:
        print('Err deep_pars')
        pass
    finally:
        print('Done deep_pars')

    print('Время выполнения deep_pars: ' + str(datetime.now() - start))
    return deepcont


def clear_string(string):
    # очищаем результат от html разметки с помощью регулярных выражений
    start_def: datetime = datetime.now()

    cleanr = {
        '</p>', '<p>', '</li>', '<li>', '</ul>', '<ul>', '</div>', '<div>', '</strong>', '<strong>',
        '</em>', '<em>', '</br>', '<br>', '<br/>', '  '
    }
    for cl in cleanr:
        string = re.sub(cl, ' ', string)

    a = 0
    for i in list(string):
        a += 1
        if i == ">":
            string = string[a: -1]

    print('Время выполнения clear_string: ' + str(datetime.now() - start_def))
    return string


def file_writer_win32(jobs, number_repetitions, deepcont, full_path):
    # записываем результаты парсинга (jobs), глубокого парсинга (deepcont) в файл по пути full_path
    start_def: datetime = datetime.now()
    ExcelApp.app_open()  # открываем Excel в скрытом режиме

    iRow: int = 0
    try:
        if os.path.exists(full_path):  # если файл с данными уже был то
            os.remove(full_path)  # удаляем предедущие данные если они были
            ExcelApp.file_create(full_path)  # создаём файл с данными
        else:
            ExcelApp.file_create(full_path)  # создаём файл с данными

        try:
            print('начало file_writer_win32')
            spisok: list = []

            try:
                wb = com_client.Dispatch("Excel.Application").Workbooks.Open(full_path)
                print('Книга создана')

                iRow: int = 0
                ws_list1_List: list = []
                for job in jobs:
                    iRow += 1
                    wb.Worksheets('Лист1').Cells(iRow, 1).Value = iRow  # - 2
                    wb.Worksheets('Лист1').Cells(iRow, 2).Value = job['title']
                    wb.Worksheets('Лист1').Cells(iRow, 3).Value = job['href']
                    wb.Worksheets('Лист1').Cells(iRow, 4).Value = job['company']
                    wb.Worksheets('Лист1').Cells(iRow, 5).Value = job['content']

                iRow: int = 0
                for cont in deepcont:
                    iRow += 1
                    wb.Worksheets('Лист2').Cells(iRow, 1).Value = iRow  # - 2
                    wb.Worksheets('Лист2').Cells(iRow, 2).Value = cont['deepcontent']
                    spisok.append(cont['deepcontent'])

                opts_ex, resutlLsts = list_spliter(spisok, number_repetitions)

                iRow: int = 0
                # список с исключениями
                for opt in opts_ex:
                    iRow += 1
                    wb.Worksheets('Лист3').Cells(iRow, 1).Value = iRow  # - 2
                    wb.Worksheets('Лист3').Cells(iRow, 2).Value = opt

                # iRow: int = 0
                # список без исключений
                # for resutlLst in resutlLsts:
                #    iRow += 1
                #    wb.Worksheets('Лист3').Cells(iRow, 6).Value = iRow  # - 2
                #    wb.Worksheets('Лист3').Cells(iRow, 7).Value = resutlLst
            except:
                print('Книга не создана')
                ExcelApp.app_close()

        except:
            print('Не книга не создана')
            return

    except:
        print('file_writer_win32 не сработал')

    finally:
        ExcelApp.app_close()
        finish = datetime.now()
        print('Время выполнения file_writer_win32: ' + str(finish - start_def))


def list_spliter(ws_list,number_repetitions):
    # очищаем результат от мусора (удаляем ненужные знаки и пробелы)
    print('Начало ws_list1_List_spliter')
    resutlLst: list = []
    for lpart in ws_list:

        if len(lpart) > 0:
            from lxml.doctestcompare import strip
            lpart = strip(lpart)
            Listadd: list = []
            for line in lpart.split(" "):  # цикл
                line = line.replace(",", "").replace(".", "").replace("—", "")
                line = line.replace("?", "").replace("!", "").replace(":", "")
                line = line.replace("(", "").replace(")", "").replace(";", "")
                line = line.lower()  # приводим все символы в нижний регистр lower
                if line != "":
                    Listadd.append(line)
            resutlLst.extend(Listadd)
        else:
            print("Пустой список Листа1")
    resutlLst.sort()
    # Если слово повторяется 3 и более раз то заносим в opts
    opts: list = []
    opts = [item for item in set(resutlLst) if resutlLst.count(item) >= number_repetitions]
    opts.sort()

    ex_dic: list = []
    # исключаем слова из списка исключений
    with open('exception_dictionary.txt', "r", encoding='utf-8') as file:
        for line in file:
            ex_dic.append(line[0:-1])
    file.close()

    opts_ex: list = []
    for item in set(opts).difference(ex_dic):
        opts_ex.append(item)
    opts_ex.sort()

    return opts_ex, resutlLst


# запрос браузера
headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/80.0.3987.106 Safari/537.36'}
# основной запрос
base_url: str = f'https://krasnodar.hh.ru/search/vacancy?clusters=true&area=53&enable_snippets=true&salary=&st' \
                f'=searchVacancy&text=%D0%A1%D0%BF%D0%B5%D1%86%D0%B8%D0%B0%D0%BB%D0%B8%D1%81%D1%82+%D0%BF%D0%BE+%D0' \
                f'%BE%D1%85%D1%80%D0%B0%D0%BD%D0%B5+%D1%82%D1%80%D1%83%D0%B4%D0%B0&from=suggest_post '
# колличество повторов ключевых слов
number_repetitions: int = 3
# файл с результатами

full_path = r'C:\Users\DeusEx\PycharmProjects\ParserHH\parsed_jobs.xlsx'

jobs = HH_parse(base_url, headers)  # парсим HH.ru по base_url
deep_cont = deep_pars(jobs, headers)  # заходим страницы HH.ru из выдачи jobs и парсим текст из них
file_writer_win32(jobs, number_repetitions, deep_cont, full_path)  # записываем результаты в parsed_jobs.xlsx
print('Готово')
