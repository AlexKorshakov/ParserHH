import csv
from datetime import datetime
from typing import Union
import requests
import re
import win32com.client as com_client
from bs4 import BeautifulSoup as bs

headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.106 Safari/537.36'}
base_url = f'https://krasnodar.hh.ru/search/vacancy?area=53&clusters=true&employment=full&enable_snippets=true&items_on_page=100&label=not_from_agency&no_magic=true&schedule=fullDay&text=Python&only_with_salary=true&salary=115000&from=cluster_compensation&showClusters=true'

def HH_parse(base_url, headers):
    jobs = []
    urls = []
    urls.append(base_url)
    session = requests.Session()
    request = session.get(base_url, headers=headers)
    if request.status_code == 200:
        start: datetime = datetime.now()
        soup = bs(request.content, 'lxml')
        try:
            pagination = soup.find_all('a', attrs={'data-qa': 'pager-page'})
            count = int(pagination[-1].text)
            for i in range(count):
                url = base_url + '&page={i}'
                #url = f'https://krasnodar.hh.ru/search/vacancy?clusters=true&area=53&items_on_page=50&no_magic=true&enable_snippets=true&salary=&st=searchVacancy&text=Python&page={i}'
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
                    text2 = div.find('div',attrs={'data-qa': 'vacancy-vacancy-serp__vacancy_snippet_requirement'}).text
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
        finish = datetime.now()
        print('Всего:' + str(len(jobs)) + ' ' + 'Время выполнения lxml: ' + str(finish - start))
    else:
        print('Error or Done ' + str(request.status_code))
    return jobs


def file_writer(jobs):
    with open('parsed_jobs.csv', 'w') as file:
        a_pen = csv.writer(file)
        a_pen.writerow((" Номер вакансии ", " Название вакансии ", " URL ", " Название Компании ",
                        " Описание ", " Подробное описание "))
        for job in jobs:
            a_pen.writerow((job['rowNom'], job['title'], job['href'],
                            job['company'], job['content']))


def deep_pars(jobs, headers):
    deepcont=[]
    try:
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
    return deepcont


def clear_string(string):
    cleanr = {
        '</p>', '<p>', '</li>', '<li>', '</ul>', '<ul>', '</div>', '<div>', '</strong>', '<strong>',
        '</em>', '<em>', '</br>', '<br>', '<br/>', '  '
    }
    for cl in cleanr:
        string = re.sub(cl, '', string)
    #print(string)
    return string


def file_writer_win32(jobs, deepcont):
    #print(jobs)
    try:
        excel = com_client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        path = r'C:\Users\DeusEx\PycharmProjects\ParserHH\parsed_jobs.xlsx'
        wb = excel.Workbooks.Open(path)
        try:
            ws_list1 = wb.Worksheets('Лист1')
            ws_list2 = wb.Worksheets('Лист2')
            #ws_list1.Range('A1:F200').Select()
            #ws_list1.Selection.Delete()
            #ws = wb.Worksheets.Add()
            iRow = 2
            for job in jobs:
                iRow: Union[int, iRow] = iRow + 1
                ws_list1.Cells(iRow, 1).Value = iRow - 2 #job['rowNom']
                ws_list1.Cells(iRow, 2).Value = job['title']
                ws_list1.Cells(iRow, 3).Value = job['href']
                ws_list1.Cells(iRow, 4).Value = job['company']
                ws_list1.Cells(iRow, 5).Value = job['content']
            iRow = 2
            for dcont in deepcont:
                iRow: Union[int, iRow] = iRow + 1
                ws_list2.Cells(iRow, 1).Value = iRow - 2
                ws_list2.Cells(iRow, 2).Value = dcont['deepcontent']
        finally:
            wb.SaveAs(path)
            excel.DisplayAlerts = True
            excel.ScreenUpdating = True
            #wb.Close(True)
        print('Всего записей: ' + str(iRow - 2))
    finally:
         excel.Quit()


jobs = HH_parse(base_url, headers)
#file_writer(jobs)
deepcont = deep_pars(jobs, headers)
file_writer_win32(jobs, deepcont)
