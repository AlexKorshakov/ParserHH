import requests
import csv
from datetime import time, datetime
from bs4 import BeautifulSoup as bs

headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.106 Safari/537.36'}
base_url = 'https://krasnodar.hh.ru/search/vacancy?clusters=true&area=53&items_on_page=50&no_magic=true&enable_snippets=true&salary=&st=searchVacancy&text=Python'


def HH_parse(base_url, headers):
    jobs = []
    urls = []
    urls.append(base_url)
    session = requests.Session()
    request = session.get(base_url, headers=headers)
    if request.status_code == 200:
        start = datetime.now()
        soup = bs(request.content, 'lxml')
        try:
            pagination = soup.find_all('a', attrs={'data-qa': 'pager-page'})
            count = int(pagination[-1].text)
            for i in range(count):
                #url = f'base_url+'&page='+{i}
                url = f'https://krasnodar.hh.ru/search/vacancy?clusters=true&area=53&items_on_page=50&no_magic=true&enable_snippets=true&salary=&st=searchVacancy&text=Python&page={i}'
                if url not in urls:
                    urls.append(url)
        except:
            pass
    for url in urls:
        request = session.get(url, headers=headers)
        soup = bs(request.content, 'lxml')

        divs = soup.find_all('div', attrs={'data-qa': 'vacancy-serp__vacancy'})
        for div in divs:
            try:
                title = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-title'}).text + ' ; '
                href = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-title'})['href'] + ' ; '
                company = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-employer'}).text + ' ; '
                text1 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'}).text + ' ; '
                try:
                    text2 = div.find('div',attrs={'data-qa': 'vacancy-vacancy-serp__vacancy_snippet_requirement'}).text + ' ; '
                    content = str(text1) + ' ' + str(text2)
                except:
                    content = str(text1)
                jobs.append({
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
        a_pen.writerow((" Название вакансии; ", " URL; ", " Название Компании; ", " Описание; "))
        for job in jobs:
            a_pen.writerow((job['title'], job['href'], job['company'], job['content']))


jobs = HH_parse(base_url, headers)
file_writer(jobs)
