import requests
from bs4 import BeautifulSoup
import json

def get_polis_number(fam, im, ot, dayr):
    USER_AGENT = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) snap Chromium/76.0.3809.100 Chrome/76.0.3809.100 Safari/537.36'
    SecondName = fam
    FirstName = im
    PatronymicName = ot
    BirthDate = dayr
    # web_form_submit	 = "Отправить"
    # allowed_domains = ['fomsrd.ru']
    start_urls = 'http://fomsrd.ru/service/polis/between.php?'
    end_urls = f'SecondName={SecondName}&FirstName={FirstName}&PatronymicName={PatronymicName}&BirthDate={BirthDate}'
    url = start_urls + end_urls

    response = requests.get(url, headers={'User-Agent': USER_AGENT})
    print(response.status_code)
    soup = BeautifulSoup(response.text, 'lxml')
    print(soup.text)
    insuranceStatuses = soup.text

    if len(insuranceStatuses) == 0:
        insuranceStatuses = '{"insuranceStatuses":{"enp":"000","status":"нет ответа","smo":"нет ответа",\
        "region":"нет ответа", "spolis":"","npolis":"0","startDate":"","endDate":""},"result":0,"description":null}'

    json_string = insuranceStatuses  # будем парсить эту json строку
    parsed_string = json.loads(json_string)  # распарсенная строка

    return parsed_string


def get_begin_row_maks_excel(df):
    col_names = df.columns
    i = 0
    while df[col_names[1]][i] != "ID случая":
        i += 1
    return i - 1