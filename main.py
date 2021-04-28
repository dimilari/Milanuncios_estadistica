from requests import get
import openpyxl
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
from bs4 import BeautifulSoup
from datetime import date


town = ['malaga', 'torremolinos']  # ciudades para observar
prises = {600: 699, 700: 799, 800: 899, 900: 1000}

HOST = 'https://www.milanuncios.com/'
URL = 'https://www.milanuncios.com/alquiler-de-pisos-en-malaga-malaga/?fromSearch\
       =1&desde=700&hasta=799&dormd=3&dormh=3&banosd=2&banosh=2'
HEADERS = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:75.0) Gecko/20100101 Firefox/75.0'
}
lista = []

def get_html(url, params=None):
    r = get(url, headers=HEADERS, params=params)
    return r


def get_content(text_of_page):
    soup = BeautifulSoup(text_of_page, 'html.parser')
    items = str(soup.find('div', style="margin-bottom:4px; text-align: center; clear:both"))
    count = ''
    for i in items[231:237]:
        if i.isdigit():
            count += i
    if count == '':
        count = 0
    return int(count)

if get_html(URL).status_code == 200:
    number = 0
    n = 10
    for i in town:
        for j, jj in prises.items():
            html = get_html(f'https://www.milanuncios.com/alquiler-de-pisos-en-{i}-malaga/?fromSearch\
                =1&desde={j}&hasta={jj}&dormd=3&dormh=3&banosd=2&banosh=2')
            lista.append(get_content(html.text))
            number += 1
            n -= 1
            numb = {1: 'st', 2: 'nd', 3: 'rd'}
            if number < 4:
                suff = numb[number]
            else:
                suff = 'th'
            print(f'Parsing of {number}-{suff} page'+' '*n+' ...')
    for i in town:
        html = get_html(f'https://www.milanuncios.com/alquiler-de-pisos-en-{i}-malaga/?fromSearch\
            =1&desde=600&demanda=n&dormd=4&dormh=4&banosd=2')
        lista.append(get_content(html.text))
        number += 1
        n -= 1
        print(f'Parsing of {number}-th page'+' '*n+' ...')
else:
    print('Error')
print(lista)


book = openpyxl.load_workbook(filename='Estadistica.xlsm', read_only=False, keep_vba=True)
sheet = book.active
next_row = sheet.max_row + 1
medium_border = Border(left=Side(style='medium'),
                       right=Side(style='medium'),
                       top=Side(style='medium'),
                       bottom=Side(style='medium'))

if sheet[sheet.max_row][0].value != date.today().strftime('%d.%m.%Y'):

    lista.insert(0, date.today().strftime('%d.%m.%Y'))
    sheet.append(lista)

    sheet.row_dimensions[next_row].height = 23.6
    sheet[next_row][0].fill = PatternFill(fill_type='solid', start_color='FFB6FCC5')
    sheet[next_row][9].fill = PatternFill(fill_type='solid', start_color='FFF2F2F2')
    sheet[next_row][10].fill = PatternFill(fill_type='solid', start_color='FFF2F2F2')
    sheet[next_row][0].font = Font(size=10)
    for i in range(1, 12):
        sheet.cell(row=next_row, column=i).border = medium_border
        sheet.cell(row=next_row, column=i).alignment = Alignment(horizontal='center', vertical='center')
    for i in range(1, 11):
        sheet[next_row][i].font = Font(size=18)

# сохраняем файл
book.save('Estadistica.xlsm')

# открываем файл для визуального просмотра
import os
os.startfile('Estadistica.xlsm')
