import openpyxl
import requests
from bs4 import BeautifulSoup

read_book = openpyxl.open('price_url.xlsx', read_only=True)
sheet_read = read_book.active

new_book = openpyxl.Workbook()
new_sheet = new_book.active

clean_url = []

# читає колонки і стовпці з xlsx
for row in sheet_read.iter_rows(min_row=5, max_row=15, min_col=2, max_col=15):
    for cell in row:
        url = cell.value
        # видаляє зайві символи з комірки xlsx щоб отримати посилання
        url = url[12:-9]
        clean_url.append(url)


def gerbor_kiev(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'price-number s-product-price'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('грн.', '')

        # переводимо назад в формулу для ексель
        formula = '=HYPERLINK("' + url + '";"' + price + '")'

        new_sheet.cell(row=5, column=next_column).value = formula
        next_column += 1


gerbor_kiev(clean_url[0:14])


def brwland(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('div', {'class': 'price product__price'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('грн.', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'
        new_sheet.cell(row=6, column=next_column).value = formula
        next_column += 1


brwland(clean_url[14:28])


def vashamebel(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'itemprop': 'price'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('грн', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'
        new_sheet.cell(row=7, column=next_column).value = formula
        next_column += 1


vashamebel(clean_url[28:42])


def mebel_mebel(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('div', {'itemprop': 'price'})
        price = price.text
        formula = '=HYPERLINK("' + url + '";"' + price + '")'
        new_sheet.cell(row=8, column=next_column).value = formula
        next_column += 1


mebel_mebel(clean_url[42:56])


def abcmebli(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'price'})
        price = price.text
        price = price.replace('\xa0', '')
        price = price.replace('грн', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'
        new_sheet.cell(row=9, column=next_column).value = formula
        next_column += 1


abcmebli(clean_url[56:70])


def mebelok(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'price-new'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('грн.', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'
        new_sheet.cell(row=10, column=next_column).value = formula
        next_column += 1


mebelok(clean_url[70:82])


def maxmebel(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('b', {'class': 'int'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('\xa0', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'

        new_sheet.cell(row=11, column=next_column).value = formula
        next_column += 1


maxmebel(clean_url[84:98])


def moyamebel(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'fn-price'})
        price = price.text
        price = price.replace(' ', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'

        new_sheet.cell(row=12, column=next_column).value = formula
        next_column += 1


moyamebel(clean_url[98:112])


def brw_kiev(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'price nowrap'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('грн.', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'

        new_sheet.cell(row=13, column=next_column).value = formula
        next_column += 1


brw_kiev(clean_url[112:126])


def shurup(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'price-new'})
        price = price.text
        # price = price.replace(' ', '')
        price = price.replace('грн', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'

        new_sheet.cell(row=14, column=next_column).value = formula
        next_column += 1


shurup(clean_url[126:140])


def mebel_online(clean_url):
    next_column = 2
    for url in clean_url:
        print(url)
        source = requests.get(url)
        main_text = source.text
        soup = BeautifulSoup(main_text, features="html.parser")
        price = soup.find('span', {'class': 'pr-price'})
        price = price.text
        price = price.replace(' ', '')
        price = price.replace('грн.', '')
        formula = '=HYPERLINK("' + url + '";"' + price + '")'

        new_sheet.cell(row=15, column=next_column).value = formula
        next_column += 1


mebel_online(clean_url[140:154])

new_book.save('new_price_url.xlsx')
new_book.close()
