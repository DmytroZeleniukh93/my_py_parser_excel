import openpyxl
import requests
import time
from bs4 import BeautifulSoup

# Відкриває існуючий xlsx файл з url
read_book = openpyxl.open('shop_url.xlsx', read_only=True)
sheet_read = read_book.active
# Відкриває файл з html тегами
read_book2 = openpyxl.open('settings.xlsx', read_only=True)
sheet_read2 = read_book2.active
# Створює новий xlsx файл в який буде записоно нові дані
new_book = openpyxl.Workbook()
new_sheet = new_book.active

start_row = int(input('Початковий рядок: '))
end_row = int(input('Кінцейвий рядок: '))
start_col = int(input('Початкова колонка: '))
end_col = int(input('Кінцева колонка: '))

# Читає всі комірки з тегами і записує
def read_html_tags(start_row, end_row):
    all_tags = []
    for row in sheet_read2.iter_rows(min_row=start_row, max_row=end_row, min_col=2, max_col=4):
        for cell in row:
            read_cell = cell.value
            all_tags.append(read_cell)
    return all_tags


all_tags = read_html_tags(start_row, end_row)


def get_clean_price(price):
    price = price.replace(' ', '')
    price = price.replace('грн', '')
    price = price.replace('грн.', '')
    price = price.replace('\xa0', '')
    price = price.replace('.', '')
    price = price.replace('₴', '')
    price = price.replace('\n', '')
    price = price.replace('\t', '')
    price = price.replace('.', '')
    price = price.replace(':', '')
    price = price.replace('Цена', '')
    return price


def result(start_row, end_row, start_col, end_col):
    #active_column = 2
    tag1 = 0
    tag2 = 1
    tag3 = 2
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
    for row in sheet_read.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
        for cell in row:
            read_cell = cell.value
            if read_cell == 'x':
                new_sheet.cell(row=start_row, column=start_col).value = 'x'  # винести за if
            else:
                url = read_cell
                print(url)
                try:
                    source = requests.get(url, headers=headers)
                    main_text = source.content.decode()
                    soup = BeautifulSoup(main_text, features="html.parser")
                    price = soup.find(all_tags[tag1], {all_tags[tag2]: all_tags[tag3]})
                    price = price.text
                    price = get_clean_price(price)
                    formula = f'=HYPERLINK("{url}";"{price}")'
                    new_sheet.cell(row=start_row, column=start_col).value = formula

                except BaseException:
                    print('Error: ' + url)
                    url_not_work = f'=HYPERLINK("{url}";"!404!")'
                    new_sheet.cell(row=start_row, column=start_col).value = url_not_work

            start_col += 1

            if start_col == 17:  # пофіксити
                start_row += 1
                start_col = 2
                tag1 += 3
                tag2 += 3
                tag3 += 3


result(start_row, end_row, start_col, end_col)

name_file = time.asctime()
name_file = name_file[3:10]

new_book.save(f'price_url {name_file}.xlsx')
new_book.close()
