import openpyxl
import requests
from bs4 import BeautifulSoup

# Відкриває існуючий xlsx файл з url
read_book = openpyxl.open('shop_url.xlsx', read_only=True)
sheet_read = read_book.active
# Відкриває файл з html тегами
read_book2 = openpyxl.open('html_tags.xlsx', read_only=True)
sheet_read2 = read_book2.active
# Створює новий xlsx файл в який буде записоно нові дані
new_book = openpyxl.Workbook()
new_sheet = new_book.active


# Читає всі комірки з тегами і записує
def read_html_tags():
    all_tags = []
    for row in sheet_read2.iter_rows(min_row=5, max_row=15, min_col=2, max_col=4):
        for cell in row:
            read_cell = cell.value
            all_tags.append(read_cell)
    return all_tags


all_tags = read_html_tags()


def calkueted():
    next_column = 2
    next_row = 5
    tag1 = 0
    tag2 = 1
    tag3 = 2
    for row in sheet_read.iter_rows(min_row=5, max_row=15, min_col=2, max_col=15):
        for cell in row:
            read_cell = cell.value
            if read_cell == 'x':
                new_sheet.cell(row=next_row, column=next_column).value = 'NO'
            elif read_cell != 'x':
                url = read_cell
                print(url)
                try:
                    source = requests.get(url)
                    main_text = source.text
                    soup = BeautifulSoup(main_text, features="html.parser")
                    price = soup.find(all_tags[tag1], {all_tags[tag2]: all_tags[tag3]})
                    price = price.text
                    # Забирає з ціни зайві символи
                    price = price.replace(' ', '')
                    price = price.replace('грн', '')
                    price = price.replace('грн.', '')
                    price = price.replace('\xa0', '')
                    price = price.replace('.', '')
                    formula = '=HYPERLINK("' + url + '";"' + price + '")'
                    new_sheet.cell(row=next_row, column=next_column).value = formula

                except AttributeError:
                    print('Помилка')
                    url_not_work = '=HYPERLINK("' + url + '";"' + '!404!' + '")'
                    new_sheet.cell(row=next_row, column=next_column).value = url_not_work


            next_column += 1

            if next_column == 16:
                next_row += 1
                next_column = 2
                tag1 += 3
                tag2 += 3
                tag3 += 3

calkueted()
new_book.save('new_price_url.xlsx')
new_book.close()
