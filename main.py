import openpyxl
import requests
import time
from bs4 import BeautifulSoup


class ScrapPrices:
    replace_list = [' ', 'грн', '.', 'грн.', '\xa0', '₴', '\n', '\t', 'Цена', ':']

    def __init__(self):
        self.read_settings = openpyxl.open('settings.xlsx', read_only=True)
        self.sheet_settings = self.read_settings.active
        self.all_tags = []
        self.read_shop_url = openpyxl.open('shop_url.xlsx', read_only=True)
        self.shop_url = self.read_shop_url.active
        self.new_book = openpyxl.Workbook()
        self.new_sheet = self.new_book.active

    def get_tags(self, start_row, end_row):
        for row in self.sheet_settings.iter_rows(min_row=start_row, max_row=end_row, min_col=2, max_col=4):
            for cell in row:
                read_cell = cell.value
                self.all_tags.append(read_cell)
                print(read_cell)  # добавити фічу забрати пробіл в кінці

    def get_result(self, start_row, end_row, start_col, end_col):
        tag1 = 0
        tag2 = 1
        tag3 = 2
        headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
        for row in self.shop_url.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
            for cell in row:
                read_cell = cell.value
                if read_cell == 'x':
                    self.new_sheet.cell(row=start_row, column=start_col).value = 'x'
                else:
                    url = read_cell
                    print(url)
                    try:
                        source = requests.get(url, headers=headers)
                        main_text = source.content.decode()
                        soup = BeautifulSoup(main_text, features="html.parser")
                        price = soup.find(self.all_tags[tag1], {self.all_tags[tag2]: self.all_tags[tag3]})
                        price = self.__get_clean_price(price.text)
                        to_write = f'=HYPERLINK("{url}";"{price}")'
                        self.new_sheet.cell(row=start_row, column=start_col).value = to_write
                    except Exception:
                        print('Error: ' + url)
                        to_write = f'=HYPERLINK("{url}";"!404!")'
                        self.new_sheet.cell(row=start_row, column=start_col).value = to_write
                start_col += 1
                if start_col == end_col + 1:
                    start_row += 1
                    start_col = 2
                    tag1 += 3
                    tag2 += 3
                    tag3 += 3

    def __get_clean_price(self, price):
        for to_replace in self.replace_list:
            price = price.replace(to_replace, '')
        return price

    def __close_files(self):
        self.read_settings.close()
        self.read_shop_url.close()
        self.new_book.close()

    def save(self):
        self.new_book.save('class_price_url.xlsx')
        self.new_book.close()
        self.__close_files()


if __name__ == "__main__":
    scrap_price = ScrapPrices()
    scrap_price.get_tags(5, 5)
    scrap_price.get_result(5, 5, 2, 4)
    scrap_price.save()

'''
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
'''
