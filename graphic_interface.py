from tkinter import *
import openpyxl
import requests
from bs4 import BeautifulSoup
from tkinter.ttk import Progressbar


class Settings: # створити ініт
    def read_settings(self):
        self.settings_book = openpyxl.open('settings.xlsx', read_only=True)
        self.settings_sheet = self.settings_book.active
        self.row = self.get_text()  # get_text()
        self.start_row = self.row[0]
        self.end_row = self.row[1]
        self.all_tags = []
        for row in self.settings_sheet.iter_rows(min_row=self.start_row, max_row=self.end_row, min_col=2, max_col=4):
            for cell in row:
                read_cell = cell.value
                self.all_tags.append(read_cell)
        return print(self.all_tags)

    def read_url(self):
        tag1 = 0
        tag2 = 1
        tag3 = 2
        self.tags = self.all_tags
        self.read_book = openpyxl.open('shop_url.xlsx', read_only=True)
        self.sheet_read = self.read_book.active
        self.new_book = openpyxl.Workbook()
        self.new_sheet = self.new_book.active

        self.s_col = 2  # ------------

        headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
        for row in self.sheet_read.iter_rows(min_row=self.start_row, max_row=self.end_row, min_col=self.s_col, max_col=16):
            for cell in row:
                read_cell = cell.value
                if read_cell == 'x':
                    self.new_sheet.cell(row=self.start_row, column=self.s_col).value = 'x'  # винести за if
                else:
                    url = read_cell
                    print(url)
                    try:
                        source = requests.get(url, headers=headers)
                        main_text = source.content.decode()
                        soup = BeautifulSoup(main_text, features="html.parser")
                        price = soup.find(self.all_tags[tag1], {self.all_tags[tag2]: self.all_tags[tag3]})
                        price = price.text
                        #price = get_clean_price(price)
                        formula = f'=HYPERLINK("{url}";"{price}")'
                        self.new_sheet.cell(row=self.start_row, column=self.s_col).value = formula

                    except BaseException:
                        print('Error: ' + url)
                        url_not_work = f'=HYPERLINK("{url}";"!404!")'
                        self.new_sheet.cell(row=self.start_row, column=self.s_col).value = url_not_work

                self.s_col += 1
                print(self.s_col)
                if self.s_col == 17:  # пофіксити
                    print('yes')
                    self.start_row += 1
                    self.s_col = 2
                    tag1 += 3
                    tag2 += 3
                    tag3 += 3

    def save_new(self):
        self.new_book.save('qwerty.xlsx')
        self.new_book.close()


class Window(Settings):
    def __init__(self):
        self.root = Tk()
        self.root.title('Перевірка цін з конкурентами')
        self.root.geometry('600x400+500+200')
        self.start_row = Entry(self.root)  # для вводу тексту
        self.end_row = Entry(self.root)

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        Label(self.root, text='Start row').pack(anchor=NW)
        self.start_row.pack(anchor=NW)
        Label(self.root, text='End row').pack(anchor=NW)
        self.end_row.pack(anchor=NW)
        Progressbar(self.root, length=300).pack()
        Button(self.root, text='Пуск', command=self.button_action).pack()
        Button(self.root, text='Зберегти', command=self.get_save).pack()

    def get_save(self):
        self.save_new()

    def get_text(self):
        all_entry = [int(self.start_row.get()), int(self.end_row.get())]
        return all_entry

    def button_action(self):
        self.read_settings()
        self.read_url()


if __name__ == "__main__":
    window = Window()
    window.run()
