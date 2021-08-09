import openpyxl
from graphic_interface import *

class Test():
    def __init__(self):
        self.settings_book = openpyxl.open('settings.xlsx', read_only=True)
        self.settings_sheet = self.settings_book.active

    def read_settings(self):
        self.start_row = x.get_text()
        #self.start_row = int(input('Початковий рядок: '))
        self.end_row = int(input('Кінцейвий рядок: '))

        self.all_tags = []
        for row in self.settings_sheet.iter_rows(min_row=self.start_row, max_row=self.end_row, min_col=2, max_col=4):
            for cell in row:
                read_cell = cell.value
                self.all_tags.append(read_cell)
        return self.all_tags


x = Window()
