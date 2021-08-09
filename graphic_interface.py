from tkinter import *
import openpyxl

class Settings:
    def read_settings(self):
        self.settings_book = openpyxl.open('settings.xlsx', read_only=True)
        self.settings_sheet = self.settings_book.active
        self.row = self.get_text()
        self.start_row = self.row[0]
        self.end_row = self.row[1]
        self.all_tags = []
        for row in self.settings_sheet.iter_rows(min_row=self.start_row, max_row=self.end_row, min_col=2, max_col=4):
            for cell in row:
                read_cell = cell.value
                self.all_tags.append(read_cell)
        return print(self.all_tags)


class Window(Settings):
    def __init__(self):
        self.root = Tk()
        self.root.geometry('600x400+500+200')
        self.start_row = Entry(self.root)  # для вводу тексту
        self.end_row = Entry(self.root)

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        self.start_row.pack()
        self.end_row.pack()
        Button(self.root, text='Пуск', command=self.button_action).pack()

    def get_text(self):
        all_entry = [int(self.start_row.get()), int(self.end_row.get())]
        return all_entry

    def button_action(self):
        self.read_settings()


if __name__ == "__main__":
    window = Window()
    window.run()

