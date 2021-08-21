from tkinter import *
from tkinter.ttk import Progressbar
from main import ScrapPrices


class Window(ScrapPrices):
    def __init__(self):
        super().__init__()
        self.canvas = Tk()
        self.canvas.title('Перевірка цін з конкурентами')
        self.canvas.geometry('600x400+500+200')
        self.canvas.resizable(True, True)
        self.label_url = Label(self.canvas)
        Label(self.canvas, text='Info: ').place(x=30, y=30)
        self.pb = Progressbar(self.canvas, length=534)
        Label(self.canvas, text='Початок рядка: ').place(x=30, y=160)
        default_start_row = StringVar(self.canvas, value='4')
        self.start_row = Entry(self.canvas, width=8, textvariable=default_start_row)
        Label(self.canvas, text='Початок колонки: ').place(x=197, y=100)
        self.start_col = Entry(self.canvas, width=16).place(x=200, y=120)
        Label(self.canvas, text='Кінець колонки: ').place(x=310, y=100)
        self.start_col = Entry(self.canvas, width=16).place(x=313, y=120)
        Label(self.canvas, text='Кінець рядка: ').place(x=30, y=200)
        self.start_end = Entry(self.canvas, width=8).place(x=130, y=200)

    def run(self):
        self.draw_widgets()
        self.canvas.mainloop()

    def draw_widgets(self):
        Button(self.canvas, width=10, height=1, text='Пуск', command=self.button_action).place(x=487, y=100)
        self.label_url.place(x=60, y=30)
        self.pb.place(x=33, y=50)
        self.start_row.place(x=130, y=160)

    def button_action(self):
        self.get_tags(4, 5)
        self.get_result(4, 5, 2, 7) # порахувати колонки і стовпці і їх помножити
        self.save()
        print(self.start_row.get())



if __name__ == "__main__":
    window = Window()
    window.run()
