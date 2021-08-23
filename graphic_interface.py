from tkinter import *
from tkinter.ttk import Progressbar
from main import ScrapPrices
from tkinter import messagebox


class Window(ScrapPrices):
    def __init__(self):
        super().__init__()
        self.root = Tk()
        self.root.title('Перевірка цін з конкурентами')
        self.root.geometry('600x370+500+200')
        self.root.resizable(False, False)
        self.canvas = Canvas(self.root)
        self.my_image = PhotoImage(file='excel_image.png')
        self.canvas.create_image(210, 190, image=self.my_image)
        self.label_url = Label(self.root)
        Label(self.root, text='Info: ').place(x=30, y=30)
        self.pb = Progressbar(self.root, length=534)
        Label(self.root, text='Почати з ряду: ').place(x=30, y=197)
        default_start_row = StringVar(self.root, value='4')
        self.start_row = Entry(self.root, width=8, textvariable=default_start_row)
        Label(self.root, text='Закінчити рядом: ').place(x=30, y=277)
        default_end_row = StringVar(self.root, value='26')
        self.end_row = Entry(self.root, width=7, textvariable=default_end_row)
        Label(self.root, text='Почати з колонки: ').place(x=234, y=120)
        default_start_col = StringVar(self.root, value='2')
        self.start_col = Entry(self.root, width=13, textvariable=default_start_col)
        Label(self.root, text='Закінчити колонкою: ').place(x=404, y=120)
        default_end_col = StringVar(self.root, value='16')
        self.end_col = Entry(self.root, width=13, textvariable=default_end_col)

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        self.canvas.place(x=135, y=50)
        self.label_url.place(x=60, y=30)
        self.pb.place(x=33, y=50)
        Button(self.root, width=10, height=1, text='Пуск', command=self.button_action).place(x=487, y=90)
        self.start_row.place(x=134, y=197)
        self.end_row.place(x=134, y=277)
        self.start_col.place(x=237, y=140)
        self.end_col.place(x=407, y=140)

    def button_action(self):
        self.get_nums()
        self.test()
        self.get_tags(self.s_row, self.e_row)
        self.get_result(self.s_row, self.e_row, self.s_col, self.e_col)
        self.save()
        messagebox.showinfo('Info', 'Сканування завершено')
        self.root.destroy()

    def get_nums(self):
        self.s_row = int(self.start_row.get())
        self.e_row = int(self.end_row.get())
        self.s_col = int(self.start_col.get())
        self.e_col = int(self.end_col.get())

    def test(self):
        x = (self.e_row - self.s_row) + 1
        y = (self.e_col - self.s_col) + 1
        self.res = 100 / (x * y)


if __name__ == "__main__":
    window = Window()
    window.run()
