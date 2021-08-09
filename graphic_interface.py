from tkinter import *
from test_class import *

class Window(Test):
    def __init__(self):
        self.root = Tk()
        self.root.title('Перевірка цін конкурентів')
        self.root.geometry('600x400+500+200')
        self.root.resizable(False, False)

    def run(self):
        self.draw_widgets()
        self.root.mainloop()

    def draw_widgets(self):
        Button(self.root, text='Пуск', command=self.button_action).pack()  # кнопка може приймати метод або функцію

    def button_action(self):
        print(test.read_settings())



if __name__ == "__main__":
    test = Test()
    window = Window()
    window.run()
