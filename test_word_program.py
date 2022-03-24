from tkinter import *

class Test_word:
    def __init__(self, master):
        self.master = master
        master.title("test word program")
        
root = Tk()
my_gui = Test_word(root)
root.mainloop()