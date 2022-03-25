from sre_constants import GROUPREF_EXISTS
from tkinter import *
from openpyxl import *
from random import *
from tkinter import filedialog
import tkinter.messagebox as msgbox

class Test_word:
    def __init__(self, main):
        self.groupe = []
        self.length_word = 0
        self.main = main
        main.title("test word program")

        self.list_file = Listbox(main, selectmode="extended", height=0, width=30)
        self.list_file.grid(row=0, column=0, columnspan=2)
        self.list_file.insert(END, "테스트할 xlsx파일을 추가하세요")

        self.mean_search = Entry(main, width=30)
        self.mean_search.grid(row=3, column=0, columnspan=2)
        self.mean_search.insert(END, "뜻을 입력하세요")    

        self.word_search = Entry(main, width=30)
        self.word_search.grid(row=3, column=2, columnspan=2)
        self.word_search.insert(END, "단어를 입력하세요")

        btn_add_file = Button(main, text="파일추가", command=self.add_file)
        btn_del_file = Button(main, text="파일삭제", command=self.del_file)
        btn_trans = Button(main, text="테스트할 단어파일 변환", command=self.trans_file)

        btn_mean_search = Button(main, text="뜻으로 단어 찾기", command=self.cmd_mean_search)
        btn_word_search = Button(main, text="단어로 뜻 찾기", command=self.cmd_word_search)

        btn_test_start_word = Button(main, text="불어 단어로 테스트 시작", command=self.test_start_word)
        btn_test_start_mean = Button(main, text="한글 뜻으로 테스트 시작", command=self.test_start_mean)

        btn_add_file.grid(row=0, column=2, sticky=W+E)
        btn_del_file.grid(row=0, column=3, sticky=W+E)
        btn_trans.grid(row=1, column=0, columnspan=4, sticky=W+E)

        btn_mean_search.grid(row=4, column=0, columnspan=2, sticky=W+E)
        btn_word_search.grid(row=4, column=2, columnspan=2, sticky=W+E)

        btn_test_start_word.grid(row=2, column=2, columnspan=2, sticky=W+E)
        btn_test_start_mean.grid(row=2, column=0, columnspan=2, sticky=W+E)
    
    def add_file(self):
        files = filedialog.askopenfilenames(initialdir="/", \
            title="파일을 선택하세요",\
            filetypes=(("Excel 통합 문서", "*.xlsx"), ("모든 파일", "*.*"))) #무조건 튜플로 해야함
        for file in files:
            self.list_file.insert(END, file)
        print(files)
    def del_file(self):
        selected_file = self.list_file.curselection()
        self.list_file.delete(selected_file)
    def trans_file(self):
        self.groupe.clear()
        self.length_word = 0
        selected_file= []
        index_selected_file = self.list_file.curselection()
        for i in index_selected_file:
            t = self.list_file.get(i)
            selected_file.append(t)
            #테스트
        for file in selected_file:
            file = file.replace("/", "\\")
            wb = load_workbook(file)
            ws = wb.active
            for word_row in ws.iter_rows(min_row=2, min_col=2, max_col=4, values_only=True):
                self.groupe.append(word_row)
            self.length_word = self.length_word+len(list(ws.rows))
        print(self.groupe)
        print(self.length_word)
        selected_file.clear()
        return self.groupe, self.length_word
    def cmd_mean_search(self):
        search = self.mean_search.get()
        num_row_mean = 0
        while num_row_mean<self.length_word-1:
            if search == self.groupe[num_row_mean][2]:
                print(self.groupe[num_row_mean][:2])
                break
            num_row_mean += 1
            if num_row_mean == self.length_word-1:
                print("찾는 단어가 없습니다")
                break
    def cmd_word_search(self):
        search = self.word_search.get()
        num_row_word = 0
        while num_row_word<self.length_word-1:
            if search == self.groupe[num_row_word][1]:
                print(self.groupe[num_row_word][2])
                break
            num_row_word += 1
            if num_row_word == self.length_word-1:
                print("찾는 뜻이 없습니다")
                break
    def test_start_word(self):
        response = msgbox.askokcancel("확인 메세지", "불어 단어를 보고 뜻을 찾는 테스트를 하시겠습니까?")
        if response == 1 :
            new_window = Toplevel(root)
            new_window.title("불어 단어로 테스트")
            new_window.geometry()
            shuffle(self.groupe)
    def test_start_mean(self):
        response = msgbox.askokcancel("확인 메세지", "뜻을 보고 단어를 찾는 테스트를 하시겠습니까?")
        if response == 1:
            shuffle(self.groupe)
            # Test()
# class Test(Test_word):
#     def __init__(self):
#         self.new_window = Toplevel(root)
#         self.new_window.title("뜻으로 테스트")
#         self.new_window.geometry("480x300")
#         self.btn_new_window_start = Button(self.new_window, text="테스트 시작", command=self.start_test)
#         self.btn_new_window_start.pack()
#         self.new_window.mainloop()
#     def start_test(self):
#         super.
#         self.x = 0
#         self.btn_new_window_start.destroy()
#         self.display_word = StringVar()
#         self.display_word.set("{0}".format(self.groupe[self.x][2]))
#         self.word = Label(self.new_window, relief="groove", height=5, width=20, padx=10, pady=10, textvariable=self.display_word).pack()
#         self.answer = Entry(self.new_window, width=30).pack()
#         self.new_window.bind("<Return>", self.next_word)
#     def next_word(self, event):
#         self.x += 1
#         self.display_word.set("{0}".format(self.groupe[self.x][2]))

root = Tk()
my_gui = Test_word(root)
root.mainloop()