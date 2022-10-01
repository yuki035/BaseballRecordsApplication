import tkinter

class Application(tkinter.Frame):
    def __init__(self, root=None):
        super().__init__(root, width=380, height=280,
                         borderwidth=1, relief='groove')
        self.root = root
        self.pack()
        self.pack_propagate(0)
        self.create_widgets()

    def create_widgets(self):
        # 閉じるボタン
        quit_btn = tkinter.Button(self)
        quit_btn['text'] = '閉じる'
        quit_btn['command'] = self.root.destroy
        quit_btn.pack(side='bottom')
        
        # チーム成績生成ボタン
        team_record_btn = tkinter.Button(self, text="チーム成績を生成",
                                         command=self.test)
        team_record_btn.place(x= 140, y = 100)
        
        # 個人成績生成ボタン
        player_record_btn = tkinter.Button(self, text="個人成績を生成",
                                           command=self.test)
        player_record_btn.place(x = 140, y = 150)
        
    def test(self):
        print('ボタンが押された')
        

root = tkinter.Tk()
root.title('成績自動生成アプリ')
root.geometry('400x300')
app = Application(root=root)
app.mainloop()