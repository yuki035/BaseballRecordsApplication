import tkinter
import teamrecord
import playerrecord
import threading
import time

class Application(tkinter.Frame):
    def __init__(self, root=None):
        super().__init__(root, width=280, height=160,
                         borderwidth=4, relief='groove')
        self.root = root
        self.pack()
        self.pack_propagate(0)
        self.create_widgets()

    def create_widgets(self):
        # 閉じるボタン
        quit_btn = tkinter.Button(self)
        quit_btn['text'] = '閉じる'
        quit_btn['command'] = self.root.destroy
        quit_btn.pack(side='bottom', pady=3)
        
        # チーム成績生成ボタン
        self.team_record_btn = tkinter.Button(self, text="チーム成績を生成",
                                              command=lambda:self.button_click(self.make_teamrecord_with_status))
        self.team_record_btn.place(x=40, y=30, width=100)
        
        # プログラムの状況をメッセージ出力
        self.team_record_status_label = tkinter.Label(self)
        self.team_record_status_label.place(x=150, y=32)
        self.team_record_status_label['text'] = '実行待ち'
        
        # 個人成績生成ボタン
        self.player_record_btn = tkinter.Button(self, text="個人成績を生成", 
                                                command=lambda:self.button_click(self.make_playerrecord_with_status))
        self.player_record_btn.place(x=40, y=75, width=100)
        
        # プログラムの状況をメッセージ出力
        self.player_record_status_label = tkinter.Label(self)
        self.player_record_status_label.place(x=150, y=77)
        self.player_record_status_label['text'] = '実行待ち' 
    
    def button_click(self, func):
        thread = threading.Thread(target=func)
        thread.start()  
        
    def make_teamrecord_with_status(self):
        self.team_record_btn['state'] = tkinter.DISABLED
        self.team_record_btn.update()
        self.team_record_status_label['text'] = '実行中'
        try:
            teamrecord.main()
        except:
            self.team_record_status_label['text'] = 'エラー'
        else:
            self.team_record_status_label['text'] = '完了'
        finally:
            self.team_record_btn['state'] = 'normal'
            time.sleep(3)
            self.team_record_status_label['text'] = '実行待ち'
            
    def make_playerrecord_with_status(self):
        self.player_record_btn['state'] = tkinter.DISABLED
        self.player_record_btn.update()
        self.player_record_status_label['text'] = '実行中'
        try:
            playerrecord.main()
        except:
            self.player_record_status_label['text'] = 'エラー'
        else:
            self.player_record_status_label['text'] = '完了'
        finally:
            self.player_record_btn['state'] = tkinter.NORMAL
            time.sleep(3)
            self.player_record_status_label['text'] = '実行待ち'

def main():
    root = tkinter.Tk()
    root.title('成績自動生成アプリ')
    root.geometry('350x170')
    app = Application(root=root)
    app.mainloop()

if __name__ == "__main__":
    main()