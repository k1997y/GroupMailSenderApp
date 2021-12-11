import tkinter as tk
from tkinter.scrolledtext import ScrolledText
import win32com.client


# GUIを提供するクラス
class App:
    TITLE = "一斉送信アプリ"
    WINDOW_SIZE = "720x720"
    OFFSET = 30

    def __init__(self):
        # Mailerのオブジェクト作成
        mailer = Mailer()

        self.root = tk.Tk()  # ウィンドウ作成
        self.root.geometry(self.WINDOW_SIZE)

        # タイトル設定
        self.root.title(self.TITLE)

        # 宛先のテキストボックスのリスト
        self.address_textbox_list = []

        # 宛先
        label_to = tk.Label(self.root, text="宛先", font=("normal", 14, "bold"))
        label_to.place(x=20, y=20)
        self.button_add_to = tk.Button(self.root, text="追加", command=self.add_address_textbox)
        self.button_add_to.place(x=20, y=60)

        # 件名
        label_title = tk.Label(self.root, text="件名", font=("normal", 14, "bold"))
        label_title.place(x=20, y=self.OFFSET + 150)
        self.textbox_title = tk.Entry(width=70)
        self.textbox_title.place(x=80, y=self.OFFSET + 155)

        # 本文
        label_body = tk.Label(self.root, text="本文", font=("normal", 14, "bold"))
        label_body.place(x=20, y=self.OFFSET + 200)
        self.textbox_body = ScrolledText(self.root, font=("normal", 10), height=15, width=40)
        self.textbox_body.place(x=20, y=self.OFFSET + 250)

        # 送信前確認するかどうかのチェックボックス
        self.checkbutton_value = tk.BooleanVar()
        self.checkbutton_value.set(True)
        self.checkbutton_prechecked = tk.Checkbutton(self.root,
                                                     text="送信前に確認する",
                                                     variable=self.checkbutton_value,
                                                     onvalue=True,
                                                     offvalue=False)
        self.checkbutton_prechecked.place(x=100, y=500)

        # 送信ボタン
        self.button_send = tk.Button(self.root, text="送信", height=5, width=10,
                                     command=lambda: mailer.send_group_mail(self))
        self.button_send.place(x=200, y=550)

    # アプリ起動
    def mainloop(self):
        self.root.mainloop()

    # ボタンを押すと宛先を入力できるテキストボックスを増やす
    def add_address_textbox(self):
        # テキストボックス作成
        textbox = tk.Entry(width=40)
        textbox.place(x=80, y=self.OFFSET)

        # テキストボックスのリストに追加
        self.address_textbox_list.append(textbox)

        self.OFFSET += 30


# メールに関する機能をまとめたクラス
class Mailer:
    # メールを一斉送信する
    def send_group_mail(self, app):
        # 宛先を取得しリストに格納する
        address_list = []
        for address in app.address_textbox_list:
            address_list.append(address.get())

        # 件名を取得する
        subject = app.textbox_title.get()

        # 本文を取得する
        mail_body = app.textbox_body.get("1.0", "end-1c")

        # 送信前に確認するか否か
        is_prechecked = app.checkbutton_value.get()

        # 一斉送信
        for address in address_list:
            self.send_mail(address, subject, mail_body, is_prechecked)

    # メールを送信する
    def send_mail(self, address, subject, body_string, is_prechecked):
        # メールオブジェクトの作成
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.to = address
        mail.subject = subject
        mail.bodyFormat = 1
        mail.body = body_string

        if is_prechecked:
            mail.display(is_prechecked)
        else:
            mail.Send()

