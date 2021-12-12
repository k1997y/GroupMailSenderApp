import threading
import tkinter as tk
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
import win32com.client
import pandas as pd
import pyperclip


# GUIを提供するクラス
class App:
    TITLE = "一斉送信アプリ"
    WINDOW_SIZE = "720x720"
    OFFSET = 60

    def __init__(self):
        # Mailerのオブジェクト作成
        mailer = Mailer()

        # ファイルマネージャーの初期化
        filemanager=FileManager()
        filemanager.initialize()

        self.root = tk.Tk()  # ウィンドウ作成
        self.root.geometry(self.WINDOW_SIZE)

        # タイトル設定
        self.root.title(self.TITLE)

        # 宛先のテキストボックスのリスト
        self.address_textbox_list = []
        # 宛先名のテキストボックスのリスト
        self.address_name_textbox_list = []

        # 宛先メールアドレス
        label_to = tk.Label(self.root, text="宛先", font=("normal", 14, "bold"))
        label_to.place(x=20, y=20)
        self.button_add_to = tk.Button(self.root, text="追加", command=self.add_address_textbox)
        self.button_add_to.place(x=20, y=60)

        # 宛先名
        label_to_name = tk.Label(self.root, text="宛先名", font=("normal", 14, "bold"))
        label_to_name.place(x=350, y=20)

        # 件名(メールタイトル)
        label_title = tk.Label(self.root, text="件名", font=("normal", 14, "bold"))
        label_title.place(x=20, y=self.OFFSET + 120)
        self.textbox_title = tk.Entry(width=45)
        self.textbox_title.place(x=80, y=self.OFFSET + 125)

        # コース
        label_course = tk.Label(self.root,text="コース名",font=("normal",14,"bold"))
        label_course.place(x=20,y=self.OFFSET+160)
        self.textbox_course=tk.Entry(width = 70)
        self.textbox_course.place(x=120,y=self.OFFSET+165)

        # 本文
        label_body = tk.Label(self.root, text="本文", font=("normal", 14, "bold"))
        label_body.place(x=20, y=self.OFFSET + 230)
        self.textbox_body = ScrolledText(self.root, font=("normal", 10), height=15, width=70)
        # self.textbox_body.insert("1.0", "{名前} 様")
        self.textbox_body.place(x=20, y=self.OFFSET + 280)
        # 本文にコピペできるようにする
        # self.textbox_body.bind("<Control-v>",paste_string)
        # 本文ファイルを取得するためのボタン設置
        self.button_get_message_file = tk.Button(self.root,
                                                 text="インポート",
                                                 width=10,
                                                 command=self.import_mail_body)
        self.button_get_message_file.place(x=100, y=self.OFFSET + 230)

        # 送信前確認するかどうかのチェックボックス
        self.checkbutton_value = tk.BooleanVar()
        self.checkbutton_value.set(True)
        self.checkbutton_prechecked = tk.Checkbutton(self.root,
                                                     text="送信前に確認する",
                                                     variable=self.checkbutton_value,
                                                     onvalue=True,
                                                     offvalue=False)
        self.checkbutton_prechecked.place(x=100, y=560)

        # 送信ボタン
        self.button_send = tk.Button(self.root, text="送信", height=5, width=10,
                                     command=lambda: mailer.send_group_mail(self))
        self.button_send.place(x=200, y=600)

        # ペースト機能実装
        self.root.bind("<Control-v>",self.paste_string)

    # アプリ起動
    def mainloop(self):
        self.root.mainloop()

    # ボタンを押すと宛先を入力できるテキストボックスを増やす
    # 同時に宛先名を入力するボックスも追加
    def add_address_textbox(self):
        # 宛先テキストボックス作成
        textbox = tk.Entry(width=30)
        textbox.place(x=80, y=self.OFFSET)
        # 宛先名テキストボックス作成
        textbox_name = tk.Entry(width=30)
        textbox_name.place(x=350, y=self.OFFSET)

        # 宛先テキストボックスのリストに追加
        self.address_textbox_list.append(textbox)
        # 宛先名テキストボックスのリストに追加
        self.address_name_textbox_list.append(textbox_name)

        self.OFFSET += 30

    # 本文をインポートする
    def import_mail_body(self):
        filepath = import_file()

        with open(filepath, encoding="utf-8") as f:
            str = f.readlines()

        # 本文のテキストボックスに挿入
        i = 1
        for string in str:
            tmp = "{}.0".format(i)
            self.textbox_body.insert(tmp, string)
            i += 1

    # クリップボードからペーストする
    def paste_string(self):
        # フォーカスされている要素を取得
        element = self.root.focus_get()

        # それがテキストボックスならば、そこにペーストする
        if isinstance(element,tk.Entry):
            str = pyperclip.paste()
            element.insert(tk.END,str)


# メールに関する機能をまとめたクラス
class Mailer:
    # メールを一斉送信する
    def send_group_mail(self, app):
        # 宛先を取得しリストに格納する
        # address_list = []
        # for address in app.address_textbox_list:
        #     address_list.append(address.get())

        # 宛先名を取得しリストに格納する
        # address_name_list = []
        # for address_name in app.address_name_textbox_list:
        #     address_name_list.append(address_name.get())

        # 件名を取得する
        subject = app.textbox_title.get()

        # 本文を取得する
        mail_body = app.textbox_body.get("1.0", "end-1c")

        # 送信前に確認するか否か
        is_prechecked = app.checkbutton_value.get()

        # 一斉送信
        # i = 0
        # for address in address_list:
        #     self.send_mail(address, address_name_list[i], subject, mail_body, is_prechecked)

        filemanager = FileManager()
        filemanager.initialize()

        # 本文を送信先ごとに書き換える
        for i in range(filemanager.column_number):
            message = ""
            # 1行目に社名を追加
            message += (filemanager.company_list[i] + "\n")


            # 2行目に担当者1を追加
            # 担当者が存在しない場合はスルーする
            if not pd.isna(filemanager.person1_list[i]):
                message += filemanager.person1_list[i]

            # 担当者が2人いれば追加する
            if not pd.isna(filemanager.person2_list[i]):
                message += ", " + filemanager.person2_list[i]
            # 「様」を付ける
            message += " 様\n\n"

            mail_body = message + mail_body

            self.send_mail(filemanager.address_list[i], subject, mail_body, is_prechecked)

    # メールを送信する
    def send_mail(self,
                  address,
                  subject,
                  body_string,
                  is_prechecked):
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


# 本文をファイルブラウザで読み込む機能
def import_file():
    return filedialog.askopenfilename()


# ファイル操作をまとめたクラス(シングルトン)
class FileManager:
    # ファイル名
    ADDRESS_LIST_PATH = "送信先リスト.xlsx"
    # アドレスの列名
    ADDRESS_COL_NAME = "メールアドレス"
    # 社名の列名
    COMPANY_COL_NAME = "社名"
    # 担当者1の列名
    PERSON_1 = "担当者1"
    # 担当者2の列名
    PERSON_2 = "担当者2"

    # シングルトンオブジェクト
    __singleton = None
    # 初期化済みか否かのフラグ
    __is_initialized = False

    # 行数
    column_number = 0

    # アドレスのリスト
    address_list = []
    # 社名のリスト
    company_list = []
    # 担当者1のリスト
    person1_list = []
    # 担当者2のリスト
    person2_list = []

    # シングルトンを作成して返す
    def __new__(cls, *args, **kwargs):
        if cls.__singleton == None:
            cls.__singleton = super(FileManager,cls).__new__(cls)
        return cls.__singleton

    def initialize(self):
        # 初期化済みなら何もしない
        if self.__is_initialized:
            return

        # 送信先リストのexcelファイルからデータフレームを作成
        df = pd.read_excel(self.ADDRESS_LIST_PATH,sheet_name=0,header=0)
        # 行数の取得
        self.column_number = len(df)

        # アドレス、社名、担当者1、担当者2のリストを作成
        for address in df[self.ADDRESS_COL_NAME]:
            self.address_list.append(address)
        for company in df[self.COMPANY_COL_NAME]:
            self.company_list.append(company)
        for person1 in df[self.PERSON_1]:
            self.person1_list.append(person1)
        for person2 in df[self.PERSON_2]:
            self.person2_list.append(person2)

        # 初期化フラグをオンにする
        self.__is_initialized = True

    # def __init__(self):
    #     df = pd.read_excel(self.ADDRESS_LIST_PATH, sheet_name=0, header=0)
    #
    #     # 行数取得
    #     self.column_number = len(df)
    #
    #     # アドレス、社名、担当者1, 担当者2のリストを作成
    #     self.address_list = []
    #     self.company_name_list = []
    #     self.person1_list = []
    #     self.person2_list = []
    #
    #     # アドレスのリストに値を追加
    #     for address in df[self.ADDRESS_COL_NAME]:
    #         self.address_list.append(address)
    #     for company in df[self.COMPANY_COL_NAME]:
    #         self.company_name_list.append(company)
    #     for person1 in df[self.PERSON_1]:
    #         self.person1_list.append(person1)
    #     for person2 in df[self.PERSON_2]:
    #         self.person2_list.append(person2)
