#!/usr/bin/env python
#-*- coding: utf8 -*-
import sys
import tkinter as tk
import tkinter.messagebox as tkm
import os, tkinter.filedialog
from tkinter import ttk
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
from openpyxl import Workbook
import numpy as np
from openpyxl import utils
import copy
import tournament_making
from tkinter import font
from tkinter import *
import Excel_to_pdf
import pdf_to_image
from pdf2image import convert_from_path
from PIL import Image
import slackweb
import date_make
import tournament_set
import tournament_start
import tournament_process
import threading
import shutil
import subprocess


class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master =master
        self.my_font = font.Font(self.master,size=30)
        self.my_small_font = font.Font(self.master,size=10)
        self.red="#FF0000"
        self.yellow="#FFFF00"
        self.blue="#0000FF"
        self.green="#008000"
        self.white="#FFFFFF"
        self.black="#000000"
        #self.pack()
        # ウインドウのタイトルを定義する
        master.title('LEDmatrix for Scoreboard')
        # ここでウインドウサイズを定義する
        master.geometry('550x400+0+0')
        master.configure(bg=self.white)
        self.create_widgets()

    def create_widgets(self):
        self.set_main_buttons()

    #メインウィンドウの配置
    def set_main_buttons(self):
        ttk.Label(self.master, text="トーナメントの作成や試合結果の集計を行います。",font=self.my_small_font).grid(row=0, column=0)
        tk.Button(self.master, text="トーナメント準備管理", command=self.sub_window_making,font=self.my_font).grid(row=1, column=0)
        ttk.Label(self.master, text="トーナメントの進行や各コートへの連絡を行います。",font=self.my_small_font).grid(row=2, column=0)
        tk.Button(self.master, text="トーナメントの進行", command=self.sub_window_process,font=self.my_font).grid(row=3, column=0)
        ttk.Label(self.master, text="LEDmatrixの操作を行えます。",font=self.my_small_font).grid(row=4, column=0)
        tk.Button(self.master, text="LED Matrixの操作", command=self.sub_window_ledmatrix,font=self.my_font).grid(row=5, column=0)
        ttk.Label(self.master, text="アプリを終了します。",font=self.my_small_font).grid(row=4, column=1)
        tk.Button(self.master, text="終了", command=self.master.destroy,font=self.my_font).grid(row=5, column=1)

    def sub_window_making(self):
        #サブウィンドウ生成
        self.sub_win_make = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win_make.geometry("700x500+550+0")
        self.sub_win_make.title("トーナメントの準備と管理")
        ttk.Label(self.sub_win_make, text="参加者のリストから全トーナメントと\n参加者のリストの作成をします。",font=self.my_small_font).grid(row=0, column=0)
        tk.Button(self.sub_win_make, text="全トーナメント作成", command=self.create_tournament,font=self.my_font).grid(row=1, column=0)
        ttk.Label(self.sub_win_make, text="全トーナメントと全参加者のリストから\n日付ごとのデータを作成をします。",font=self.my_small_font).grid(row=2, column=0)
        tk.Button(self.sub_win_make, text="日付毎トーナメント作成", command=self.tournament_select,font=self.my_font).grid(row=3, column=0)
        ttk.Label(self.sub_win_make, text="日付ごとのデータから\n当日に必要な管理表の作成をします。",font=self.my_small_font).grid(row=4, column=0)
        tk.Button(self.sub_win_make, text="管理表作成", command=self.default_make,font=self.my_font).grid(row=5, column=0)
        ttk.Label(self.sub_win_make, text="試合の結果から各選手の\nポイントを計算し、順位表を作成します。",font=self.my_small_font).grid(row=6, column=0)
        tk.Button(self.sub_win_make, text="結果計算", command=self.sub_window_making,font=self.my_font).grid(row=7, column=0)
        ttk.Label(self.sub_win_make, text="エクセルファイルを\nPDFに変換します。",font=self.my_small_font).grid(row=0, column=1)
        tk.Button(self.sub_win_make, text="PDF変換", command=self.excel_to_pdf,font=self.my_font).grid(row=1, column=1)
        ttk.Label(self.sub_win_make, text="アプリを終了します。",font=self.my_small_font).grid(row=6, column=1)
        tk.Button(self.sub_win_make, text="終了", command=self.master.destroy,font=self.my_font).grid(row=7, column=1)

    def sub_window_process(self):
        #サブウィンドウ生成
        self.sub_win_process = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win_process.geometry("750x500+500+0")
        self.sub_win_process.title("トーナメントの進行")
        ttk.Label(self.sub_win_process, text="トーナメントの進行を行います",font=self.my_small_font).grid(row=2, column=0)
        tk.Button(self.sub_win_process, text="トーナメントの進行", command=self.tournament_progress,font=self.my_font).grid(row=3, column=0)
        ttk.Label(self.sub_win_process, text="トーナメント表画像を出力します",font=self.my_small_font).grid(row=4, column=0)
        tk.Button(self.sub_win_process, text="トーナメント表画像取得", command=self. progress_sub_window_txt,font=self.my_font).grid(row=5, column=0)
        ttk.Label(self.sub_win_process, text="コート表の画像を出力します。",font=self.my_small_font).grid(row=6, column=0)
        tk.Button(self.sub_win_process, text="コート表画像取得", command=self.show_court,font=self.my_font).grid(row=7, column=0)
        ttk.Label(self.sub_win_process, text="指示を送ります。",font=self.my_small_font).grid(row=8, column=0)
        tk.Button(self.sub_win_process, text="指示出し", command=self.sub_window_making,font=self.my_font).grid(row=9, column=0)
        ttk.Label(self.sub_win_process, text="得点板の初期設定を行います。",font=self.my_small_font).grid(row=0, column=1)
        tk.Button(self.sub_win_process, text="初期設定", command=self.scoreboard_setting,font=self.my_font).grid(row=1, column=1)
        ttk.Label(self.sub_win_process, text="得点板機能を起動します",font=self.my_small_font).grid(row=0, column=0)
        tk.Button(self.sub_win_process, text="得点板機能起動", command=lambda: self.slack_income("-scb on"),font=self.my_font).grid(row=1, column=0)
        ttk.Label(self.sub_win_process, text="得点板機能を終了します",font=self.my_small_font).grid(row=2, column=1)
        tk.Button(self.sub_win_process, text="得点板機能終了", command=lambda: self.slack_income("-scb off"),font=self.my_font).grid(row=3, column=1)
        ttk.Label(self.sub_win_process, text="次回以降の試合を表示します",font=self.my_small_font).grid(row=4, column=1)
        tk.Button(self.sub_win_process, text="次回試合取得", command=self.show_next,font=self.my_font).grid(row=5, column=1)
        ttk.Label(self.sub_win_process, text="アプリを終了します。",font=self.my_small_font).grid(row=8, column=1)
        tk.Button(self.sub_win_process, text="終了", command=self.master.destroy,font=self.my_font).grid(row=9, column=1)


    def sub_window_ledmatrix(self):
        #サブウィンドウ生成
        self.sub_win_led = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win_led.geometry("700x500+550+0")
        self.sub_win_led.title("LED matrix操作")
        ttk.Label(self.sub_win_led, text="ニュース、カレンダー情報\n電車時刻表を表示します。",font=self.my_small_font).grid(row=6, column=0)
        tk.Button(self.sub_win_led, text="ニュース表示", command=lambda: self.slack_income("-news"),font=self.my_font).grid(row=7, column=0)
        ttk.Label(self.sub_win_led, text="天気予報、時計の表示をします。",font=self.my_small_font).grid(row=8, column=0)
        tk.Button(self.sub_win_led, text="時計表示", command=lambda: self.slack_income("-clock"),font=self.my_font).grid(row=9, column=0)
        ttk.Label(self.sub_win_led, text="画像を撮影します",font=self.my_small_font).grid(row=10, column=0)
        tk.Button(self.sub_win_led, text="画像撮影", command=lambda: self.slack_income("-photo"),font=self.my_font).grid(row=11, column=0)
        ttk.Label(self.sub_win_led, text="URLの画像を表示します",font=self.my_small_font).grid(row=3, column=0)
        self.url_box=tk.Entry(self.sub_win_led,width=30)
        self.url_box.grid(row=5, column=0)
        tk.Button(self.sub_win_led, text="URL送信", command=lambda: self.slack_income("-img"+self.url_box.get()),font=self.my_font).grid(row=4, column=0)
        ttk.Label(self.sub_win_led, text="24文字以下のテキストを表示します",font=self.my_small_font).grid(row=0, column=1)
        self.print_box=tk.Entry(self.sub_win_led,width=30)
        self.print_box.grid(row=2, column=1)
        tk.Button(self.sub_win_led, text="テキスト表示", command=lambda: self.slack_income("-print"+self.print_box.get()),font=self.my_font).grid(row=1, column=1)
        ttk.Label(self.sub_win_led, text="テキストをスクロールします",font=self.my_small_font).grid(row=3, column=1)
        self.scroll_box=tk.Entry(self.sub_win_led,width=30)
        self.scroll_box.grid(row=5, column=1)
        tk.Button(self.sub_win_led, text="テキストスクロール", command=lambda: self.slack_income("-scroll"+self.scroll_box.get()),font=self.my_font).grid(row=4, column=1)
        ttk.Label(self.sub_win_led, text="Slackでメッセージを送信します",font=self.my_small_font).grid(row=0, column=0)
        self.slack_box=tk.Entry(self.sub_win_led,width=30)
        self.slack_box.grid(row=2, column=0)
        tk.Button(self.sub_win_led, text="Slack送信", command=lambda: self.slack_income(self.slack_box.get()),font=self.my_font).grid(row=1, column=0)
        ttk.Label(self.sub_win_led, text="LEDmatrixを再起動します。",font=self.my_small_font).grid(row=6, column=1)
        tk.Button(self.sub_win_led, text="再起動", command=lambda: self.slack_income("-reboot"),font=self.my_font).grid(row=7, column=1)
        ttk.Label(self.sub_win_led, text="LEDmatrixをシャットダウンします。",font=self.my_small_font).grid(row=8, column=1)
        tk.Button(self.sub_win_led, text="シャットダウン", command=lambda: self.slack_income("-shutdown"),font=self.my_font).grid(row=9, column=1)
        ttk.Label(self.sub_win_led, text="アプリを終了します。",font=self.my_small_font).grid(row=10, column=1)
        tk.Button(self.sub_win_led, text="終了", command=self.master.destroy,font=self.my_font).grid(row=11, column=1)


    def slack_income(self,something):
        slack = slackweb.Slack(url="https://hooks.slack.com/services/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")
        slack.notify(text=something)
    
    #トーナメント作成
    def create_tournament(self):
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('トーナメントの作成', '参加者ファイルを選択してください')
        # 選択したファイルのパスを取得
        self.filepath = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath+"で正しいですか？")
        self.making_sub_window_radio()

    #シングルス/ダブルスの選択画面
    def making_sub_window_radio(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="トーナメントの作成をします。",font=self.my_font).place(x=100,y=100)
        # チェック有無変数
        self.var = tkinter.IntVar()
        # value=0のラジオボタンにチェックを入れる
        self.var.set(0)
        # ラジオボタン作成
        rdo1 = tkinter.Radiobutton(self.sub_win, value=0, variable=self.var, text='シングルス')
        rdo1.place(x=250, y=150)
        rdo2 = tkinter.Radiobutton(self.sub_win, value=1, variable=self.var, text='ダブルス')
        rdo2.place(x=250, y=180)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.making_decision_radio)
        button.place(x=250, y=210)

    #シングルスかダブルスかを決定
    def making_decision_radio(self):
        if self.var.get()==1:
            self.category='ダブルス'
        else:
            self.category='シングルス'
        self.sub_win.destroy()
        self.making_sub_window_txt()
        
    #メイン会場の入力
    def making_sub_window_txt(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="メイン会場を入力してください",font=self.my_font).place(x=100,y=100)
        # ラベル
        self.lbl = tkinter.Label(self.sub_win,text='メイン会場')
        self.lbl.place(x=100, y=150)
        # テキストボックス
        self.txt = tkinter.Entry(self.sub_win,width=20)
        self.txt.place(x=200, y=150)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.making_decision_text)
        button.place(x=250, y=180)

    #メイン会場の決定
    def making_decision_text(self):
        self.main_place=self.txt.get()
        self.sub_win.destroy()
        self.making_sub_window_lot_txt()

    #個別会場の入力
    def making_sub_window_lot_txt(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("1600x1400")
        ttk.Label(self.sub_win, text="個別会場を入力してください",font=self.my_font).place(x=400,y=30)
        self.tex=[]
        self.rank_date_list=tournament_making.make_rank_date_list(self.filepath)
        for i in range(len(self.rank_date_list)):
            # ラベル
            lbl = tkinter.Label(self.sub_win,text=self.rank_date_list[i])
            lbl.place(x=30+300*(i%4), y=70+30*(i/4))
            # テキストボックス
            self.tex.append(tkinter.Entry(self.sub_win,width=20))
            self.tex[i].place(x=100+300*(i%4), y=70+30*(i/4))

            # テキストボックスに文字をセット
            self.tex[i].insert(tkinter.END,self.main_place)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.making_decision_lot_text)
        button.place(x=600, y=600)

    #個別会場の決定
    def making_decision_lot_text(self):
        self.place=[]
        for i in range(len(self.rank_date_list)):
            self.place.append(self.tex[i].get())

        self.sub_win.destroy()
        self.making_sub_window_txt_2()

    #ファイル名の入力
    def making_sub_window_txt_2(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="ファイル名を決めてください",font=self.my_font).place(x=100,y=100)
        # ラベル
        self.lbl = tkinter.Label(self.sub_win,text='保存ファイル名(.xlsx)')
        self.lbl.place(x=100, y=150)
        # テキストボックス
        self.txt = tkinter.Entry(self.sub_win,width=20)
        self.txt.place(x=200, y=150)
        self.txt.insert(tkinter.END,".xlsx")
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.making_decision_text_2)
        button.place(x=250, y=180)
    #ファイル名の決定
    def making_decision_text_2(self):
        self.file=self.txt.get()
        self.sub_win.destroy()
        self.file.rstrip(".xlsx")
        print(self.file)
        tournament_making.main(self.filepath,self.category,self.place,self.file)

    #日付ごとのトーナメント作成
    def tournament_select(self):
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('トーナメント表の選択', 'トーナメント表を選択してください')
        # 選択したファイルのパスを取得
        self.filepath1 = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath1+"で正しいですか？")
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('メンバー表の選択', 'メンバー表を選択してください')
        # 選択したファイルのパスを取得
        self.filepath2 = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath2+"で正しいですか？")
        self.sub_window_radio2()

    def sub_window_radio2(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="日付を選択してください。",font=self.my_font).place(x=100,y=100)
        # チェック有無変数
        self.var = tkinter.IntVar()
        # value=0のラジオボタンにチェックを入れる
        self.var.set(0)
        self.rdo=[]
        self.date_cate_list=date_make.date_print(self.filepath1,self.filepath2)
        for i in range(len(self.date_cate_list)):
            # ラジオボタン作成
            self.rdo.append(tkinter.Radiobutton(self.sub_win, value=i, variable=self.var, text=self.date_cate_list[i]))
            self.rdo[i].place(x=250, y=150+30*i)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command =self.decision_tournament_select )
        button.place(x=250, y=150+30*len(self.date_cate_list))
    
    def decision_tournament_select(self):
        self.sub_win.destroy()
        date_make.date_making(self.filepath1,self.filepath2,self.date_cate_list[self.var.get()])

    def excel_to_pdf(self):
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('pdf化オプション', 'pdf化したいものを選んでください')
        # 選択したファイルのパスを取得
        self.filepath = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath+"で正しいですか？")
        Excel_to_pdf.main(self.filepath)

    #日付のトーナメント選択
    def default_make(self):
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('日付トーナメント表の選択', '日付トーナメント表を選択してください')
        # 選択したファイルのパスを取得
        self.filepath1 = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath1+"で正しいですか？")
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('日付メンバー表の選択', '日付メンバー表を選択してください')
        # 選択したファイルのパスを取得
        self.filepath2 = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath2+"で正しいですか？")
        self.select_sub_window_txt()

    #コート数の入力
    def select_sub_window_txt(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="コート数を入力してください",font=self.my_font).place(x=100,y=100)
        # ラベル
        self.lbl = tkinter.Label(self.sub_win,text='コート数')
        self.lbl.place(x=100, y=150)
        # テキストボックス
        self.court = tkinter.Entry(self.sub_win,width=20)
        self.court.place(x=200, y=150)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.select_decision_text)
        button.place(x=250, y=180)
    
    def select_decision_text(self):
        tournament_set.main(self.filepath1,self.filepath2,int(self.court.get()))
        self.sub_win.destroy()
    
    def scoreboard_setting(self):
        # 選択候補を拡張子jpgに絞る（絞らない場合は *.jpg → *）
        filetype = [("", "*.xlsx")]
        dirpath = os.path.abspath(os.path.dirname(__file__))
        tk.messagebox.showinfo('管理表の選択', '管理表を選択してください')
        # 選択したファイルのパスを取得
        self.filepath1 = tkinter.filedialog.askopenfilename(filetypes = filetype, initialdir = dirpath)
        tk.messagebox.askokcancel("ファイルの確認",self.filepath1+"で正しいですか？")
        self.start_sub_window_txt()
    

    def start_sub_window_txt(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="得点板に設定するコート番号を\n入力してください",font=self.my_font).place(x=50,y=50)
        # ラベル
        self.lbl = tkinter.Label(self.sub_win,text='コート番号')
        self.lbl.place(x=100, y=150)
        # テキストボックス
        self.courtnum = tkinter.Entry(self.sub_win,width=20)
        self.courtnum.place(x=200, y=150)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.start_decision_text)
        button.place(x=250, y=180)

    def start_decision_text(self):
        tournament_start.main(self.filepath1,int(self.courtnum.get()))
        self.sub_win.destroy()

    def tournament_progress(self):
        """
        if os.path.exists('C:\\Users\\mech-user\\Desktop\\IoT\\pdf'):
            shutil.rmtree("./pdf")
        os.mkdir("./pdf")
        if os.path.exists('C:\\Users\\mech-user\\Desktop\\IoT\\pi'):
            shutil.rmtree("./pi")
        os.mkdir("./pi")
        if os.path.exists('C:\\Users\\mech-user\\Desktop\\IoT\\試合番号データ'):
            shutil.rmtree("./試合番号データ")
        os.mkdir("./試合番号データ")
        """
        subprocess.Popen("python3 tournament_process.py",shell=True)
        print("ok")


    
    def progress_sub_window_txt(self):
        #サブウィンドウ生成
        self.sub_win = Toplevel()
        #サブウィンドウの画面サイズ
        self.sub_win.geometry("600x400")
        ttk.Label(self.sub_win, text="表示したいランク名を\n入力してください",font=self.my_font).place(x=50,y=50)
        # ラベル
        self.lbl = tkinter.Label(self.sub_win,text='ランク名')
        self.lbl.place(x=100, y=150)
        # テキストボックス
        self.courtnum = tkinter.Entry(self.sub_win,width=20)
        self.courtnum.place(x=200, y=150)
        #Button生成
        button = ttk.Button(self.sub_win,text = '決定',width = str('決定'),command = self.show_tournament)
        button.place(x=250, y=180)

    def show_tournament(self):
        if os.path.exists("C:\\Users\\mech-user\\Desktop\\IoT\\photo\\tournament"+self.courtnum.get()+".png"):
            img = Image.open("C:\\Users\\mech-user\\Desktop\\IoT\\photo\\tournament"+self.courtnum.get()+".png")
            img.show()
        self.sub_win.destroy()

    def show_court(self):
        # 元となる画像の読み込み
        if os.path.exists('C:\\Users\\mech-user\\Desktop\\IoT\\photo\\court.png'):
            img = Image.open('C:\\Users\\mech-user\\Desktop\\IoT\\photo\\court.png')
            img.show()


    def show_next(self):
        # 元となる画像の読み込み
        if os.path.exists('C:\\Users\\mech-user\\Desktop\\IoT\\photo\\next.png'):
            img = Image.open('C:\\Users\\mech-user\\Desktop\\IoT\\photo\\next.png')
            img.show()



    

root = tkinter.Tk()

app = Application(master=root)

app.mainloop()