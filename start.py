import ShareWork
import iSelenium
import Grab_Web
import process_file
import os
import datetime
import warnings
import tkinter as tk

from dateutil.relativedelta import relativedelta
from tkinter import messagebox, ttk
from tkcalendar import Calendar, DateEntry
from datetime import datetime, timedelta

# 用户账号密码文档数组
users_info = {
    "dayday.manager": {
        "password": "Aa113322*",
        "folder": "dayday"
    },
    "k88foodcourt.manager": {
        "password": "K88k88**",
        "folder": "k88"
    },
    "boobowpor.restau.manager.78": {
        "password": "Uco888**",
        "folder": "tarchong"
    },
    "user123": {
        "password": "abc123",
        "folder": "user123"
    }
}

ParentPath = os.getcwd()

# 登录逻辑
def login():
    user = entry_id.get()
    pwd = entry_pwd.get()

    if user in users_info and pwd == users_info[user]["password"]:
        folder = users_info[user]["folder"]
        userInfo = [user, pwd, folder]
        messagebox.showinfo("登录成功", f"欢迎 {user}！")
        root.withdraw()  # 隐藏登录窗口
        show_date_window(ParentPath, userInfo)  # 打开日期选择窗口

    else:
        messagebox.showerror("登录失败", "用户名或密码错误")

# 日期选择界面
def show_date_window(path, userInfo):
    date_win = tk.Toplevel()
    date_win.title("选择日期范围")
    date_win.geometry("300x250")

    # 起始日期
    tk.Label(date_win, text="开始日期：", font=("Arial", 12)).pack(pady=5)
    start_cal = DateEntry(date_win, width=15, background='darkblue',
                          foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
    start_cal.pack(pady=5)

    # 结束日期
    tk.Label(date_win, text="结束日期：", font=("Arial", 12)).pack(pady=5)
    end_cal = DateEntry(date_win, width=15, background='darkblue',
                        foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')
    end_cal.pack(pady=5)

    def getDate():
        start_date = start_cal.get_date()
        end_date = end_cal.get_date()
        # 生成日期数组
        date_list = []
        date_list_yymmdd = []
        current_date = start_date
        while current_date <= end_date:
            date_list.append(current_date.strftime("%a, %d %b %Y"))
            date_list_yymmdd.append(current_date.strftime("%Y-%m-%d"))
            current_date += timedelta(days=1)
        return date_list, date_list_yymmdd

    # messagebox.showinfo("选择的日期", f"起始日期: {start_str}\n结束日期: {end_str}")

    # 下载按钮
    def confirm_download():
        date_list, date_list_yymmdd = getDate()
        Grab_Web.start(path, userInfo, date_list, date_list_yymmdd)

    # 处理按键
    def confirm_process():
        date_list, date_list_yymmdd = getDate()
        process_file.start(path, userInfo, date_list_yymmdd)

    # 上传按键
    def confirm_upload():
        date_list, date_list_yymmdd = getDate()
        print(date_list_yymmdd)
    #     Grab_Web.start(path, userInfo, date_list)

    # 一键处理按键
    def confirm_oneflow():
        date_list, date_list_yymmdd = getDate()
        Grab_Web.start(path, userInfo, date_list, date_list_yymmdd)
        process_file.start(path, userInfo, date_list_yymmdd)

    # 按键上部分
    frame_top = tk.Frame(date_win)
    frame_top.pack(pady=10)
    tk.Button(frame_top, text="Grab Sales下载", command=confirm_download).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_top, text="Grab Sales处理", command=confirm_process).pack(side=tk.LEFT, padx=5)
    # tk.Button(frame_top, text="Grab Sales上传", command=confirm_upload).pack(side=tk.LEFT, padx=5)

    # 按键下部分
    frame_bottom = tk.Frame(date_win)
    frame_bottom.pack(pady=10)
    tk.Button(frame_bottom, text="一键处理", command=confirm_oneflow).pack()

# ---------------- 主界面 ----------------
root = tk.Tk()
root.title("Grab Sales 登录界面")
root.geometry("300x220")

# 用户名
tk.Label(root, text="用户名：", font=("Arial", 12)).pack(pady=5)
entry_id = tk.Entry(root, font=("Arial", 12))
entry_id.pack(pady=5)

# 密码
tk.Label(root, text="密码：", font=("Arial", 12)).pack(pady=5)
entry_pwd = tk.Entry(root, show="*", font=("Arial", 12))
entry_pwd.pack(pady=5)

# 勾选显示密码
def toggle_password():
    if show_pwd_var.get():
        entry_pwd.config(show="")
    else:
        entry_pwd.config(show="*")

show_pwd_var = tk.BooleanVar()
chk_show_pwd = tk.Checkbutton(root, text="显示密码", variable=show_pwd_var, command=toggle_password)
chk_show_pwd.pack()

# 登录按钮
tk.Button(root, text="登录", command= login, font=("Arial", 12)).pack(pady=10)

root.mainloop()
