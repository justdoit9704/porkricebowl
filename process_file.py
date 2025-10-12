import pdfplumber
import os
import re
import warnings

import ShareWork
import ShareWork as shareWork
import pandas as pd
from collections import Counter
from contextlib import redirect_stderr
from datetime import datetime, timedelta

warnings.simplefilter(action="ignore", category=pd.errors.SettingWithCopyWarning)

def start(curPath, userInfo, date_list_yymmdd):
    main_json = os.path.join(curPath, 'Main_Config.json')
    JSON = shareWork.getJsonData(main_json)
    min_date = (datetime.strptime(min(date_list_yymmdd), "%Y-%m-%d")).strftime("%Y%m%d")
    max_date = (datetime.strptime(max(date_list_yymmdd), "%Y-%m-%d")).strftime("%Y%m%d")
    user_path = os.path.join(os.path.join(curPath, 'resources'), userInfo[2])

    dfs = [] # 空的df

    print("Start Processing PDF File......")

    for date in date_list_yymmdd:
        date = (datetime.strptime(date, "%Y-%m-%d")).strftime("%Y%m%d")
        print(date + "\nProcessing Reports......")

        all_text = ""  # 用于存放整个 PDF 的文字

        for i in range(3):
            i+=1
            if os.path.exists(JSON["Path"]["PDF"].format(user_path) + "\\{}_{}.pdf".format(date,i)):
                # 读取PDF文字
                with open(os.devnull, "w") as f, redirect_stderr(f): # 临时屏蔽 stderr
                    with pdfplumber.open(JSON["Path"]["PDF"].format(user_path) + "\\{}_{}.pdf".format(date,i)) as pdf:
                        for page in pdf.pages:
                            page_text = page.extract_text()
                            if page_text:  # 避免 None
                                all_text += page_text + "\n"  # 每页加换行

                # 正则：(\d+)x\s* 捕获数量
                # ([A-R]|S[1-6]|Omelette Egg|Rice) 捕获标记，不带 ')'
                pattern = r'(\d+)x\s*(Omelette Egg|Rice|S[1-6]|[A-R])'

                # 找到所有匹配
                matches = re.findall(pattern, all_text)

                # 根据数量拆分
                result = []
                for count, letter in matches:
                    result.extend([letter] * int(count))

                counted = Counter(result)
                print(counted)
                # print(page_text)

                # 构建字段
                df_dict = {'Date': date}

                # Normal (A,C,E,J,L)
                df_dict['Normal'] = sum(counted.get(c, 0) for c in ['A', 'C', 'E', 'J', 'L'])
                # Egg (B,D,F,K,M)
                df_dict['Egg'] = sum(counted.get(c, 0) for c in ['B', 'D', 'F', 'K', 'M'])

                # Q, R
                df_dict['Q'] = counted.get('Q', 0)
                df_dict['R'] = counted.get('R', 0)

                # Omelette Egg, Rice
                df_dict['Omelette Egg'] = counted.get('Omelette Egg', 0)
                df_dict['Rice'] = counted.get('Rice', 0)

                # S = sum(S1-S6)
                df_dict['S'] = sum(counted.get(f'S{i}', 0) for i in range(1, 7))

                # A-R（排除 Q,R）
                for c in 'ABCDEFGHIJKLMNOP':
                    if c not in ['Q', 'R']:
                        df_dict[c] = counted.get(c, 0)

                # S1-S6
                for i in range(1, 7):
                    key = f'S{i}'
                    df_dict[key] = counted.get(key, 0)

                # 转成 DataFrame
                df = pd.DataFrame([df_dict])
                dfs.append(df)
            else:
                continue
    print("Completed Reports......")

    # csv file 处理
    excel_path = JSON["Path"]["Excel"].format(user_path) + "\\{}_{}.csv".format(min_date, max_date)
    print("Excel path is: {}".format(excel_path))
    df = pd.read_csv(excel_path)  # 使用绝对路径最好
    # 转换成 datetime 类型
    df["date"] = pd.to_datetime(df["Updated On"], errors="coerce")

    # 格式化成 yyyy-mm-dd 字符串
    df["date"] = df["date"].dt.strftime("%Y%m%d")
    df.rename(columns={'Amount':'Sales', 'Discount (Merchant-Funded)':'Discount', 'Marketing success fee':'Marketing fees', 'Total':'Net Amount'}, inplace=True)

    # 'Advertisement'
    df_adver = df[df[('Category')] == 'Advertisement']
    df_adver["Marketing fees"] = df_adver["Net Amount"]
    df_adver["Sales"] = 0

    # 'Adjustment'
    # df_adjust = df[(df[('Category')] == 'Adjustment') | (df[('Category')] == 'Adjustment ')]
    df_adjust = df[df['Category'].str.strip() == 'Adjustment']
    if df_adjust.empty:
        pass
    else:
        df_adjust["Marketing fees"] = df_adjust["Net Amount"]
        df_adjust["Sales"] = 0

    # print(df_adjust)

    # 把原本的Sales 负数变成0
    df['Sales'] = df['Sales'].clip(lower=0)


    # 把负数变正数
    df['Discount'] = df['Discount'] * -1
    df['Marketing fees'] = df['Marketing fees'] * -1
    df_adver["Marketing fees"] = df_adver['Marketing fees'] * -1

    df = df[df["Category"].isin(["Payment", "Payment "])]

    if df_adjust.empty:
        df_f = pd.concat([df, df_adver], ignore_index=True)
    else:
        df_adjust["Marketing fees"] = df_adjust['Marketing fees'] * -1
        df_f = pd.concat([df, df_adver, df_adjust], ignore_index=True)

    # 筛选 'Transferred'
    df_f = df_f[df_f[('Status')].isin(['Transferred','Completed'])][['date','Sales','Discount','Net Sales', 'Marketing fees', 'Order commission', 'Net Amount']]
    # group by date 统计 'Sales', 'Discount', 'Marketing fees', 'Net Amount'
    df_f = df_f.groupby(['date'])[['Sales', 'Discount', 'Marketing fees', 'Net Amount']].sum().reset_index()

    # Deduction = (sales - discount) * deduction%
    df_f['Deduction'] = ((df_f['Sales'] - df_f['Discount']) * JSON["Deduction"]).round(2)
    # SST = deduction * SST%
    df_f['SST'] = (df_f['Deduction'] * JSON["SST"]).round(2)

    # 字段排序
    df_f = df_f[['date','Sales','Discount','Deduction','SST','Marketing fees','Net Amount']]

    # print(df_f[df_f[('date')] == '2025-09-18'])
    # print(df_f)

    if not dfs:
        print("No PDF file......")
    else:
        # order集合
        df_all = pd.concat(dfs, ignore_index=True, sort=False)
        df_all = df_all.groupby(["Date"])[
            ['Normal', 'Egg', 'Q', 'R', 'Omelette Egg', 'Rice', 'S', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
             'K', 'L', 'M', 'N', 'O', 'P', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6']].sum().reset_index()

        # order + sales 集合
        df_merged = pd.merge(df_f, df_all, how="left", left_on="date",  right_on="Date")

        # df_merged = df_merged.drop(columns=["Date"])

        df_merged = df_merged[['date', 'Normal', 'Egg', 'Q', 'R', 'Omelette Egg', 'Rice', 'S', 'Sales', 'Discount', 'Deduction', 'SST', 'Marketing fees', 'Net Amount', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'S1', 'S2', 'S3', 'S4', 'S5', 'S6']]

        # 构建文件名
        ShareWork.createFolder(JSON["Path"]["Reports"].format(user_path))
        file_name = JSON["Path"]["Reports"].format(user_path) + "\\Grab Reports {}-{}.xlsx".format(min_date,max_date)

        # 保存 DataFrame
        df_merged.to_excel(file_name, index=False)
        print("Combined All Reports......")
        print("File Location: {}".format(file_name))


# curPath = os.getcwd()
# # main_json = os.path.join(curPath, 'Main_Config.json')
# # JSON = shareWork.getJsonData(main_json)
# start(curPath, "", ["2025-08-13","2025-08-19"])