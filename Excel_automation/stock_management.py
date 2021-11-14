import pandas as pd
from glob import glob
from collections import defaultdict

import os
import shutil

import json    
import smtplib
from email.mime.text import MIMEText
from email.utils import formatdate
from datetime import date

# 注文表の集計
def update_order(order, filepath):
    # 注文表読み込み
    df_order = pd.read_excel(filepath)
    for key, value in df_order.to_dict().items():
        order[key] += value[0]
    # ファイルの移動
    shutil.move(filepath, 'sources/order_old/')
    return order

def main():
    print('ーーーーー開始ーーーーー')
    # 注文情報の収集
    print('注文情報の収集')
    order = defaultdict(int)
    filepaths_order = glob('sources/order_new/order*.xlsx')
    for filepath in filepaths_order:
        order = update_order(order, filepath)
    print('Done')

    # 最新の在庫量の確認
    print('最新の在庫量の確認')
    filepath_stock = 'sources/stock.xlsx'
    df_stock = pd.read_excel(filepath_stock)

    stock = df_stock.iloc[-1:, 2:]
    stock = stock.reset_index(drop=True)

    order = pd.DataFrame(order, index=[0], columns=stock.columns)
    order = order.fillna(0)

    updated = stock - order
    print('Done')


    # 追加注文品の算出
    print('追加注文品の算出')
    filepath_master = 'sources/master.xlsx'
    df_master = pd.read_excel(filepath_master)

    threshold = df_master.iloc[:1 , 1:]
    shortage_columns = updated[updated < threshold].dropna(axis=1).columns

    df_shortage = df_master.iloc[1:, 1:][shortage_columns]

    order_text = ''
    for key, value in df_shortage.to_dict().items():
        order_text += f'{key}を{value[1]}本、'
    print('Done')


    # Gmailで注文
    print('Gmailで注文')
    with open('sources/secret.json') as f:
        address_password = json.load(f)

    from_addr = address_password['ADDRESS']
    password = address_password['PASSWORD']

    subject = '発注依頼'
    body = f'佐藤さん、在庫数が足りなくなってしまったため、{order_text}発注してください。'
    to_addr = 'gigpayown@macr2.com'

    # SMTPサーバに接続
    smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
    smtpobj.starttls()
    smtpobj.login(from_addr, password)

    # メッセージ（メール）の作成
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = from_addr
    msg['To'] = to_addr
    msg['Date'] = formatdate()

    # 作成したメールを送信
    smtpobj.send_message(msg)
    smtpobj.close()
    print('Done')


    # 在庫表の更新
    print('在庫表の更新')
    today = date.today()
    updated['日付'] = today
    updated['曜日'] = today.strftime('%a')

    pd.concat([df_stock, updated], sort=False).reset_index(drop=True).to_excel(filepath_stock, index=False)
    pd.read_excel(filepath_stock)
    print('ーーーーー終了ーーーーー')

if __name__ == '__main__':
    main()