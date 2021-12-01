import os
import pymysql
import xlwings as xw
import pandas as pd

path_dir = os.getcwd() + "/KRX데이터_LP체결내역"

# 경로에 있는 파일 리스트화
file_list = os.listdir(path_dir)

# mysql_address = "localhost"
mysql_address = "10.238.116.29"
mysql_port = 3306
mysql_user = 'admin'
mysql_pw = 'admin'
mysql_schemas = 'passive'
mysql_charset = 'utf8'

for file_name in file_list:
    book = xw.Book(path_dir + '/' + file_name)

    # 읽어온 데이터프레임의 index를 파일의 거래 날짜로 설정
    xlsfile = book.sheets(1).used_range.options(pd.DataFrame).value

    insert_data = []

    for i in range(0, len(xlsfile)):
        date = xlsfile.index[i]
        cat_product = xlsfile['증권그룹'].iloc[i]
        ticker = xlsfile['종목코드'].iloc[i]
        name = xlsfile['종목명'].iloc[i]
        corp = xlsfile['회원사명'].iloc[i]
        lp_sell_vol = xlsfile['LP매도거래량'].iloc[i]
        lp_sell_amt = xlsfile['LP매도거래대금'].iloc[i]
        lp_buy_vol = xlsfile['LP매수거래량'].iloc[i]
        lp_buy_amt = xlsfile['LP매수거래대금'].iloc[i]

        insert_data.append([int(date), cat_product, ticker, name, corp, 
            float(lp_sell_vol), float(lp_sell_amt), float(lp_buy_vol), float(lp_buy_amt)])

    sql = "INSERT INTO dat_etf_lphist(\
            date, cat_product, ticker, name, corp, lp_sell_vol, lp_sell_amt, lp_buy_vol, lp_buy_amt)\
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"

    conn = pymysql.connect(
                            host        = mysql_address, 
                            port        = mysql_port, 
                            user        = mysql_user, 
                            password    = mysql_pw, 
                            db          = mysql_schemas, 
                            charset     = mysql_charset
                            )

    cursor = conn.cursor()
    cursor.executemany(sql, insert_data)
    conn.commit()
    conn.close()
    
    print(file_name + " Done")
