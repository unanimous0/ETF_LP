import os
from typing import final
import pymysql
import xlwings as xw
import pandas as pd

# 코드/실행 파일이 속한 폴더의 경로
foler_dir = os.path.dirname(os.path.abspath(__file__))
# print(foler_dir)

# 디폴트 경로 변경
os.chdir(foler_dir)

# ETF LP 체결내역 데이터 경로
path_dir = os.getcwd() + "/KRX데이터_LP체결내역"
print(path_dir)

# 해당 경로에 있는 파일 리스트화
file_list = os.listdir(path_dir)

# 데이터베이스 정보
mysql_address = "localhost"
# mysql_address = "10.238.116.29"
mysql_port = 3306
mysql_user = 'root'
mysql_pw = 'admin'
mysql_db = 'etf_data'
mysql_charset = 'utf8'

# 데이터베이스 연결
conn = pymysql.connect(
                        host        = mysql_address, 
                        port        = mysql_port, 
                        user        = mysql_user, 
                        password    = mysql_pw, 
                        db          = mysql_db, 
                        charset     = mysql_charset
                        )

# ETF LP 체결내역 데이터 -> 데이터베이스화
try:

    with conn.cursor() as cursor:
        
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

            sql = "INSERT INTO etf_lp_trading_data(\
                    Date, Cat, Ticker, Name, LPComp, LP_Sell_vol, LP_Sell_Val, LP_Buy_Vol, LP_Buy_Val)\
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"

            cursor.executemany(sql, insert_data)
            conn.commit()
            
            print(f"DB Processing -> {file_name} -> Done")
            
        cursor.close()
        
except Exception as ex:
    print(ex)
    
finally:
    conn.close()
    print("DB Connection Closed.")
